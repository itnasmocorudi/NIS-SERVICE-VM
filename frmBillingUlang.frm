VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.DLL"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmBillingUlang 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Mencetak Ulang Billing"
   ClientHeight    =   10725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10725
   ScaleWidth      =   11280
   Begin VB.CommandButton cmdTutup7 
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
      Height          =   300
      Left            =   17760
      Style           =   1  'Graphical
      TabIndex        =   142
      Top             =   6240
      Visible         =   0   'False
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer7 
      Height          =   8055
      Left            =   17640
      TabIndex        =   143
      Top             =   5760
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
   Begin VB.CommandButton cmdTutup6 
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
      Height          =   300
      Left            =   13920
      Style           =   1  'Graphical
      TabIndex        =   144
      Top             =   5040
      Visible         =   0   'False
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer6 
      Height          =   8055
      Left            =   13800
      TabIndex        =   141
      Top             =   4560
      Visible         =   0   'False
      Width           =   3090
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
   Begin VB.CommandButton Command1 
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
      Height          =   300
      Left            =   13920
      Style           =   1  'Graphical
      TabIndex        =   139
      Top             =   4080
      Visible         =   0   'False
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer5 
      Height          =   8055
      Left            =   13800
      TabIndex        =   140
      Top             =   3600
      Visible         =   0   'False
      Width           =   3090
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
   Begin VB.CommandButton cmdTutup2 
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
      Height          =   300
      Left            =   13920
      Style           =   1  'Graphical
      TabIndex        =   112
      Top             =   3120
      Visible         =   0   'False
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer2 
      Height          =   8055
      Left            =   13800
      TabIndex        =   111
      Top             =   2640
      Visible         =   0   'False
      Width           =   3090
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
   Begin VB.CommandButton cmdSettingPrinter3 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Setting Printer"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   14880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   138
      Top             =   2160
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton cmdTutup3 
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
      Height          =   300
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   114
      Top             =   2160
      Visible         =   0   'False
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer3 
      Height          =   8055
      Left            =   13800
      TabIndex        =   115
      Top             =   1680
      Visible         =   0   'False
      Width           =   3180
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
   Begin VB.CommandButton cmdTutup4 
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
      Height          =   300
      Left            =   13920
      Style           =   1  'Graphical
      TabIndex        =   136
      Top             =   1320
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton cmdSettingPrinter4 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Setting Printer"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   15000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   137
      Top             =   1320
      Visible         =   0   'False
      Width           =   1395
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer4 
      Height          =   4455
      Left            =   13800
      TabIndex        =   135
      Top             =   840
      Visible         =   0   'False
      Width           =   3255
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
      Height          =   300
      Left            =   14040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton cmdSettingPrinter 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Setting Printer"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   15120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   125
      Top             =   480
      Visible         =   0   'False
      Width           =   1395
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   8055
      Left            =   13800
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   3090
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
   Begin VB.CommandButton cmdCall 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Pang&gil Ulang Customer"
      Height          =   375
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   126
      Top             =   0
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   810
      Left            =   600
      TabIndex        =   40
      Top             =   2400
      Width           =   12885
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   36
         Left            =   0
         TabIndex        =   47
         Top             =   480
         Width           =   3975
         _ExtentX        =   7011
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   35
         Left            =   0
         TabIndex        =   46
         Top             =   225
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
      Begin VB.OptionButton opWO 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Regular"
         Enabled         =   0   'False
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
         Index           =   0
         Left            =   720
         TabIndex        =   43
         Top             =   0
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton opWO 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Program"
         Enabled         =   0   'False
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
         Index           =   1
         Left            =   1680
         TabIndex        =   42
         Top             =   0
         Width           =   975
      End
      Begin VB.OptionButton opWO 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Warranty"
         Enabled         =   0   'False
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
         Index           =   2
         Left            =   2760
         TabIndex        =   41
         Top             =   0
         Width           =   1095
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   20
         Left            =   8280
         TabIndex        =   44
         Top             =   360
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Mask            =   "CCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   19
         Left            =   6675
         TabIndex        =   45
         Top             =   45
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   37
         Left            =   4800
         TabIndex        =   48
         Top             =   45
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   38
         Left            =   6675
         TabIndex        =   49
         Top             =   360
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   31
         Left            =   8280
         TabIndex        =   50
         Top             =   60
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Mask            =   "CCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   43
         Left            =   4800
         TabIndex        =   51
         Top             =   360
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Mask            =   "99-99-9999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   47
         Left            =   10410
         TabIndex        =   123
         Top             =   60
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Mask            =   "CCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   48
         Left            =   10410
         TabIndex        =   127
         Top             =   360
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Mask            =   ">99-99-9999"
         PromptChar      =   ">"
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Retur"
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
         Left            =   9885
         TabIndex        =   128
         Top             =   405
         Width           =   405
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Faktur"
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
         Left            =   9825
         TabIndex        =   124
         Top             =   105
         Width           =   465
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. WO"
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
         Left            =   7590
         TabIndex        =   60
         Top             =   405
         Width           =   570
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T SA."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6180
         TabIndex        =   59
         Top             =   90
         Width           =   390
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipe WO"
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
         TabIndex        =   58
         Top             =   30
         Width           =   615
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "KM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4440
         TabIndex        =   57
         Top             =   90
         Width           =   210
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mech"
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
         Left            =   6195
         TabIndex        =   56
         Top             =   405
         Width           =   375
      End
      Begin VB.Label lblPDS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PDS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   12285
         TabIndex        =   55
         Top             =   60
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Book"
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
         Left            =   7815
         TabIndex        =   54
         Top             =   105
         Width           =   345
      End
      Begin VB.Label Label19 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4080
         TabIndex        =   53
         Top             =   405
         Width           =   570
      End
      Begin VB.Label LblReturn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RTJ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   12285
         TabIndex        =   52
         Top             =   420
         Visible         =   0   'False
         Width           =   465
      End
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   270
      Left            =   2950
      TabIndex        =   107
      Top             =   105
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox text 
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   105
      Width           =   1335
      _ExtentX        =   2355
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
      Mask            =   ">9999999999"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox text 
      Height          =   285
      Index           =   1
      Left            =   5880
      TabIndex        =   108
      Top             =   90
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   12640511
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
      Mask            =   ">99-99-9999"
      PromptChar      =   ">"
   End
   Begin VB.CheckBox ChkCtkRate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rate tercetak"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   12120
      TabIndex        =   106
      Top             =   120
      Width           =   1380
   End
   Begin VB.CheckBox chCtkDisc 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Discount Tidak tercetak"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   9960
      TabIndex        =   105
      Top             =   120
      Width           =   2100
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   6015
      LargeChange     =   800
      Left            =   13080
      Max             =   2600
      SmallChange     =   600
      TabIndex        =   1
      Top             =   3240
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Height          =   6015
      Left            =   480
      ScaleHeight     =   5955
      ScaleWidth      =   12915
      TabIndex        =   61
      Top             =   3240
      Width           =   12975
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8655
         Left            =   0
         TabIndex        =   62
         Top             =   0
         Width           =   12615
         Begin VB.TextBox txtCatatan 
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1485
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   134
            Top             =   6675
            Width           =   7320
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   375
            Left            =   10560
            TabIndex        =   132
            Top             =   5580
            Width           =   2055
         End
         Begin VB.CheckBox ChCopy 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Copy"
            Height          =   375
            Left            =   9720
            TabIndex        =   131
            Top             =   8160
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CommandButton TblCetak 
            BackColor       =   &H00DCFCFB&
            Caption         =   "&Cetak"
            Height          =   375
            Left            =   10680
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   8160
            Width           =   1815
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFFF&
            Height          =   540
            Left            =   120
            TabIndex        =   89
            Top             =   7440
            Width           =   12405
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   32
               Left            =   1920
               TabIndex        =   90
               Top             =   165
               Width           =   3540
               _ExtentX        =   6244
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
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   33
               Left            =   11355
               TabIndex        =   91
               Top             =   165
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   12640511
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
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Suplay Slip/Keterangan"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   93
               Top             =   180
               Width           =   1755
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "PRC"
               Height          =   195
               Left            =   10920
               TabIndex        =   92
               Top             =   210
               Width           =   330
            End
         End
         Begin VB.Frame Frame9 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1170
            Left            =   10560
            TabIndex        =   82
            Top             =   4800
            Width           =   2055
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   26
               Left            =   915
               TabIndex        =   83
               Top             =   120
               Visible         =   0   'False
               Width           =   570
               _ExtentX        =   1005
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   16777215
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
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
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   27
               Left            =   720
               TabIndex        =   84
               Top             =   480
               Width           =   1260
               _ExtentX        =   2223
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.00##"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   28
               Left            =   720
               TabIndex        =   85
               Top             =   795
               Width           =   1260
               _ExtentX        =   2223
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PromptChar      =   ">"
            End
            Begin VB.Label Label45 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Material"
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
               TabIndex        =   88
               Top             =   840
               Width           =   570
            End
            Begin VB.Label Label44 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sublet"
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
               TabIndex        =   87
               Top             =   525
               Width           =   450
            End
            Begin VB.Label Label34 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Discount                   %"
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
               TabIndex        =   86
               Top             =   165
               Visible         =   0   'False
               Width           =   1635
            End
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1290
            Left            =   10560
            TabIndex        =   75
            Top             =   2640
            Width           =   2055
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   23
               Left            =   915
               TabIndex        =   76
               Top             =   120
               Visible         =   0   'False
               Width           =   570
               _ExtentX        =   1005
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   16777215
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
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
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   24
               Left            =   480
               TabIndex        =   77
               Top             =   480
               Width           =   1500
               _ExtentX        =   2646
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.00##"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   25
               Left            =   480
               TabIndex        =   78
               Top             =   840
               Width           =   1500
               _ExtentX        =   2646
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.00##"
               PromptChar      =   ">"
            End
            Begin VB.Label Label41 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Oli"
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
               TabIndex        =   81
               Top             =   840
               Width           =   180
            End
            Begin VB.Label Label38 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Part"
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
               Left            =   120
               TabIndex        =   80
               Top             =   525
               Width           =   300
            End
            Begin VB.Label Label30 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Discount                   %"
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
               Left            =   120
               TabIndex        =   79
               Top             =   165
               Visible         =   0   'False
               Width           =   1635
            End
         End
         Begin VB.CommandButton cmdCetak 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Cetak"
            Height          =   390
            Left            =   11040
            Style           =   1  'Graphical
            TabIndex        =   74
            Top             =   8640
            Width           =   1890
         End
         Begin VB.TextBox txtTotalRp 
            BackColor       =   &H00C0E0FF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   1485
            TabIndex        =   73
            Top             =   7080
            Width           =   11055
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1290
            Left            =   10560
            TabIndex        =   65
            Top             =   240
            Width           =   2055
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   21
               Left            =   915
               TabIndex        =   66
               Top             =   120
               Visible         =   0   'False
               Width           =   570
               _ExtentX        =   1005
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   16777215
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
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
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   22
               Left            =   720
               TabIndex        =   67
               Top             =   480
               Width           =   1260
               _ExtentX        =   2223
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
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
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   39
               Left            =   720
               TabIndex        =   68
               Top             =   840
               Width           =   1260
               _ExtentX        =   2223
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
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
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Body R"
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
               Left            =   105
               TabIndex        =   71
               Top             =   885
               Width           =   510
            End
            Begin VB.Label Label37 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Labour"
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
               Left            =   120
               TabIndex        =   70
               Top             =   525
               Width           =   495
            End
            Begin VB.Label Label29 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Discount                   %"
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
               Left            =   120
               TabIndex        =   69
               Top             =   165
               Visible         =   0   'False
               Width           =   1635
            End
         End
         Begin VB.CommandButton cetakSublet 
            BackColor       =   &H0080FF80&
            Caption         =   "Cetak Invoice Sublet"
            Height          =   300
            Left            =   8760
            MaskColor       =   &H008080FF&
            Style           =   1  'Graphical
            TabIndex        =   64
            Top             =   4800
            Width           =   1815
         End
         Begin VB.CommandButton cetakAdd 
            BackColor       =   &H0080FF80&
            Caption         =   "Cetak Invoice Add Job"
            Height          =   300
            Left            =   8760
            MaskColor       =   &H008080FF&
            Style           =   1  'Graphical
            TabIndex        =   63
            Top             =   240
            Width           =   1815
         End
         Begin TrueOleDBGrid70.TDBGrid JobGrid 
            Height          =   2295
            Left            =   0
            TabIndex        =   72
            Top             =   240
            Width           =   10575
            _ExtentX        =   18653
            _ExtentY        =   4048
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
            Columns(4).NumberFormat=   "Standard"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).DataField=   ""
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).DataField=   ""
            Columns(6).NumberFormat=   "Standard"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).DataField=   ""
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).DataField=   ""
            Columns(8).NumberFormat=   "Standard"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   4
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
            Columns.Count   =   15
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   14215660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=15"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8196"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
            Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=8196"
            Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(16)=   "Column(3).Width=2725"
            Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2646"
            Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=8196"
            Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(21)=   "Column(4).Width=2725"
            Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=8196"
            Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(26)=   "Column(5).Width=2725"
            Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2646"
            Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=8196"
            Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(31)=   "Column(6).Width=2725"
            Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2646"
            Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=8196"
            Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(36)=   "Column(7).Width=2725"
            Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=2646"
            Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=8196"
            Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(41)=   "Column(8).Width=2725"
            Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=2646"
            Splits(0)._ColumnProps(44)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(45)=   "Column(9).Width=2725"
            Splits(0)._ColumnProps(46)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(47)=   "Column(9)._WidthInPix=2646"
            Splits(0)._ColumnProps(48)=   "Column(9)._ColStyle=8196"
            Splits(0)._ColumnProps(49)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(50)=   "Column(10).Width=2725"
            Splits(0)._ColumnProps(51)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(52)=   "Column(10)._WidthInPix=2646"
            Splits(0)._ColumnProps(53)=   "Column(10).Order=11"
            Splits(0)._ColumnProps(54)=   "Column(11).Width=2725"
            Splits(0)._ColumnProps(55)=   "Column(11).DividerColor=0"
            Splits(0)._ColumnProps(56)=   "Column(11)._WidthInPix=2646"
            Splits(0)._ColumnProps(57)=   "Column(11).Order=12"
            Splits(0)._ColumnProps(58)=   "Column(12).Width=2725"
            Splits(0)._ColumnProps(59)=   "Column(12).DividerColor=0"
            Splits(0)._ColumnProps(60)=   "Column(12)._WidthInPix=2646"
            Splits(0)._ColumnProps(61)=   "Column(12).Order=13"
            Splits(0)._ColumnProps(62)=   "Column(13).Width=2725"
            Splits(0)._ColumnProps(63)=   "Column(13).DividerColor=0"
            Splits(0)._ColumnProps(64)=   "Column(13)._WidthInPix=2646"
            Splits(0)._ColumnProps(65)=   "Column(13).Order=14"
            Splits(0)._ColumnProps(66)=   "Column(14).Width=2725"
            Splits(0)._ColumnProps(67)=   "Column(14).DividerColor=0"
            Splits(0)._ColumnProps(68)=   "Column(14)._WidthInPix=2646"
            Splits(0)._ColumnProps(69)=   "Column(14).Order=15"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            Caption         =   "JOB ORDER"
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   12577002
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=Tahoma"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
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
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=70,.parent=13,.locked=-1"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=67,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=68,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=69,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=66,.parent=13,.locked=-1"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=63,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=64,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=65,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=62,.parent=13,.locked=-1"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=58,.parent=13,.locked=-1"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.locked=-1"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13,.locked=-1"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=46,.parent=13,.locked=-1"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=43,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=44,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=45,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=28,.parent=13,.locked=-1"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=25,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=26,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=27,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=74,.parent=13"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=71,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=72,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=73,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=32,.parent=13,.locked=-1"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=29,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=30,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=31,.parent=17"
            _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=78,.parent=13"
            _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=14"
            _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=15"
            _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=17"
            _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=82,.parent=13"
            _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=79,.parent=14"
            _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=80,.parent=15"
            _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=81,.parent=17"
            _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=86,.parent=13"
            _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=83,.parent=14"
            _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=84,.parent=15"
            _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=85,.parent=17"
            _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=90,.parent=13"
            _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=87,.parent=14"
            _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=88,.parent=15"
            _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=89,.parent=17"
            _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=94,.parent=13"
            _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=91,.parent=14"
            _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=92,.parent=15"
            _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=93,.parent=17"
            _StyleDefs(96)  =   "Named:id=33:Normal"
            _StyleDefs(97)  =   ":id=33,.parent=0"
            _StyleDefs(98)  =   "Named:id=34:Heading"
            _StyleDefs(99)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(100) =   ":id=34,.wraptext=-1"
            _StyleDefs(101) =   "Named:id=35:Footing"
            _StyleDefs(102) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(103) =   "Named:id=36:Selected"
            _StyleDefs(104) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(105) =   "Named:id=37:Caption"
            _StyleDefs(106) =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(107) =   "Named:id=38:HighlightRow"
            _StyleDefs(108) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(109) =   "Named:id=39:EvenRow"
            _StyleDefs(110) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(111) =   "Named:id=40:OddRow"
            _StyleDefs(112) =   ":id=40,.parent=33"
            _StyleDefs(113) =   "Named:id=41:RecordSelector"
            _StyleDefs(114) =   ":id=41,.parent=34"
            _StyleDefs(115) =   "Named:id=42:FilterBar"
            _StyleDefs(116) =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid70.TDBGrid partGrid 
            Height          =   2055
            Left            =   0
            TabIndex        =   94
            Top             =   2640
            Width           =   10575
            _ExtentX        =   18653
            _ExtentY        =   3625
            _LayoutType     =   0
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
            Columns.Count   =   5
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   14215660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=5"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(9)=   "Column(2).Width=2725"
            Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2646"
            Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(13)=   "Column(3).Width=2725"
            Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2646"
            Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(17)=   "Column(4).Width=2725"
            Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowDelete     =   -1  'True
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            Caption         =   "PART ORDER"
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   12577002
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=208,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bold=0,.fontsize=900,.italic=0"
            _StyleDefs(10)  =   ":id=4,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(11)  =   ":id=4,.fontname=Tahoma"
            _StyleDefs(12)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(13)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(14)  =   ":id=2,.fontname=Tahoma"
            _StyleDefs(15)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(16)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(17)  =   ":id=3,.fontname=Tahoma"
            _StyleDefs(18)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(19)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(20)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(21)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(22)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(23)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(24)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(25)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(26)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
            _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
            _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
            _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
            _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=32,.parent=13"
            _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14"
            _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
            _StyleDefs(58)  =   "Named:id=33:Normal"
            _StyleDefs(59)  =   ":id=33,.parent=0"
            _StyleDefs(60)  =   "Named:id=34:Heading"
            _StyleDefs(61)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(62)  =   ":id=34,.wraptext=-1"
            _StyleDefs(63)  =   "Named:id=35:Footing"
            _StyleDefs(64)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(65)  =   "Named:id=36:Selected"
            _StyleDefs(66)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(67)  =   "Named:id=37:Caption"
            _StyleDefs(68)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(69)  =   "Named:id=38:HighlightRow"
            _StyleDefs(70)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(71)  =   "Named:id=39:EvenRow"
            _StyleDefs(72)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(73)  =   "Named:id=40:OddRow"
            _StyleDefs(74)  =   ":id=40,.parent=33"
            _StyleDefs(75)  =   "Named:id=41:RecordSelector"
            _StyleDefs(76)  =   ":id=41,.parent=34"
            _StyleDefs(77)  =   "Named:id=42:FilterBar"
            _StyleDefs(78)  =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid70.TDBGrid subletGrid 
            Height          =   1815
            Left            =   0
            TabIndex        =   95
            Top             =   4800
            Width           =   10575
            _ExtentX        =   18653
            _ExtentY        =   3201
            _LayoutType     =   0
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
            Columns(8)._VlistStyle=   4
            Columns(8)._MaxComboItems=   5
            Columns(8).DataField=   ""
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   9
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   14215660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=9"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(9)=   "Column(2).Width=2725"
            Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2646"
            Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(13)=   "Column(3).Width=2725"
            Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2646"
            Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(17)=   "Column(4).Width=2725"
            Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(21)=   "Column(5).Width=2725"
            Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=2646"
            Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(25)=   "Column(6).Width=2725"
            Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=2646"
            Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(29)=   "Column(7).Width=2725"
            Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=2646"
            Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(33)=   "Column(8).Width=2725"
            Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=2646"
            Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowDelete     =   -1  'True
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            Caption         =   "SUBLET ORDER"
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   12577002
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=208,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bold=0,.fontsize=975,.italic=0"
            _StyleDefs(10)  =   ":id=4,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(11)  =   ":id=4,.fontname=Tahoma"
            _StyleDefs(12)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(13)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(14)  =   ":id=2,.fontname=Tahoma"
            _StyleDefs(15)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(16)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(17)  =   ":id=3,.fontname=Tahoma"
            _StyleDefs(18)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(19)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(20)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(21)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(22)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(23)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(24)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(25)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(26)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=62,.parent=13"
            _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=58,.parent=13"
            _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=54,.parent=13"
            _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
            _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=46,.parent=13"
            _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
            _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=28,.parent=13"
            _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=14"
            _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=15"
            _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=17"
            _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=32,.parent=13"
            _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=29,.parent=14"
            _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=30,.parent=15"
            _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=31,.parent=17"
            _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
            _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
            _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
            _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
            _StyleDefs(70)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
            _StyleDefs(71)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
            _StyleDefs(72)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
            _StyleDefs(73)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
            _StyleDefs(74)  =   "Named:id=33:Normal"
            _StyleDefs(75)  =   ":id=33,.parent=0"
            _StyleDefs(76)  =   "Named:id=34:Heading"
            _StyleDefs(77)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(78)  =   ":id=34,.wraptext=-1"
            _StyleDefs(79)  =   "Named:id=35:Footing"
            _StyleDefs(80)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(81)  =   "Named:id=36:Selected"
            _StyleDefs(82)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(83)  =   "Named:id=37:Caption"
            _StyleDefs(84)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(85)  =   "Named:id=38:HighlightRow"
            _StyleDefs(86)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(87)  =   "Named:id=39:EvenRow"
            _StyleDefs(88)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(89)  =   "Named:id=40:OddRow"
            _StyleDefs(90)  =   ":id=40,.parent=33"
            _StyleDefs(91)  =   "Named:id=41:RecordSelector"
            _StyleDefs(92)  =   ":id=41,.parent=34"
            _StyleDefs(93)  =   "Named:id=42:FilterBar"
            _StyleDefs(94)  =   ":id=42,.parent=33"
         End
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   30
            Left            =   1485
            TabIndex        =   96
            Top             =   6735
            Width           =   5655
            _ExtentX        =   9975
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
         Begin MSMask.MaskEdBox text 
            Height          =   405
            Index           =   29
            Left            =   10725
            TabIndex        =   97
            Top             =   6600
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   714
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   42
            Left            =   11355
            TabIndex        =   98
            Top             =   6120
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox MaskEdBox2 
            Height          =   270
            Left            =   12405
            TabIndex        =   99
            Top             =   6120
            Width           =   75
            _ExtentX        =   132
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   41
            Left            =   2040
            TabIndex        =   145
            Top             =   8040
            Width           =   3540
            _ExtentX        =   6244
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
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Penerima Unit"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   146
            Top             =   8055
            Width           =   1755
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "__________________    +"
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
            Left            =   10680
            TabIndex        =   104
            Top             =   6360
            Width           =   1920
         End
         Begin VB.Label Label39 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Grand Total"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   8910
            TabIndex        =   103
            Top             =   6652
            Width           =   1530
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Terbilang (Rupiah)"
            Height          =   195
            Left            =   60
            TabIndex        =   102
            Top             =   7200
            Width           =   1305
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Catatan"
            Height          =   255
            Left            =   405
            TabIndex        =   101
            Top             =   6780
            Width           =   945
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Materai"
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
            Left            =   10710
            TabIndex        =   100
            Top             =   6165
            Width           =   540
         End
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   2970
      Left            =   480
      TabIndex        =   5
      Top             =   360
      Width           =   13125
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   2
         Left            =   930
         TabIndex        =   6
         Top             =   240
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   4
         Left            =   920
         TabIndex        =   7
         Top             =   530
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   55
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   5
         Left            =   930
         TabIndex        =   8
         Top             =   820
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   6
         Left            =   930
         TabIndex        =   9
         Top             =   1110
         Width           =   3975
         _ExtentX        =   7011
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   7
         Left            =   930
         TabIndex        =   10
         Top             =   1400
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Mask            =   ">9999-9999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   10
         Left            =   5670
         TabIndex        =   11
         Top             =   210
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Mask            =   "9999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   12
         Left            =   5670
         TabIndex        =   12
         Top             =   510
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   15
         Left            =   5670
         TabIndex        =   13
         Top             =   1110
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   14
         Left            =   5670
         TabIndex        =   14
         Top             =   810
         Width           =   3375
         _ExtentX        =   5953
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   13
         Left            =   8250
         TabIndex        =   15
         Top             =   510
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   3
         Left            =   960
         TabIndex        =   16
         Top             =   240
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
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
         Mask            =   ">CCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   16
         Left            =   9750
         TabIndex        =   17
         Top             =   210
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   17
         Left            =   9750
         TabIndex        =   18
         Top             =   495
         Width           =   3015
         _ExtentX        =   5318
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   11
         Left            =   7830
         TabIndex        =   19
         Top             =   195
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Mask            =   "99-99-9999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   8
         Left            =   3180
         TabIndex        =   20
         Top             =   1400
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Mask            =   ">9999-9999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   9
         Left            =   930
         TabIndex        =   21
         Top             =   1680
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Mask            =   ">9999-9999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   18
         Left            =   9750
         TabIndex        =   22
         Top             =   810
         Width           =   3015
         _ExtentX        =   5318
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   40
         Left            =   2520
         TabIndex        =   23
         Top             =   240
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   34
         Left            =   9750
         TabIndex        =   24
         Top             =   1110
         Width           =   3015
         _ExtentX        =   5318
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   44
         Left            =   5670
         TabIndex        =   116
         Top             =   1400
         Width           =   3375
         _ExtentX        =   5953
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   45
         Left            =   5670
         TabIndex        =   118
         Top             =   1680
         Width           =   3375
         _ExtentX        =   5953
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   46
         Left            =   9750
         TabIndex        =   121
         Top             =   1400
         Width           =   3015
         _ExtentX        =   5318
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
      Begin VB.Label lblPWC 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PWC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   285
         Left            =   12405
         TabIndex        =   133
         Top             =   1800
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dealer"
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
         Left            =   9255
         TabIndex        =   122
         Top             =   1445
         Width           =   465
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PKP"
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
         Left            =   5250
         TabIndex        =   119
         Top             =   1725
         Width           =   270
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NPWP"
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
         Left            =   5085
         TabIndex        =   117
         Top             =   1445
         Width           =   435
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Engine."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9180
         TabIndex        =   39
         Top             =   855
         Width           =   540
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CDMA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   345
         TabIndex        =   38
         Top             =   1725
         Width           =   435
      End
      Begin VB.Label Label55 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DO TAM"
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
         Left            =   7095
         TabIndex        =   37
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Frame."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9210
         TabIndex        =   36
         Top             =   540
         Width           =   510
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9075
         TabIndex        =   35
         Top             =   255
         Width           =   645
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   135
         TabIndex        =   34
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Warna"
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
         Left            =   5040
         TabIndex        =   33
         Top             =   1155
         Width           =   480
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
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
         Left            =   5100
         TabIndex        =   32
         Top             =   525
         Width           =   420
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Th. Mobil"
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
         Left            =   4875
         TabIndex        =   31
         Top             =   255
         Width           =   645
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PSTN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   405
         TabIndex        =   30
         Top             =   1445
         Width           =   375
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kota"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   450
         TabIndex        =   29
         Top             =   1155
         Width           =   330
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   285
         TabIndex        =   28
         Top             =   855
         Width           =   495
      End
      Begin VB.Label Label27 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   375
         TabIndex        =   27
         Top             =   555
         Width           =   405
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MOU"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9375
         TabIndex        =   26
         Top             =   1155
         Width           =   345
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GSM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2760
         TabIndex        =   25
         Top             =   1445
         Width           =   315
      End
   End
   Begin VB.CommandButton cmdFaktur 
      BackColor       =   &H00DCFCFB&
      Caption         =   "Cetak &Faktur Pajak"
      Height          =   390
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   113
      Top             =   9360
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9855
      Left            =   0
      Picture         =   "frmBillingUlang.frx":0000
      ScaleHeight     =   9855
      ScaleWidth      =   495
      TabIndex        =   129
      Top             =   0
      Width           =   495
   End
   Begin ucXPButton.XPButton tblKeluar 
      Height          =   495
      Left            =   12360
      TabIndex        =   130
      Top             =   9360
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
      MICON           =   "frmBillingUlang.frx":23C3
      PICN            =   "frmBillingUlang.frx":23DF
      PICH            =   "frmBillingUlang.frx":28B3
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   9495
      Left            =   0
      Picture         =   "frmBillingUlang.frx":2DFE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
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
      Left            =   5175
      TabIndex        =   109
      Top             =   135
      Width           =   570
   End
   Begin VB.Label lblRetur 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RETUR INVOICE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   3120
      TabIndex        =   120
      Top             =   90
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. Invoice"
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
      Left            =   585
      TabIndex        =   110
      Top             =   150
      Width           =   825
   End
End
Attribute VB_Name = "frmBillingUlang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Job ID          :
'Tgl. Mulai      : 04 November 2009
'Created Coding  : Thania
'Tgl. Modifikasi : Januari-Maret 2012 - Acep
'=======================================

Dim simpan As Boolean
Dim rsTemp As New ADODB.Recordset
Dim rsPart As New ADODB.Recordset
Dim rsJob As New ADODB.Recordset
Dim rsSublet As New ADODB.Recordset
Dim rsHlp As New ADODB.Recordset
Dim objBilling As New BClsBilling
Dim subJob As Double
Dim nDuoInv As Boolean
Dim vNamaTabel As String
Dim nRC As Boolean
Dim nHeavy As Boolean
Dim Report As Object
Dim rpt As Object

' UNTUK INVOICE KEDUA
Dim gKdPlgDealer2 As String
Dim gKodePlg2 As String
Dim gInvCust2 As String
Dim gInvAlmt2 As String
Dim gInvKota2 As String
Dim gInvNPWP2 As String
Dim gInvPKP2 As String
Dim vTipeCustomer As String
Dim vTglInv As Date
Dim support As New clsFormDealer

Private txt5 As String
Dim metB As Double

Private rsCetakNewBilling As New ADODB.Recordset

Dim nilaiPPnInv     As Double
Dim nilaiDppInv     As Double
Dim nilaiTotalInv   As Double

Private Sub doCreatePart()
1        On Error GoTo doCreatePart_Error
2         If rsPart.State Then rsPart.Close
3         Set rsPart = Nothing
4         With rsPart
              'If .State Then .Close
5             .Fields.Append "KODEITEM", adVarChar, 50
6             .Fields.Append "NAMAITEM", adVarChar, 200
7             .Fields.Append "QTY", adInteger
8             .Fields.Append "HARGA", adDouble
9             .Fields.Append "OLI", adBoolean
10            .Fields.Append "EXTDISC", adDouble
11            .Open
12        End With

13       'On Error GoTo 0
14       Exit Sub

doCreatePart_Error:

15        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure doCreatePart of Form frmBillingUlang")
End Sub

Private Sub doHeadPart()
1        On Error GoTo doHeadPart_Error

2         With partGrid
3             .Columns(0).Width = 1800
4             .Columns(0).Caption = "KODE PART"
5             .Columns(0).DataField = "KODEITEM"
6             .Splits(0).Columns(0).Locked = True
              
7             .Columns(1).Width = 2500
8             .Columns(1).Caption = "NAMA PART"
9             .Columns(1).DataField = "NAMAITEM"
10            .Splits(0).Columns(1).Locked = True
              
11            .Columns(2).Width = 800
12            .Columns(2).Caption = "QTY"
13            .Columns(2).DataField = "QTY"
14            .Splits(0).Columns(2).Locked = True
              
15            .Columns(3).Visible = False
16            .Columns(3).Caption = "HARGA"
17            .Columns(3).DataField = "HRG"
18            .Splits(0).Columns(3).Locked = True
              
19            .Columns(4).Visible = False
20            .Columns(4).Caption = "OLI"
21            .Columns(4).DataField = "OLI"
22            .Splits(0).Columns(4).Locked = True
23        End With

24       'On Error GoTo 0
25       Exit Sub

doHeadPart_Error:

26        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure doHeadPart of Form frmBillingUlang")
End Sub

Private Sub doCreateJob()
   On Error GoTo doCreateJob_Error
    If rsJob.State Then rsJob.Close
    'Set rsJob = Nothing
    With rsJob
        'If .State Then .Close
        .Fields.Append "JOB", adVarChar, 50
        .Fields.Append "RATE", adDouble
        .Fields.Append "ADDJOB", adBoolean
        .Fields.Append "HRG", adDouble
        .Fields.Append "DISC", adDouble
        .Fields.Append "STLDISC", adDouble
        .Fields.Append "FREEC", adBoolean
        If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(text(43), "DD/MM/YYYY")) >= 0 Then
            .Fields.Append "IDPROGRAM", adVarChar, 20
            .Fields.Append "PROGRAM", adVarChar, 50
            .Fields.Append "DISCMAX", adDouble
            .Fields.Append "ISFP", adBoolean
            .Fields.Append "PRICE", adDouble
            .Fields.Append "discrp", adDouble
        End If
        .Open
    End With

   'On Error GoTo 0
   Exit Sub

doCreateJob_Error:

    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure doCreateJob of Form frmBillingUlang")
End Sub

Private Sub doHeadJob()
   On Error GoTo doHeadJob_Error

    With jobGrid
        If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(text(43), "DD/MM/YYYY")) >= 0 Then
            .Columns(2).Width = 2000
            .Columns(2).Caption = "PROGRAM"
            .Columns(2).DataField = "PROGRAM"
            .Columns(2).Visible = True
            .Splits(0).Columns(2).Locked = True
            
            .Columns(10).Width = 0
            .Columns(10).Caption = "ACTRATE"
            .Columns(10).DataField = "ACTRT"
            .Columns(10).Visible = False
            .Splits(0).Columns(10).Locked = True
            
            
            .Columns(11).Width = 0
            .Columns(11).Caption = "IDPROGRRAM"
            .Columns(11).DataField = "IDPROGRRAM"
            .Columns(11).Visible = False
            .Splits(0).Columns(11).Locked = True
            
            .Columns(12).Width = 0
            .Columns(12).Caption = "DISCMAX"
            .Columns(12).DataField = "DISCMAX"
            .Columns(12).Visible = False
            .Splits(0).Columns(12).Locked = True
            
            .Columns(13).Width = 0
            .Columns(13).Caption = "ISFP"
            .Columns(13).DataField = "ISFP"
            .Columns(13).Visible = False
            .Splits(0).Columns(13).Locked = True
            
            .Columns(14).Width = 0
            .Columns(14).Caption = "PRICE"
            .Columns(14).DataField = "PRICE"
            .Columns(14).Visible = False
            .Splits(0).Columns(14).Locked = True
        Else
            
            .Columns(2).Width = 0
            .Columns(2).Caption = "ACTRATE"
            .Columns(2).DataField = "ACTRT"
            .Columns(2).Visible = False
            .Splits(0).Columns(2).Locked = True
                    
            .Columns(10).Visible = False
            .Columns(11).Visible = False
            .Columns(12).Visible = False
            .Columns(13).Visible = False
            .Columns(14).Visible = False
        End If
        .Columns(0).Width = 4300
        .Columns(0).Caption = "JOB ORDER"
        .Columns(0).DataField = "JOB"
        .Splits(0).Columns(0).Locked = True
        
        .Columns(1).Width = 600
        .Columns(1).Caption = "RATE"
        .Columns(1).DataField = "RATE"
        .Splits(0).Columns(1).Locked = True
        
'        .Columns(2).Width = 0
'        .Columns(2).Caption = "ACTRATE"
'        .Columns(2).DataField = "ACTRT"
'        .Columns(2).Visible = False
'        .Splits(0).Columns(2).Locked = True
        
        .Columns(3).Width = 0
        .Columns(3).Caption = "GROUP"
        .Columns(3).DataField = "GRP"
        .Columns(3).Visible = False
        .Splits(0).Columns(3).Locked = True
        
        .Columns(4).Width = 0
        .Columns(4).Caption = "MEKANIK"
        .Columns(4).DataField = "MECH"
        .Columns(4).Visible = False
        .Splits(0).Columns(4).Locked = True
        
        .Columns(5).Width = 0
        .Columns(5).Caption = "Additional Job"
        .Columns(5).DataField = "ADDJOB"
        .Columns(5).Visible = False
        .Splits(0).Columns(5).Locked = True
        
        .Columns(6).Width = 1200
        .Columns(6).Caption = "HARGA"
        .Columns(6).DataField = "HRG"
        .Splits(0).Columns(6).Locked = True
               
        .Columns(7).Width = 600
        .Columns(7).Caption = "DISC"
        .Columns(7).DataField = "DISC"
        
        .Columns(8).Width = 1200
        .Columns(8).Caption = "SUB TOTAL"
        .Columns(8).DataField = "STLDISC"
        .Splits(0).Columns(8).Locked = True
        
        .Columns(9).Width = 400
        .Columns(9).Caption = "C"
        .Columns(9).DataField = "FREEC"
        .Columns(9).Alignment = dbgCenter
        .Splits(0).Columns(9).Locked = True
        .Columns(9).Visible = False
        If nDuoInv Then
            .Columns(9).Visible = True
        End If
    End With

   'On Error GoTo 0
   Exit Sub

doHeadJob_Error:

    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure doHeadJob of Form frmBillingUlang")
End Sub

Private Sub doHeadSublet()
    
   On Error GoTo doHeadSublet_Error

    With subletGrid
        .Columns(0).Width = 1500
        .Columns(0).Caption = "NO NOTA"
        .Splits(0).Columns(0).Locked = True
        
        .Columns(1).Width = 1800
        .Columns(1).Caption = "SUPPLIER"
        .Columns(1).Visible = True
        .Splits(0).Columns(1).Locked = True
        
        .Columns(2).Width = 3500
        .Columns(2).Caption = "NAMA BARANG"
        .Splits(0).Columns(2).Locked = True
        
        .Columns(3).Width = 1200
        .Columns(3).Caption = "KODE SATUAN"
        .Splits(0).Columns(3).Locked = True
        
        .Columns(4).Width = 800
        .Columns(4).Caption = "JUMLAH"
        .Splits(0).Columns(4).Locked = True
        
        .Columns(5).Width = 800
        .Columns(5).Caption = "VALID"
        .Columns(5).Visible = False
        .Splits(0).Columns(5).Locked = True
        
        .Columns(6).Width = 1200
        .Columns(6).Caption = "HARGA"
        .Columns(6).NumberFormat = "#,#0"
        .Splits(0).Columns(6).Locked = True
        
'        .Columns(7).Width = 800
'        .Columns(7).Caption = "DISC"
'        .Columns(7).NumberFormat = "#,#0"
'        .Splits(0).Columns(7).Locked = True
        
        
        .Columns(8).Width = 800
        .Columns(8).Caption = "price"
        .Columns(8).NumberFormat = "#,#0"
        .Splits(0).Columns(8).Locked = True
        .Columns(8).Visible = False
        
        
        If Log.standarisasiSublet = 1 Then
            .Columns(7).Width = 800
            .Columns(7).Caption = "DISC"
            .Columns(7).DataField = "diskon"
            .Splits(0).Columns(7).Locked = True
            .Columns(7).Visible = True
        Else
            .Columns(7).Visible = False
        End If
        
        If Log.freepart = 1 Then
            .Columns(8).Width = 800
            .Columns(8).Caption = "Cust"
            .Columns(8).DataField = "freec"
            .Splits(0).Columns(8).Locked = True

            If isInvoiceC Then
                .Columns(8).Visible = True
            Else
                .Columns(8).Visible = False
            End If
        Else
            .Columns(8).Visible = False
        End If
    End With

   'On Error GoTo 0
   Exit Sub

doHeadSublet_Error:

    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure doHeadSublet of Form frmBillingUlang")
End Sub

Private Sub hapusForm()
1        On Error GoTo hapusForm_Error

2         Call objO.DeleteData("DELETE TEMPJOBSUBLET WHERE NO_WO = '" & text(0) & "'")
          
3         For i = 1 To 46
4             Select Case i
                  Case 1, 11, 43
5                     text(i).Mask = "99-99-9999"
6                 Case 10, 21, 22, 23, 24, 25, 26, 27, 28, 29, 39, 42
7                     text(i) = 0
8                 Case 41
9                 Case Else
10                    text(i) = ""
11            End Select
12        Next
          
13        VScroll1.value = 0
14        opWO(0).value = True

15        If rsJob.State Then rsJob.Close
16        Set rsJob = Nothing
17        Set jobGrid.DataSource = rsJob
          
18        If rsPart.State Then rsPart.Close
19        Set rsPart = Nothing
20        Set partGrid.DataSource = rsPart
          
21        Set subletGrid.DataSource = Nothing
22        adaAddJob = False
23        cetakSublet.Visible = False
24        cetakAdd.Visible = False
25        lblRetur.Visible = False
          
26        vTipeCustomer = ""
27        gKodePlg = ""
28        gInvNo = ""
29        gInvCust = ""
30        gInvAlmt = ""
31        gInvKota = ""
32        gInvNPWP = ""
33        nDuoInv = False

34        nilaiPPnInv = 10
35        nilaiDppInv = 100
36        nilaiTotalInv = nilaiDppInv + nilaiPPnInv
          
37        If Log.DIDX = 1 Then
38            Label35(1).Visible = True
39            text(41).Visible = True
40            text(41) = ""
41        Else
42            Label35(1).Visible = False
43            text(41).Visible = False
44        End If
          
         'On Error GoTo 0
45       Exit Sub

hapusForm_Error:

46        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure hapusForm of Form frmBillingUlang")
End Sub

Private Sub cetakAdd_Click()
Set Report = New rptAddInvoice
Dim rsRepo As ADODB.Recordset
    
    If rsRepo.State Then rsRepo.Close
    sSQL = "SELECT A.NO_WO,NOINV, A.NO_CHASSIS, NO_POLISI, FMDL, KLPKERJA, REP, RTE, MEKANIK, OFFN, A.DISC, CUSTOMER, CUST_ADRES, KOTA, NO_ENGINE " _
         & "FROM HIST_JOBADD A WITH(NOLOCK) INNER JOIN T_INVOICE B WITH(NOLOCK) ON A.NO_WO = B.NO_WO AND A.NO_CHASSIS = B.NO_CHASSIS " _
         & "WHERE NOINV = '" & text(0) & "' AND A.NO_WO = '" & text(20) & "' ORDER BY A.NO_WO, KLPKERJA, REP"
    Set rsRepo = objO.FetchData(sSQL)
    Report.database.SetDataSource rsRepo
    Report.txtPT.SetText Trim(gPerush)
    Report.TxtAlamat.SetText Trim(gAlamat)
    Report.TxtKota.SetText Trim(gKotaCbg)
    Report.txtTelp.SetText "Telp. " & Trim(gTelp)
    Report.TxtFax.SetText ""  '"Fax. "
    Report.txtUser.SetText Trim(gSA)
    Report.txtKot.SetText Trim(gKotaCbg)
    Report.txtKabeng.SetText Trim(gSecmanager)

    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
        
        .Top = 0
        .Left = asxPowerBanner1.Left + asxPowerBanner1.Width + 20
        .Width = Frame2.Width + 100
        .Height = Me.Height
        
        cmdTutup.Top = .Top + 430
        cmdTutup.Left = .Left
        cmdTutup.Visible = True
        .Visible = True
    End With
    Screen.MousePointer = vbDefault
    cmdTutup.SetFocus
End Sub

Private Sub cetakSublet_Click()
Set Report = New rptSublet
Dim rsRepo As ADODB.Recordset
    
    ''' ASLI NICH !!!
'    sSQL = "SELECT CUSTOMER, CUST_ADRES, KOTA, A.NO_WO, A.NOINV, B.NO_POLISI, A.NO_CHASSIS, NO_ENGINE, NAMABARANG, HARGA, QTY, KETERANGAN " _
'         & "FROM HIST_SUBLET A INNER JOIN T_INVOICE B ON A.NO_WO = B.NO_WO AND A.NOINV = B.NOINV " _
'         & "WHERE A.NOINV = '" & text(0) & "' AND A.NO_WO = '" & text(20) & "' ORDER BY A.NO_WO, NAMABARANG "
    '#########################################################'
    '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..'
    '#########################################################'
    sSQL = "SELECT CUSTOMER, CUST_ADRES, KOTA, A.NO_WO, A.NOINV, B.NO_POLISI, A.NO_CHASSIS, NO_ENGINE, NAMABARANG, HARGA, QTY, KETERANGAN " _
         & "FROM HIST_SUBLET A WITH(NOLOCK) INNER JOIN T_INVOICE B WITH(NOLOCK) ON A.NO_WO = B.NO_WO AND A.NOINV = B.NOINV  " _
         & "WHERE A.NOINV LIKE '" & text(0) & "%' AND B.NOINV NOT LIKE '%-B' AND A.NO_WO = '" & text(20) & "' ORDER BY A.NO_WO, NAMABARANG "
    
    Set rsRepo = objO.FetchData(sSQL)
    Report.database.SetDataSource rsRepo
    Report.txtPT.SetText Trim(gPerush)
    Report.TxtAlamat.SetText Trim(gAlamat)
    Report.TxtKota.SetText Trim(gKotaCbg)
    Report.txtTelp.SetText "Telp. " & Trim(gTelp)
    Report.TxtFax.SetText ""  '"Fax. "
    Report.txtKotaTgl.SetText Trim(gKotaCbg) & ", " & TglDDMMMMYYYY(Format(text(1), "DD/MM/YYYY"), 1)
    Report.txtKabeng.SetText Trim(gSecmanager)

    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
        
        .Top = 0
        .Left = asxPowerBanner1.Left + asxPowerBanner1.Width + 20
        .Width = Frame2.Width + 100
        .Height = TblKeluar.Top + TblKeluar.Height
        
        cmdTutup.Top = .Top + 430
        cmdTutup.Left = .Left
        cmdTutup.Visible = True
        
        cmdSettingPrinter.Top = cmdTutup.Top
        cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 50
        cmdSettingPrinter.Visible = True
        .Visible = True
    End With
    Screen.MousePointer = vbDefault
    cmdTutup.SetFocus
End Sub

Private Sub cmdCall_Click()
    If Trim(text(16)) <> "" Then
        ' PANGGIL CUSTOMER
        Call PanggilCust("AwalBilling.wav", Replace(text(16), "-", ""), "AkhirBilling.wav", False, "", "")
    Else
        MsgBox "Tampilkan Invoice Customer yang akan dipanggil "
        text(0).SelLength = Len(text(0))
        text(0).SetFocus
    End If
End Sub

Private Sub cmdSettingPrinter_Click()
    Report.PrinterSetup Me.hWnd
    CRViewer1.Refresh
End Sub

Private Sub cmdSettingPrinter3_Click()
    Report.PrinterSetup Me.hWnd
    CRViewer3.Refresh
End Sub

Private Sub cmdSettingPrinter4_Click()
    Report.PrinterSetup Me.hWnd
    CRViewer4.Refresh
    
  SaveSetting "SUPPLYSLIP-SIMPLIFY", "Startup", "PaperSize", Report.PaperSize
End Sub

Private Sub cmdTutup_Click()
Dim vNamaTabel As String

    vNamaTabel = "INVOICE" & Trim(text(0))
    sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ) " _
         & " drop table [tempdb]..[" & vNamaTabel & "]"
    Call objO.UpdateData(sSQL)
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
End Sub

Private Sub cmdTutup2_Click()
Dim vNamaTabel As String

    vNamaTabel = "INVOICE2-" & Trim(text(0))
    sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ) " _
         & " drop table [tempdb]..[" & vNamaTabel & "]"
    Call objO.UpdateData(sSQL)
    CRViewer2.Visible = False
    cmdTutup2.Visible = False
End Sub

Private Sub cmdTutup3_Click()
    CRViewer3.Visible = False
    cmdTutup3.Visible = False
    cmdSettingPrinter3.Visible = False
End Sub

Private Sub cmdTutup4_Click()
    CRViewer4.Visible = False
    cmdTutup4.Visible = False
    cmdSettingPrinter4.Visible = False
End Sub

Private Sub cmdTutup6_Click()
    CRViewer6.Visible = False
    cmdTutup6.Visible = False
End Sub

Private Sub cmdTutup7_Click()
    CRViewer7.Visible = False
    cmdTutup7.Visible = False
End Sub

Private Sub Command1_Click()
    CRViewer5.Visible = False
    Command1.Visible = False
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Resize()
    asxPowerBanner1.Height = Me.Height - 500
End Sub


Private Sub TblCetak_Click()
    'Call hapusForm
   On Error GoTo TblCetak_Click_Error

    If rsTemp.State Then rsTemp.Close
    ''' ASLI NICH !!!
    'pSQL = "SELECT * FROM T_INVOICE WHERE NOINV = '" & Trim(text(Index)) & "' "
    '#########################################################'
    '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..'
    '#########################################################'
    Dim rsCetakNew As ADODB.Recordset
    
    Set rsCetakNew = objO.FetchData("SELECT * FROM T_INVOICE WITH(NOLOCK) WHERE NOINV LIKE '" & Left(Trim(text(Index)), 8) & "%' ")
    If rsCetakNew.EOF Then
        gMainKeyCari = text(Index)
        gSearchOption = "INVULANG"
        frmCari.show (vbModal)
    Else
        If Log.newCetak = 1 Then
            rsCetakNew.MoveFirst
            Dim No As Integer
            No = 1
            Call cetakInv1New(NulSID(rsCetakNew!noInv), No)
            Call cetakSupplySlipSimplify
            Call cetakSubletSimplify
            rsCetakNew.MoveNext
            No = No + 1
            If Not rsCetakNew.EOF Then
                Call cetakInv2New(NulSID(rsCetakNew!noInv), No)
            End If
            If Not rsCetakNew.EOF Then
             rsCetakNew.MoveNext
            No = No + 1
            If Not rsCetakNew.EOF Then
                Call cetakInv3New(NulSID(rsCetakNew!noInv), No)
                Call cetakSupplySlipSimplifyC(NulSID(rsCetakNew!noInv))
                Call cetakSubletSimplifyC
            End If
            End If
            'Call cetakInv3New(NulSID(Left(Trim(Text(Index)), 8)) & "-C", No + 1)
        End If
    End If
    If rsTemp.State <> adStateClosed Then rsTemp.Close
    If Log.newCetak = 0 Then
        Call cetakInv1
        If nDuoInv Then
            Call cetakInv2
        End If
    End If

   'On Error GoTo 0
   Exit Sub

TblCetak_Click_Error:

    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure TblCetak_Click of Form frmBillingUlang")
End Sub

Private Sub TblCetak_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 27 Then
        text(0).SetFocus
    End If
End Sub

Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub Text_GotFocus(Index As Integer)
Dim DiTab As Boolean

    DiTab = False
    If text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            text(Index).Mask = ""
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
    Select Case KeyCode
        Case vbKeyF12
            Select Case Index
                Case 0
                    gMainKeyCari = text(Index)
                    gSearchOption = "INVULANG"
                    frmCari.show (vbModal)
            End Select
        Case vbKeyEscape
            If simpan = False Then
                If MsgBox("Data Belum disimpan, menampilkan data baru ?", vbYesNo, "Konfirmasi") = vbYes Then
                    Call hapusForm
                    text(0).SelLength = Len(text(0))
                    text(0).SetFocus
                End If
            End If
    End Select
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpCase(KeyAscii)
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
    
    Select Case Index
        Case 0
            text(Index).text = Replace(text(Index).text, "-", "")
            If text(Index) <> "" Then
                boleh = True
                If simpan = False Then
                    If MsgBox("Data Belum disimpan, menampilkan data baru ?", vbYesNo, "Konfirmasi") = vbNo Then
                        boleh = False
                    End If
                End If
                If boleh Then
                    Call hapusForm
                    If rsTemp.State Then rsTemp.Close
                    ''' ASLI NICH !!!
                    'pSQL = "SELECT * FROM T_INVOICE WHERE NOINV = '" & Trim(text(Index)) & "' "
                    '#########################################################'
                    '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..'
                    '#########################################################'
                    Set rsTemp = objO.FetchData("SELECT * FROM T_INVOICE WITH(NOLOCK) WHERE NOINV LIKE '" & Left(Trim(text(Index)), 8) & "%' ")
                    If rsTemp.EOF Then
                        gMainKeyCari = text(Index)
                        gSearchOption = "INVULANG"
                        frmCari.show (vbModal)
                    Else
                        Call panggilINV(Trim(text(Index)))
                    End If
                    If rsTemp.State <> adStateClosed Then rsTemp.Close
                    simpan = False
                End If
            End If
    End Select
End Sub

Private Function returnJob(labelReturn As Object) As Boolean
          
1        On Error GoTo returnJob_Error

2         returnJob = False
3         If labelReturn.Visible Then
4             returnJob = True
5         End If

6        'On Error GoTo 0
7        Exit Function

returnJob_Error:

8         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure returnJob of Form frmBillingUlang")
End Function

Private Function giveDiscount(discountValue As String, discountText As Object, statusText As Boolean)
          
1        On Error GoTo giveDiscount_Error

2         discountText = discountValue
3         discountText.Enabled = statusText
4         jobGrid.Splits(0).Columns(6).Locked = Not statusText

5        'On Error GoTo 0
6        Exit Function

giveDiscount_Error:

7         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure giveDiscount of Form frmBillingUlang")
End Function

Private Sub panggilINV(noInv As String)
      Dim rsInvoice As New ADODB.Recordset
      Dim rsBantu As New ADODB.Recordset
      Dim isTerlibatFreeSvcHO     As Boolean

1        On Error GoTo panggilINV_Error

2         If rsInvoice.State Then rsInvoice.Close
          ''' ASLI NICH !!!
      '    Set rsInvoice = oConn.Execute(" SELECT *, (SELECT GSM FROM CUSTDAT WHERE NO_CHASSIS = T_INVOICE.NO_CHASSIS) AS GSM, " _
      '                                & "(SELECT CDMA FROM CUSTDAT WHERE NO_CHASSIS = T_INVOICE.NO_CHASSIS) AS CDMA, " _
      '                                & "(SELECT DOTAM FROM CUSTDAT WHERE NO_CHASSIS = T_INVOICE.NO_CHASSIS) AS DOTAM, " _
      '                                & "(SELECT NPWP FROM CUSTDAT WHERE NO_CHASSIS = T_INVOICE.NO_CHASSIS) AS NPWP, " _
      '                                & "(SELECT PKP FROM CUSTDAT WHERE NO_CHASSIS = T_INVOICE.NO_CHASSIS) AS PKP, " _
      '                                & "(SELECT FS_NM_DEALER FROM CUSTDAT WHERE NO_CHASSIS = T_INVOICE.NO_CHASSIS) AS NAMADEALER " _
      '                                & " FROM T_INVOICE WHERE NOINV = '" & Trim(noinv) & "'")
          
          '#########################################################'
          '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..'
          '#########################################################'
      '    Set rsInvoice = objO.FetchData(" SELECT *, (SELECT GSM FROM CUSTDAT WHERE NO_CHASSIS = T_INVOICE.NO_CHASSIS) AS GSM, " _
      '                                & "(SELECT CDMA FROM CUSTDAT WHERE NO_CHASSIS = T_INVOICE.NO_CHASSIS) AS CDMA, " _
      '                                & "(SELECT DOTAM FROM CUSTDAT WHERE NO_CHASSIS = T_INVOICE.NO_CHASSIS) AS DOTAM, " _
      '                                & "(SELECT NPWP FROM CUSTDAT WHERE NO_CHASSIS = T_INVOICE.NO_CHASSIS) AS NPWP, " _
      '                                & "(SELECT PKP FROM CUSTDAT WHERE NO_CHASSIS = T_INVOICE.NO_CHASSIS) AS PKP, " _
      '                                & "(SELECT FS_NM_DEALER FROM CUSTDAT WHERE NO_CHASSIS = T_INVOICE.NO_CHASSIS) AS NAMADEALER " _
      '                                & " FROM T_INVOICE WHERE NOINV LIKE '" & Trim(noinv) & "%' ORDER BY NOINV")
3         Set rsInvoice = objO.FetchData("SELECT A.*, B.GSM AS GSM, B.CDMA AS CDMA, " & _
                                      "B.DOTAM AS DOTAM, B.NPWP AS NPWP, B.PKP AS PKP, B.FS_NM_DEALER AS NAMADEALER " & _
                                      "FROM T_INVOICE A WITH(NOLOCK) INNER Join CUSTDAT B WITH(NOLOCK) on A.NO_CHASSIS = B.NO_CHASSIS WHERE A.NOINV LIKE '" & Trim(noInv) & "%' ORDER BY A.NOINV")
4        metB = 0
          
5         isTerlibatFreeSvcHO = False
          
6         If Not rsInvoice.EOF Then
          
7             cmdCall.Visible = True
8             If Trim(rsInvoice("FS_LOG_DEL")) <> "" Then
9                 MsgBox "Invoice ini sudah diretur !!!"
10                lblRetur.Visible = True
11                cmdCall.Visible = False
12            End If
              
              
13                nilaiPPnInv = NulSID(rsInvoice("PPn"))
14                nilaiDppInv = 100
15                nilaiTotalInv = nilaiDppInv + nilaiPPnInv
                  
              
16            text(1) = TglDDMMYYYY(rsInvoice("INVDATE"))
17            text(2) = Trim(rsInvoice("FS_KD_PLG"))
18            text(3) = Trim(rsInvoice("FS_KD_WIL"))
19            text(4) = Trim(rsInvoice("CUSTOMER"))
20            txt5 = Trim(rsInvoice("CUST_ADRES"))
21            text(5) = Left(Trim(rsInvoice("CUST_ADRES")), 60)
22            text(6) = Trim(rsInvoice("KOTA"))
23            text(7) = Replace(Trim(rsInvoice("TELPON")), "-", "")
24            text(8) = Replace(Trim(rsInvoice("GSM")), "-", "-")
25            text(9) = Replace(Trim(rsInvoice("CDMA")), "-", "-")
26            text(10) = Trim(rsInvoice("FS_TH_MOBIL"))
27            text(11) = TglDDMMYYYY(rsInvoice("DOTAM"))
28            vTglInv = TglDDMMYYYY(rsInvoice("INVDATE"))
29            text(12) = Trim(rsInvoice("MDL"))
30            text(13) = Trim(rsInvoice("MODEL"))
31            text(14) = Trim(rsInvoice("FMDL"))
32            text(15) = Trim(rsInvoice("COLOR"))
33            text(16) = Trim(rsInvoice("NO_POLISI"))
34            text(17) = Trim(rsInvoice("NO_CHASSIS"))
35            text(18) = Trim(rsInvoice("NO_ENGINE"))
36            text(19) = Trim(rsInvoice("FRONT"))
37            text(20) = Trim(rsInvoice("NO_WO"))
38            txtCatatan.text = Trim(rsInvoice("FS_CATATAN"))
39            text(31) = Trim(rsInvoice("NOFAK"))
40            text(32) = Trim(rsInvoice("SLIP"))
41            text(33) = Trim(rsInvoice("FS_RC"))
42            nRC = IIf(Trim(text(33)) <> "", True, False)
43            text(44) = Trim(rsInvoice("NPWP"))
44            text(45) = Trim(rsInvoice("PKP"))
45            text(46) = Trim(rsInvoice("NAMADEALER"))
46            text(47) = Trim(rsInvoice("NOFAKTUR"))
47            n = text(48).Mask
48            text(48).Mask = ""
49            Select Case TglDDMMYYYY(rsInvoice("FD_RETUR"))
                  Case "01-01-3000"
50                    text(48) = ""
51                Case "01-01-1900"
52                    text(48) = ""
53                Case Else
54                    text(48) = TglDDMMYYYY(rsInvoice("FD_RETUR"))
55            End Select
56            text(48).Mask = n
              
              '' 20190625DEV1
57            If Log.lockFreeSVCHO = 1 And DateDiff("D", Format(Log.tglFreeSVCHO, "DD/MM/YYYY"), Format(isnow, "DD/MM/YYYY")) >= 0 Then
58                 Set rsHlp = objO.FetchData("SELECT top 1 * FROM [" & dbReplikasiHO & "]..summary WITH (NOLOCK) WHERE TipeTr = 'NK' and tanggal >= '" & Format(Log.tglFreeSVCHO, "yyyy-mm-dd") & "' and NoRangka = '" & Nul(rsInvoice("NO_CHASSIS")) & "' order by  Tanggal desc")
59                 If Not rsHlp.EOF Then
60                    isTerlibatFreeSvcHO = True
61                 End If
62            End If
              
63            text(34) = Trim(rsInvoice("MOU"))
64            text(35) = Trim(rsInvoice("KODETIPE"))
65            n = IIf(Trim(rsInvoice("TPSVC")) = "P", 1, IIf(Trim(rsInvoice("TPSVC")) = "W", 2, 0))
66            opWO(n).value = True
67            LblReturn.Visible = IIf(rsTemp("FN_RETURN") = "1", True, False)
68            lblPDS.Visible = IIf(rsTemp("PDS") = "1", True, False)
              ' ==============================================================================================
              ' 08 Maret 2017 | Thania | penambahan informasi PWC utk cetakkan invoice & Cetak Ulang
              ' ==============================================================================================
69                Call panggilWarranty(Trim(rsInvoice("NO_WO")))
              ' ==============================================================================================
              
              
70            Set rsHlp = objBilling.searchDtlOrder(rsTemp("NO_WO"), rsTemp("NO_CHASSIS"))
              'oConn.Execute("SELECT * FROM JOBDTL WHERE NO_WO = '" & rsTemp("NO_WO") & "' AND NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "'")
71            If Not rsHlp.EOF Then
72                nHeavy = IIf(rsHlp("HEAVYUSE") <> 0, True, False)
73            End If
74            rsHlp.Close
              
              'If objBilling.isDuoInvoice(rsInvoice("KODETIPE")) Then nDuoInv = True
75            nDuoInv = False
76            vTipeCustomer = ""
77            text(40) = Trim(rsInvoice("STATUS"))
78            If returnJob(LblReturn) Then
79                If typeJob(Trim(rsTemp("TPSVC"))) = "PROGRAM" Then
80                    opWO(1).value = True
                       
81                    Set rsHlp = objBilling.searchProgramDetail(rsTemp("KODETIPE"))
82                    If Not rsHlp.EOF Then
83                        text(35) = rsHlp("ID_PROGRAM")
84                        text(36) = rsHlp("PROGRAM")
85                        text(35).Visible = True
86                        text(36).Visible = True
87                    End If
88                    rsHlp.Close
89                    If objBilling.isDuoInvoice(rsTemp("KODETIPE")) Then
90                        nDuoInv = True
                          'RTJ WO BERKEMBANG
                          '' 20190625DEV1
91                        If Log.lockFreeSVCHO = 1 And isTerlibatFreeSvcHO = True And DateDiff("D", Format(Log.tglFreeSVCHO, "DD/MM/YYYY"), Format(isnow, "DD/MM/YYYY")) >= 0 Then
92                            Set rsHlp = objO.FetchDataACC("SELECT KodeLgn AS KDPLG,NamaLgn AS FS_NM_DEALER,NamaLgn AS INIS_DEALER,Alamat1 + ' ' + Alamat2  AS FS_ALAMAT,Alamat3 AS FS_KOTA,Npwp AS FS_NPWP,1 AS INTERNAL FROM Armst WITH (NOLOCK) WHERE KodeLgn = '" & NulSID(gKdPlgHo) & "'")
93                        Else
94                            Set rsHlp = objBilling.takeKodeDealer(gKdCbg)
95                        End If
                          ''Set rsHlp = objBilling.takeKodeDealer(gKdCbg)
96                        If Not rsHlp.EOF Then
97                            gKdPlgDealer2 = Trim(rsHlp("KDPLG"))
98                            gKodePlg2 = Trim(rsHlp("KDPLG"))
99                            gInvCust2 = Trim(rsHlp("FS_NM_DEALER"))
100                           gInvAlmt2 = Trim(rsHlp("FS_ALAMAT"))
101                           gInvKota2 = Trim(rsHlp("FS_KOTA"))
102                           gInvNPWP2 = Trim(rsHlp("FS_NPWP"))
                              
103                           If Trim(rsHlp("INTERNAL")) = 1 Then
104                               If Trim(gNpwp) = Trim(gInvNPWP2) Then
105                                   vTipeCustomer = "AFFNONPPN"
106                               Else
107                                   vTipeCustomer = "AFFPPN"
108                               End If
109                           Else
110                               vTipeCustomer = "EXT"
111                           End If
112                       End If
113                       rsHlp.Close
114                       Call giveDiscount("100", text(21), False)
115                       txtCatatan.text = "A/N " & text(4) & "  [" & text(2) & "]"
116                   End If
117               Else
                      ' HANYA RTJ
118                   Call giveDiscount("100", text(21), False)
119                   vTipeCustomer = "INT"
120               End If
121           Else
122               If typeJob(Trim(rsTemp("TPSVC"))) = "WARRANTY" Then
123                   opWO(2).value = True
                       
124                   Set rsHlp = objBilling.searchWarrantyDetail(rsTemp("KODETIPE"))
125                   If Not rsHlp.EOF Then
126                       text(35) = rsHlp("ID_WARRANTY")
127                       text(36) = rsHlp("WARRANTY")
128                       text(35).Visible = True
129                       text(36).Visible = True
130                   End If
131                   rsHlp.Close
132                   Call giveDiscount("0", text(21), True)
133                   txtCatatan.text = "A/N " & text(4) & "  [" & text(2) & "]"
                       
134                   Set rsHlp = objO.FetchData("SELECT * FROM T_PROFIL WITH(NOLOCK)")
135                   If Not rsHlp.EOF Then
136                       gKdPlgDealer = Trim(rsHlp("FN_HOKDPLG"))
137                       gKodePlg = Trim(rsHlp("FN_HOKDPLG"))
138                       gInvCust = Trim(rsHlp("FN_HO"))
139                       gInvAlmt = Trim(rsHlp("FN_HOADRES"))
140                       gInvKota = Trim(rsHlp("FN_HOKOTA"))
141                       gInvNPWP = Trim(rsHlp("FN_HONPWP"))
142                       gInvPKP = Trim(rsHlp("FN_HOPKP"))
143                       vTipeCustomer = "AFFPPN"
144                   End If
145                   rsHlp.Close
146               ElseIf typeJob(Trim(rsTemp("TPSVC"))) = "PROGRAM" Then
147                   opWO(1).value = True
                       
148                   Set rsHlp = objBilling.searchProgramDetail(rsTemp("KODETIPE"))
149                   If Not rsHlp.EOF Then
150                       text(35) = rsHlp("ID_PROGRAM")
151                       text(36) = rsHlp("PROGRAM")
152                       text(35).Visible = True
153                       text(36).Visible = True
154                   End If
155                   rsHlp.Close
156                   If objBilling.isDuoInvoice(rsTemp("KODETIPE")) Then
157                       nDuoInv = True
158                       If rsTemp.RecordCount > 1 Then
159                           rsTemp.MoveNext
160                           metB = rsTemp!met
161                           rsTemp.MovePrevious
162                       Else
163                           metB = rsTemp!met
164                       End If
                          
                          'SBI WO BERKEMBANG
165                       If rsTemp("FS_RC") = "R001" Then
                              '' 20190625DEV1
166                           If Log.lockFreeSVCHO = 1 And isTerlibatFreeSvcHO = True And DateDiff("D", Format(Log.tglFreeSVCHO, "DD/MM/YYYY"), Format(isnow, "DD/MM/YYYY")) >= 0 Then
167                               Set rsHlp = objO.FetchDataACC("SELECT KodeLgn AS KDPLG,NamaLgn AS FS_NM_DEALER,NamaLgn AS INIS_DEALER,Alamat1 + ' ' + Alamat2  AS FS_ALAMAT,Alamat3 AS FS_KOTA,Npwp AS FS_NPWP,1 AS INTERNAL,PKP AS FS_PKP FROM Armst WITH (NOLOCK) WHERE KodeLgn = '" & NulSID(gKdPlgHo) & "'")
168                           Else
169                               Set rsHlp = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
170                           End If
                              ''Set rsHlp = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
171                           If Not rsHlp.EOF Then
172                               gKdPlgDealer2 = Trim(rsHlp("KDPLG"))
173                               gKodePlg2 = Trim(rsHlp("KDPLG"))
174                               gInvCust2 = Trim(rsHlp("FS_NM_DEALER"))
175                               gInvAlmt2 = Trim(rsHlp("FS_ALAMAT"))
176                               gInvKota2 = Trim(rsHlp("FS_KOTA"))
177                               gInvNPWP2 = Trim(rsHlp("FS_NPWP"))
178                               gInvPKP2 = Trim(rsHlp("FS_PKP"))
                                  
179                               If Trim(rsHlp("INTERNAL")) = 1 Then
180                                   If Trim(gNpwp) = Trim(gInvNPWP2) Then
181                                       vTipeCustomer = "AFFNONPPN"
182                                   Else
183                                       vTipeCustomer = "AFFPPN"
184                                   End If
185                               Else
186                                   vTipeCustomer = "EXT"
187                               End If
188                           End If
189                           rsHlp.Close
190                           Call giveDiscount("100", text(21), False)
191                           txtCatatan.text = "A/N " & text(4) & "  [" & text(2) & "]"
192                       Else
                              '' 20190625DEV1
193                           If Log.lockFreeSVCHO = 1 And isTerlibatFreeSvcHO = True And DateDiff("D", Format(Log.tglFreeSVCHO, "DD/MM/YYYY"), Format(isnow, "DD/MM/YYYY")) >= 0 Then
194                               Set rsHlp = objO.FetchDataACC("SELECT KodeLgn AS KDPLG,NamaLgn AS FS_NM_DEALER,NamaLgn AS INIS_DEALER,Alamat1 + ' ' + Alamat2  AS FS_ALAMAT,Alamat3 AS FS_KOTA,Npwp AS FS_NPWP,1 AS INTERNAL,PKP AS FS_PKP FROM Armst WITH (NOLOCK) WHERE KodeLgn = '" & NulSID(gKdPlgHo) & "'")
195                           Else
196                               Set rsHlp = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
197                           End If
                              ''Set rsHlp = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
198                           If Not rsHlp.EOF Then
199                               gKdPlgDealer2 = Trim(Nul(rsHlp("KDPLG")))
200                               gKodePlg2 = Trim(Nul(rsHlp("KDPLG")))
201                               gInvCust2 = Trim(Nul(rsHlp("FS_NM_DEALER")))
202                               gInvAlmt2 = Trim(Nul(rsHlp("FS_ALAMAT")))
203                               gInvKota2 = Trim(Nul(rsHlp("FS_KOTA")))
204                               gInvNPWP2 = Trim(Nul(rsHlp("FS_NPWP")))
205                               gInvPKP2 = Trim(Nul(rsHlp("FS_PKP")))
                                  
206                               If Trim(rsHlp("INTERNAL")) = 1 Then
207                                   If Trim(gNpwp) = Trim(gInvNPWP2) Then
208                                       vTipeCustomer = "AFFNONPPN"
209                                   Else
210                                       vTipeCustomer = "AFFPPN"
211                                   End If
212                               Else
213                                   vTipeCustomer = "EXT"
214                               End If
215                           End If
216                           rsHlp.Close
217                           Call giveDiscount("100", text(21), False)
218                           txtCatatan.text = "A/N " & text(4) & "  [" & text(2) & "]"
219                       End If
220                   Else
221                       Call giveDiscount(objBilling.takeDiscountProgram(rsTemp("KODETIPE")), text(21), False)
222                   End If
223               ElseIf typeJob(rsTemp("TPSVC")) = "REGULAR" Then
224                   If lblPDS.Visible Then
225                       Call giveDiscount("0", text(21), False)
226                       vTipeCustomer = "INT"
227                   ElseIf rsTemp("FS_RC") = "R001" Then
                          ' HANYA SBI
228                       woStat = "SBI"
                          '' 20190625DEV1
229                       If Log.lockFreeSVCHO = 1 And isTerlibatFreeSvcHO = True And DateDiff("D", Format(Log.tglFreeSVCHO, "DD/MM/YYYY"), Format(isnow, "DD/MM/YYYY")) >= 0 Then
230                           Set rsHlp = objO.FetchDataACC("SELECT KodeLgn AS KDPLG,NamaLgn AS FS_NM_DEALER,NamaLgn AS INIS_DEALER,Alamat1 + ' ' + Alamat2  AS FS_ALAMAT,Alamat3 AS FS_KOTA,Npwp AS FS_NPWP,1 AS INTERNAL,PKP AS FS_PKP FROM Armst WITH (NOLOCK) WHERE KodeLgn = '" & NulSID(gKdPlgHo) & "'")
231                       Else
232                           Set rsHlp = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
233                       End If
                          ''Set rsHlp = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
234                       If Not rsHlp.EOF Then
235                           gKdPlgDealer = IIf(Trim(rsHlp("KDPLG")) = "", "", Trim(rsHlp("KDPLG")))
236                           gKodePlg = IIf(Trim(rsHlp("KDPLG")) = "", Trim(text(2)), Trim(rsHlp("KDPLG")))
237                           gInvCust = IIf(Trim(rsHlp("KDPLG")) = "", Trim(text(4)), Trim(rsHlp("FS_NM_DEALER")))
238                           gInvAlmt = IIf(Trim(rsHlp("KDPLG")) = "", Trim(text(5)), Trim(rsHlp("FS_ALAMAT")))
239                           gInvKota = IIf(Trim(rsHlp("KDPLG")) = "", Trim(text(6)), Trim(rsHlp("FS_KOTA")))
240                           gInvNPWP = IIf(Trim(rsHlp("KDPLG")) = "", Trim(text(44)), Trim(rsHlp("FS_NPWP")))
241                           gInvPKP = IIf(Trim(rsHlp("KDPLG")) = "", Trim(text(45)), Trim(rsHlp("FS_PKP")))
                              
242                           If Trim(rsHlp("INTERNAL")) = 1 Then
243                               If Trim(gNpwp) = Trim(gInvNPWP) Then
244                                   text(40) = "AFFNONPPN"
245                               Else
246                                   text(40) = "AFFPPN"
247                               End If
248                           Else
249                               text(40) = "EXT"
250                           End If
251                       End If
252                       rsHlp.Close
253                       Call giveDiscount("0", text(21), False)
254                       txtCatatan.text = "A/N " & text(4) & "  [" & text(2) & "]"
255                   ElseIf objBilling.isMOUCustomer(gKdCbg, rsTemp("MDL"), rsTemp("NO_CHASSIS")) Then
                           
256                       Set rsHlp = objBilling.searchMOUDetail(gKdCbg, rsTemp("MDL"), rsTemp("NO_CHASSIS"))
257                       If Not rsHlp.EOF Then
258                           text(34) = rsHlp("NO_MOU")
259                       End If
260                       rsHlp.Close
261                       Call giveDiscount(objBilling.takeDiscountMOU(gKdCbg, rsTemp("MDL"), rsTemp("NO_CHASSIS")), text(21), False)
262                   Else
263                       Call giveDiscount("0", text(21), True)
264                   End If
265               End If
266           End If

267           If rsBantu.State Then rsBantu.Close
268           If n = 1 Then
                  
269               If Log.RPLHistori = 0 Then
270                   Set rsBantu = objO.FetchData("SELECT PROGRAM AS JENIS, * FROM HO_PROGRAM WITH(NOLOCK) WHERE ID_PROGRAM = '" & Trim(rsInvoice("KODETIPE")) & "'")
271               Else
272                   Set rsBantu = objO.FetchData("SELECT PROGRAM AS JENIS, * FROM [" & dbReplikasiHO & "]..HO_PROGRAM WITH(NOLOCK) WHERE ID_PROGRAM = '" & Trim(rsInvoice("KODETIPE")) & "'")
273               End If
274           ElseIf n = 2 Then
                  
275               If Log.RPLHistori = 0 Then
276                   Set rsBantu = objO.FetchData("SELECT WARRANTY AS JENIS, * FROM HO_WARRANTY WITH(NOLOCK) WHERE ID_WARRANTY = '" & Trim(rsInvoice("KODETIPE")) & "'")
277               Else
278                   Set rsBantu = objO.FetchData("SELECT WARRANTY AS JENIS, * FROM [" & dbReplikasiHO & "]..HO_WARRANTY WITH(NOLOCK) WHERE ID_WARRANTY = '" & Trim(rsInvoice("KODETIPE")) & "'")
279               End If
280           End If
281           If n > 0 Then
282               If Not rsBantu.EOF Then
283                   text(36) = Trim(rsBantu("JENIS"))
284               End If
285               If rsBantu.State <> adStateClosed Then rsBantu.Close
286           End If
              
287           text(37) = CDbl(rsInvoice("KM"))
288           text(38) = Trim(rsInvoice("MECH"))
289           text(43) = TglDDMMYYYY(rsInvoice("DATE"))
              
              ' PANGGIL JOB ORDER
              Dim hlp As String
              Dim hlp2 As String
              Dim hlp3 As String
              Dim hlp4 As String
              Dim hlp5 As String
              Dim hlp6 As String
              Dim nilPPN As Double
              
290           subJob = 0
291           If rsJob.State Then rsJob.Close
292           If rsJob.State = adStateClosed Then doCreateJob
              
293           If NulSID(rsInvoice!job) < 3 And Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(text(43), "DD/MM/YYYY")) >= 0 Then
294               If typeJob(Trim(rsTemp("TPSVC"))) <> "WARRANTY" Then
295                   Call loadJobDetailMultiple("", NulSID(rsInvoice("NO_CHASSIS")))
296                   GoTo lanjutJobOrder
297               End If
298           End If
299           For i = 1 To 7
300               hlp = "REP" & i
301               hlp2 = "RTE" & i
302               hlp3 = "DISC" & i
303               hlp4 = "OFF" & i
304               hlp5 = "FREEC" & i
305               hlp6 = "VALID" & i
                  
306               If Trim(rsInvoice(hlp)) <> "" Then
307                   nilPPN = (nilaiDppInv / nilaiTotalInv) * CDbl(rsInvoice(hlp4))
308                   With rsJob
309                       .AddNew
310                       !job = Trim(rsInvoice(hlp))
311                       !Rate = CDbl(rsInvoice(hlp2))
312                       !addJob = False
313                       !hrg = CDbl(rsInvoice(hlp4)) '- IIf(Trim(text(40)) = "AFFNONPPN", nilPPN, 0)
314                       !disc = CDbl(rsInvoice(hlp3))
315                       !STLDISC = Round(!hrg * ((100 - !disc) / 100))
316                       !FREEC = IIf(Trim(rsInvoice(hlp5)) = 1, True, False)
317                       .Update
318                       text(22) = CDbl(text(22)) + CDbl(!STLDISC)
319                       subJob = subJob + IIf(!FREEC, 0, CDbl(!hrg))
320                   End With
321               End If
322           Next
              
              ' DARI ADDJOB
323           If lblRetur.Visible Then
324               Set rsBantu = objO.FetchData("SELECT * FROM HIST_JOBADD WITH(NOLOCK) WHERE NO_WO = '" & Trim(rsInvoice("NO_WO")) & "' AND No_INV = '" & text(0) & "'") '   FD_TGL_DEL = '" & Format(rsInvoice("FD_RETUR"), "MM/DD/YYYY HH:NN") & "'")
325           Else
326               Set rsBantu = objO.FetchData("SELECT * FROM HIST_JOBADD WITH(NOLOCK) WHERE NO_WO = '" & Trim(rsInvoice("NO_WO")) & "' AND FS_LOG_DEL = ''")
327           End If
328           If Not rsBantu.EOF Then
329               If rsJob.State = adStateClosed Then doCreateJob
330               Do While Not rsBantu.EOF
331                   nilPPN = (nilaiDppInv / nilaiTotalInv) * CDbl(rsBantu("OFFN"))
332                   With rsJob
333                       .AddNew
334                       !job = Trim(rsBantu("REP"))
335                       !Rate = CDbl(rsBantu("RTE")) 'IIf(Trim(rsBantu("VALID")) <> "", 0, CDbl(rsBantu("RTE")))
336                       !addJob = True
337                       !hrg = CDbl(rsBantu("OFFN")) '- IIf(Trim(text(40)) = "AFFNONPPN", nilPPN, 0) 'IIf(Trim(rsBantu("VALID")) <> "", 0, CDbl(rsBantu("OFFN")))
338                       !disc = CDbl(rsBantu("DISC"))
339                       !STLDISC = CDbl(!hrg * ((100 - !disc) / 100))
340                       !FREEC = IIf(Trim(rsBantu("FREEC")) = 1, True, False)
341                       .Update
342                       text(22) = CDbl(text(22)) + CDbl(!STLDISC)
343                       subJob = subJob + IIf(!FREEC, 0, CDbl(!hrg))
344                   End With
345                   rsBantu.MoveNext
346               Loop
                  'cetakAdd.Visible = True
347           End If
lanjutJobOrder:
348           If rsBantu.State <> adStateClosed Then rsBantu.Close
349           Set jobGrid.DataSource = rsJob
350           Call doHeadJob
              
              ' PANGGIL PART
351           text(24) = 0
352           text(25) = 0
353           Set rsBantu = objO.FetchData("SELECT distinct * FROM T_PART WITH(NOLOCK) WHERE NO_WO = '" & Trim(rsInvoice("NO_WO")) & "' AND NOINV LIKE '" & Left(Trim(rsInvoice("NOINV")), 8) & "%'") 'AND FS_LOG_DEL = ''")
354           If Not rsBantu.EOF Then
355               If rsPart.State Then rsPart.Close
356               If rsPart.State = adStateClosed Then doCreatePart
357               Do While Not rsBantu.EOF
358                   With rsPart
359                       .AddNew
360                       !KodeItem = Trim(rsBantu("KODEITEM"))
361                       !NAMAITEM = Trim(rsBantu("KETERANGAN"))
362                       !qty = CDbl(rsBantu("QTY"))
363                       !Harga = CDbl(rsBantu("HARGA"))
364                       !oli = IIf(Trim(rsBantu("OLI")), True, False)
365                       !EXTDISC = CDbl(rsBantu("FN_EXT_DISKON"))
366                       .Update
367                       If Not !oli Then
368                           text(24) = CDbl(text(24)) + CDbl(!Harga)
369                       Else
370                           text(25) = CDbl(text(25)) + CDbl(!Harga)
371                       End If
372                   End With
373                   rsBantu.MoveNext
374               Loop
375           End If
376           If rsBantu.State <> adStateClosed Then rsBantu.Close
377           Set partGrid.DataSource = rsPart
378           Call doHeadPart
              'Text(24) = CDbl(rsInvoice("PART_G")) + CDbl(rsInvoice("PART_NG"))
              'Text(25) = CDbl(rsInvoice("OIL_G")) + CDbl(rsInvoice("OIL_NG"))
              
              ' PANGGIL SUBLET
379           If rsSublet.State Then rsSublet.Close
              '(SELECT SUM(HARGA) FROM HIST_SUBLET WHERE NO_WO = A.NO_WO AND NOINV = A.NOINV AND NO_CHASSIS = A.NO_CHASSIS GROUP BY NO_WO) AS HARGA
              

380           If Log.freepart = 1 Then
381               Set rsSublet = objO.FetchData("SELECT a.NONOTA, a.SUPPLIER, a.NAMABARANG, a.KODESATUAN, a.QTY, 'TRUE' AS VALID, isnull(case when A.KDPROMO = '' then B.Hargajual * " & nilaiTotalInv & " / " & nilaiDppInv & "  else e.HARGAJUAL end,0) AS HARGA,ISNULL(D.DISCbilling,0) diskon , case when right(noinv,1) = 'C' then 'FALSE' ELSE 'TRUE' END as FREEC FROM HIST_SUBLET A WITH(NOLOCK)  LEFT JOIN [ReplBengkel99]..MSTSUBLETPROMO e WITH (NOLOCK) ON e.KDPROMO = A.KDPROMO  LEFT JOIN [" & dbReplikasiHO & "]..MSTSUBLET B WITH (NOLOCK) ON A.KODEITEM = B.KDSUBLET AND B.FS_LOG_DEL = '' LEFT JOIN [vo_discsublet] D with (nolock) on a.NO_WO = d.NO_WO and a.NO_CHASSIS = d.NO_CHASSIS and d.kodeitem = A.KODEITEM AND D.FS_LOG_DEL = ''  WHERE a.NO_WO = '" & NulSID(rsInvoice("NO_WO")) & "' AND a.NOINV like '" & NulSID(text(0)) & "%'")
382           Else
383               If Log.standarisasiSublet = 1 Then
384                   Set rsSublet = objO.FetchData("SELECT a.NONOTA, a.SUPPLIER, a.NAMABARANG, a.KODESATUAN, a.QTY, 'TRUE' AS VALID, case when A.KDPROMO = '' then B.Hargajual * " & nilaiTotalInv & " / " & nilaiDppInv & " else e.HARGAJUAL end AS HARGA,ISNULL(D.DISCbilling,0) diskon,HARGA AS PRICE FROM HIST_SUBLET A WITH(NOLOCK) LEFT JOIN [ReplBengkel99]..MSTSUBLETPROMO e WITH (NOLOCK) ON e.KDPROMO = A.KDPROMO   LEFT JOIN [" & dbReplikasiHO & "]..MSTSUBLET B WITH (NOLOCK) ON A.KODEITEM = B.KDSUBLET AND B.FS_LOG_DEL = '' LEFT JOIN [vo_discsublet] D with (nolock) on a.NO_WO = d.NO_WO and a.NO_CHASSIS = d.NO_CHASSIS and d.kodeitem = A.KODEITEM AND D.FS_LOG_DEL = '' WHERE a.NO_WO = '" & Trim(rsInvoice("NO_WO")) & "' AND a.NOINV LIKE '" & Left(Trim(rsInvoice("NOINV")), 8) & "%'")
385               Else
386                   Set rsSublet = objO.FetchData("SELECT NONOTA, SUPPLIER, NAMABARANG, KODESATUAN, QTY, 'TRUE' AS VALID, HARGA AS HARGA,0 diskon,harga as price FROM HIST_SUBLET A WITH(NOLOCK) WHERE NO_WO = '" & Trim(rsInvoice("NO_WO")) & "' AND NOINV LIKE '" & Left(Trim(rsInvoice("NOINV")), 8) & "%'")
387               End If
388           End If
389           If Not rsSublet.EOF Then
390               Set subletGrid.DataSource = rsSublet
391               Call doHeadSublet
392               Do While Not rsSublet.EOF
393                   If Log.freepart = 1 Then
394                       If rsSublet!FREEC Then
                              'text(27) = CDbl(text(27)) + ((CDbl(rsSublet("price")) * CDbl(rsSublet("QTY")))) '
395                           text(27) = CDbl(text(27)) + ((CDbl(rsSublet("HARGA")) * CDbl(rsSublet("QTY"))) * ((100 - CDbl(rsSublet!diskon)) / 100))
396                       End If
397                   Else
398                       text(27) = CDbl(text(27)) + SpyRound(rsSublet!Price * rsSublet!qty) ''SpyRound(((CDbl(rsSublet("HARGA")) * ((100 - rsSublet!diskon) / 100) * CDbl(rsSublet("QTY")))), 0)
399                   End If
400                   rsSublet.MoveNext
401               Loop
402               text(27) = SpyRound(text(27))
403               cetakSublet.Visible = True
404           End If
              
405           text(39) = CDbl(rsInvoice("BODY"))
406           text(28) = CDbl(rsInvoice("MTRL"))
407           text(42) = CDbl(rsInvoice("MET"))
408           text(29) = CDbl(text(22)) + CDbl(text(24)) + CDbl(text(25)) + CDbl(text(27)) + CDbl(text(28)) + CDbl(text(39)) + CDbl(text(42)) 'CDbl(rsInvoice("TOTAL"))
              'Text(29) = Round(Text(29) - IIf(nDuoInv, CDbl(Text(42)), 0), 0)
409            text(29) = (text(29) - IIf(nDuoInv, CDbl(text(42)), 0))
410            text(29) = objO.FetchData_String("SELECT ROUND(" & Replace(text(29), ",", ".") & ",0)")

              
411           txtTotalRp = Angka2Huruf(text(29), 1)
              
412           If Log.DIDX = 1 Then
413               If rsBantu.State <> adStateClosed Then rsBantu.Close
                  
414               sSQL = "SELECT NAMAPENERIMA FROM InvoiceDX WITH (NOLOCK) WHERE NO_WO = '" & NulSID(text(20)) & "' AND NO_CHASSIS = '" & NulSID(text(17)) & "' AND NOINV = '" & NulSID(text(0)) & "' "
415               Set rsBantu = objO.FetchData(sSQL)
416               If Not rsBantu.EOF Then
417                   text(41) = NulSID(rsBantu!namapenerima)
418               End If
                  
419           End If
420       End If
421       rsInvoice.Close

         'On Error GoTo 0
422      Exit Sub

panggilINV_Error:

423       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure panggilINV of Form frmBillingUlang")
End Sub

Private Function typeJob(statusJob As String) As String
          
1        On Error GoTo typeJob_Error

2         If statusJob = "W" Then
3             typeJob = "WARRANTY"
4         ElseIf statusJob = "P" Then
5             typeJob = "PROGRAM"
6         Else
7             typeJob = "REGULAR"
8         End If

9        'On Error GoTo 0
10       Exit Function

typeJob_Error:

11        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure typeJob of Form frmBillingUlang")
End Function

Private Sub VScroll1_Change()
'    Frame6.Top = -VScroll1.value + 20
    If VScroll1.value > 0 Then
        Frame6.Top = -VScroll1.value + 20
    Else
        Frame6.Top = 0
    End If
End Sub

Private Sub VScroll1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        text(0).SetFocus
    End If
End Sub

Private Sub VScroll1_Scroll()
    Frame6.Top = -VScroll1.value + 20
End Sub

Private Function createTemp(vNamaTabel As String) As Boolean
      Dim rsHlp As New ADODB.Recordset

1        On Error GoTo createTemp_Error

2         If Get_Connectiontemp("tempdb", gUser, gPwd, gServerDb) Then
3             createTemp = False
               
4             Set rsHlp = objO.FetchData("select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ")
5             If Not rsHlp.EOF Then
6                 createTemp = True
7             Else
8                 Call BuatTblInv(vNamaTabel)
9                 createTemp = True
10            End If
11        End If

12       'On Error GoTo 0
13       Exit Function

createTemp_Error:

14        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure createTemp of Form frmBillingUlang")
End Function

Public Sub BuatTblInv(vNamaTabel As String)
   On Error GoTo BuatTblInv_Error

    sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ) " _
         & " drop table [tempdb]..[" & vNamaTabel & "]"
    Call objO.UpdateData(sSQL)
    
    sSQL = "CREATE TABLE [tempdb]..[" & vNamaTabel & "] ( " _
         & "[REV_CETAK] [decimal] (18,2) NOT NULL DEFAULT 0, "
    sSQL = sSQL & "[REV_HITUNG] [decimal] (18,2) NOT NULL DEFAULT 0, " _
         & "[NO_CHASSIS] [varchar] (25) NOT NULL DEFAULT '', " _
         & "[NOINV] [varchar] (20) NOT NULL DEFAULT '', " _
         & "[INVDATE] [datetime] NOT NULL DEFAULT '01/01/1900', " _
         & "[FS_KD_PLG] [varchar] (15) NOT NULL DEFAULT '', " _
         & "[CUSTOMER] [varchar] (50) NOT NULL DEFAULT '', " _
         & "[CUST_ADRES] [varchar] (250) NOT NULL DEFAULT '', " _
         & "[KOTA] [varchar] (100) NOT NULL DEFAULT '', " _
         & "[NPWP] [varchar] (50) NOT NULL DEFAULT '', " _
         & "[NO_POLISI] [varchar] (15) NOT NULL DEFAULT '', " _
         & "[FMDL] [varchar] (200) NOT NULL DEFAULT '', " _
         & "[NO_ENGINE] [varchar] (15) NOT NULL DEFAULT '', " _
         & "[NO_WO] [varchar] (15) NOT NULL DEFAULT ''," _
         & "[FRONT] [varchar] (10) NOT NULL DEFAULT '', " _
         & "[STATUS] [varchar] (10) NOT NULL DEFAULT '', " _
         & "[MECH] [varchar] (10) NOT NULL DEFAULT '', " _
         & "[FS_LOG_ADD] [varchar] (10) NOT NULL DEFAULT '', " _
         & "[SBLT] [decimal] (18,2) NOT NULL DEFAULT 0, " _
         & "[PART] [decimal] (18,2) NOT NULL DEFAULT 0, " _
         & "[OLI] [decimal] (18,2) NOT NULL DEFAULT 0, " _
         & "[MTRL] [decimal] (18,2) NOT NULL DEFAULT 0, " _
         & "[MET] [decimal] (18,2) NOT NULL DEFAULT 0, " _
         & "[TOTAL] [decimal] (18,2) NOT NULL DEFAULT 0, " _
         & "[SLIP] [varchar] (100) NOT NULL DEFAULT '', "
    sSQL = sSQL + "[FS_CATATAN] [varchar] (100) NOT NULL DEFAULT '', " _
                & "[REP] [varchar] (100) NOT NULL DEFAULT '', " _
                & "[KLPKERJA] [varchar] (5) NOT NULL DEFAULT '', " _
                & "[RTE] [varchar] (5) NOT NULL DEFAULT '', " _
                & "[OFF] [decimal] (18,2) NOT NULL DEFAULT 0, " _
                & "[DISC] [decimal] (18,4) NOT NULL DEFAULT 0, " _
                & "[CTKDISC] [char] (1) NOT NULL DEFAULT 0, " _
                & "[BP] [decimal] (18,2) NOT NULL DEFAULT 0,[FN_PAJAK] [decimal] (18,2) NOT NULL DEFAULT 0) ON [PRIMARY]"
    Call objO.UpdateData(sSQL)

   'On Error GoTo 0
   Exit Sub

BuatTblInv_Error:

    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure BuatTblInv of Form frmBillingUlang")
End Sub

Private Function hitungINV2(sTipeCust As String, vNamaTabel As String) As Double
          Dim total As Double
          Dim rsTamp As New ADODB.Recordset
          Dim totalPPn As Double
1        On Error GoTo hitungINV2_Error

2         If Get_Connectiontemp("tempdb", gUser, gPwd, gServerDb) Then
3             sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ) " _
                   & " drop table [tempdb]..[" & vNamaTabel & "]"
4             objO.UpdateData (sSQL)
5         End If
          
          ' INVOICE
6         total = 0
          ''' ASLI NICH !!!
      '    sSQL = "SELECT '" & Trim(gInvNo) & "' AS NOINV, INVDATE, '" & Trim(gKodePlg) & "' AS FS_KD_PLG, '" & Trim(gInvCust) & "' AS CUSTOMER, '" & Trim(gInvAlmt) & "' AS CUST_ADRES, '" & Trim(gInvKota) & "' AS KOTA, '" & Trim(gInvNPWP) & "' AS NPWP, NO_POLISI, FMDL, NO_CHASSIS, NO_ENGINE, NO_WO, FRONT, STATUS, MECH, FS_LOG_ADD, " _
      '         & "SBLT, PART_G+PART_NG AS PART, OIL_G+OIL_NG AS OLI, MTRL, MET, BODY, SLIP, FS_CATATAN, TOTAL, " _
      '         & "REP1, REP2, REP3, REP4, REP5, REP6, REP7, REP8, RTE1, RTE2, RTE3, RTE4, RTE5, RTE6, RTE7, RTE8, " _
      '         & "KLPKERJA1 , KLPKERJA2, KLPKERJA3, KLPKERJA4, KLPKERJA5, KLPKERJA6, KLPKERJA7, KLPKERJA8, " _
      '         & "Off1 , Off2, Off3, Off4, Off5, Off6, Off7, Off8, DISC1, DISC2, DISC3, DISC4, DISC5, DISC6, DISC7, DISC8, " _
      '         & "FREEC1, FREEC2, FREEC3, FREEC4, FREEC5, FREEC6, FREEC7, FREEC8 " _
      '         & "FROM T_INVOICE WHERE NOINV = '" & Trim(text(0)) & "' AND NO_WO = '" & Trim(text(20)) & "'"
          '#########################################################'
          '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..'
          '#########################################################'
7         sSQL = "SELECT '" & Trim(gInvNo) & "' AS NOINV, INVDATE, '" & Trim(gKodePlg) & "' AS FS_KD_PLG, '" & Trim(gInvCust) & "' AS CUSTOMER, '" & Trim(gInvAlmt) & "' AS CUST_ADRES, '" & Trim(gInvKota) & "' AS KOTA, '" & Trim(gInvNPWP) & "' AS NPWP, NO_POLISI, FMDL, NO_CHASSIS, NO_ENGINE, NO_WO, FRONT, STATUS, MECH, FS_LOG_ADD, " _
               & "SBLT, PART_G+PART_NG AS PART, OIL_G+OIL_NG AS OLI, MTRL, MET, BODY, SLIP, FS_CATATAN, TOTAL, " _
               & "REP1, REP2, REP3, REP4, REP5, REP6, REP7, REP8, RTE1, RTE2, RTE3, RTE4, RTE5, RTE6, RTE7, RTE8, " _
               & "KLPKERJA1 , KLPKERJA2, KLPKERJA3, KLPKERJA4, KLPKERJA5, KLPKERJA6, KLPKERJA7, KLPKERJA8, " _
               & "Off1 , Off2, Off3, Off4, Off5, Off6, Off7, Off8, DISC1, DISC2, DISC3, DISC4, DISC5, DISC6, DISC7, DISC8, " _
               & "FREEC1, FREEC2, FREEC3, FREEC4, FREEC5, FREEC6, FREEC7, FREEC8, fn_pajak " _
               & "FROM T_INVOICE WITH(NOLOCK) WHERE NOINV = '" & Trim(text(0)) & IIf(nDuoInv, "-B", "") & "' AND NO_WO = '" & Trim(text(20)) & "' ORDER BY NOINV"
               
8          totalPPn = 0
9         Set rsTamp = objO.FetchData(sSQL)
10        If Not rsTamp.EOF Then
11            If createTemp(vNamaTabel) Then
12                onlySublet = True
13                total = total + Replace(CDbl(rsTamp("MET")), ",", ".")
14                totalPPn = CDbl(rsTamp!fn_pajak)
                  
15                Do While Not rsTamp.EOF
16                    For i = 1 To 8
17                        hlp1 = "REP" & i
18                        hlp2 = "RTE" & i
19                        hlp3 = "OFF" & i
20                        hlp4 = "DISC" & i
21                        hlp6 = "KLPKERJA" & i
22                        hlp7 = "FREEC" & i
                          
23                        If Trim(rsTamp(hlp1)) <> "" Then
      '                        jPPN = Round((nilaiDppInv / nilaiTotalInv,0)  ''10 / 110) * CDbl(rsTamp(hlp3)))
      '' 2016-03-24 || jppn diambil dari nilai off dikurangi diskon baru diambil dpp
24                                  jPPN = Round((nilaiDppInv / nilaiTotalInv) * (CDbl(rsTamp(hlp3)) * ((100 - CDbl(rsTamp(hlp4))) / 100)))
                                    
                              '--:: DATA AWAL ::--
      '                        oconn.Execute ("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & Trim(rsTamp("NO_CHASSIS")) & "','" & Trim(rsTamp("NOINV")) & "','" & Format(rsTamp("INVDATE"), "MM/DD/YYYY") & "','" & Trim(rsTamp("FS_KD_PLG")) & "','" & Trim(rsTamp("CUSTOMER")) & "','" & Trim(rsTamp("CUST_ADRES")) & "', " _
      '                                     & "'" & Trim(rsTamp("KOTA")) & "','" & Trim(rsTamp("NPWP")) & "','" & Trim(rsTamp("NO_POLISI")) & "','" & Trim(rsTamp("FMDL")) & "','" & Trim(rsTamp("NO_ENGINE")) & "','" & Trim(rsTamp("NO_WO")) & "','" & Trim(rsTamp("FRONT")) & "', " _
      '                                     & "'" & Trim(rsTamp("STATUS")) & "','" & Trim(rsTamp("MECH")) & "','" & Trim(rsTamp("FS_LOG_ADD")) & "','" & CDbl(rsTamp("SBLT")) & "','" & CDbl(rsTamp("PART")) & "','" & CDbl(rsTamp("OLI")) & "','" & CDbl(rsTamp("MTRL")) & "', " _
      '                                     & "'" & CDbl(rsTamp("MET")) & "','" & Trim(rsTamp("SLIP")) & "','" & Trim(rsTamp("FS_CATATAN")) & "','" & Trim(rsTamp(hlp1)) & "','" & Trim(rsTamp(hlp6)) & "','" & IIf(ChkCtkRate, CDbl(rsTamp(hlp2)), "") & "','" & Round(CDbl(rsTamp(hlp3)) - IIf(sTipeCust = "AFFNONPPN", jPPN, 0)) & "','" & Replace(100 - IIf(Trim(rsTamp(hlp6)) = "SBE" Or Trim(rsTamp(hlp6)) = "SBI", IIf(Trim(rsTamp(hlp7)) = 1, 0, CDbl(rsTamp(hlp4))), 0), ",", ".") & "','" & IIf(chCtkDisc, 0, 1) & "','" & CDbl(rsTamp("BODY")) & "') ")
      '                        total = total + Round(CDbl(rsTamp(hlp3)) - IIf(sTipeCust = "AFFNONPPN", jPPN, 0) - ((CDbl(rsTamp(hlp3)) - IIf(sTipeCust = "AFFNONPPN", jPPN, 0)) * (1 - IIf(Trim(rsTamp(hlp6)) = "SBE" Or Trim(rsTamp(hlp6)) = "SBI", (IIf(Trim(rsTamp(hlp7)) = 1, 0, CDbl(rsTamp(hlp4))) / 100), 0))))
                              '--:: 25 AGUSTUS 2010 ; DISPLAY JASA ALL ::--
                              ''' ASLI NICH !!!
      '                        oConn.Execute ("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & Trim(rsTamp("NO_CHASSIS")) & "','" & Trim(rsTamp("NOINV")) & "','" & Format(rsTamp("INVDATE"), "MM/DD/YYYY") & "','" & Trim(rsTamp("FS_KD_PLG")) & "','" & Trim(rsTamp("CUSTOMER")) & "','" & Trim(rsTamp("CUST_ADRES")) & "', " _
      '                                     & "'" & Trim(rsTamp("KOTA")) & "','" & Trim(rsTamp("NPWP")) & "','" & Trim(rsTamp("NO_POLISI")) & "','" & Trim(rsTamp("FMDL")) & "','" & Trim(rsTamp("NO_ENGINE")) & "','" & Trim(rsTamp("NO_WO")) & "','" & Trim(rsTamp("FRONT")) & "', " _
      '                                     & "'" & Trim(rsTamp("STATUS")) & "','" & Trim(rsTamp("MECH")) & "','" & Trim(rsTamp("FS_LOG_ADD")) & "','" & CDbl(rsTamp("SBLT")) & "','" & CDbl(rsTamp("PART")) & "','" & CDbl(rsTamp("OLI")) & "','" & CDbl(rsTamp("MTRL")) & "', " _
      '                                     & "'" & CDbl(rsTamp("MET")) & "','" & Trim(rsTamp("SLIP")) & "','" & Trim(rsTamp("FS_CATATAN")) & "','" & Trim(rsTamp(hlp1)) & "','" & Trim(rsTamp(hlp6)) & "','" & IIf(ChkCtkRate, CDbl(rsTamp(hlp2)), "") & "','" & Round(CDbl(rsTamp(hlp3)) - IIf(sTipeCust = "AFFNONPPN", jPPN, 0)) & "','" & Replace(100 - IIf(Trim(rsTamp(hlp7)) = 1, 0, CDbl(rsTamp(hlp4))), ",", ".") & "','" & IIf(chCtkDisc, 0, 1) & "','" & CDbl(rsTamp("BODY")) & "') ")
      '                        total = total + Round(CDbl(rsTamp(hlp3)) - IIf(sTipeCust = "AFFNONPPN", jPPN, 0) - ((CDbl(rsTamp(hlp3)) - IIf(sTipeCust = "AFFNONPPN", jPPN, 0)) * (1 - (IIf(Trim(rsTamp(hlp7)) = 1, 0, CDbl(rsTamp(hlp4))) / 100))))
                              '#########################################################'
                              '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..'
                              '#########################################################'
      '                        Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES (0,0,'" & Trim(rsTamp("NO_CHASSIS")) & "','" & Trim(rsTamp("NOINV")) & "','" & Format(rsTamp("INVDATE"), "MM/DD/YYYY") & "','" & Trim(rsTamp("FS_KD_PLG")) & "','" & Trim(rsTamp("CUSTOMER")) & "','" & Trim(rsTamp("CUST_ADRES")) & "', " _
      '                                     & "'" & Trim(rsTamp("KOTA")) & "','" & Trim(rsTamp("NPWP")) & "','" & Trim(rsTamp("NO_POLISI")) & "','" & Trim(rsTamp("FMDL")) & "','" & Trim(rsTamp("NO_ENGINE")) & "','" & Trim(rsTamp("NO_WO")) & "','" & Trim(rsTamp("FRONT")) & "', " _
      '                                     & "'" & Trim(rsTamp("STATUS")) & "','" & Trim(rsTamp("MECH")) & "','" & Trim(rsTamp("FS_LOG_ADD")) & "','" & RepRp(CDbl(rsTamp("SBLT"))) & "','" & RepRp(CDbl(rsTamp("PART"))) & "','" & RepRp(CDbl(rsTamp("OLI"))) & "','" & RepRp(CDbl(rsTamp("MTRL"))) & "', " _
      '                                     & "'" & RepRp(CDbl(metB)) & "'," & RepRp(CDbl(rsTamp("TOTAL"))) & ",'" & Trim(rsTamp("SLIP")) & "','" & Trim(rsTamp("FS_CATATAN")) & "','" & Trim(rsTamp(hlp1)) & "','" & Trim(rsTamp(hlp6)) & "','" & IIf(ChkCtkRate, CDbl(rsTamp(hlp2)), "") & "','" & Round(CDbl(rsTamp(hlp3)) - IIf(sTipeCust = "AFFNONPPN", jPPN, 0)) & "','" & setKoma(CDbl(rsTamp(hlp4))) & "','" & IIf(chCtkDisc, 0, 1) & "','" & RepRp(CDbl(rsTamp("BODY"))) & "'," & RepRp(CDbl(rsTamp("fn_pajak"))) & ") ")
25                            Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES (0,0,'" & Trim(rsTamp("NO_CHASSIS")) & "','" & Trim(rsTamp("NOINV")) & "','" & Format(rsTamp("INVDATE"), "MM/DD/YYYY") & "','" & Trim(rsTamp("FS_KD_PLG")) & "','" & Trim(rsTamp("CUSTOMER")) & "','" & Trim(rsTamp("CUST_ADRES")) & "', " _
                                           & "'" & Trim(rsTamp("KOTA")) & "','" & Trim(rsTamp("NPWP")) & "','" & Trim(rsTamp("NO_POLISI")) & "','" & Trim(rsTamp("FMDL")) & "','" & Trim(rsTamp("NO_ENGINE")) & "','" & Trim(rsTamp("NO_WO")) & "','" & Trim(rsTamp("FRONT")) & "', " _
                                           & "'" & Trim(rsTamp("STATUS")) & "','" & Trim(rsTamp("MECH")) & "','" & Trim(rsTamp("FS_LOG_ADD")) & "','" & RepRp(CDbl(rsTamp("SBLT"))) & "','" & RepRp(CDbl(rsTamp("PART"))) & "','" & RepRp(CDbl(rsTamp("OLI"))) & "','" & RepRp(CDbl(rsTamp("MTRL"))) & "', " _
                                           & "'" & RepRp(CDbl(metB)) & "'," & RepRp(CDbl(rsTamp("TOTAL"))) & ",'" & Trim(rsTamp("SLIP")) & "','" & Trim(rsTamp("FS_CATATAN")) & "','" & Trim(rsTamp(hlp1)) & "','" & Trim(rsTamp(hlp6)) & "','" & IIf(ChkCtkRate, CDbl(rsTamp(hlp2)), "") & "','" & Round(CDbl(rsTamp(hlp3)) - IIf(sTipeCust = "AFFNONPPN", jPPN, 0)) & "','" & setKoma(CDbl(rsTamp(hlp4))) & "','" & IIf(chCtkDisc, 0, 1) & "','" & RepRp(CDbl(rsTamp("BODY"))) & "'," & RepRp(CDbl(totalPPn)) & ") ")
26                            total = total + Round(CDbl(rsTamp(hlp3)) - IIf(sTipeCust = "AFFNONPPN", jPPN, 0) - ((CDbl(rsTamp(hlp3)) - IIf(sTipeCust = "AFFNONPPN", jPPN, 0)) * (CDbl(rsTamp(hlp4)) / 100)))
27                            onlySublet = False
28                        End If
29                    Next
                      
30                    If onlySublet Then
      '                    Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES (0,0,'" & Trim(rsTamp("NO_CHASSIS")) & "','" & Trim(rsTamp("NOINV")) & "','" & Format(rsTamp("INVDATE"), "MM/DD/YYYY") & "','" & Trim(rsTamp("FS_KD_PLG")) & "','" & Trim(rsTamp("CUSTOMER")) & "','" & Trim(rsTamp("CUST_ADRES")) & "', " _
      '                                 & "'" & Trim(rsTamp("KOTA")) & "','" & Trim(rsTamp("NPWP")) & "','" & Trim(rsTamp("NO_POLISI")) & "','" & Trim(rsTamp("FMDL")) & "','" & Trim(rsTamp("NO_ENGINE")) & "','" & Trim(rsTamp("NO_WO")) & "','" & Trim(rsTamp("FRONT")) & "', " _
      '                                 & "'" & Trim(rsTamp("STATUS")) & "','" & Trim(rsTamp("MECH")) & "','" & Trim(rsTamp("FS_LOG_ADD")) & "','" & RepRp(CDbl(rsTamp("SBLT"))) & "','" & RepRp(CDbl(rsTamp("PART"))) & "','" & RepRp(CDbl(rsTamp("OLI"))) & "','" & RepRp(CDbl(rsTamp("MTRL"))) & "', " _
      '                                 & "'" & RepRp(CDbl(metB)) & "'," & RepRp(CDbl(rsTamp("TOTAL"))) & ",'" & Trim(rsTamp("SLIP")) & "','" & Trim(rsTamp("FS_CATATAN")) & "','','','0','0','0','" & IIf(chCtkDisc, 0, 1) & "','" & RepRp(CDbl(rsTamp("BODY"))) & "'," & RepRp(CDbl(rsTamp("fn_pajak"))) & ") ")
31                         Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES (0,0,'" & Trim(rsTamp("NO_CHASSIS")) & "','" & Trim(rsTamp("NOINV")) & "','" & Format(rsTamp("INVDATE"), "MM/DD/YYYY") & "','" & Trim(rsTamp("FS_KD_PLG")) & "','" & Trim(rsTamp("CUSTOMER")) & "','" & Trim(rsTamp("CUST_ADRES")) & "', " _
                                       & "'" & Trim(rsTamp("KOTA")) & "','" & Trim(rsTamp("NPWP")) & "','" & Trim(rsTamp("NO_POLISI")) & "','" & Trim(rsTamp("FMDL")) & "','" & Trim(rsTamp("NO_ENGINE")) & "','" & Trim(rsTamp("NO_WO")) & "','" & Trim(rsTamp("FRONT")) & "', " _
                                       & "'" & Trim(rsTamp("STATUS")) & "','" & Trim(rsTamp("MECH")) & "','" & Trim(rsTamp("FS_LOG_ADD")) & "','" & RepRp(CDbl(rsTamp("SBLT"))) & "','" & RepRp(CDbl(rsTamp("PART"))) & "','" & RepRp(CDbl(rsTamp("OLI"))) & "','" & RepRp(CDbl(rsTamp("MTRL"))) & "', " _
                                       & "'" & RepRp(CDbl(metB)) & "'," & RepRp(CDbl(rsTamp("TOTAL"))) & ",'" & Trim(rsTamp("SLIP")) & "','" & Trim(rsTamp("FS_CATATAN")) & "','','','0','0','0','" & IIf(chCtkDisc, 0, 1) & "','" & RepRp(CDbl(rsTamp("BODY"))) & "'," & RepRp(CDbl(totalPPn)) & ") ")
32                    End If
33                    rsTamp.MoveNext
34                Loop
35            End If
36        End If
37        If rsTamp.State <> adStateClosed Then rsTamp.Close
          
          ' ADD JOB
           
38        If lblRetur.Visible Then
              'AND B.NOINV NOT LIKE '%-B'
39            Set rsTamp = objO.FetchData("SELECT '" & Trim(gInvNo) & "' AS NOINV, INVDATE, '" & Trim(gKodePlg) & "' AS FS_KD_PLG, '" & Trim(gInvCust) & "' AS CUSTOMER, '" & Trim(gInvAlmt) & "' AS CUST_ADRES, '" & Trim(gInvKota) & "' AS KOTA, '" & Trim(gInvNPWP) & "' AS NPWP, NO_POLISI, FMDL, A.NO_CHASSIS, NO_ENGINE, A.NO_WO, FRONT, STATUS, MECH, FS_LOG_ADD, " _
                                       & "SBLT, PART_G+PART_NG AS PART, OIL_G+OIL_NG AS OLI, MTRL, MET, BODY, SLIP, FS_CATATAN, TOTAL,fn_pajak, A.* FROM HIST_JOBADD A WITH(NOLOCK) INNER JOIN T_INVOICE B WITH(NOLOCK) ON A.NO_WO = B.NO_WO AND A.NO_INV = B.NOINV AND A.NO_CHASSIS = B.NO_CHASSIS AND A.FS_LOG_DEL = B.FS_LOG_DEL  AND A.FD_TGL_DEL = B.FD_RETUR WHERE A.NO_WO = '" & Trim(text(20)) & "' AND A.NO_CHASSIS = '" & Trim(text(17)) & "' AND B.NOINV LIKE '" & Trim(text(0)) & "%'")
40        Else
              'AND B.NOINV NOT LIKE '%-B'
41            Set rsTamp = objO.FetchData("SELECT '" & Trim(gInvNo) & "' AS NOINV, INVDATE, '" & Trim(gKodePlg) & "' AS FS_KD_PLG, '" & Trim(gInvCust) & "' AS CUSTOMER, '" & Trim(gInvAlmt) & "' AS CUST_ADRES, '" & Trim(gInvKota) & "' AS KOTA, '" & Trim(gInvNPWP) & "' AS NPWP, NO_POLISI, FMDL, A.NO_CHASSIS, NO_ENGINE, A.NO_WO, FRONT, STATUS, MECH, FS_LOG_ADD, " _
                                       & "SBLT, PART_G+PART_NG AS PART, OIL_G+OIL_NG AS OLI, MTRL, MET, BODY, SLIP, FS_CATATAN, TOTAL,fn_pajak, A.* FROM HIST_JOBADD A WITH(NOLOCK) INNER JOIN T_INVOICE B WITH(NOLOCK) ON A.NO_WO = B.NO_WO AND A.NO_INV = B.NOINV AND A.NO_CHASSIS = B.NO_CHASSIS AND A.FS_LOG_DEL = B.FS_LOG_DEL  WHERE A.NO_WO = '" & Trim(text(20)) & "' AND A.NO_CHASSIS = '" & Trim(text(17)) & "' AND A.FS_LOG_DEL = '' AND B.NOINV LIKE '" & Trim(text(0)) & "%'")
42        End If
43        If Not rsTamp.EOF Then
44            If createTemp(vNamaTabel) Then
45                Do While Not rsTamp.EOF
46                    jPPN = Round((nilaiDppInv / nilaiTotalInv) * CDbl(rsTamp("OFFN")))
      '                oconn.Execute ("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & Trim(rsTamp("NO_CHASSIS")) & "','" & Trim(rsTamp("NOINV")) & "','" & Format(rsTamp("INVDATE"), "MM/DD/YYYY") & "','" & Trim(rsTamp("FS_KD_PLG")) & "','" & Trim(rsTamp("CUSTOMER")) & "','" & Trim(rsTamp("CUST_ADRES")) & "', " _
      '                             & "'" & Trim(rsTamp("KOTA")) & "','" & Trim(rsTamp("NPWP")) & "','" & Trim(rsTamp("NO_POLISI")) & "','" & Trim(rsTamp("FMDL")) & "','" & Trim(rsTamp("NO_ENGINE")) & "','" & Trim(rsTamp("NO_WO")) & "','" & Trim(rsTamp("FRONT")) & "', " _
      '                             & "'" & Trim(rsTamp("STATUS")) & "','" & Trim(rsTamp("MECH")) & "','" & Trim(rsTamp("FS_LOG_ADD")) & "','" & CDbl(rsTamp("SBLT")) & "','" & CDbl(rsTamp("PART")) & "','" & CDbl(rsTamp("OLI")) & "','" & CDbl(rsTamp("MTRL")) & "', " _
      '                             & "'" & CDbl(rsTamp("MET")) & "','" & Trim(rsTamp("SLIP")) & "','" & Trim(rsTamp("FS_CATATAN")) & "','Add. " & Trim(rsTamp("REP")) & "','" & Trim(rsTamp("KLPKERJA")) & "','" & IIf(ChkCtkRate, CDbl(rsTamp("RTE")), "") & "','" & Round(CDbl(rsTamp("OFFN")) - IIf(sTipeCust = "AFFNONPPN", jPPN, 0)) & "','" & Replace(100 - IIf(Trim(rsTamp("KLPKERJA")) = "SBE" Or Trim(rsTamp("KLPKERJA")) = "SBI", IIf(Trim(rsTamp("FREEC")) = 1, 0, CDbl(rsTamp("DISC"))), 0), ",", ".") & "','" & IIf(chCtkDisc, 0, 1) & "','" & CDbl(rsTamp("BODY")) & "') ")
      '                total = total + Round((CDbl(rsTamp("OFFN")) - IIf(sTipeCust = "AFFNONPPN", jPPN, 0)) - ((CDbl(rsTamp("OFFN")) - IIf(sTipeCust = "AFFNONPPN", jPPN, 0)) * (1 - IIf(Trim(rsTamp("KLPKERJA")) = "SBE" Or Trim(rsTamp("KLPKERJA")) = "SBI", (IIf(Trim(rsTamp("FREEC")) = 1, 0, CDbl(rsTamp("DISC"))) / 100), 0))))
      '                Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES (0,0,'" & Trim(rsTamp("NO_CHASSIS")) & "','" & Trim(rsTamp("NOINV")) & "','" & Format(rsTamp("INVDATE"), "MM/DD/YYYY") & "','" & Trim(rsTamp("FS_KD_PLG")) & "','" & Trim(rsTamp("CUSTOMER")) & "','" & Trim(rsTamp("CUST_ADRES")) & "', " _
      '                             & "'" & Trim(rsTamp("KOTA")) & "','" & Trim(rsTamp("NPWP")) & "','" & Trim(rsTamp("NO_POLISI")) & "','" & Trim(rsTamp("FMDL")) & "','" & Trim(rsTamp("NO_ENGINE")) & "','" & Trim(rsTamp("NO_WO")) & "','" & Trim(rsTamp("FRONT")) & "', " _
      '                             & "'" & Trim(rsTamp("STATUS")) & "','" & Trim(rsTamp("MECH")) & "','" & Trim(rsTamp("FS_LOG_ADD")) & "','" & RepRp(CDbl(rsTamp("SBLT"))) & "','" & RepRp(CDbl(rsTamp("PART"))) & "','" & RepRp(CDbl(rsTamp("OLI"))) & "','" & RepRp(CDbl(rsTamp("MTRL"))) & "', " _
      '                             & "'" & RepRp(CDbl(metB)) & "'," & RepRp(CDbl(rsTamp("TOTAL"))) & ",'" & Trim(rsTamp("SLIP")) & "','" & Trim(rsTamp("FS_CATATAN")) & "','Add. " & Trim(rsTamp("REP")) & "','" & Trim(rsTamp("KLPKERJA")) & "','" & IIf(ChkCtkRate, CDbl(rsTamp("RTE")), "") & "','" & Round(CDbl(rsTamp("OFFN")) - IIf(sTipeCust = "AFFNONPPN", jPPN, 0)) & "','" & Replace(100 - IIf(Trim(rsTamp("FREEC")) = 1, 0, CDbl(rsTamp("DISC"))), ",", ".") & "','" & IIf(chCtkDisc, 0, 1) & "','" & RepRp(CDbl(rsTamp("BODY"))) & "'," & RepRp(CDbl(rsTamp("fn_pajak"))) & ") ")
                                   
47                   Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES (0,0,'" & Trim(rsTamp("NO_CHASSIS")) & "','" & Trim(rsTamp("NOINV")) & "','" & Format(rsTamp("INVDATE"), "MM/DD/YYYY") & "','" & Trim(rsTamp("FS_KD_PLG")) & "','" & Trim(rsTamp("CUSTOMER")) & "','" & Trim(rsTamp("CUST_ADRES")) & "', " _
                                   & "'" & Trim(rsTamp("KOTA")) & "','" & Trim(rsTamp("NPWP")) & "','" & Trim(rsTamp("NO_POLISI")) & "','" & Trim(rsTamp("FMDL")) & "','" & Trim(rsTamp("NO_ENGINE")) & "','" & Trim(rsTamp("NO_WO")) & "','" & Trim(rsTamp("FRONT")) & "', " _
                                   & "'" & Trim(rsTamp("STATUS")) & "','" & Trim(rsTamp("MECH")) & "','" & Trim(rsTamp("FS_LOG_ADD")) & "','" & RepRp(CDbl(rsTamp("SBLT"))) & "','" & RepRp(CDbl(rsTamp("PART"))) & "','" & RepRp(CDbl(rsTamp("OLI"))) & "','" & RepRp(CDbl(rsTamp("MTRL"))) & "', " _
                                   & "'" & RepRp(CDbl(metB)) & "'," & RepRp(CDbl(rsTamp("TOTAL"))) & ",'" & Trim(rsTamp("SLIP")) & "','" & Trim(rsTamp("FS_CATATAN")) & "','Add. " & Trim(rsTamp("REP")) & "','" & Trim(rsTamp("KLPKERJA")) & "','" & IIf(ChkCtkRate, CDbl(rsTamp("RTE")), "") & "','" & Round(CDbl(rsTamp("OFFN")) - IIf(sTipeCust = "AFFNONPPN", jPPN, 0)) & "','" & Replace(100 - IIf(Trim(rsTamp("FREEC")) = 1, 0, CDbl(rsTamp("DISC"))), ",", ".") & "','" & IIf(chCtkDisc, 0, 1) & "','" & RepRp(CDbl(rsTamp("BODY"))) & "'," & RepRp(CDbl(totalPPn)) & ") ")
48                    total = total + Round((CDbl(rsTamp("OFFN")) - IIf(sTipeCust = "AFFNONPPN", jPPN, 0)) - ((CDbl(rsTamp("OFFN")) - IIf(sTipeCust = "AFFNONPPN", jPPN, 0)) * (1 - (IIf(Trim(rsTamp("FREEC")) = 1, 0, CDbl(rsTamp("DISC"))) / 100))))
49                    rsTamp.MoveNext
50                Loop
51            End If
52        End If
53        If rsTamp.State <> adStateClosed Then rsTamp.Close
          
54        hitungINV2 = total

         'On Error GoTo 0
55       Exit Function

hitungINV2_Error:

56        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure hitungINV2 of Form frmBillingUlang")
End Function

Private Sub cetakInv2()
        'Dim rpt As New RptFormInvoiceN2_1      ' INV KGW   '23 AGUSTUS 2010 GANTI FORMAT BARU
          'Dim rpt As New RptFormInvoiceN2_2      ' inv mlati
          'Dim rpt As New RptFormInvoiceN2_3      ' inv ALL
          'Dim rpt As Object
          
          
1        On Error GoTo cetakInv2_Error

2           Select Case gPrintCetak
               Case "4"
3                 Set rpt = New RptFormInvoiceN2_4           ' inv blanko kosong
4                 rpt.txtDealer.SetText support.kopOrder(1)
5                 rpt.TxtAlamat.SetText support.kopOrder(2)
6                 rpt.txtTelp.SetText support.kopOrder(3)
7                 rpt.txtBank.SetText support.kopBank(0)
8              Case Else
               '            Set rpt = New RptFormInvoiceN2_3           ' inv ALL
9              If DateDiff("d", Format(Log.StartAffPpn, "mm/dd/yyyy"), Format(isnow, "mm/dd/yyyy")) >= 0 Then
10                 If Trim(gNpwp) = Trim(gInvNPWP2) Then
11                     Set rpt = New RptFormInvoiceN2_3_nonPPn
12                 Else
13                     Set rpt = New RptFormInvoiceN2_3           ' inv ALL
14                 End If
15             Else
16                 Set rpt = New RptFormInvoiceN2_3           ' inv ALL
17             End If
18          End Select
          Dim rsRepo As ADODB.Recordset
          Dim sTipeCust As String
          Dim vNamaTabel As String
          Dim gcat As String
          'Dim onlySublet As Boolean
          Dim total As Double
          
          'On Error GoTo gagalCetak
          
          ''2015-10-23 || ambil nama customer dari t_invoice
19        Set rsRepo = objO.FetchData("SELECT CUSTOMER,CUST_ADRES,KOTA,STATUS,FS_KD_PLG,NPWP FROM T_INVOICE WHERE NOINV = '" & text(0) & IIf(nDuoInv, "-B", "") & "'")
20        If Not rsRepo.EOF Then
21          rsRepo.MoveFirst
            
22          sTipeCust = Trim(rsRepo!Status) ''IIf(vTipeCustomer = "", text(40), vTipeCustomer)
23          gKodePlg = Trim(rsRepo!FS_KD_PLG) ''IIf(Trim(gKodePlg2) = "", text(2), gKodePlg2)
24          gInvNo = text(0) & IIf(nDuoInv, "-B", "") & "/" & IIf(vTipeCustomer = "", text(40), vTipeCustomer)
25          gInvTgl = vTglInv 'TglDDMMYYYY(isnow) '
26          gInvCust = Trim(rsRepo!customer) ''IIf(gInvCust2 = "", text(4), gInvCust2)
27          gInvAlmt = Trim(rsRepo!cust_adres) ''IIf(gInvAlmt2 = "", txt5, gInvAlmt2)
28          gInvKota = "" 'Trim(rsRepo!Kota) ''IIf(gInvKota2 = "", text(6) & "-" & text(7), gInvKota2) 'vNoTelp
29          gInvNPWP = IIf(gInvNPWP2 = "", "", gInvNPWP2) 'Text(14)
30        Else
31          sTipeCust = IIf(vTipeCustomer = "", text(40), vTipeCustomer)
32          gKodePlg = IIf(Trim(gKodePlg2) = "", text(2), gKodePlg2)
33          gInvNo = text(0) & IIf(nDuoInv, "-B", "") & "/" & IIf(vTipeCustomer = "", text(40), vTipeCustomer)
34          gInvTgl = vTglInv 'TglDDMMYYYY(isnow) '
35          gInvCust = IIf(gInvCust2 = "", text(4), gInvCust2)
36          gInvAlmt = IIf(gInvAlmt2 = "", txt5, gInvAlmt2)
37          gInvKota = "" 'IIf(gInvKota2 = "", text(6) & "-" & text(7), gInvKota2) 'vNoTelp
38          gInvNPWP = IIf(gInvNPWP2 = "", "", gInvNPWP2) 'Text(14)
39        End If
          
      '    sTipeCust = IIf(vTipeCustomer = "", text(40), vTipeCustomer)
      '    gKodePlg = IIf(Trim(gKodePlg2) = "", text(2), gKodePlg2)
      '    gInvNo = text(0) & IIf(nDuoInv, "-B", "") & "/" & IIf(vTipeCustomer = "", text(40), vTipeCustomer)
      '    gInvTgl = vTglInv 'TglDDMMYYYY(isnow) '
      '    gInvCust = IIf(gInvCust2 = "", text(4), gInvCust2)
      '    gInvAlmt = IIf(gInvAlmt2 = "", txt5, gInvAlmt2)
      '    gInvKota = IIf(gInvKota2 = "", text(6) & "-" & text(7), gInvKota2) 'vNoTelp
      '    gInvNPWP = IIf(gInvNPWP2 = "", "", gInvNPWP2) 'Text(14)
40        vNamaTabel = "INVOICE2-" & Trim(text(0))
          
41        total = hitungINV2(sTipeCust, vNamaTabel)
        
42        Set rsRepo = objO.FetchData("SELECT DISTINCT * FROM [TEMPDB]..[" & vNamaTabel & "] ORDER BY KLPKERJA DESC, REP DESC, [OFF]")
43        rpt.database.SetDataSource rsRepo
          
44        If nRC Then
        Dim rsHlp As New ADODB.Recordset
        Dim nBulan As String
        Dim nSbe As String
45      nSbe = ""
         
46      Set rsHlp = objO.FetchData("SELECT CASE WHEN JT_TGL < JT_BLN THEN JT_TGL ELSE JT_BLN END AS JT_TGL, SBENEXT FROM M_ESTIMASI WITH(NOLOCK) WHERE NORANGKA = '" & Trim(text(17)) & "' AND TGLSVCNOW = '" & Format(text(43), "MM/DD/YYYY") & "'")
47      If Not rsHlp.EOF Then
48          nBulan = GetNamaBulan(Month(rsHlp("JT_TGL"))) & " " & year(rsHlp("JT_TGL"))
49          nSbe = Trim(rsHlp("SBENEXT"))
50      End If
51      rsHlp.Close
        
52      If Trim(nSbe) <> "" Then
            'gcat = IIf(nHeavy, "Kembali untuk service berkala setelah 5.000 KM atau bulan " & nBulan, "Kembali untuk service berkala setelah 10.000 KM atau bulan " & nBulan)
53          gcat = "Kembali untuk service berkala " & Format(CDbl(nSbe), "#,#0") & " KM atau bulan " & nBulan
54          gcat = gcat + IIf(gcat <> "", " (mana yang tercapai lebih dahulu).", "")
55      End If
56        End If
          Dim rsTerbilang As ADODB.Recordset
          Dim totalTerbilang As Double
          
          'Set rsTerbilang = objO.FetchData("SELECT NOINV, ROUND(SBLT+PART+OLI+SUM([OFF]*((100-DISC)/100)),0) AS REV from [TEMPDB]..[" & vNamaTabel & "] GROUP  BY NOINV, SBLT, PART, OLI")
          'Set rsTerbilang = objO.FetchData("SELECT total, NOINV, ROUND(SBLT+PART+OLI+MET+SUM([OFF]*((100-DISC)/100)),0) AS REV from [TEMPDB]..[" & vNamaTabel & "] GROUP  BY NOINV, SBLT, PART, OLI, MET,total")
          
          '------------------------------------------------
          '2015-07-30 | GANTI QUERY PENGAMBILAN TERBILANG -
          '------------------------------------------------
          
57        Set rsTerbilang = objO.FetchData("SELECT distinct met,fn_pajak,total,A.noinv, a.rev+SBLT+round(PART+OLI,0)+MET AS REV FROM (SELECT NOINV, ROUND(SUM([OFF]*((100-DISC)/100)),0) AS REV from [TEMPDB]..[" & vNamaTabel & "] GROUP  BY NOINV) A INNER JOIN [TEMPDB]..[" & vNamaTabel & "] B ON A.NOINV = B.NOINV WHERE SBLT+PART+OLI = 0 and DISC = 0")
58        totalTerbilang = 0
59        If Not rsTerbilang.EOF Then
60      Do While Not rsTerbilang.EOF
61          totalTerbilang = totalTerbilang + rsTerbilang("REV")
62          rsTerbilang.MoveNext
63      Loop
64      rsTerbilang.MoveFirst
65      If totalTerbilang <> CDbl(rsTerbilang("total")) Then
           ' totalTerbilang = CDbl(rsTerbilang("total"))
66      End If
67        End If
          
         ' If DateDiff("d", Format(Log.StartAffPpn, "mm/dd/yyyy"), Format(isnow, "mm/dd/yyyy")) >= 0 Then
      '  If Trim(gNpwp) = Trim(gInvNPWP2) Then
            ''hilangkan ppn afiliasi
            ''2016-03-23 || revisi pajak menggunakan fn_pajak
            'totalTerbilang = totalTerbilang - CDbl(rsTerbilang!MET)
          '  totalTerbilang = Round((totalTerbilang * nilaiDppInv / nilaiTotalInv,0)  ''10 / 11), 0) ''totalTerbilang * nilaiDppInv / nilaiTotalInv,0)  ''10 / 11
            
      '  End If
       '   End If
68        rpt.txtTerbilang.SetText Trim(Angka2Huruf(totalTerbilang, 1))
          
          'rpt.txtTerbilang.SetText Trim(Angka2Huruf(Round(total, 0), 1))
69        rpt.txtManager.SetText Trim(gSecmanager)
70        If ChCopy.value = 1 Then
71      rpt.txtKet.SetText "COPY INVOICE, " & Trim(gcat)
72        Else
73      rpt.txtKet.SetText Trim(gcat)
74        End If
          ' ==============================================================================================
          ' 08 Maret 2017 | Thania | penambahan informasi PWC utk cetakkan invoice & Cetak Ulang
          ' ==============================================================================================
              'rpt.txtRetur.SetText IIf(lblRetur.Visible, "DIBATALKAN", "")
75            rpt.txtRetur.SetText IIf(lblRetur.Visible, "DIBATALKAN", "") & IIf(lblPWC.Visible, " PWC", "")
          ' ==============================================================================================
          
76        rpt.txtJT.SetText "Jatuh Tempo : " & JTBayar(gKodePlg, text(1))
77        rpt.txtInfoMaterai.SetText ""
78        rpt.txtLunasMaterai.SetText ""
79        If Log.fakturMaterai = 1 Then
80            Set rsTemp = objO.FetchData("SELECT * FROM t_invoiceMaterai WITH (NOLOCK) WHERE NOINV = '" & text(0) & IIf(nDuoInv, "-B", "") & "'")
81            If Not rsTemp.EOF Then
82                rpt.txtInfoMaterai.SetText "Ijin Pembubuhan Materai No. " & NulSID(rsTemp!NOIJIN)
83                rpt.txtLunasMaterai.SetText "BEA MATERAI LUNAS " & NulSID(rsTemp!nilai)
84            End If
85        End If
          
86        With CRViewer2
87            .ReportSource = rpt
88            .ViewReport
89            .Zoom 100
90            .Top = 0
91            .Left = asxPowerBanner1.Left + asxPowerBanner1.Width + 20
92            .Width = Frame2.Width + 100
93            .Height = asxPowerBanner1.Height
              
94            cmdTutup2.Top = .Top + 430
95            cmdTutup2.Left = .Left
96            cmdTutup2.Visible = True
97            .Visible = True
98        End With
99        cmdTutup2.Visible = True
100       Screen.MousePointer = vbDefault
101       Exit Sub

gagalCetak:
102       MsgBox "Invoice kedua, tidak bisa dicetak ulang !"
103       Exit Sub

         'On Error GoTo 0
104      Exit Sub

cetakInv2_Error:

105       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cetakInv2 of Form frmBillingUlang")
End Sub

Private Function hitungINV1(vNamaTabel As String) As Double
          Dim total As Double
          Dim rsTamp As ADODB.Recordset

1        On Error GoTo hitungINV1_Error

2         If Get_Connectiontemp("tempdb", gUser, gPwd, gServerDb) Then
3             sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ) " _
                   & " drop table [tempdb]..[" & vNamaTabel & "]"
4             Call objO.UpdateData(sSQL)
5         End If
          
          ''' ASLI NICH !!!
      '    sSQL = "SELECT '" & Trim(gInvNo) & "' AS NOINV, INVDATE, '" & Trim(gKodePlg) & "' AS FS_KD_PLG, '" & Trim(gInvCust) & "' AS CUSTOMER, '" & Trim(gInvAlmt) & "' AS CUST_ADRES, '" & Trim(gInvKota) & "' AS KOTA, '" & Trim(gInvNPWP) & "' AS NPWP, NO_POLISI, FMDL, NO_CHASSIS, NO_ENGINE, NO_WO, FRONT, STATUS, MECH, FS_LOG_ADD, " _
      '         & "SBLT, PART_G+PART_NG AS PART, OIL_G+OIL_NG AS OLI, MTRL, MET, BODY, SLIP, FS_CATATAN, TOTAL, " _
      '         & "REP1, REP2, REP3, REP4, REP5, REP6, REP7, REP8, RTE1, RTE2, RTE3, RTE4, RTE5, RTE6, RTE7, RTE8, " _
      '         & "KLPKERJA1 , KLPKERJA2, KLPKERJA3, KLPKERJA4, KLPKERJA5, KLPKERJA6, KLPKERJA7, KLPKERJA8, " _
      '         & "Off1 , Off2, Off3, Off4, Off5, Off6, Off7, Off8, DISC1, DISC2, DISC3, DISC4, DISC5, DISC6, DISC7, DISC8 " _
      '         & "FROM T_INVOICE WHERE NOINV = '" & Trim(text(0)) & "' AND NO_WO = '" & Trim(text(20)) & "'"
          '#########################################################'
          '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..'
          '#########################################################'
6         sSQL = "SELECT '" & Trim(gInvNo) & "' AS NOINV, INVDATE, '" & Trim(gKodePlg) & "' AS FS_KD_PLG, '" & Trim(gInvCust) & "' AS CUSTOMER, '" & Trim(gInvAlmt) & "' AS CUST_ADRES, '" & Trim(gInvKota) & "' AS KOTA, '" & Trim(gInvNPWP) & "' AS NPWP, NO_POLISI, FMDL, NO_CHASSIS, NO_ENGINE, NO_WO, FRONT, STATUS, MECH, FS_LOG_ADD, " _
               & "SBLT, PART_G+PART_NG AS PART, OIL_G+OIL_NG AS OLI, MTRL, MET, BODY, SLIP, FS_CATATAN, TOTAL, " _
               & "REP1, REP2, REP3, REP4, REP5, REP6, REP7, REP8, RTE1, RTE2, RTE3, RTE4, RTE5, RTE6, RTE7, RTE8, " _
               & "KLPKERJA1 , KLPKERJA2, KLPKERJA3, KLPKERJA4, KLPKERJA5, KLPKERJA6, KLPKERJA7, KLPKERJA8, " _
               & "Off1 , Off2, Off3, Off4, Off5, Off6, Off7, Off8, DISC1, DISC2, DISC3, DISC4, DISC5, DISC6, DISC7, DISC8,fn_pajak,ppn  " _
               & "FROM T_INVOICE WITH(NOLOCK) WHERE NOINV = '" & Trim(text(0)) & IIf(nDuoInv, "-A", "") & "' AND NO_WO = '" & Trim(text(20)) & "' ORDER BY NOINV"
            
           '' 2016-05-18 || total dan fn_pajak diambil dari invoice a
           Dim totalA As Double
           Dim ppnA As Double
           
7         Set rsTamp = objO.FetchData(sSQL)
8         If Not rsTamp.EOF Then
9             If createTemp(vNamaTabel) Then
10                onlySublet = True
11                total = Round(total + CDbl(rsTamp("PART")) + CDbl(rsTamp("OLI")) + CDbl(rsTamp("MTRL")) + CDbl(rsTamp("SBLT")) + IIf(nDuoInv, 0, CDbl(rsTamp("MET"))) + CDbl(rsTamp("BODY")))
                  
12                totalA = CDbl(rsTamp!total)
13                ppnA = CDbl(rsTamp!fn_pajak)
                  
14                nilaiPPnInv = NulSID(rsTamp("PPn"))
15                nilaiDppInv = 100
16                nilaiTotalInv = nilaiDppInv + nilaiPPnInv
                  
17                Do While Not rsTamp.EOF
18                    For i = 1 To 8
19                        hlp1 = "REP" & i
20                        hlp2 = "RTE" & i
21                        hlp3 = "OFF" & i
22                        hlp4 = "DISC" & i
23                        hlp5 = "FREEC" & i
24                        hlp6 = "KLPKERJA" & i
                          
25                        If Trim(rsTamp(hlp1)) <> "" Then
26                            jPPN = Round((nilaiDppInv / nilaiTotalInv) * CDbl(rsTamp(hlp3)))
                              ''' ASLI NICH !!!
      '                        oConn.Execute ("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & Trim(rsTamp("NO_CHASSIS")) & "','" & Trim(rsTamp("NOINV")) & "','" & Format(rsTamp("INVDATE"), "MM/DD/YYYY") & "','" & Trim(rsTamp("FS_KD_PLG")) & "','" & Trim(rsTamp("CUSTOMER")) & "','" & Trim(rsTamp("CUST_ADRES")) & "', " _
      '                                     & "'" & Trim(rsTamp("KOTA")) & "','" & Trim(rsTamp("NPWP")) & "','" & Trim(rsTamp("NO_POLISI")) & "','" & Trim(rsTamp("FMDL")) & "','" & Trim(rsTamp("NO_ENGINE")) & "','" & Trim(rsTamp("NO_WO")) & "','" & Trim(rsTamp("FRONT")) & "', " _
      '                                     & "'" & Trim(rsTamp("STATUS")) & "','" & Trim(rsTamp("MECH")) & "','" & Trim(rsTamp("FS_LOG_ADD")) & "','" & CDbl(rsTamp("SBLT")) & "','" & CDbl(rsTamp("PART")) & "','" & CDbl(rsTamp("OLI")) & "','" & CDbl(rsTamp("MTRL")) & "', " _
      '                                     & "'" & IIf(nDuoInv, 0, CDbl(rsTamp("MET"))) & "','" & Trim(rsTamp("SLIP")) & "','" & Trim(rsTamp("FS_CATATAN")) & "','" & Trim(rsTamp(hlp1)) & "','" & Trim(rsTamp(hlp6)) & "','" & IIf(ChkCtkRate, CDbl(rsTamp(hlp2)), "") & "','" & Round(CDbl(rsTamp(hlp3)) - IIf(Trim(text(40)) = "AFFNONPPN", jPPN, 0)) & "','" & Replace(CDbl(rsTamp(hlp4)), ",", ".") & "','" & IIf(chCtkDisc, 0, 1) & "','" & CDbl(rsTamp("BODY")) & "') ")
      '                        total = total + Round((CDbl(rsTamp(hlp3)) - IIf(Trim(text(40)) = "AFFNONPPN", jPPN, 0)) - ((CDbl(rsTamp(hlp3)) - IIf(Trim(text(40)) = "AFFNONPPN", jPPN, 0)) * (CDbl(rsTamp(hlp4)) / 100)))
                              '#########################################################'
                              '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..'
                              '#########################################################'
      '                        Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES (round((" & Round(CDbl(rsTamp(hlp3)) - IIf(Trim(Text(40)) = "AFFNONPPN", jPPN, 0)) & " * (1-(" & RepRp(CDbl(rsTamp(hlp4))) & " * 0.01))),0,1),((" & Round(CDbl(rsTamp(hlp3)) - IIf(Trim(Text(40)) = "AFFNONPPN", jPPN, 0)) & " * (1-(" & RepRp(CDbl(rsTamp(hlp4))) & " * 0.01)))),'" & Trim(rsTamp("NO_CHASSIS")) & "','" & Trim(rsTamp("NOINV")) & "','" & Format(rsTamp("INVDATE"), "MM/DD/YYYY") & "','" & Trim(rsTamp("FS_KD_PLG")) & "','" & Trim(rsTamp("CUSTOMER")) & "','" & Trim(rsTamp("CUST_ADRES")) & "', " _
      '                                     & "'" & Trim(rsTamp("KOTA")) & "','" & Trim(rsTamp("NPWP")) & "','" & Trim(rsTamp("NO_POLISI")) & "','" & Trim(rsTamp("FMDL")) & "','" & Trim(rsTamp("NO_ENGINE")) & "','" & Trim(rsTamp("NO_WO")) & "','" & Trim(rsTamp("FRONT")) & "', " _
      '                                     & "'" & Trim(rsTamp("STATUS")) & "','" & Trim(rsTamp("MECH")) & "','" & Trim(rsTamp("FS_LOG_ADD")) & "'," & RepRp(CDbl(rsTamp("SBLT"))) & "," & RepRp(CDbl(rsTamp("PART"))) & "," & RepRp(CDbl(rsTamp("OLI"))) & "," & RepRp(CDbl(rsTamp("MTRL"))) & ", " _
      '                                     & " " & IIf(nDuoInv, 0, RepRp(CDbl(rsTamp("MET")))) & "," & RepRp(CDbl(rsTamp("TOTAL"))) & ",'" & Trim(rsTamp("SLIP")) & "','" & Trim(rsTamp("FS_CATATAN")) & "','" & Trim(rsTamp(hlp1)) & "','" & Trim(rsTamp(hlp6)) & "','" & IIf(ChkCtkRate, CDbl(rsTamp(hlp2)), "") & "'," & Round(CDbl(rsTamp(hlp3)) - IIf(Trim(Text(40)) = "AFFNONPPN", jPPN, 0)) & "," & RepRp(CDbl(rsTamp(hlp4))) & ",'" & IIf(chCtkDisc, 0, 1) & "'," & CDbl(rsTamp("BODY")) & "," & RepRp(CDbl(rsTamp("fn_pajak"))) & ") ")
                                           
27                            Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES (round((" & Round(CDbl(rsTamp(hlp3)) - IIf(Trim(text(40)) = "AFFNONPPN", jPPN, 0)) & " * (1-(" & RepRp(CDbl(rsTamp(hlp4))) & " * 0.01))),0,1),((" & Round(CDbl(rsTamp(hlp3)) - IIf(Trim(text(40)) = "AFFNONPPN", jPPN, 0)) & " * (1-(" & RepRp(CDbl(rsTamp(hlp4))) & " * 0.01)))),'" & Trim(rsTamp("NO_CHASSIS")) & "','" & Trim(rsTamp("NOINV")) & "','" & Format(rsTamp("INVDATE"), "MM/DD/YYYY") & "','" & Trim(rsTamp("FS_KD_PLG")) & "','" & Trim(rsTamp("CUSTOMER")) & "','" & Trim(rsTamp("CUST_ADRES")) & "', " _
                                           & "'" & Trim(rsTamp("KOTA")) & "','" & Trim(rsTamp("NPWP")) & "','" & Trim(rsTamp("NO_POLISI")) & "','" & Trim(rsTamp("FMDL")) & "','" & Trim(rsTamp("NO_ENGINE")) & "','" & Trim(rsTamp("NO_WO")) & "','" & Trim(rsTamp("FRONT")) & "', " _
                                           & "'" & Trim(rsTamp("STATUS")) & "','" & Trim(rsTamp("MECH")) & "','" & Trim(rsTamp("FS_LOG_ADD")) & "'," & RepRp(CDbl(rsTamp("SBLT"))) & "," & RepRp(CDbl(rsTamp("PART"))) & "," & RepRp(CDbl(rsTamp("OLI"))) & "," & RepRp(CDbl(rsTamp("MTRL"))) & ", " _
                                           & " " & IIf(nDuoInv, 0, RepRp(CDbl(rsTamp("MET")))) & "," & RepRp(CDbl(rsTamp("TOTAL"))) & ",'" & Trim(rsTamp("SLIP")) & "','" & Trim(rsTamp("FS_CATATAN")) & "','" & Trim(rsTamp(hlp1)) & "','" & Trim(rsTamp(hlp6)) & "','" & IIf(ChkCtkRate, CDbl(rsTamp(hlp2)), "") & "'," & Round(CDbl(rsTamp(hlp3)) - IIf(Trim(text(40)) = "AFFNONPPN", jPPN, 0)) & "," & RepRp(CDbl(rsTamp(hlp4))) & ",'" & IIf(chCtkDisc, 0, 1) & "'," & CDbl(rsTamp("BODY")) & "," & RepRp(CDbl(rsTamp("fn_pajak"))) & ") ")
                                           
28                            total = Round(total + Round((CDbl(rsTamp(hlp3)) - IIf(Trim(text(40)) = "AFFNONPPN", jPPN, 0)) - ((CDbl(rsTamp(hlp3)) - IIf(Trim(text(40)) = "AFFNONPPN", jPPN, 0)) * (CDbl(rsTamp(hlp4)) / 100))), 0)
29                            onlySublet = False
30                        End If
31                    Next
                      
                      
32                    If onlySublet Then
33                        Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES (0,0,'" & Trim(rsTamp("NO_CHASSIS")) & "','" & Trim(rsTamp("NOINV")) & "','" & Format(rsTamp("INVDATE"), "MM/DD/YYYY") & "','" & Trim(rsTamp("FS_KD_PLG")) & "','" & Trim(rsTamp("CUSTOMER")) & "','" & Trim(rsTamp("CUST_ADRES")) & "', " _
                                       & "'" & Trim(rsTamp("KOTA")) & "','" & Trim(rsTamp("NPWP")) & "','" & Trim(rsTamp("NO_POLISI")) & "','" & Trim(rsTamp("FMDL")) & "','" & Trim(rsTamp("NO_ENGINE")) & "','" & Trim(rsTamp("NO_WO")) & "','" & Trim(rsTamp("FRONT")) & "', " _
                                       & "'" & Trim(rsTamp("STATUS")) & "','" & Trim(rsTamp("MECH")) & "','" & Trim(rsTamp("FS_LOG_ADD")) & "'," & RepRp(CDbl(rsTamp("SBLT"))) & "," & RepRp(CDbl(rsTamp("PART"))) & "," & RepRp(CDbl(rsTamp("OLI"))) & "," & RepRp(CDbl(rsTamp("MTRL"))) & ", " _
                                       & " " & IIf(nDuoInv, 0, RepRp(CDbl(rsTamp("MET")))) & "," & RepRp(CDbl(rsTamp("TOTAL"))) & ",'" & Trim(rsTamp("SLIP")) & "','" & Trim(rsTamp("FS_CATATAN")) & "','','','0',0,0,'" & IIf(chCtkDisc, 0, 1) & "'," & CDbl(rsTamp("BODY")) & "," & RepRp(CDbl(rsTamp("fn_pajak"))) & ") ")
34                    End If
35                    rsTamp.MoveNext
36                Loop
37            End If
38        End If
39        If rsTamp.State <> adStateClosed Then rsTamp.Close
          
          ' ADD JOB
           
40        If lblRetur.Visible Then
              'AND B.NOINV NOT LIKE '%-B'
41            Set rsTamp = objO.FetchData("SELECT '" & Trim(gInvNo) & "' AS NOINV, INVDATE, '" & Trim(gKodePlg) & "' AS FS_KD_PLG, '" & Trim(gInvCust) & "' AS CUSTOMER, '" & Trim(gInvAlmt) & "' AS CUST_ADRES, '" & Trim(gInvKota) & "' AS KOTA, '" & Trim(gInvNPWP) & "' AS NPWP, NO_POLISI, FMDL, A.NO_CHASSIS, NO_ENGINE, A.NO_WO, FRONT, STATUS, MECH, FS_LOG_ADD, " _
                                       & "SBLT, PART_G+PART_NG AS PART, OIL_G+OIL_NG AS OLI, MTRL, MET, BODY, SLIP, FS_CATATAN,fn_pajak, TOTAL, A.* FROM HIST_JOBADD A WITH(NOLOCK) INNER JOIN T_INVOICE B WITH(NOLOCK) ON A.NO_WO = B.NO_WO AND A.NO_INV = B.NOINV AND A.NO_CHASSIS = B.NO_CHASSIS AND A.FS_LOG_DEL = B.FS_LOG_DEL  AND A.FD_TGL_DEL = B.FD_RETUR WHERE A.NO_WO = '" & Trim(text(20)) & "' AND A.NO_CHASSIS = '" & Trim(text(17)) & "' AND B.NOINV LIKE '" & Trim(text(0)) & "%'")
42        Else
              'AND B.NOINV NOT LIKE '%-B'
43            Set rsTamp = objO.FetchData("SELECT '" & Trim(gInvNo) & "' AS NOINV, INVDATE, '" & Trim(gKodePlg) & "' AS FS_KD_PLG, '" & Trim(gInvCust) & "' AS CUSTOMER, '" & Trim(gInvAlmt) & "' AS CUST_ADRES, '" & Trim(gInvKota) & "' AS KOTA, '" & Trim(gInvNPWP) & "' AS NPWP, NO_POLISI, FMDL, A.NO_CHASSIS, NO_ENGINE, A.NO_WO, FRONT, STATUS, MECH, FS_LOG_ADD, " _
                                       & "SBLT, PART_G+PART_NG AS PART, OIL_G+OIL_NG AS OLI, MTRL, MET, BODY, SLIP, FS_CATATAN, TOTAL ,fn_pajak , A.* FROM HIST_JOBADD A WITH(NOLOCK) INNER JOIN T_INVOICE B WITH(NOLOCK) ON A.NO_WO = B.NO_WO AND A.NO_INV = B.NOINV AND A.NO_CHASSIS = B.NO_CHASSIS AND A.FS_LOG_DEL = B.FS_LOG_DEL  WHERE A.NO_WO = '" & Trim(text(20)) & "' AND A.NO_CHASSIS = '" & Trim(text(17)) & "' AND A.FS_LOG_DEL = '' AND B.NOINV LIKE '" & Trim(text(0)) & "%'")
44        End If
45        If Not rsTamp.EOF Then
46            If createTemp(vNamaTabel) Then
47                Do While Not rsTamp.EOF
48                    jPPN = Round((nilaiDppInv / nilaiTotalInv) * CDbl(rsTamp("OFFN")))
                      
49                    Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES (round((" & Round(CDbl(rsTamp("OFFN")) - IIf(Trim(text(40)) = "AFFNONPPN", jPPN, 0)) & " * (1-(" & Replace(CDbl(rsTamp("DISC")), ",", ".") & " * 0.01))),0,1),((" & Round(CDbl(rsTamp("OFFN")) - IIf(Trim(text(40)) = "AFFNONPPN", jPPN, 0)) & " * (1-(" & Replace(CDbl(rsTamp("DISC")), ",", ".") & " * 0.01)))),'" & Trim(rsTamp("NO_CHASSIS")) & "','" & Trim(rsTamp("NOINV")) & "','" & Format(rsTamp("INVDATE"), "MM/DD/YYYY") & "','" & Trim(rsTamp("FS_KD_PLG")) & "','" & Trim(rsTamp("CUSTOMER")) & "','" & Trim(rsTamp("CUST_ADRES")) & "', " _
                                   & "'" & Trim(rsTamp("KOTA")) & "','" & Trim(rsTamp("NPWP")) & "','" & Trim(rsTamp("NO_POLISI")) & "','" & Trim(rsTamp("FMDL")) & "','" & Trim(rsTamp("NO_ENGINE")) & "','" & Trim(rsTamp("NO_WO")) & "','" & Trim(rsTamp("FRONT")) & "', " _
                                   & "'" & Trim(rsTamp("STATUS")) & "','" & Trim(rsTamp("MECH")) & "','" & Trim(rsTamp("FS_LOG_ADD")) & "', " & RepRp(CDbl(rsTamp("SBLT"))) & " , " & RepRp(CDbl(rsTamp("PART"))) & " , " & RepRp(CDbl(rsTamp("OLI"))) & " , " & RepRp(CDbl(rsTamp("MTRL"))) & " , " _
                                   & " " & RepRp(IIf(nDuoInv, 0, CDbl(rsTamp("MET")))) & "," & RepRp(CDbl(rsTamp("TOTAL"))) & ",'" & Trim(rsTamp("SLIP")) & "','" & Trim(rsTamp("FS_CATATAN")) & "','Add. " & Trim(rsTamp("REP")) & "','" & Trim(rsTamp("KLPKERJA")) & "','" & IIf(ChkCtkRate, CDbl(rsTamp("RTE")), "") & "'," & Round(CDbl(rsTamp("OFFN")) - IIf(Trim(text(40)) = "AFFNONPPN", jPPN, 0)) & "," & Replace(CDbl(rsTamp("DISC")), ",", ".") & ",'" & IIf(chCtkDisc, 0, 1) & "'," & CDbl(rsTamp("BODY")) & "," & RepRp(CDbl(rsTamp("fn_pajak"))) & ") ")
50                    total = total + Round((CDbl(rsTamp("OFFN")) - IIf(Trim(text(40)) = "AFFNONPPN", jPPN, 0)) - ((CDbl(rsTamp("OFFN")) - IIf(Trim(text(40)) = "AFFNONPPN", jPPN, 0)) * (CDbl(rsTamp("DISC")) / 100)))
51                    rsTamp.MoveNext
52                Loop
53            End If
54        End If
55        If rsTamp.State <> adStateClosed Then rsTamp.Close
          
56        hitungINV1 = total

         'On Error GoTo 0
57       Exit Function

hitungINV1_Error:

58        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure hitungINV1 of Form frmBillingUlang")
End Function

Private Sub cetakInv1()
          'Dim rpt As New RptFormInvoiceN1_1      'INV KGW       '23 AGUSTUS 2010 GANTI FORMAT BARU
          'Dim rpt As New RptFormInvoiceN1_2      'inv mlati
          'Dim rpt As New RptFormInvoiceN1_3      'inv ALL
          'Dim rpt As Object
1        On Error GoTo cetakInv1_Error

2         Select Case gPrintCetak
              Case "4"
3                 Set rpt = New RptFormInvoiceN1_4           ' inv blanko kosong
4                 rpt.txtDealer.SetText support.kopOrder(1)
5                 rpt.TxtAlamat.SetText support.kopOrder(2)
6                 rpt.txtTelp.SetText support.kopOrder(3)
7                 rpt.txtBank.SetText support.kopBank(0)
8             Case Else
9                 Set rpt = New RptFormInvoiceN1_3           ' inv ALL
10        End Select
          
          Dim rsRepo As New ADODB.Recordset
          Dim vNamaTabel As String
          Dim gcat As String
          Dim infoBill As String
          Dim total As Double
          
          ''2017-10-26
          ''--------------------------------------------------------------------
11        If Log.loyalityCustomer = 1 Then
12            Set rsTemp = objO.FetchData("SELECT * FROM RPLNASAPP..MASTERUSER WHERE isnull(idUser,'') = '" & Trim(text(2)) & "@apps.nasmoco" & "'")
13            If Not rsTemp.EOF Then
14                rpt.txtLoyalityUser.SetText Trim(rsTemp!idUser)
15                rpt.txtKeteranganLoyality.SetText Trim(Log.ketNasApp)
16                rpt.txtLoyalityPassword.SetText Trim(rsTemp!Password)
17            End If
18        End If
          
          ''********************************************************************
          '' 2017-06-16 || PENANDA UNTUK UNIT NCM DITAMBAHKAN KETERANGAN NOTIF
          ''====================================================================
19        If Log.notifInvNCM = 1 Then
20            rpt.boxNCM.Suppress = True
21            rpt.txtNCM1.Suppress = True
22            rpt.txtNCM2.Suppress = True
      '        rpt.txtNCM3.Suppress = True
      '        rpt.txtNCM4.Suppress = True
      '        rpt.txtNCM5.Suppress = True
23            If isJualNRM = True Then
24                Set rsTemp = objO.FetchData("SELECT * FROM notifInvoice WITH (NOLOCK) WHERE tipe = 'NCM'")
25                If Not rsTemp.EOF Then
                  
26                    rpt.boxNCM.Suppress = False
27                    rpt.txtNCM1.Suppress = False
28                    rpt.txtNCM1.SetText Trim(rsTemp!Text1)
                      
29                    rpt.txtNCM2.Suppress = False
30                    rpt.txtNCM2.SetText Trim(rsTemp!Text2)
                      
      '                rpt.txtNCM3.Suppress = False
      '                rpt.txtNCM3.SetText Trim(rsTemp!Text3)
      '
      '                rpt.txtNCM4.Suppress = False
      '                rpt.txtNCM4.SetText Trim(rsTemp!Text4)
      '
      '                rpt.txtNCM5.Suppress = False
      '                rpt.txtNCM5.SetText Trim(rsTemp!Text5)
                   
31                End If
32            Else
33                Set rsTemp = objO.FetchData("SELECT * FROM notifInvoice WITH (NOLOCK) WHERE tipe = 'NonNCM'")
34                If Not rsTemp.EOF Then
                  
35                    rpt.boxNCM.Suppress = False
36                    rpt.txtNCM1.Suppress = False
37                    rpt.txtNCM1.SetText Trim(rsTemp!Text1)
                      
38                    rpt.txtNCM2.Suppress = False
39                    rpt.txtNCM2.SetText Trim(rsTemp!Text2)
40                End If
41            End If
42        End If
          
      '    If Log.notifInvNonNCM = 1 Then
      '        If isJualNRM = False Then
      '            Set rsTemp = objO.FetchData("SELECT * FROM notifInvoice WITH (NOLOCK) WHERE tipe = 'NonNCM'")
      '            If Not rsTemp.EOF Then
      '
      '                rpt.boxNCM.Suppress = False
      '                rpt.txtNCM1.Suppress = False
      '                rpt.txtNCM1.SetText Trim(rsTemp!Text1)
      '
      '                rpt.txtNCM2.Suppress = False
      '                rpt.txtNCM2.SetText Trim(rsTemp!Text2)
      '
      '                rpt.txtNCM3.Suppress = False
      '                rpt.txtNCM3.SetText Trim(rsTemp!Text3)
      '
      '                rpt.txtNCM4.Suppress = False
      '                rpt.txtNCM4.SetText Trim(rsTemp!Text4)
      '
      '                rpt.txtNCM5.Suppress = False
      '                rpt.txtNCM5.SetText Trim(rsTemp!Text5)
      '
      '            End If
      '        Else
      '            rpt.boxNCM.Suppress = True
      '            rpt.txtNCM1.Suppress = True
      '            rpt.txtNCM2.Suppress = True
      '            rpt.txtNCM3.Suppress = True
      '            rpt.txtNCM4.Suppress = True
      '            rpt.txtNCM5.Suppress = True
      '        End If
      '    End If
          
43        On Error GoTo gagalCetak
44        gKodePlg = IIf(Trim(gKodePlg) = "", text(2), gKodePlg)
45        gInvNo = text(0) & IIf(nDuoInv, "-A", "") & "/" & text(40)
46        gInvTgl = TglDDMMYYYY(isnow)
47        gInvCust = IIf(gInvCust = "", text(4), gInvCust)
48        gInvAlmt = IIf(gInvAlmt = "", txt5, gInvAlmt)
49        gInvKota = "" 'IIf(gInvKota = "", text(6) & "-" & text(7), gInvKota) 'vNoTelp
50        gInvNPWP = IIf(gInvNPWP = "", text(44), gInvNPWP) 'Text(14)
51        vNamaTabel = "INVOICE" & Trim(text(0))
52        total = hitungINV1(vNamaTabel)
          'Set rsRepo = objO.FetchData("SELECT DISTINCT * FROM [TEMPDB]..[" & vNamaTabel & "] ORDER BY KLPKERJA DESC, REP DESC, [OFF]")
53        Set rsRepo = objO.FetchData("SELECT DISTINCT *,(SBLT+PART+OLI+MTRL+MET) as nonLabor FROM [TEMPDB]..[" & vNamaTabel & "] ORDER BY disc desc , (SBLT+PART+OLI+MTRL+MET) asc, KLPKERJA DESC, [OFF]") ''objO.FetchData("SELECT DISTINCT * FROM [TEMPDB]..[" & vNamaTabel & "] ORDER BY KLPKERJA , REP , [OFF]")
54            rsRepo.MoveLast
55            If rsRepo.Fields(18).value = 0 And rsRepo.Fields(19).value = 0 And rsRepo.Fields(20).value = 0 Then
56                Set rsRepo = objO.FetchData("SELECT DISTINCT * FROM [TEMPDB]..[" & vNamaTabel & "] ORDER BY KLPKERJA DESC, REP DESC, [OFF]")
57            Else
58                rsRepo.MoveFirst
59            End If
          
60        rpt.database.SetDataSource rsRepo
          
61        If nRC Then
              Dim rsHlp As New ADODB.Recordset
              Dim nBulan As String
              Dim nSbe As String
62            nSbe = ""
               
63            Set rsHlp = objO.FetchData("SELECT CASE WHEN JT_TGL < JT_BLN THEN JT_TGL ELSE JT_BLN END AS JT_TGL, SBENEXT FROM M_ESTIMASI WITH(NOLOCK) WHERE NORANGKA = '" & Trim(text(17)) & "' AND TGLSVCNOW = '" & Format(text(43), "MM/DD/YYYY") & "'")
64            If Not rsHlp.EOF Then
65                nBulan = GetNamaBulan(Month(rsHlp("JT_TGL"))) & " " & year(rsHlp("JT_TGL"))
66                nSbe = Trim(rsHlp("SBENEXT"))
67            End If
68            rsHlp.Close
              
69            If Trim(nSbe) <> "" Then
                  'gcat = IIf(nHeavy, "Kembali untuk service berkala setelah 5.000 KM atau bulan " & nBulan, "Kembali untuk service berkala setelah 10.000 KM atau bulan " & nBulan)
70                gcat = "Kembali untuk service berkala " & Format(CDbl(nSbe), "#,#0") & " KM atau bulan " & nBulan
71                gcat = gcat + IIf(gcat <> "", " (mana yang tercapai lebih dahulu).", "")
72            End If
73        End If
          Dim rsTerbilang As ADODB.Recordset
          Dim totalTerbilang As Double

          'Set rsTerbilang = objO.FetchData("SELECT NOINV, ROUND(SBLT+PART+OLI+SUM([OFF]*((100-DISC)/100)),0) AS REV from [TEMPDB]..[" & vNamaTabel & "] GROUP  BY NOINV, SBLT, PART, OLI")
          'Set rsTerbilang = objO.FetchData("SELECT total,NOINV, ROUND(SBLT+PART+OLI+MET+SUM([OFF]*((100-DISC)/100)),0) AS REV from [TEMPDB]..[" & vNamaTabel & "] GROUP  BY NOINV, SBLT, PART, OLI, MET,total")
          
          '---------------------------------------------
          '2015-07-30 | GANTI QUERY TERBILANG          -
          '---------------------------------------------
74        If nDuoInv = True Or (text(24) > 0 And text(25) > 0) Or text(27) > 0 Then
              ''Set rsTerbilang = objO.FetchData("SELECT distinct total,A.noinv, a.rev+SBLT+round(PART+OLI,0)+MET AS REV FROM (SELECT NOINV, ROUND(SUM([OFF]*((100-DISC)/100)),0) AS REV from [TEMPDB]..[" & vNamaTabel & "] GROUP  BY NOINV) A INNER JOIN [TEMPDB]..[" & vNamaTabel & "] B ON A.NOINV = B.NOINV")
              ''2016-05-20 || pergantian query
75            Set rsTerbilang = objO.FetchData("SELECT NOINV, ROUND(SUM([OFF]*((100-DISC)/100)) + sblt + Part + oli + met,0) AS REV from [TEMPDB]..[" & vNamaTabel & "] GROUP  BY NOINV,sblt,part,oli,met")
76        Else
77            Set rsTerbilang = objO.FetchData("SELECT total,NOINV, ROUND(SBLT+PART+OLI+MET+SUM([OFF]*((100-DISC)/100)),0) AS REV from [TEMPDB]..[" & vNamaTabel & "] GROUP  BY NOINV, SBLT, PART, OLI, MET,total")
78        End If

79        If Not rsTerbilang.EOF Then
80            Do While Not rsTerbilang.EOF
81                totalTerbilang = totalTerbilang + rsTerbilang("REV")
82                rsTerbilang.MoveNext
83            Loop
84            rsTerbilang.MoveFirst
      '53            If totalTerbilang <> CDbl(rsTerbilang("total")) Then
      '54                totalTerbilang = CDbl(rsTerbilang("total"))
'55            End If
85        End If
86        rpt.txtTerbilang.SetText Trim(Angka2Huruf(totalTerbilang, 1))
          ''2016-09-06 || text booking online
87        Set rsHlp = objO.FetchData("SELECT ISNULL(INFOBILL,'') AS INFO FROM T_PROFIL WITH (NOLOCK)")
88        If Not rsHlp.EOF Then
89          infoBill = Trim(rsHlp!info)
90        End If
          'rpt.txtTerbilang.SetText Trim(Angka2Huruf(Round(total, 0), 1)) 'Trim(txtTotalRp)
91        rpt.txtManager.SetText Trim(gSecmanager)
92        If ChCopy.value = 1 Then
93            rpt.txtKet.SetText "COPY INVOICE, " & Trim(gcat)
94        Else
95            rpt.txtKet.SetText Trim(gcat)
96        End If
97        rpt.txtInfoBilling.SetText Trim(infoBill)
98        rpt.txtRetur.SetText IIf(lblRetur.Visible, "DIBATALKAN", "") & IIf(lblPWC.Visible, "PWC", "")
99        rpt.txtJT.SetText "Jatuh Tempo : " & JTBayar(gKodePlg, text(1))
          
100       rpt.txtInfoMaterai.SetText ""
101       rpt.txtLunasMaterai.SetText ""
          
102       If Log.fakturMaterai = 1 Then
103           Set rsTemp = objO.FetchData("SELECT * FROM t_invoiceMaterai WITH (NOLOCK) WHERE NOINV = '" & text(0) & IIf(nDuoInv, "-A", "") & "'")
104           If Not rsTemp.EOF Then
105               rpt.txtInfoMaterai.SetText "Ijin Pembubuhan Materai No. " & NulSID(rsTemp!NOIJIN)
106               rpt.txtLunasMaterai.SetText "BEA MATERAI LUNAS " & NulSID(rsTemp!nilai)
107           End If
108       End If
          
109       With CRViewer1
110           .ReportSource = rpt
111           .ViewReport
112           .Zoom 100
113           .Top = 0
114           .Left = asxPowerBanner1.Left + asxPowerBanner1.Width + 20
115           .Width = Frame2.Width + 100
116           .Height = asxPowerBanner1.Height
              
117           cmdTutup.Top = .Top + 430
118           cmdTutup.Left = .Left
119           cmdTutup.Visible = True
120           cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 5
121           cmdSettingPrinter.Top = cmdTutup.Top
122           cmdSettingPrinter.Visible = True
123           .Visible = True
124       End With
125       cmdTutup.Visible = True
126       Screen.MousePointer = vbDefault
127       Exit Sub
          
gagalCetak:
128       MsgBox "Invoice tidak dapat dicetak ulang !"
129       Exit Sub

         'On Error GoTo 0
130      Exit Sub

cetakInv1_Error:

131       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cetakInv1 of Form frmBillingUlang")
End Sub

' ==============================================================================================
' 08 Maret 2017 | Thania | penambahan informasi PWC utk cetakkan invoice & Cetak Ulang
' ==============================================================================================
Private Sub panggilWarranty(noWO As String)
1         On Error GoTo panggilWarranty_error
          
          Dim sPWC As Integer
          
2         sPWC = objO.FetchData_Integer("SELECT PWC FROM DIAGAPPROVAL WHERE NO_WO = '" & Trim(noWO) & "' AND ALASAN_TDKSETUJU = ''")
3         If sPWC = 1 Then
4             lblPWC.Visible = True
5         Else
6             lblPWC.Visible = False
7         End If
8         Exit Sub
          
panggilWarranty_error:
9         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure returnJob of Form frmBilling")
End Sub

' ==============================================================================================

'' ===========================================================================
'' 20180906 SIMPLIFY PROSES
'' ===========================================================================

Private Sub createTableNewBilling()
1        On Error GoTo createTableNewBilling_Error
2         If rsCetakNewBilling.State Then rsCetakNewBilling.Close
3         Set rsCetakNewBilling = Nothing
          
4         With rsCetakNewBilling
5             .Fields.Append "namaperusahaan", adVarChar, 200
6             .Fields.Append "alamatperusahaan", adVarChar, 200
7             .Fields.Append "teleponperusahaan", adVarChar, 200
8             .Fields.Append "emailperusahaan", adVarChar, 200
9             .Fields.Append "alamatperusahaanpajak", adVarChar, 200
10            .Fields.Append "npwpperusahaan", adVarChar, 200
11            .Fields.Append "pkpperusahaan", adVarChar, 200
12            .Fields.Append "noinv", adVarChar, 20
13            .Fields.Append "kodelgn", adVarChar, 30
14            .Fields.Append "customer", adVarChar, 100
15            .Fields.Append "customeralamat", adVarChar, 200
16            .Fields.Append "nopolisi", adVarChar, 12
17            .Fields.Append "model", adVarChar, 50
18            .Fields.Append "norangka", adVarChar, 25
19            .Fields.Append "nomesin", adVarChar, 15
20            .Fields.Append "noorder", adVarChar, 40
21            .Fields.Append "namasa", adVarChar, 40
22            .Fields.Append "tipecustomer", adVarChar, 40
23            .Fields.Append "mekanik", adVarChar, 40
24            .Fields.Append "printed", adVarChar, 40
              ''.Fields.Append "billing", adDate
              ''20191108 REQ MB THANIA BILLING DIBUAT INISIAL FRONT
25            .Fields.Append "billing", adVarChar, 10
26            .Fields.Append "job", adVarChar, 50
27            .Fields.Append "jobdetail", adVarChar, 100
28            .Fields.Append "jobrate", adDouble
29            .Fields.Append "jobharga", adDouble
30            .Fields.Append "jobdisc", adVarChar, 10
31            .Fields.Append "jobsubtotal", adDouble
32            .Fields.Append "dppJasa", adDouble
33            .Fields.Append "dppPart", adDouble
34            .Fields.Append "dppSublet", adDouble
35            .Fields.Append "dppTotal", adDouble
36            .Fields.Append "ppnTotal", adDouble
37            .Fields.Append "grandTotal", adDouble
38            .Fields.Append "kodeqr", adVarChar, 200
39            .Open
40        End With

41       On Error GoTo 0
42       Exit Sub
createTableNewBilling_Error:
43        MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure createTableNewBilling of Form frmBilling"
End Sub

Private Sub cetakInv1New(noInv As String, No As Integer)
          Dim rsRepo As ADODB.Recordset
          Dim rsAddJob As ADODB.Recordset
          Dim rsProfil As ADODB.Recordset
          Dim rsRinciSBE As ADODB.Recordset
          'Dim rpt As New rtpNewInvoice
          Dim hlp1 As String
          Dim hlp2 As String
          Dim hlp3 As String
          Dim hlp4 As String
          Dim hlp6 As String
          Dim gcat As String
          Dim infoBill As String
          Dim no_wo As String
          Dim nilaiMAT As Double
          Dim tglInv  As String
          Dim kodeLgnCetak, kodeQR, kodeCabangVA As String
          
          
          Dim onlySublet      As Boolean
          Dim isSBE           As Boolean
          Dim isFixPrice      As Boolean
          Dim discFixPrice    As Double
          Dim total           As Double
          Dim isSNL           As Boolean
          
          'Dim rpt As Object
1         If Log.dpplain = 1 Then
2             isSNL = objO.FetchData_Boolean("SELECT * FROM T_InvoicePPn WHERE NOINV = '" & noInv & "' ")
3         Else
4             isSNL = False
5         End If
          
6         kodeCabangVA = objO.FetchData_String("select KodeVA from ReplVA..Master_KodeVA where KodeDealer = '" & Log.kodeCabang & "'")
          
7         If isSNL Then
8             Set rpt = New rtpNewInvoiceSNL
9         Else
10            Set rpt = New rtpNewInvoice
11        End If
          
12        kodeQR = objO.FetchData_String("select UnixID from [" & dbReplikasiVA & "]..QRIS_Invoice where KodeDealer = '" & Log.kodeCabang & "' and NoInv = 'E-" & noInv & "'")
          
13        If Log.notifInvNCM = 1 Then
14            rpt.boxNCM.Suppress = True
15            rpt.txtNCM1.Suppress = True
16            rpt.txtNCM2.Suppress = True
              
17            If isJualNRM = True Then
18                Set rsTemp = objO.FetchData("SELECT * FROM notifInvoice WITH (NOLOCK) WHERE tipe = 'NCM'")
19                If Not rsTemp.EOF Then
                  
20                    rpt.boxNCM.Suppress = False
21                    rpt.txtNCM1.Suppress = False
22                    rpt.txtNCM1.SetText Trim(rsTemp!Text1)
                      
23                    rpt.txtNCM2.Suppress = False
24                    rpt.txtNCM2.SetText Trim(rsTemp!Text2)
                   
25                End If
26            Else
27                Set rsTemp = objO.FetchData("SELECT * FROM notifInvoice WITH (NOLOCK) WHERE tipe = 'NonNCM'")
28                If Not rsTemp.EOF Then
                  
29                    rpt.boxNCM.Suppress = False
30                    rpt.txtNCM1.Suppress = False
31                    rpt.txtNCM1.SetText Trim(rsTemp!Text1)
                      
32                    rpt.txtNCM2.Suppress = False
33                    rpt.txtNCM2.SetText Trim(rsTemp!Text2)
34                End If
35            End If
36        End If
37        nilaiMAT = 0
38        createTableNewBilling
39        sSQL = "select a.NOINV, a.NO_WO,a.FS_KD_PLG,a.CUSTOMER,a.CUST_ADRES,"
40        sSQL = sSQL & " a.NO_POLISI,a.NO_CHASSIS,a.FMDL,a.NO_ENGINE,a.NO_WO,a.FRONT,A.STATUS,A.MECH,A.INVDATE,A.INVTIME,"
41        sSQL = sSQL & " A.REP1,A.REP2,A.REP3,A.REP4,A.REP5,A.REP6,A.REP7,A.REP8,"
42        sSQL = sSQL & " A.RTE1,A.RTE2,A.RTE3,A.RTE4,A.RTE5,A.RTE6,A.RTE7,A.RTE8,"
43        sSQL = sSQL & " A.OFF1,A.OFF2,A.OFF3,A.OFF4,A.OFF5,A.OFF6,A.OFF7,A.OFF8,"
44        sSQL = sSQL & " A.DISC1,A.DISC2,A.DISC3,A.DISC4,A.DISC5,A.DISC6,A.DISC7,A.DISC8,"
45        sSQL = sSQL & " PART_G,PART_NG,MTRL,OIL_G,OIL_NG,MET,SBLT,LABOR,BODY,TOTAL-(FN_PAJAK+met) AS DPP,FN_PAJAK,TOTAL,FS_RC"
46        sSQL = sSQL & " from T_INVOICE a with (nolock) "
47        sSQL = sSQL & " where NOINV = '" & noInv & "'"
48        sSQL = sSQL & " order by INVDATE desc"
49        Set rsRepo = objO.FetchData(sSQL)
50        sSQL = "SELECT A.* FROM HIST_JOBADD A WITH(NOLOCK) INNER JOIN T_INVOICE B WITH(NOLOCK) ON A.NO_WO = B.NO_WO AND A.NO_INV = B.NOINV AND A.NO_CHASSIS = B.NO_CHASSIS AND A.FS_LOG_DEL = B.FS_LOG_DEL  WHERE A.NO_WO = '" & Trim(text(20)) & "' AND A.NO_CHASSIS = '" & Trim(text(17)) & "' AND A.FS_LOG_DEL = '' AND B.NOINV LIKE '" & Trim(text(0)) & "%'"
51        Set rsAddJob = objO.FetchData(sSQL)
          
52        sSQL = "select FS_NM_PERUSH,FS_ALM_PERUSH,'No Telpon : ' + FS_TELP + ', Fax : ' + FS_FAX AS FS_TELP,'Email : ' + FS_EMAIL as FS_EMAIL,'NPWP : ' + FS_NPWP + ' Tanggal : ' as FS_NPWP,'PKP : ' + FS_PKP as FS_PKP,bankname1,bankname2,bankname3,bankno1,bankno2,bankno3 from t_profil"
53        Set rsProfil = objO.FetchData(sSQL)
54        If rsProfil.EOF Then
55            MsgBox "Gagal memuat data profil dealer", vbInformation
56            Exit Sub
57        End If
58        If Not rsRepo.EOF Then
59        kodeLgnCetak = rsRepo!FS_KD_PLG
60            nilaiMAT = rsRepo!met
61            rsRepo.MoveFirst
62            isSBE = False
63            no_wo = NulSID(rsRepo!no_wo)
64            tglInv = TglDDMMMMYYYY(rsRepo!InvDate, 1) & " " & rsRepo!INVTIME
65            If NulSID(rsRepo!FS_RC) <> "" Then
66                isSBE = True
67            End If
68            onlySublet = True
69            Do While Not rsRepo.EOF
70                For i = 1 To 8
71                    hlp1 = "REP" & i
72                    hlp2 = "RTE" & i
73                    hlp3 = "OFF" & i
74                    hlp4 = "DISC" & i
                      
                      '' -------------------------------------------------
                      '' 20200316 || cetakan invoice untuk fix price ditampilkan disc
75                    isFixPrice = False
76                    discFixPrice = 0
77                    If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(text(43), "DD/MM/YYYY")) >= 0 Then
78                        If objO.FetchData_String("select no_wo from t_invoicejobdetail where fs_log_del = '' and isFP = 1 and rep = '" & NulSID(rsRepo(hlp1)) & "' and no_wo = '" & NulSID(rsRepo!no_wo) & "'") <> "" Then
79                            isFixPrice = True
80                            discFixPrice = objO.FetchData_Double("select disc from t_invoicejobdetail where fs_log_del = '' and isFP = 1 and rep = '" & NulSID(rsRepo(hlp1)) & "' and no_wo = '" & NulSID(rsRepo!no_wo) & "'")
81                        End If
82                    End If
                      
83                    If Trim(rsRepo(hlp1)) <> "" Then
84                        With rsCetakNewBilling
85                            .AddNew
86                            !namaperusahaan = NulSID(rsProfil!FS_NM_PERUSH)
87                            !alamatperusahaan = NulSID(rsProfil!FS_ALM_PERUSH)
88                            !teleponperusahaan = NulSID(rsProfil!FS_TELP)
89                            !emailperusahaan = NulSID(rsProfil!FS_EMAIL)
90                            !alamatperusahaanpajak = NulSID("")
91                            !npwpperusahaan = NulSID(rsProfil!FS_NPWP)
92                            !pkpperusahaan = NulSID(rsProfil!FS_PKP)
93                            !noInv = NulSID(rsRepo!noInv)
94                            !kodeLgn = objO.FetchData_String("SELECT idcust FROM [SINGLEIDORI].dbo.bridging WITH (NOLOCK) WHERE kodelgn = '" & NulSID(rsRepo!FS_KD_PLG) & "' AND kodedealer = '" & Right(gDb, 2) & "' ")
95                            !customer = NulSID(rsRepo!customer)
96                            !customeralamat = NulSID(rsRepo!cust_adres)
97                            !NoPolisi = NulSID(rsRepo!NO_POLISI)
98                            !model = NulSID(rsRepo!Fmdl)
99                            !noRangka = NulSID(rsRepo!NO_CHASSIS)
100                           !noMesin = NulSID(rsRepo!No_Engine)
101                           !noOrder = NulSID(rsRepo!no_wo)
102                           !namaSA = objO.FetchData_String("select FS_NM_PETUGAS from T_PETUGAS with (nolock) where FS_KD_PETUGAS = '" & NulSID(rsRepo!FRONT) & "'")
103                           !tipecustomer = NulSID(rsRepo!Status)
104                           !mekanik = NulSID(rsRepo!Mech)
105                           !billing = NulSID(rsRepo!FRONT) ''Format(NulSID(rsRepo!InvDate), "dd-mm-yyyy")
106                           !job = Left(IIf(InStr(NulSID(rsRepo(hlp1)), ".000") > 0, " ", "") & NulSID(rsRepo(hlp1)), 50)
107                           !jobdetail = NulSID("")
108                           !jobRate = k2t(NulSID(rsRepo(hlp2)))
109                           !jobharga = NulSID(rsRepo(hlp3))
                              '!jobdisc = IIf(chCtkDisc.value = False, NulSID(rsRepo(hlp4) & " %"), "") 'NulSID(rsRepo(hlp4) & " %")
110                           If isFixPrice = True Then
111                               !jobdisc = IIf(chCtkDisc.value = False, NulSID(discFixPrice & " %"), "")
112                           Else
113                               !jobdisc = IIf(chCtkDisc.value = False, NulSID(rsRepo(hlp4) & " %"), "")
114                           End If
115                           !jobsubtotal = k2t(NulSID((rsRepo(hlp3) * (100 - NulSID(rsRepo(hlp4)))) / 100)) ''NulSID(rsRepo(hlp3) * ((100 - NulSID(rsRepo(hlp4))) / 100))
116                           !dppJasa = Round(NulA(rsRepo!labor) * nilaiDppInv / nilaiTotalInv, 0) ''nilaiDppInv / nilaiTotalInv,0)  ''10 / 11, 0)
117                           !dpppart = Round((NulA(rsRepo!PART_G) + NulA(rsRepo!PART_NG) + NulA(rsRepo!MTRL) + NulA(rsRepo!OIL_G) + NulA(rsRepo!OIL_NG)) * nilaiDppInv / nilaiTotalInv, 0) ''nilaiDppInv / nilaiTotalInv,0)  ''10 / 11, 0)
118                           !dppsublet = Round(NulA(rsRepo!sblt) * nilaiDppInv / nilaiTotalInv, 0) ''nilaiDppInv / nilaiTotalInv,0)  ''10 / 11, 0)
119                           !dpptotal = NulA(!dppJasa + !dpppart + !dppsublet) ''NulA(rsRepo!dpp)
120                           !PPNtotal = NulA(rsRepo!fn_pajak)
121                           !grandTotal = NulA(!dpptotal + !PPNtotal + nilaiMAT) ''NulA(rsRepo!total)
122                           !kodeQR = NulSID(kodeQR)
123                           .Update
124                       End With
125                       onlySublet = False
126                       If isSBE Then
                              Dim nilaisbe As Integer
127                           nilaisbe = CInt(Right(NulSID(rsRepo!FS_RC), 3))
128                           If InStr(NulSID(rsRepo(hlp1)), nilaisbe & ".000") > 0 Then
129                               Set rsRinciSBE = objO.FetchData("select rinciansb as detail from rinciansbe where tipesb = 'R" & Format(nilaisbe, String(3, "0")) & "' order by Status")
130                               If Not rsRinciSBE.EOF Then
131                                   rsRinciSBE.MoveFirst
132                                   Do While Not rsRinciSBE.EOF
133                                       With rsCetakNewBilling
134                                           .AddNew
135                                           !namaperusahaan = NulSID(rsProfil!FS_NM_PERUSH)
136                                           !alamatperusahaan = NulSID(rsProfil!FS_ALM_PERUSH)
137                                           !teleponperusahaan = NulSID(rsProfil!FS_TELP)
138                                           !emailperusahaan = NulSID(rsProfil!FS_EMAIL)
139                                           !alamatperusahaanpajak = NulSID("")
140                                           !npwpperusahaan = NulSID(rsProfil!FS_NPWP)
141                                           !pkpperusahaan = NulSID(rsProfil!FS_PKP)
142                                           !noInv = NulSID(rsRepo!noInv)
143                                           !kodeLgn = objO.FetchData_String("SELECT idcust FROM [SINGLEIDORI].dbo.bridging WITH (NOLOCK) WHERE kodelgn = '" & NulSID(rsRepo!FS_KD_PLG) & "' AND kodedealer = '" & Right(gDb, 2) & "' ")
144                                           !customer = NulSID(rsRepo!customer)
145                                           !customeralamat = NulSID(rsRepo!cust_adres)
146                                           !NoPolisi = NulSID(rsRepo!NO_POLISI)
147                                           !model = NulSID(rsRepo!Fmdl)
148                                           !noRangka = NulSID(rsRepo!NO_CHASSIS)
149                                           !noMesin = NulSID(rsRepo!No_Engine)
150                                           !noOrder = NulSID(rsRepo!no_wo)
151                                           !namaSA = objO.FetchData_String("select FS_NM_PETUGAS from T_PETUGAS with (nolock) where FS_KD_PETUGAS = '" & NulSID(rsRepo!FRONT) & "'")
152                                           !tipecustomer = NulSID(rsRepo!Status)
153                                           !mekanik = NulSID(rsRepo!Mech)
154                                           !billing = NulSID(rsRepo!FRONT) ''NulSID(rsRepo!InvDate) ''Format(NulSID(rsRepo!InvDate), "dd-mm-yyyy")
155                                           !job = Left(" " & NulSID(rsRepo(hlp1)), 50)
156                                           !jobdetail = NulSID(rsRinciSBE!Detail)
157                                           !jobRate = k2t(NulSID(rsRepo(hlp2)))
158                                           !jobharga = NulSID(rsRepo(hlp3))
                                              ''!jobdisc = IIf(chCtkDisc.value = False, NulSID(rsRepo(hlp4) & " %"), "") 'NulSID(rsRepo(hlp4) & " %")
159                                           If isFixPrice = True Then
160                                               !jobdisc = IIf(chCtkDisc.value = False, NulSID(discFixPrice & " %"), "")
161                                           Else
162                                               !jobdisc = IIf(chCtkDisc.value = False, NulSID(rsRepo(hlp4) & " %"), "")
163                                           End If
164                                           !jobsubtotal = k2t(NulSID(rsRepo(hlp3) * ((100 - NulSID(rsRepo(hlp4))) / 100)))
165                                           !dppJasa = Round(NulA(rsRepo!labor) * nilaiDppInv / nilaiTotalInv, 0) ''nilaiDppInv / nilaiTotalInv,0)  ''10 / 11, 0)
166                                           !dpppart = Round((NulA(rsRepo!PART_G) + NulA(rsRepo!PART_NG) + NulA(rsRepo!MTRL) + NulA(rsRepo!OIL_G) + NulA(rsRepo!OIL_NG)) * nilaiDppInv / nilaiTotalInv, 0) ''nilaiDppInv / nilaiTotalInv,0)  ''10 / 11, 0)
167                                           !dppsublet = Round(NulA(rsRepo!sblt) * nilaiDppInv / nilaiTotalInv, 0) ''nilaiDppInv / nilaiTotalInv,0)  ''10 / 11, 0)
168                                           !dpptotal = NulA(!dppJasa + !dpppart + !dppsublet) ''NulA(rsRepo!dpp)
169                                           !PPNtotal = NulA(rsRepo!fn_pajak)
170                                           !grandTotal = NulA(!dpptotal + !PPNtotal + nilaiMAT) ''NulA(rsRepo!total)
171                                           !kodeQR = NulSID(kodeQR)
172                                           .Update
173                                       End With
174                                       rsRinciSBE.MoveNext
175                                   Loop
176                               End If
177                           End If
178                       End If
179                   End If
180               Next
                  
181               If onlySublet Then
182                   With rsCetakNewBilling
183                       .AddNew
184                       !namaperusahaan = NulSID(rsProfil!FS_NM_PERUSH)
185                       !alamatperusahaan = NulSID(rsProfil!FS_ALM_PERUSH)
186                       !teleponperusahaan = NulSID(rsProfil!FS_TELP)
187                       !emailperusahaan = NulSID(rsProfil!FS_EMAIL)
188                       !alamatperusahaanpajak = NulSID("")
189                       !npwpperusahaan = NulSID(rsProfil!FS_NPWP)
190                       !pkpperusahaan = NulSID(rsProfil!FS_PKP)
191                       !noInv = NulSID(rsRepo!noInv)
192                       !kodeLgn = objO.FetchData_String("SELECT idcust FROM [SINGLEIDORI].dbo.bridging WITH (NOLOCK) WHERE kodelgn = '" & NulSID(rsRepo!FS_KD_PLG) & "' AND kodedealer = '" & Right(gDb, 2) & "' ")
193                       !customer = NulSID(rsRepo!customer)
194                       !customeralamat = NulSID(rsRepo!cust_adres)
195                       !NoPolisi = NulSID(rsRepo!NO_POLISI)
196                       !model = NulSID(rsRepo!Fmdl)
197                       !noRangka = NulSID(rsRepo!NO_CHASSIS)
198                       !noMesin = NulSID(rsRepo!No_Engine)
199                       !noOrder = NulSID(rsRepo!no_wo)
200                       !namaSA = objO.FetchData_String("select FS_NM_PETUGAS from T_PETUGAS with (nolock) where FS_KD_PETUGAS = '" & NulSID(rsRepo!FRONT) & "'")
201                       !tipecustomer = NulSID(rsRepo!Status)
202                       !mekanik = NulSID(rsRepo!Mech)
203                       !billing = NulSID(rsRepo!FRONT) ''NulSID(rsRepo!InvDate) ''Format(NulSID(rsRepo!InvDate), "dd-mm-yyyy")
204                       !job = NulSID("Order Sublet")
205                       !jobdetail = NulSID("")
206                       !jobRate = 0 'NulSID(rsRepo(hlp2))
207                       !jobharga = 0 'NulSID(rsRepo(hlp3))
208                       !jobdisc = 0 'NulSID(rsRepo(hlp4) & " %")
209                       !jobsubtotal = 0 ' NulSID(rsRepo(hlp3) * ((100 - NulSID(rsRepo(hlp4))) / 100))
210                       !dppJasa = Round(NulA(rsRepo!labor) * nilaiDppInv / nilaiTotalInv, 0) ''nilaiDppInv / nilaiTotalInv,0)  ''10 / 11, 0)
211                       !dpppart = Round((NulA(rsRepo!PART_G) + NulA(rsRepo!PART_NG) + NulA(rsRepo!MTRL) + NulA(rsRepo!OIL_G) + NulA(rsRepo!OIL_NG)) * nilaiDppInv / nilaiTotalInv, 0) ''nilaiDppInv / nilaiTotalInv,0)  ''10 / 11, 0)
212                       !dppsublet = Round(NulA(rsRepo!sblt) * nilaiDppInv / nilaiTotalInv, 0) ''nilaiDppInv / nilaiTotalInv,0)  ''10 / 11, 0)
213                       !dpptotal = NulA(!dppJasa + !dpppart + !dppsublet) ''NulA(rsRepo!dpp)
214                       !PPNtotal = NulA(rsRepo!fn_pajak)
215                       !grandTotal = NulA(!dpptotal + !PPNtotal + nilaiMAT) ''NulA(rsRepo!total)
216                       !kodeQR = NulSID(kodeQR)
217                       .Update
218                   End With
219               End If
220               rsRepo.MoveNext
221           Loop
              
222           If Not rsAddJob.EOF Then
223               rsAddJob.MoveFirst
224               rsRepo.MoveFirst
225               Do While Not rsAddJob.EOF
                      '' -------------------------------------------------
                      '' 20200316 || cetakan invoice untuk fix price ditampilkan disc
226                   isFixPrice = False
227                   discFixPrice = 0
228                   If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(text(43), "DD/MM/YYYY")) >= 0 Then
229                       If objO.FetchData_String("select no_wo from t_invoicejobdetail where fs_log_del = '' and isFP = 1 and rep = '" & NulSID(rsAddJob!rep) & "' and no_wo = '" & NulSID(rsRepo!no_wo) & "'") <> "" Then
230                           isFixPrice = True
231                           discFixPrice = objO.FetchData_Double("select disc from t_invoicejobdetail where fs_log_del = '' and isFP = 1 and rep = '" & NulSID(rsAddJob!rep) & "' and no_wo = '" & NulSID(rsRepo!no_wo) & "'")
232                       End If
233                   End If
234                   With rsCetakNewBilling
235                       .AddNew
236                       !namaperusahaan = NulSID(rsProfil!FS_NM_PERUSH)
237                       !alamatperusahaan = NulSID(rsProfil!FS_ALM_PERUSH)
238                       !teleponperusahaan = NulSID(rsProfil!FS_TELP)
239                       !emailperusahaan = NulSID(rsProfil!FS_EMAIL)
240                       !alamatperusahaanpajak = NulSID("")
241                       !npwpperusahaan = NulSID(rsProfil!FS_NPWP)
242                       !pkpperusahaan = NulSID(rsProfil!FS_PKP)
243                       !noInv = NulSID(rsRepo!noInv)
244                       !kodeLgn = objO.FetchData_String("SELECT idcust FROM [SINGLEIDORI].dbo.bridging WITH (NOLOCK) WHERE kodelgn = '" & NulSID(rsRepo!FS_KD_PLG) & "' AND kodedealer = '" & Right(gDb, 2) & "' ")
245                       !customer = NulSID(rsRepo!customer)
246                       !customeralamat = NulSID(rsRepo!cust_adres)
247                       !NoPolisi = NulSID(rsRepo!NO_POLISI)
248                       !model = NulSID(rsRepo!Fmdl)
249                       !noRangka = NulSID(rsRepo!NO_CHASSIS)
250                       !noMesin = NulSID(rsRepo!No_Engine)
251                       !noOrder = NulSID(rsRepo!no_wo)
252                       !namaSA = objO.FetchData_String("select FS_NM_PETUGAS from T_PETUGAS with (nolock) where FS_KD_PETUGAS = '" & NulSID(rsRepo!FRONT) & "'")
253                       !tipecustomer = NulSID(rsRepo!Status)
254                       !mekanik = NulSID(rsRepo!Mech)
255                       !billing = NulSID(rsRepo!FRONT) ''NulSID(rsRepo!InvDate) ''Format(NulSID(rsRepo!InvDate), "dd-mm-yyyy")
256                       !job = Left("Add. " & NulSID(rsAddJob!rep), 50)
257                       !jobdetail = NulSID("")
258                       !jobRate = k2t(NulSID(rsAddJob!rte))
259                       !jobharga = NulSID(rsAddJob!offn)
                          ''!jobdisc = IIf(chCtkDisc.value = False, NulSID(rsAddJob!disc) & " %", "") 'NulSID(rsAddJob!disc) & " %"
260                       If isFixPrice = True Then
261                           !jobdisc = IIf(chCtkDisc.value = False, NulSID(discFixPrice & " %"), "")
262                       Else
263                           !jobdisc = IIf(chCtkDisc.value = False, NulSID(rsAddJob!disc) & " %", "")
264                       End If
265                       !jobsubtotal = Round(NulSID(rsAddJob!offn * ((100 - NulSID(rsAddJob!disc)) / 100)))
266                       !dppJasa = Round(NulA(rsRepo!labor) * nilaiDppInv / nilaiTotalInv, 0) ''nilaiDppInv / nilaiTotalInv,0)  ''10 / 11, 0)
267                       !dpppart = Round((NulA(rsRepo!PART_G) + NulA(rsRepo!PART_NG) + NulA(rsRepo!MTRL) + NulA(rsRepo!OIL_G) + NulA(rsRepo!OIL_NG)) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
268                       !dppsublet = Round(NulA(rsRepo!sblt) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
269                       !dpptotal = NulA(!dppJasa + !dpppart + !dppsublet) ''NulA(rsRepo!dpp)
270                       !PPNtotal = NulA(rsRepo!fn_pajak)
271                       !grandTotal = NulA(!dpptotal + !PPNtotal + nilaiMAT) ''NulA(rsRepo!total)
272                       !kodeQR = NulSID(kodeQR)
273                       .Update
274                   End With
275                   If isSBE Then
                          Dim nilaisbe2 As Integer
276                       nilaisbe2 = CInt(Right(NulSID(rsRepo!FS_RC), 3))
277                       If InStr(NulSID(rsRepo(hlp1)), nilaisbe2 & ".000") > 0 Then
278                           Set rsRinciSBE = objO.FetchData("select rinciansb as detail from rinciansbe where tipesb = 'R" & Format(nilaisbe2, String(3, "0")) & "' order by Status")
279                           If Not rsRinciSBE.EOF Then
280                               rsRinciSBE.MoveFirst
281                               Do While Not rsRinciSBE.EOF
282                                   With rsCetakNewBilling
283                                       .AddNew
284                                       !namaperusahaan = NulSID(rsProfil!FS_NM_PERUSH)
285                                       !alamatperusahaan = NulSID(rsProfil!FS_ALM_PERUSH)
286                                       !teleponperusahaan = NulSID(rsProfil!FS_TELP)
287                                       !emailperusahaan = NulSID(rsProfil!FS_EMAIL)
288                                       !alamatperusahaanpajak = NulSID("")
289                                       !npwpperusahaan = NulSID(rsProfil!FS_NPWP)
290                                       !pkpperusahaan = NulSID(rsProfil!FS_PKP)
291                                       !noInv = NulSID(rsRepo!noInv)
292                                       !kodeLgn = objO.FetchData_String("SELECT idcust FROM [SINGLEIDORI].dbo.bridging WITH (NOLOCK) WHERE kodelgn = '" & NulSID(rsRepo!FS_KD_PLG) & "' AND kodedealer = '" & Right(gDb, 2) & "' ")
293                                       !customer = NulSID(rsRepo!customer)
294                                       !customeralamat = NulSID(rsRepo!cust_adres)
295                                       !NoPolisi = NulSID(rsRepo!NO_POLISI)
296                                       !model = NulSID(rsRepo!Fmdl)
297                                       !noRangka = NulSID(rsRepo!NO_CHASSIS)
298                                       !noMesin = NulSID(rsRepo!No_Engine)
299                                       !noOrder = NulSID(rsRepo!no_wo)
300                                       !namaSA = objO.FetchData_String("select FS_NM_PETUGAS from T_PETUGAS with (nolock) where FS_KD_PETUGAS = '" & NulSID(rsRepo!FRONT) & "'")
301                                       !tipecustomer = NulSID(rsRepo!Status)
302                                       !mekanik = NulSID(rsRepo!Mech)
303                                       !billing = NulSID(rsRepo!FRONT) ''NulSID(rsRepo!InvDate) ''Format(NulSID(rsRepo!InvDate), "dd-mm-yyyy")
304                                       !job = Left(" " & "Add. " & NulSID(rsAddJob!rep), 50)
305                                       !jobdetail = NulSID(rsRinciSBE!Detail)
306                                       !jobRate = k2t(NulSID(rsAddJob!rte))
307                                       !jobharga = NulSID(rsAddJob!offn)
                                          ''!jobdisc = IIf(chCtkDisc.value = False, NulSID(rsAddJob!disc) & " %", "") 'NulSID(rsAddJob!disc) & " %"
308                                       If isFixPrice = True Then
309                                           !jobdisc = IIf(chCtkDisc.value = False, NulSID(discFixPrice & " %"), "")
310                                       Else
311                                           !jobdisc = IIf(chCtkDisc.value = False, NulSID(rsAddJob!disc) & " %", "")
312                                       End If
313                                       !jobsubtotal = k2t(NulSID(rsAddJob!offn * ((100 - NulSID(rsAddJob!disc)) / 100)))
314                                       !dppJasa = Round(NulA(rsRepo!labor) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
315                                       !dpppart = Round((NulA(rsRepo!PART_G) + NulA(rsRepo!PART_NG) + NulA(rsRepo!MTRL) + NulA(rsRepo!OIL_G) + NulA(rsRepo!OIL_NG)) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
316                                       !dppsublet = Round(NulA(rsRepo!sblt) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
317                                       !dpptotal = NulA(!dppJasa + !dpppart + !dppsublet) ''NulA(rsRepo!dpp)
318                                       !PPNtotal = NulA(rsRepo!fn_pajak)
319                                       !grandTotal = NulA(!dpptotal + !PPNtotal + nilaiMAT) ''NulA(rsRepo!total)
320                                       !kodeQR = NulSID(kodeQR)
321                                       .Update
322                                   End With
323                                   rsRinciSBE.MoveNext
324                               Loop
325                           End If
326                       End If
327                   End If
328                   rsAddJob.MoveNext
329               Loop
330           End If
331           gcat = ""
332           If nRC Then
                  Dim rsHlp As ADODB.Recordset
                  Dim nBulan As String
                  Dim nSbe As String
333               nSbe = ""
334               Set rsHlp = objO.FetchData("SELECT CASE WHEN JT_TGL < JT_BLN THEN JT_TGL ELSE JT_BLN END AS JT_TGL, SBENEXT FROM M_ESTIMASI WITH (NOLOCK) WHERE NORANGKA = '" & Trim(text(17)) & "' AND TGLSVCNOW = '" & Format(text(43), "MM/DD/YYYY") & "'")
335               If Not rsHlp.EOF Then
336                   nBulan = GetNamaBulan(Month(rsHlp("JT_TGL"))) & " " & year(rsHlp("JT_TGL"))
337                   nSbe = Trim(rsHlp("SBENEXT"))
338               End If
339               rsHlp.Close
340               If Trim(nSbe) <> "" Then
                      'gcat = IIf(nHeavy, "Kembali untuk service berkala setelah kelipatan 5.000 KM atau bulan " & nBulan, "Kembali untuk service berkala setelah kelipatan 10.000 KM atau bulan " & nBulan)
                      ''gcat = "Kembali untuk service berkala " & Format(CDbl(nSbe), "#,#0") & " KM atau bulan " & nBulan
341                   gcat = "Kembali untuk service berkala bulan " & nBulan & " atau " & Format(CDbl(nSbe), "#,#0") & " KM "
342                   gcat = gcat + IIf(gcat <> "", " (mana yang tercapai lebih dahulu).", "")
343               Else
344                   Call objO.InsertData("INSERT INTO ESTIMASISALAH (norangka, tglservice, kmsvc, sbesvc, kelipatan, noinv, nowo,customer, alamat,nopolisi, telpon, gsm,cdma) VALUES " & _
                                     "('" & Trim(text(17)) & "','" & Format(text(43), "MM/DD/YYYY") & "'," & CDbl(text(37)) & ",'" & Trim(text(33)) & "','" & IIf(nHeavy, 5000, 10000) & "', " & _
                                     "'" & Trim(text(20)) & "','" & Trim(Trim(text(0))) & "','" & Trim(text(4)) & "','" & Trim(txt5) & "','" & Trim(text(16)) & "','" & Trim(text(7)) & "','" & Trim(text(8)) & "','" & Trim(text(9)) & "')")
345               End If
346           End If
              
347           rpt.txtTerbilang.SetText "Terbilang (Rupiah) " & Trim(Angka2Huruf(rsCetakNewBilling!grandTotal, 1))
              ''20240829 revisi menutup kembali service
348           rpt.txtNextService.SetText "" 'gcat
              
              '' inspection ====================
349           If Log.CR7 = 1 Then
                  Dim rsInspection As ADODB.Recordset
                  Dim total1 As Double
                  Dim total2 As Double
                  Dim total3 As Double
                  Dim total4 As Double
                  
350               sSQL = "select sp1,sp2,sp3,sp4,estn1,estn2,estn3,estn4,estj1,estj2,estj3,estj4,estp1,estp2,estp3,estp4,ests1,ests2,ests3,ests4 from INSPECTION with (nolock) where NoWo = '" & no_wo & "' order by tglentry desc"
351               Set rsInspection = objO.FetchData(sSQL)
352               If Not rsInspection.EOF Then
353                   rsInspection.MoveFirst
354                   rpt.txtSaran1.SetText NulSID(rsInspection!sp1)
355                   rpt.txtSaran2.SetText NulSID(rsInspection!sp2)
356                   rpt.txtSaran3.SetText NulSID(rsInspection!sp3)
357                   rpt.txtSaran4.SetText NulSID(rsInspection!sp4)
                      
358                   If NulSID(rsInspection!sp1) <> "" Then
359                       rpt.txtWaktuSVC1.SetText Format(NulSID(rsInspection!estn1), "dd-mm-yyyy")
360                   Else
361                       rpt.txtWaktuSVC1.SetText ""
362                   End If
363                   If NulSID(rsInspection!sp2) <> "" Then
364                       rpt.txtWaktuSVC2.SetText Format(NulSID(rsInspection!estn2), "dd-mm-yyyy")
365                   Else
366                       rpt.txtWaktuSVC2.SetText ""
367                   End If
368                   If NulSID(rsInspection!sp3) <> "" Then
369                       rpt.txtWaktuSVC3.SetText Format(NulSID(rsInspection!estn3), "dd-mm-yyyy")
370                   Else
371                       rpt.txtWaktuSVC3.SetText ""
372                   End If
373                   If NulSID(rsInspection!sp4) <> "" Then
374                       rpt.txtWaktuSVC4.SetText Format(NulSID(rsInspection!estn4), "dd-mm-yyyy")
375                   Else
376                       rpt.txtWaktuSVC4.SetText ""
377                   End If
                                      
378                   If NulSID(rsInspection!sp1) <> "" Then
379                       rpt.txtEstJasa1.SetText Format(NulSID(rsInspection!estj1), "#,0")
380                   Else
381                       rpt.txtEstJasa1.SetText ""
382                   End If
383                   total1 = total1 + NulSID(rsInspection!estj1)
                      
384                   If NulSID(rsInspection!sp2) <> "" Then
385                       rpt.txtEstJasa2.SetText Format(NulSID(rsInspection!estj2), "#,0")
386                   Else
387                       rpt.txtEstJasa2.SetText ""
388                   End If
389                   total2 = total2 + NulSID(rsInspection!estj2)
                      
390                   If NulSID(rsInspection!sp3) <> "" Then
391                       rpt.txtEstJasa3.SetText Format(NulSID(rsInspection!estj3), "#,0")
392                   Else
393                       rpt.txtEstJasa3.SetText ""
394                   End If
395                   total3 = total3 + NulSID(rsInspection!estj3)
                      
396                   If NulSID(rsInspection!sp4) <> "" Then
397                       rpt.txtEstJasa4.SetText Format(NulSID(rsInspection!estj4), "#,0")
398                   Else
399                       rpt.txtEstJasa4.SetText ""
400                   End If
401                   total4 = total4 + NulSID(rsInspection!estj4)
                                                      
402                   If NulSID(rsInspection!sp1) <> "" Then
403                       rpt.txtEstPart1.SetText Format(NulSID(rsInspection!estp1), "#,0")
404                   Else
405                       rpt.txtEstPart1.SetText ""
406                   End If
407                   total1 = total1 + NulSID(rsInspection!estp1)
                                                      
408                   If NulSID(rsInspection!sp2) <> "" Then
409                       rpt.txtEstPart2.SetText Format(NulSID(rsInspection!estp2), "#,0")
410                   Else
411                       rpt.txtEstPart2.SetText ""
412                   End If
413                   total2 = total2 + NulSID(rsInspection!estp2)
                                                      
414                   If NulSID(rsInspection!sp3) <> "" Then
415                       rpt.txtEstPart3.SetText Format(NulSID(rsInspection!estp3), "#,0")
416                   Else
417                       rpt.txtEstPart3.SetText ""
418                   End If
419                   total3 = total3 + NulSID(rsInspection!estp3)
                                                      
420                   If NulSID(rsInspection!sp4) <> "" Then
421                       rpt.txtEstPart4.SetText Format(NulSID(rsInspection!estp4), "#,0")
422                   Else
423                       rpt.txtEstPart4.SetText ""
424                   End If
425                   total4 = total4 + NulSID(rsInspection!estp4)
                                      
                                                      
426                   If NulSID(rsInspection!sp1) <> "" Then
427                       rpt.txtEstSublet1.SetText Format(NulSID(rsInspection!ests1), "#,0")
428                   Else
429                       rpt.txtEstSublet1.SetText ""
430                   End If
431                   total1 = total1 + NulSID(rsInspection!ests1)
                                                      
432                   If NulSID(rsInspection!sp2) <> "" Then
433                       rpt.txtEstSublet2.SetText Format(NulSID(rsInspection!ests2), "#,0")
434                   Else
435                       rpt.txtEstSublet2.SetText ""
436                   End If
437                   total2 = total2 + NulSID(rsInspection!ests2)
                                                      
438                   If NulSID(rsInspection!sp3) <> "" Then
439                       rpt.txtEstSublet3.SetText Format(NulSID(rsInspection!ests3), "#,0")
440                   Else
441                       rpt.txtEstSublet3.SetText ""
442                   End If
443                   total3 = total3 + NulSID(rsInspection!ests3)
                                                      
444                   If NulSID(rsInspection!sp4) <> "" Then
445                       rpt.txtEstSublet4.SetText Format(NulSID(rsInspection!ests4), "#,0")
446                   Else
447                       rpt.txtEstSublet4.SetText ""
448                   End If
449                   total4 = total4 + NulSID(rsInspection!ests4)
                                      
450                   If NulSID(rsInspection!sp1) <> "" Then
451                       rpt.txtEstTotal1.SetText Format(NulSID(total1), "#,0")
452                   Else
453                       rpt.txtEstTotal1.SetText ""
454                   End If
455                   If NulSID(rsInspection!sp2) <> "" Then
456                       rpt.txtEstTotal2.SetText Format(NulSID(total2), "#,0")
457                   Else
458                       rpt.txtEstTotal2.SetText ""
459                   End If
460                   If NulSID(rsInspection!sp3) <> "" Then
461                       rpt.txtEstTotal3.SetText Format(NulSID(total3), "#,0")
462                   Else
463                       rpt.txtEstTotal3.SetText ""
464                   End If
465                   If NulSID(rsInspection!sp4) <> "" Then
466                       rpt.txtEstTotal4.SetText Format(NulSID(total4), "#,0")
467                   Else
468                       rpt.txtEstTotal4.SetText ""
469                   End If
                      
470               End If
471           End If
              '' akhir inspection ==============
              
              Dim noVirtualaAcc(3), namaBank(3), namaAkun(3) As String
              Dim iVirtual As Integer
              Dim rsVirtual As ADODB.Recordset
472           iVirtual = 0
              
473           Set rsVirtual = objO.FetchData("SELECT virtualAccountNo,bank,virtualAccountName From " & dbReplikasiVA & "..DOKU_INQUIRY where  KodeDealer = '" & Log.kodeCabang & "' and virtualAccountNo like '%" & "2" & Right("000000000" & NulSID(text(20)), 8) & "' order by urut ASC")
474           If Not rsVirtual.EOF Then
475             rsVirtual.MoveFirst
476             Do While Not rsVirtual.EOF
477               noVirtualaAcc(iVirtual) = NulSID(rsVirtual!virtualAccountNo)
478               namaBank(iVirtual) = NulSID(rsVirtual!bank)
479               namaAkun(iVirtual) = NulSID(rsVirtual!virtualAccountName)
480               iVirtual = iVirtual + 1
481               rsVirtual.MoveNext
482             Loop
483           End If
484           If noVirtualaAcc(0) = "" Then
485               rpt.Text56.SetText ""
486               rpt.Text59.SetText ""
487               rpt.Text60.SetText ""
488               rpt.Text61.SetText ""
489           End If
490           rpt.txtBankNama1.SetText noVirtualaAcc(0)
491           rpt.txtBankNama2.SetText noVirtualaAcc(1)
492           rpt.txtBankNama3.SetText noVirtualaAcc(2)
493           rpt.txtBankNama4.SetText noVirtualaAcc(3)
              
494           rpt.txtBankNo1.SetText namaBank(0)
495           rpt.txtBankNo2.SetText namaBank(1)
496           rpt.txtBankNo3.SetText namaBank(2)
497           rpt.txtBankNo4.SetText namaBank(3)
              
498           rpt.txtNama1.SetText namaAkun(0)
499           rpt.txtNama2.SetText namaAkun(1)
500           rpt.txtNama3.SetText namaAkun(2)
501           rpt.txtNama4.SetText namaAkun(3)
              
502           rpt.txtInfoMaterai.SetText ""
503           rpt.txtLunasMaterai.SetText ""
              
504           If Log.fakturMaterai = 1 Then
505               Set rsTemp = objO.FetchData("SELECT * FROM t_invoiceMaterai WITH (NOLOCK) WHERE NOINV = '" & noInv & "'")
506               If Not rsTemp.EOF Then
507                   rpt.txtInfoMaterai.SetText "Ijin Pembubuhan Materai No. " & NulSID(rsTemp!NOIJIN)
508                   rpt.txtLunasMaterai.SetText "BEA MATERAI LUNAS " & NulSID(rsTemp!nilai)
509               End If
510           End If
511           rsCetakNewBilling.MoveFirst
              
512           If isSNL Then
513               rpt.Text63.SetText Format(SpyRound((rsCetakNewBilling!dpptotal * 11) / 12), "#,###.00")
514               rpt.Text64.SetText Format(rsCetakNewBilling!dpptotal - SpyRound((rsCetakNewBilling!dpptotal * 11) / 12), "#,###.00")
515           End If
              
516           rpt.Text57.SetText NulSID(objO.FetchData_String("SELECT FS_KA_BENGKEL FROM T_PROFIL"))
517           rpt.Text8.SetText NulSID(objO.FetchData_String("SELECT FS_KOTA FROM T_PROFIL"))
              
518           rpt.txtMaterai.SetText IIf(nilaiMAT = 0, "0,00", Format(NulSID(nilaiMAT), "#,###.00")) ''Format(NulSID(nilaiMAT), "#,###.00")
519           rpt.txtTglBilling.SetText tglInv
520           rpt.Text41.SetText "PPn " '"PPn " & nilaiPPnInv & " %"
521           rpt.Text55.SetText "PPn " & nilaiPPnInv & " %"
522           rpt.database.SetDataSource rsCetakNewBilling
523           If No = 1 Then
524               With CRViewer1
525                   .ReportSource = rpt
526                   .ViewReport
527                   .Zoom 100
528                   .Top = 0
529                   .Left = asxPowerBanner1.Left + asxPowerBanner1.Width + 20
530                   .Width = Frame2.Width + 100
531                   .Height = asxPowerBanner1.Height
                      
532                   cmdTutup.Top = .Top + 430
533                   cmdTutup.Left = .Left
534                   cmdTutup.Visible = True
535                   cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 5
536                   cmdSettingPrinter.Visible = True
537                   cmdSettingPrinter.Top = cmdTutup.Top
538                   .Visible = True
539               End With
540               DoEvents
541               cmdTutup.Visible = True
542           Else
543                With CRViewer2
544                   .ReportSource = rpt
545                   .ViewReport
546                   .Zoom 100
547                   .Top = 0
548                   .Left = asxPowerBanner1.Left + asxPowerBanner1.Width + 20
549                   .Width = Frame2.Width + 100
550                   .Height = asxPowerBanner1.Height
                      
551                   cmdTutup2.Top = .Top + 430
552                   cmdTutup2.Left = .Left
553                   cmdTutup2.Visible = True
554                   .Visible = True
555               End With
556               cmdTutup2.Visible = True
557           End If
558           Screen.MousePointer = vbDefault
559           Exit Sub
              
560       Else
561           MsgBox "Data invoice tidak ditemukan", vbInformation
562       End If
End Sub

Private Sub cetakInv2New(noInv As String, No As Integer)
          Dim rsRepo As ADODB.Recordset
          Dim rsAddJob As ADODB.Recordset
          Dim rsProfil As ADODB.Recordset
          Dim rsRinciSBE As ADODB.Recordset
          'Set rpt = New rtpNewInvoice2
          Dim hlp1 As String
          Dim hlp2 As String
          Dim hlp3 As String
          Dim hlp4 As String
          Dim hlp6 As String
          Dim gcat As String
          Dim infoBill As String
          Dim no_wo As String
          Dim nilaiMAT As Double
          Dim tglInv, kodeQR, kodeLgnCetak As String
          
          Dim onlySublet  As Boolean
          Dim isSBE       As Boolean
          Dim total As Double
          Dim isSNL As Boolean
          
1         kodeQR = objO.FetchData_String("select UnixID from [" & dbReplikasiVA & "]..QRIS_Invoice where KodeDealer = '" & Log.kodeCabang & "' and NoInv = 'E-" & noInv & "'")
2         If Log.dpplain = 1 Then
3             isSNL = objO.FetchData_Boolean("SELECT * FROM T_InvoicePPn WHERE NOINV = '" & noInv & "' ")
4         Else
5             isSNL = False
6         End If
          'isSNL = objO.FetchData_Boolean("SELECT * FROM T_InvoicePPn WHERE NOINV = '" & noInv & "' ")
7         If isSNL Then
8             Set rpt = New rtpNewInvoiceSNL2
9         Else
10            Set rpt = New rtpNewInvoice2
11        End If
          
12        If Log.notifInvNCM = 1 Then
13            rpt.boxNCM.Suppress = True
14            rpt.txtNCM1.Suppress = True
15            rpt.txtNCM2.Suppress = True
              
16            If isJualNRM = True Then
17                Set rsTemp = objO.FetchData("SELECT * FROM notifInvoice WITH (NOLOCK) WHERE tipe = 'NCM'")
18                If Not rsTemp.EOF Then
                  
19                    rpt.boxNCM.Suppress = False
20                    rpt.txtNCM1.Suppress = False
21                    rpt.txtNCM1.SetText Trim(rsTemp!Text1)
                      
22                    rpt.txtNCM2.Suppress = False
23                    rpt.txtNCM2.SetText Trim(rsTemp!Text2)
                   
24                End If
25            Else
26                Set rsTemp = objO.FetchData("SELECT * FROM notifInvoice WITH (NOLOCK) WHERE tipe = 'NonNCM'")
27                If Not rsTemp.EOF Then
                  
28                    rpt.boxNCM.Suppress = False
29                    rpt.txtNCM1.Suppress = False
30                    rpt.txtNCM1.SetText Trim(rsTemp!Text1)
                      
31                    rpt.txtNCM2.Suppress = False
32                    rpt.txtNCM2.SetText Trim(rsTemp!Text2)
33                End If
34            End If
35        End If
          
36        createTableNewBilling
          ''20220609 rev nama alamat dealer penjual diambil dari t_dealer
37        sSQL = "select a.NOINV, a.NO_WO,a.FS_KD_PLG,a.CUSTOMER,a.CUST_ADRES + ' - ' + a.KOTA CUST_ADRES,"
          ''sSQL = "select a.NOINV, a.NO_WO,a.FS_KD_PLG,isnull(b.FS_NM_TAGIH,a.CUSTOMER) as CUSTOMER,isnull(b.FS_ALAMAT,a.CUST_ADRES) as CUST_ADRES,"
38        sSQL = sSQL & " a.NO_POLISI,a.NO_CHASSIS,a.FMDL,a.NO_ENGINE,a.NO_WO,a.FRONT,A.STATUS,A.MECH,A.INVDATE,A.INVtime,"
39        sSQL = sSQL & " A.REP1,A.REP2,A.REP3,A.REP4,A.REP5,A.REP6,A.REP7,A.REP8,"
40        sSQL = sSQL & " A.RTE1,A.RTE2,A.RTE3,A.RTE4,A.RTE5,A.RTE6,A.RTE7,A.RTE8,"
41        sSQL = sSQL & " A.OFF1,A.OFF2,A.OFF3,A.OFF4,A.OFF5,A.OFF6,A.OFF7,A.OFF8,"
42        sSQL = sSQL & " A.DISC1,A.DISC2,A.DISC3,A.DISC4,A.DISC5,A.DISC6,A.DISC7,A.DISC8,"
43        sSQL = sSQL & " PART_G,PART_NG,MTRL,OIL_G,OIL_NG,MET,SBLT,LABOR,BODY,TOTAL-(FN_PAJAK+MET) AS DPP,FN_PAJAK,TOTAL,FS_RC"
44        sSQL = sSQL & " from T_INVOICE a with (nolock) "
          ''20220609 rev nama alamat dealer penjual diambil dari t_dealer
          ''sSQL = sSQL & " LEFT JOIN T_DEALER b WITH (NOLOCK) ON a.FS_KD_PLG = b.KDPLG and b.FS_LOG_DEL = ''"
45        sSQL = sSQL & " where NOINV = '" & noInv & "'"
46        sSQL = sSQL & " order by INVDATE desc"
47        Set rsRepo = objO.FetchData(sSQL)
          
48        sSQL = "SELECT A.* FROM HIST_JOBADD A WITH(NOLOCK) INNER JOIN T_INVOICE B WITH(NOLOCK) ON A.NO_WO = B.NO_WO AND A.NO_INV = B.NOINV AND A.NO_CHASSIS = B.NO_CHASSIS AND A.FS_LOG_DEL = B.FS_LOG_DEL  WHERE A.NO_WO = '" & Trim(text(20)) & "' AND A.NO_CHASSIS = '" & Trim(text(17)) & "' AND A.FS_LOG_DEL = '' AND B.NOINV LIKE '" & Trim(text(0)) & "%'"
49        Set rsAddJob = objO.FetchData(sSQL)
          
50        sSQL = "select FS_NM_PERUSH,FS_ALM_PERUSH,'No Telpon : ' + FS_TELP + ', Fax : ' + FS_FAX AS FS_TELP,'Email : ' + FS_EMAIL as FS_EMAIL,'NPWP : ' + FS_NPWP + ' Tanggal : ' as FS_NPWP,'PKP : ' + FS_PKP as FS_PKP,bankname1,bankname2,bankname3,bankno1,bankno2,bankno3 from t_profil"
51        Set rsProfil = objO.FetchData(sSQL)
52        If rsProfil.EOF Then
53            MsgBox "Gagal memuat data profil dealer", vbInformation
54            Exit Sub
55        End If
56        If Not rsRepo.EOF Then
57            rsRepo.MoveFirst
58            kodeLgnCetak = rsRepo!FS_KD_PLG
59            nilaiMAT = NulSID(rsRepo!met)
60            isSBE = False
61            no_wo = NulSID(rsRepo!no_wo)
62            tglInv = TglDDMMMMYYYY(rsRepo!InvDate, 1) & " " & rsRepo!INVTIME
63            If NulSID(rsRepo!FS_RC) <> "" Then
64                isSBE = True
65            End If
66            onlySublet = True
67            Do While Not rsRepo.EOF
68                For i = 1 To 8
69                    hlp1 = "REP" & i
70                    hlp2 = "RTE" & i
71                    hlp3 = "OFF" & i
72                    hlp4 = "DISC" & i
                      
73                    If Trim(rsRepo(hlp1)) <> "" Then
74                        With rsCetakNewBilling
75                            .AddNew
76                            !namaperusahaan = NulSID(rsProfil!FS_NM_PERUSH)
77                            !alamatperusahaan = NulSID(rsProfil!FS_ALM_PERUSH)
78                            !teleponperusahaan = NulSID(rsProfil!FS_TELP)
79                            !emailperusahaan = NulSID(rsProfil!FS_EMAIL)
80                            !alamatperusahaanpajak = NulSID("")
81                            !npwpperusahaan = NulSID(rsProfil!FS_NPWP)
82                            !pkpperusahaan = NulSID(rsProfil!FS_PKP)
83                            !noInv = NulSID(rsRepo!noInv)
84                            !kodeLgn = objO.FetchData_String("SELECT idcust FROM [SINGLEIDORI].dbo.bridging WITH (NOLOCK) WHERE kodelgn = '" & NulSID(rsRepo!FS_KD_PLG) & "' AND kodedealer = '" & Right(gDb, 2) & "' ")
85                            !customer = NulSID(rsRepo!customer)
86                            !customeralamat = NulSID(rsRepo!cust_adres)
87                            !NoPolisi = NulSID(rsRepo!NO_POLISI)
88                            !model = NulSID(rsRepo!Fmdl)
89                            !noRangka = NulSID(rsRepo!NO_CHASSIS)
90                            !noMesin = NulSID(rsRepo!No_Engine)
91                            !noOrder = NulSID(rsRepo!no_wo)
92                            !namaSA = objO.FetchData_String("select FS_NM_PETUGAS from T_PETUGAS with (nolock) where FS_KD_PETUGAS = '" & NulSID(rsRepo!FRONT) & "'")
93                            !tipecustomer = NulSID(rsRepo!Status)
94                            !mekanik = NulSID(rsRepo!Mech)
95                            !billing = NulSID(rsRepo!FRONT) ''NulSID(rsRepo!InvDate) ''Format(NulSID(rsRepo!InvDate), "dd-mm-yyyy")
96                            !job = Left(IIf(InStr(NulSID(rsRepo(hlp1)), ".000") > 0, " ", "") & NulSID(rsRepo(hlp1)), 50)
97                            !jobdetail = NulSID("")
98                            !jobRate = k2t(NulSID(rsRepo(hlp2)))
99                            !jobharga = NulSID(rsRepo(hlp3))
100                           !jobdisc = IIf(chCtkDisc.value = False, NulSID(rsRepo(hlp4) & " %"), "") 'NulSID(rsRepo(hlp4) & " %")
101                           !jobsubtotal = NulSID(rsRepo(hlp3) * ((100 - NulSID(rsRepo(hlp4))) / 100))
102                           !dppJasa = Round(NulA(rsRepo!labor) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
103                           !dpppart = Round((NulA(rsRepo!PART_G) + NulA(rsRepo!PART_NG) + NulA(rsRepo!MTRL) + NulA(rsRepo!OIL_G) + NulA(rsRepo!OIL_NG)) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
104                           !dppsublet = Round(NulA(rsRepo!sblt) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
105                           !dpptotal = NulA(!dppJasa + !dpppart + !dppsublet) ''NulA(rsRepo!dpp)
106                           !PPNtotal = NulA(rsRepo!fn_pajak)
107                           !grandTotal = NulA(!dpptotal + !PPNtotal + nilaiMAT)  ''NulA(rsRepo!total)
108                           !kodeQR = NulSID(kodeQR)
109                           .Update
110                       End With
111                       onlySublet = False
112                       If isSBE Then
                              Dim nilaisbe As Integer
113                           nilaisbe = CInt(Right(NulSID(rsRepo!FS_RC), 3))
114                           If InStr(NulSID(rsRepo(hlp1)), nilaisbe & ".000") > 0 Then
115                               Set rsRinciSBE = objO.FetchData("select rinciansb as detail from rinciansbe where tipesb = 'R" & Format(nilaisbe, String(3, "0")) & "' order by Status")
116                               If Not rsRinciSBE.EOF Then
117                                   rsRinciSBE.MoveFirst
118                                   Do While Not rsRinciSBE.EOF
119                                       With rsCetakNewBilling
120                                           .AddNew
121                                           !namaperusahaan = NulSID(rsProfil!FS_NM_PERUSH)
122                                           !alamatperusahaan = NulSID(rsProfil!FS_ALM_PERUSH)
123                                           !teleponperusahaan = NulSID(rsProfil!FS_TELP)
124                                           !emailperusahaan = NulSID(rsProfil!FS_EMAIL)
125                                           !alamatperusahaanpajak = NulSID("")
126                                           !npwpperusahaan = NulSID(rsProfil!FS_NPWP)
127                                           !pkpperusahaan = NulSID(rsProfil!FS_PKP)
128                                           !noInv = NulSID(rsRepo!noInv)
129                                           !kodeLgn = objO.FetchData_String("SELECT idcust FROM [SINGLEIDORI].dbo.bridging WITH (NOLOCK) WHERE kodelgn = '" & NulSID(rsRepo!FS_KD_PLG) & "' AND kodedealer = '" & Right(gDb, 2) & "' ")
130                                           !customer = NulSID(rsRepo!customer)
131                                           !customeralamat = NulSID(rsRepo!cust_adres)
132                                           !NoPolisi = NulSID(rsRepo!NO_POLISI)
133                                           !model = NulSID(rsRepo!Fmdl)
134                                           !noRangka = NulSID(rsRepo!NO_CHASSIS)
135                                           !noMesin = NulSID(rsRepo!No_Engine)
136                                           !noOrder = NulSID(rsRepo!no_wo)
137                                           !namaSA = objO.FetchData_String("select FS_NM_PETUGAS from T_PETUGAS with (nolock) where FS_KD_PETUGAS = '" & NulSID(rsRepo!FRONT) & "'")
138                                           !tipecustomer = NulSID(rsRepo!Status)
139                                           !mekanik = NulSID(rsRepo!Mech)
140                                           !billing = NulSID(rsRepo!FRONT) ''NulSID(rsRepo!InvDate) ''Format(NulSID(rsRepo!InvDate), "dd-mm-yyyy")
141                                           !job = Left(" " & NulSID(rsRepo(hlp1)), 50)
142                                           !jobdetail = NulSID(rsRinciSBE!Detail)
143                                           !jobRate = NulSID(rsRepo(hlp2))
144                                           !jobharga = NulSID(rsRepo(hlp3))
145                                           !jobdisc = IIf(chCtkDisc.value = False, NulSID(rsRepo(hlp4) & " %"), "") 'NulSID(rsRepo(hlp4) & " %")
146                                           !jobsubtotal = NulSID(rsRepo(hlp3) * ((100 - NulSID(rsRepo(hlp4))) / 100))
147                                           !dppJasa = Round(NulA(rsRepo!labor) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
148                                           !dpppart = Round((NulA(rsRepo!PART_G) + NulA(rsRepo!PART_NG) + NulA(rsRepo!MTRL) + NulA(rsRepo!OIL_G) + NulA(rsRepo!OIL_NG)) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
149                                           !dppsublet = Round(NulA(rsRepo!sblt) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
150                                           !dpptotal = NulA(!dppJasa + !dpppart + !dppsublet) ''NulA(rsRepo!dpp)
151                                           !PPNtotal = NulA(rsRepo!fn_pajak)
152                                           !grandTotal = NulA(!dpptotal + !PPNtotal + nilaiMAT)  ''NulA(rsRepo!total)
153                                           !kodeQR = NulSID(kodeQR)
154                                           .Update
155                                       End With
156                                       rsRinciSBE.MoveNext
157                                   Loop
158                               End If
159                           End If
160                       End If
161                   End If
162               Next
                  
163               If onlySublet Then
164                   With rsCetakNewBilling
165                       .AddNew
166                       !namaperusahaan = NulSID(rsProfil!FS_NM_PERUSH)
167                       !alamatperusahaan = NulSID(rsProfil!FS_ALM_PERUSH)
168                       !teleponperusahaan = NulSID(rsProfil!FS_TELP)
169                       !emailperusahaan = NulSID(rsProfil!FS_EMAIL)
170                       !alamatperusahaanpajak = NulSID("")
171                       !npwpperusahaan = NulSID(rsProfil!FS_NPWP)
172                       !pkpperusahaan = NulSID(rsProfil!FS_PKP)
173                       !noInv = NulSID(rsRepo!noInv)
174                       !kodeLgn = objO.FetchData_String("SELECT idcust FROM [SINGLEIDORI].dbo.bridging WITH (NOLOCK) WHERE kodelgn = '" & NulSID(rsRepo!FS_KD_PLG) & "' AND kodedealer = '" & Right(gDb, 2) & "' ")
175                       !customer = NulSID(rsRepo!customer)
176                       !customeralamat = NulSID(rsRepo!cust_adres)
177                       !NoPolisi = NulSID(rsRepo!NO_POLISI)
178                       !model = NulSID(rsRepo!Fmdl)
179                       !noRangka = NulSID(rsRepo!NO_CHASSIS)
180                       !noMesin = NulSID(rsRepo!No_Engine)
181                       !noOrder = NulSID(rsRepo!no_wo)
182                       !namaSA = objO.FetchData_String("select FS_NM_PETUGAS from T_PETUGAS with (nolock) where FS_KD_PETUGAS = '" & NulSID(rsRepo!FRONT) & "'")
183                       !tipecustomer = NulSID(rsRepo!Status)
184                       !mekanik = NulSID(rsRepo!Mech)
185                       !billing = NulSID(rsRepo!FRONT) ''NulSID(rsRepo!InvDate) ''Format(NulSID(rsRepo!InvDate), "dd-mm-yyyy")
186                       !job = NulSID("Order Sublet")
187                       !jobdetail = NulSID("")
188                       !jobRate = 0 'NulSID(rsRepo(hlp2))
189                       !jobharga = 0 'NulSID(rsRepo(hlp3))
190                       !jobdisc = 0 'NulSID(rsRepo(hlp4) & " %")
191                       !jobsubtotal = 0 ' NulSID(rsRepo(hlp3) * ((100 - NulSID(rsRepo(hlp4))) / 100))
192                       !dppJasa = Round(NulA(rsRepo!labor) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
193                       !dpppart = Round((NulA(rsRepo!PART_G) + NulA(rsRepo!PART_NG) + NulA(rsRepo!MTRL) + NulA(rsRepo!OIL_G) + NulA(rsRepo!OIL_NG)) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
194                       !dppsublet = Round(NulA(rsRepo!sblt) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
195                       !dpptotal = NulA(!dppJasa + !dpppart + !dppsublet) ''NulA(rsRepo!dpp)
196                       !PPNtotal = NulA(rsRepo!fn_pajak)
197                       !grandTotal = NulA(!dpptotal + !PPNtotal + nilaiMAT)  ''NulA(rsRepo!total)
198                       !kodeQR = NulSID(kodeQR)
199                       .Update
200                   End With
201               End If
202               rsRepo.MoveNext
203           Loop
              
204           If Not rsAddJob.EOF Then
205               rsAddJob.MoveFirst
206               rsRepo.MoveFirst
207               Do While Not rsAddJob.EOF
208                   With rsCetakNewBilling
209                       .AddNew
210                       !namaperusahaan = NulSID(rsProfil!FS_NM_PERUSH)
211                       !alamatperusahaan = NulSID(rsProfil!FS_ALM_PERUSH)
212                       !teleponperusahaan = NulSID(rsProfil!FS_TELP)
213                       !emailperusahaan = NulSID(rsProfil!FS_EMAIL)
214                       !alamatperusahaanpajak = NulSID("")
215                       !npwpperusahaan = NulSID(rsProfil!FS_NPWP)
216                       !pkpperusahaan = NulSID(rsProfil!FS_PKP)
217                       !noInv = NulSID(rsRepo!noInv)
218                       !kodeLgn = objO.FetchData_String("SELECT idcust FROM [SINGLEIDORI].dbo.bridging WITH (NOLOCK) WHERE kodelgn = '" & NulSID(rsRepo!FS_KD_PLG) & "' AND kodedealer = '" & Right(gDb, 2) & "' ")
219                       !customer = NulSID(rsRepo!customer)
220                       !customeralamat = NulSID(rsRepo!cust_adres)
221                       !NoPolisi = NulSID(rsRepo!NO_POLISI)
222                       !model = NulSID(rsRepo!Fmdl)
223                       !noRangka = NulSID(rsRepo!NO_CHASSIS)
224                       !noMesin = NulSID(rsRepo!No_Engine)
225                       !noOrder = NulSID(rsRepo!no_wo)
226                       !namaSA = objO.FetchData_String("select FS_NM_PETUGAS from T_PETUGAS with (nolock) where FS_KD_PETUGAS = '" & NulSID(rsRepo!FRONT) & "'")
227                       !tipecustomer = NulSID(rsRepo!Status)
228                       !mekanik = NulSID(rsRepo!Mech)
229                       !billing = NulSID(rsRepo!FRONT) ''NulSID(rsRepo!InvDate) ''Format(NulSID(rsRepo!InvDate), "dd-mm-yyyy")
230                       !job = Left("Add. " & NulSID(rsAddJob!rep), 50)
231                       !jobdetail = NulSID("")
232                       !jobRate = k2t(NulSID(rsAddJob!rte))
233                       !jobharga = NulSID(rsAddJob!offn)
234                       !jobdisc = IIf(chCtkDisc.value = False, IIf(NulSID(rsAddJob!FREEC) = 1, 100, IIf(NulSID(rsAddJob!disc) = 100, 0, NulSID(rsAddJob!disc))) & " %", "") 'IIf(NulSID(rsAddJob!FREEC) = 1, NulSID(rsAddJob!disc), 0) & " %"
235                       !jobsubtotal = Round(NulSID(rsAddJob!offn * ((100 - NulSID(IIf(NulSID(rsAddJob!FREEC) = 1, 100, IIf(NulSID(rsAddJob!disc) = 100, 0, NulSID(rsAddJob!disc))))) / 100))) ''Round(NulSID(rsAddJob!offn * ((100 - NulSID(IIf(NulSID(rsAddJob!FREEC) = 1, NulSID(rsAddJob!disc), 0))) / 100)))
236                       !dppJasa = Round(NulA(rsRepo!labor) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
237                       !dpppart = Round((NulA(rsRepo!PART_G) + NulA(rsRepo!PART_NG) + NulA(rsRepo!MTRL) + NulA(rsRepo!OIL_G) + NulA(rsRepo!OIL_NG)) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
238                       !dppsublet = Round(NulA(rsRepo!sblt) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
239                       !dpptotal = NulA(!dppJasa + !dpppart + !dppsublet) ''NulA(rsRepo!dpp)
240                       !PPNtotal = NulA(rsRepo!fn_pajak)
241                       !grandTotal = NulA(!dpptotal + !PPNtotal + nilaiMAT)  ''NulA(rsRepo!total)
242                       !kodeQR = NulSID(kodeQR)
243                       .Update
244                   End With
245                   If isSBE Then
                          Dim nilaisbe2 As Integer
246                       nilaisbe2 = CInt(Right(NulSID(rsRepo!FS_RC), 3))
247                       If InStr(NulSID(rsRepo(hlp1)), nilaisbe2 & ".000") > 0 Then
248                           Set rsRinciSBE = objO.FetchData("select rinciansb as detail from rinciansbe where tipesb = 'R" & Format(nilaisbe2, String(3, "0")) & "' order by Status")
249                           If Not rsRinciSBE.EOF Then
250                               rsRinciSBE.MoveFirst
251                               Do While Not rsRinciSBE.EOF
252                                   With rsCetakNewBilling
253                                       .AddNew
254                                       !namaperusahaan = NulSID(rsProfil!FS_NM_PERUSH)
255                                       !alamatperusahaan = NulSID(rsProfil!FS_ALM_PERUSH)
256                                       !teleponperusahaan = NulSID(rsProfil!FS_TELP)
257                                       !emailperusahaan = NulSID(rsProfil!FS_EMAIL)
258                                       !alamatperusahaanpajak = NulSID("")
259                                       !npwpperusahaan = NulSID(rsProfil!FS_NPWP)
260                                       !pkpperusahaan = NulSID(rsProfil!FS_PKP)
261                                       !noInv = NulSID(rsRepo!noInv)
262                                       !kodeLgn = objO.FetchData_String("SELECT idcust FROM [SINGLEIDORI].dbo.bridging WITH (NOLOCK) WHERE kodelgn = '" & NulSID(rsRepo!FS_KD_PLG) & "' AND kodedealer = '" & Right(gDb, 2) & "' ")
263                                       !customer = NulSID(rsRepo!customer)
264                                       !customeralamat = NulSID(rsRepo!cust_adres)
265                                       !NoPolisi = NulSID(rsRepo!NO_POLISI)
266                                       !model = NulSID(rsRepo!Fmdl)
267                                       !noRangka = NulSID(rsRepo!NO_CHASSIS)
268                                       !noMesin = NulSID(rsRepo!No_Engine)
269                                       !noOrder = NulSID(rsRepo!no_wo)
270                                       !namaSA = objO.FetchData_String("select FS_NM_PETUGAS from T_PETUGAS with (nolock) where FS_KD_PETUGAS = '" & NulSID(rsRepo!FRONT) & "'")
271                                       !tipecustomer = NulSID(rsRepo!Status)
272                                       !mekanik = NulSID(rsRepo!Mech)
273                                       !billing = NulSID(rsRepo!FRONT) ''NulSID(rsRepo!InvDate) ''Format(NulSID(rsRepo!InvDate), "dd-mm-yyyy")
274                                       !job = Left(" " & "Add. " & NulSID(rsAddJob!rep), 50)
275                                       !jobdetail = NulSID(rsRinciSBE!Detail)
276                                       !jobRate = NulSID(rsAddJob!rte)
277                                       !jobharga = NulSID(rsAddJob!offn)
278                                       !jobdisc = IIf(chCtkDisc.value = False, IIf(NulSID(rsAddJob!FREEC) = 1, NulSID(rsAddJob!disc), 0) & " %", "") 'IIf(NulSID(rsAddJob!FREEC) = 1, NulSID(rsAddJob!disc), 0) & " %"
279                                       !jobsubtotal = NulSID(rsAddJob!offn * ((100 - NulSID(IIf(NulSID(rsAddJob!FREEC) = 1, NulSID(rsAddJob!disc), 0))) / 100))
280                                       !dppJasa = Round(NulA(rsRepo!labor) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
281                                       !dpppart = Round((NulA(rsRepo!PART_G) + NulA(rsRepo!PART_NG) + NulA(rsRepo!MTRL) + NulA(rsRepo!OIL_G) + NulA(rsRepo!OIL_NG)) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
282                                       !dppsublet = Round(NulA(rsRepo!sblt) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
283                                       !dpptotal = NulA(!dppJasa + !dpppart + !dppsublet) ''NulA(rsRepo!dpp)
284                                       !PPNtotal = NulA(rsRepo!fn_pajak)
285                                       !grandTotal = NulA(!dpptotal + !PPNtotal + nilaiMAT)  ''NulA(rsRepo!total)
286                                       !kodeQR = NulSID(kodeQR)
287                                       .Update
288                                   End With
289                                   rsRinciSBE.MoveNext
290                               Loop
291                           End If
292                       End If
293                   End If
294                   rsAddJob.MoveNext
295               Loop
296           End If
297           gcat = ""
298           If nRC Then
                  Dim rsHlp As ADODB.Recordset
                  Dim nBulan As String
                  Dim nSbe As String
299               nSbe = ""
300               Set rsHlp = objO.FetchData("SELECT CASE WHEN JT_TGL < JT_BLN THEN JT_TGL ELSE JT_BLN END AS JT_TGL, SBENEXT FROM M_ESTIMASI WITH (NOLOCK) WHERE NORANGKA = '" & Trim(text(17)) & "' AND TGLSVCNOW = '" & Format(text(43), "MM/DD/YYYY") & "'")
301               If Not rsHlp.EOF Then
302                   nBulan = GetNamaBulan(Month(rsHlp("JT_TGL"))) & " " & year(rsHlp("JT_TGL"))
303                   nSbe = Trim(rsHlp("SBENEXT"))
304               End If
305               rsHlp.Close
306               If Trim(nSbe) <> "" Then
                      'gcat = IIf(nHeavy, "Kembali untuk service berkala setelah kelipatan 5.000 KM atau bulan " & nBulan, "Kembali untuk service berkala setelah kelipatan 10.000 KM atau bulan " & nBulan)
                      ''gcat = "Kembali untuk service berkala " & Format(CDbl(nSbe), "#,#0") & " KM atau bulan " & nBulan
307                   gcat = "Kembali untuk service berkala bulan " & nBulan & " atau " & Format(CDbl(nSbe), "#,#0") & " KM "
308                   gcat = gcat + IIf(gcat <> "", " (mana yang tercapai lebih dahulu).", "")
309               Else
310                   Call objO.InsertData("INSERT INTO ESTIMASISALAH (norangka, tglservice, kmsvc, sbesvc, kelipatan, noinv, nowo,customer, alamat,nopolisi, telpon, gsm,cdma) VALUES " & _
                                     "('" & Trim(text(17)) & "','" & Format(text(43), "MM/DD/YYYY") & "'," & CDbl(text(37)) & ",'" & Trim(text(33)) & "','" & IIf(nHeavy, 5000, 10000) & "', " & _
                                     "'" & Trim(text(20)) & "','" & Trim(Trim(text(0))) & "','" & Trim(text(4)) & "','" & Trim(txt5) & "','" & Trim(text(16)) & "','" & Trim(text(7)) & "','" & Trim(text(8)) & "','" & Trim(text(9)) & "')")
311               End If
312           End If
              
313           rpt.txtTerbilang.SetText "Terbilang (Rupiah) " & Trim(Angka2Huruf(rsCetakNewBilling!grandTotal, 1))
              ''20240829 revisi menutup kembali service
314           rpt.txtNextService.SetText "" 'gcat
              
              '' inspection ====================
315           If Log.CR7 = 1 Then
                  Dim rsInspection As ADODB.Recordset
                  Dim total1 As Double
                  Dim total2 As Double
                  Dim total3 As Double
                  Dim total4 As Double
                  
316               sSQL = "select sp1,sp2,sp3,sp4,estn1,estn2,estn3,estn4,estj1,estj2,estj3,estj4,estp1,estp2,estp3,estp4,ests1,ests2,ests3,ests4 from INSPECTION with (nolock) where NoWo = '" & no_wo & "' order by tglentry desc"
317               Set rsInspection = objO.FetchData(sSQL)
318               If Not rsInspection.EOF Then
319                   rsInspection.MoveFirst
320                   rpt.txtSaran1.SetText NulSID(rsInspection!sp1)
321                   rpt.txtSaran2.SetText NulSID(rsInspection!sp2)
322                   rpt.txtSaran3.SetText NulSID(rsInspection!sp3)
323                   rpt.txtSaran4.SetText NulSID(rsInspection!sp4)
                      
324                   If NulSID(rsInspection!sp1) <> "" Then
325                       rpt.txtWaktuSVC1.SetText Format(NulSID(rsInspection!estn1), "dd-mm-yyyy")
326                   Else
327                       rpt.txtWaktuSVC1.SetText ""
328                   End If
329                   If NulSID(rsInspection!sp2) <> "" Then
330                       rpt.txtWaktuSVC2.SetText Format(NulSID(rsInspection!estn2), "dd-mm-yyyy")
331                   Else
332                       rpt.txtWaktuSVC2.SetText ""
333                   End If
334                   If NulSID(rsInspection!sp3) <> "" Then
335                       rpt.txtWaktuSVC3.SetText Format(NulSID(rsInspection!estn3), "dd-mm-yyyy")
336                   Else
337                       rpt.txtWaktuSVC3.SetText ""
338                   End If
339                   If NulSID(rsInspection!sp4) <> "" Then
340                       rpt.txtWaktuSVC4.SetText Format(NulSID(rsInspection!estn4), "dd-mm-yyyy")
341                   Else
342                       rpt.txtWaktuSVC4.SetText ""
343                   End If
                                      
344                   If NulSID(rsInspection!sp1) <> "" Then
345                       rpt.txtEstJasa1.SetText Format(NulSID(rsInspection!estj1), "#,0")
346                   Else
347                       rpt.txtEstJasa1.SetText ""
348                   End If
349                   total1 = total1 + NulSID(rsInspection!estj1)
                      
350                   If NulSID(rsInspection!sp2) <> "" Then
351                       rpt.txtEstJasa2.SetText Format(NulSID(rsInspection!estj2), "#,0")
352                   Else
353                       rpt.txtEstJasa2.SetText ""
354                   End If
355                   total2 = total2 + NulSID(rsInspection!estj2)
                      
356                   If NulSID(rsInspection!sp3) <> "" Then
357                       rpt.txtEstJasa3.SetText Format(NulSID(rsInspection!estj3), "#,0")
358                   Else
359                       rpt.txtEstJasa3.SetText ""
360                   End If
361                   total3 = total3 + NulSID(rsInspection!estj3)
                      
362                   If NulSID(rsInspection!sp4) <> "" Then
363                       rpt.txtEstJasa4.SetText Format(NulSID(rsInspection!estj4), "#,0")
364                   Else
365                       rpt.txtEstJasa4.SetText ""
366                   End If
367                   total4 = total4 + NulSID(rsInspection!estj4)
                                                      
368                   If NulSID(rsInspection!sp1) <> "" Then
369                       rpt.txtEstPart1.SetText Format(NulSID(rsInspection!estp1), "#,0")
370                   Else
371                       rpt.txtEstPart1.SetText ""
372                   End If
373                   total1 = total1 + NulSID(rsInspection!estp1)
                                                      
374                   If NulSID(rsInspection!sp2) <> "" Then
375                       rpt.txtEstPart2.SetText Format(NulSID(rsInspection!estp2), "#,0")
376                   Else
377                       rpt.txtEstPart2.SetText ""
378                   End If
379                   total2 = total2 + NulSID(rsInspection!estp2)
                                                      
380                   If NulSID(rsInspection!sp3) <> "" Then
381                       rpt.txtEstPart3.SetText Format(NulSID(rsInspection!estp3), "#,0")
382                   Else
383                       rpt.txtEstPart3.SetText ""
384                   End If
385                   total3 = total3 + NulSID(rsInspection!estp3)
                                                      
386                   If NulSID(rsInspection!sp4) <> "" Then
387                       rpt.txtEstPart4.SetText Format(NulSID(rsInspection!estp4), "#,0")
388                   Else
389                       rpt.txtEstPart4.SetText ""
390                   End If
391                   total4 = total4 + NulSID(rsInspection!estp4)
                                      
                                                      
392                   If NulSID(rsInspection!sp1) <> "" Then
393                       rpt.txtEstSublet1.SetText Format(NulSID(rsInspection!ests1), "#,0")
394                   Else
395                       rpt.txtEstSublet1.SetText ""
396                   End If
397                   total1 = total1 + NulSID(rsInspection!ests1)
                                                      
398                   If NulSID(rsInspection!sp2) <> "" Then
399                       rpt.txtEstSublet2.SetText Format(NulSID(rsInspection!ests2), "#,0")
400                   Else
401                       rpt.txtEstSublet2.SetText ""
402                   End If
403                   total2 = total2 + NulSID(rsInspection!ests2)
                                                      
404                   If NulSID(rsInspection!sp3) <> "" Then
405                       rpt.txtEstSublet3.SetText Format(NulSID(rsInspection!ests3), "#,0")
406                   Else
407                       rpt.txtEstSublet3.SetText ""
408                   End If
409                   total3 = total3 + NulSID(rsInspection!ests3)
                                                      
410                   If NulSID(rsInspection!sp4) <> "" Then
411                       rpt.txtEstSublet4.SetText Format(NulSID(rsInspection!ests4), "#,0")
412                   Else
413                       rpt.txtEstSublet4.SetText ""
414                   End If
415                   total4 = total4 + NulSID(rsInspection!ests4)
                                      
416                   If NulSID(rsInspection!sp1) <> "" Then
417                       rpt.txtEstTotal1.SetText Format(NulSID(total1), "#,0")
418                   Else
419                       rpt.txtEstTotal1.SetText ""
420                   End If
421                   If NulSID(rsInspection!sp2) <> "" Then
422                       rpt.txtEstTotal2.SetText Format(NulSID(total2), "#,0")
423                   Else
424                       rpt.txtEstTotal2.SetText ""
425                   End If
426                   If NulSID(rsInspection!sp3) <> "" Then
427                       rpt.txtEstTotal3.SetText Format(NulSID(total3), "#,0")
428                   Else
429                       rpt.txtEstTotal3.SetText ""
430                   End If
431                   If NulSID(rsInspection!sp4) <> "" Then
432                       rpt.txtEstTotal4.SetText Format(NulSID(total4), "#,0")
433                   Else
434                       rpt.txtEstTotal4.SetText ""
435                   End If
                      
436               End If
437           End If
              '' akhir inspection ==============
                               
              Dim noVirtualaAcc(3), namaBank(3), namaAkun(3) As String
              Dim iVirtual As Integer
              Dim rsVirtual As ADODB.Recordset
438           iVirtual = 0
              
439           Set rsVirtual = objO.FetchData("SELECT virtualAccountNo,bank,virtualAccountName From " & dbReplikasiVA & "..DOKU_INQUIRY where  KodeDealer = '" & Log.kodeCabang & "' and virtualAccountNo like '%" & "0" & Right("000000000" & NulSID(text(20)), 8) & "' order by urut ASC")
440           If Not rsVirtual.EOF Then
441             rsVirtual.MoveFirst
442             Do While Not rsVirtual.EOF
443               noVirtualaAcc(iVirtual) = NulSID(rsVirtual!virtualAccountNo)
444               namaBank(iVirtual) = NulSID(rsVirtual!bank)
445               namaAkun(iVirtual) = NulSID(rsVirtual!virtualAccountName)
446               iVirtual = iVirtual + 1
447               rsVirtual.MoveNext
448             Loop
449           End If
450           If noVirtualaAcc(0) = "" Then
451               rpt.Text56.SetText ""
452               rpt.Text59.SetText ""
453               rpt.Text60.SetText ""
454               rpt.Text61.SetText ""
455           End If
456           rpt.txtBankNama1.SetText noVirtualaAcc(0)
457           rpt.txtBankNama2.SetText noVirtualaAcc(1)
458           rpt.txtBankNama3.SetText noVirtualaAcc(2)
459           rpt.txtBankNama4.SetText noVirtualaAcc(3)
              
460           rpt.txtBankNo1.SetText namaBank(0)
461           rpt.txtBankNo2.SetText namaBank(1)
462           rpt.txtBankNo3.SetText namaBank(2)
463           rpt.txtBankNo4.SetText namaBank(3)
              
464           rpt.txtNama1.SetText namaAkun(0)
465           rpt.txtNama2.SetText namaAkun(1)
466           rpt.txtNama3.SetText namaAkun(2)
467           rpt.txtNama4.SetText namaAkun(3)
468           rpt.txtInfoMaterai.SetText ""
469           rpt.txtLunasMaterai.SetText ""
              
470           If Log.fakturMaterai = 1 Then
471               Set rsTemp = objO.FetchData("SELECT * FROM t_invoiceMaterai WITH (NOLOCK) WHERE NOINV = '" & noInv & "'")
472               If Not rsTemp.EOF Then
473                   rpt.txtInfoMaterai.SetText "Ijin Pembubuhan Materai No. " & NulSID(rsTemp!NOIJIN)
474                   rpt.txtLunasMaterai.SetText "BEA MATERAI LUNAS " & NulSID(rsTemp!nilai)
475               End If
476           End If
477           rsCetakNewBilling.MoveFirst
478           If isSNL Then
479               rpt.Text63.SetText Format(SpyRound((rsCetakNewBilling!dpptotal * 11) / 12), "#,###.00")
480               rpt.Text64.SetText Format(rsCetakNewBilling!dpptotal - SpyRound((rsCetakNewBilling!dpptotal * 11) / 12), "#,###.00")
481           End If
              
482           rpt.Text57.SetText NulSID(objO.FetchData_String("SELECT FS_KA_BENGKEL FROM T_PROFIL"))
483           rpt.Text8.SetText NulSID(objO.FetchData_String("SELECT FS_KOTA FROM T_PROFIL"))
484           rpt.txtMaterai.SetText IIf(nilaiMAT = 0, "0,00", Format(NulSID(nilaiMAT), "#,###.00"))
485           rpt.txtTglBilling.SetText tglInv
486           rpt.Text41.SetText "PPn " '"PPn " & nilaiPPnInv & " %"
487           rpt.Text55.SetText "PPn " & nilaiPPnInv & " %"
488           rpt.database.SetDataSource rsCetakNewBilling
489           If No = 1 Then
490               With CRViewer1
491                   .ReportSource = rpt
492                   .ViewReport
493                   .Zoom 100
494                   .Top = 0
495                   .Left = asxPowerBanner1.Left + asxPowerBanner1.Width + 20
496                   .Width = Frame2.Width + 100
497                   .Height = asxPowerBanner1.Height
                      
498                   cmdTutup.Top = .Top + 430
499                   cmdTutup.Left = .Left
500                   cmdTutup.Visible = True
501                   .Visible = True
502               End With
503               cmdTutup.Visible = True
504           Else
505                With CRViewer2
506                   .ReportSource = rpt
507                   .ViewReport
508                   .Zoom 100
509                   .Top = 0
510                   .Left = asxPowerBanner1.Left + asxPowerBanner1.Width + 20
511                   .Width = Frame2.Width + 100
512                   .Height = asxPowerBanner1.Height
                      
513                   cmdTutup2.Top = .Top + 430
514                   cmdTutup2.Left = .Left
515                   cmdTutup2.Visible = True
516                   .Visible = True
517               End With
518               DoEvents
519               cmdTutup2.Visible = True
520           End If
521           Screen.MousePointer = vbDefault
522           Exit Sub
              
523       Else
524           MsgBox "Data invoice tidak ditemukan", vbInformation
525       End If
End Sub


Private Sub cetakInv3New(noInv As String, No As Integer)
          Dim rsRepo As ADODB.Recordset
          Dim rsAddJob As ADODB.Recordset
          Dim rsProfil As ADODB.Recordset
          Dim rsRinciSBE As ADODB.Recordset
          Dim rpt As New rtpNewInvoice3
          Dim hlp1 As String
          Dim hlp2 As String
          Dim hlp3 As String
          Dim hlp4 As String
          Dim hlp6 As String
          Dim gcat As String
          Dim infoBill As String
          Dim no_wo As String
          Dim nilaiMAT As Double
          Dim tglInv  As String
          
          Dim onlySublet  As Boolean
          Dim isSBE       As Boolean
          Dim total As Double
          
1         If Log.notifInvNCM = 1 Then
2             rpt.boxNCM.Suppress = True
3             rpt.txtNCM1.Suppress = True
4             rpt.txtNCM2.Suppress = True
              
5             If isJualNRM = True Then
6                 Set rsTemp = objO.FetchData("SELECT * FROM notifInvoice WITH (NOLOCK) WHERE tipe = 'NCM'")
7                 If Not rsTemp.EOF Then
                  
8                     rpt.boxNCM.Suppress = False
9                     rpt.txtNCM1.Suppress = False
10                    rpt.txtNCM1.SetText Trim(rsTemp!Text1)
                      
11                    rpt.txtNCM2.Suppress = False
12                    rpt.txtNCM2.SetText Trim(rsTemp!Text2)
                   
13                End If
14            Else
15                Set rsTemp = objO.FetchData("SELECT * FROM notifInvoice WITH (NOLOCK) WHERE tipe = 'NonNCM'")
16                If Not rsTemp.EOF Then
                  
17                    rpt.boxNCM.Suppress = False
18                    rpt.txtNCM1.Suppress = False
19                    rpt.txtNCM1.SetText Trim(rsTemp!Text1)
                      
20                    rpt.txtNCM2.Suppress = False
21                    rpt.txtNCM2.SetText Trim(rsTemp!Text2)
22                End If
23            End If
24        End If
          
25        createTableNewBilling
          ''20220609 rev nama alamat dealer penjual diambil dari t_dealer
26        sSQL = "select a.NOINV, a.NO_WO,a.FS_KD_PLG,a.CUSTOMER,a.CUST_ADRES,"
          ''sSQL = "select a.NOINV, a.NO_WO,a.FS_KD_PLG,isnull(b.FS_NM_TAGIH,a.CUSTOMER) as CUSTOMER,isnull(b.FS_ALAMAT,a.CUST_ADRES) as CUST_ADRES,"
          
27        sSQL = sSQL & " a.NO_POLISI,a.NO_CHASSIS,a.FMDL,a.NO_ENGINE,a.NO_WO,a.FRONT,A.STATUS,A.MECH,A.INVDATE,A.INVtime,"
28        sSQL = sSQL & " A.REP1,A.REP2,A.REP3,A.REP4,A.REP5,A.REP6,A.REP7,A.REP8,"
29        sSQL = sSQL & " A.RTE1,A.RTE2,A.RTE3,A.RTE4,A.RTE5,A.RTE6,A.RTE7,A.RTE8,"
30        sSQL = sSQL & " A.OFF1,A.OFF2,A.OFF3,A.OFF4,A.OFF5,A.OFF6,A.OFF7,A.OFF8,"
31        sSQL = sSQL & " A.DISC1,A.DISC2,A.DISC3,A.DISC4,A.DISC5,A.DISC6,A.DISC7,A.DISC8,"
32        sSQL = sSQL & " PART_G,PART_NG,MTRL,OIL_G,OIL_NG,MET,SBLT,LABOR,BODY,TOTAL-(FN_PAJAK+MET) AS DPP,FN_PAJAK,TOTAL,FS_RC"
33        sSQL = sSQL & " from T_INVOICE a with (nolock) "
          
          ''20220609 rev nama alamat dealer penjual diambil dari t_dealer
          ''sSQL = sSQL & " LEFT JOIN T_DEALER b WITH (NOLOCK) ON a.FS_KD_PLG = b.KDPLG and b.FS_LOG_DEL = ''"
          
34        sSQL = sSQL & " where NOINV = '" & noInv & "'"
35        sSQL = sSQL & " order by INVDATE desc"
36        Set rsRepo = objO.FetchData(sSQL)
          
37        sSQL = "SELECT A.* FROM HIST_JOBADD A WITH(NOLOCK) INNER JOIN T_INVOICE B WITH(NOLOCK) ON A.NO_WO = B.NO_WO AND A.NO_INV = B.NOINV AND A.NO_CHASSIS = B.NO_CHASSIS AND A.FS_LOG_DEL = B.FS_LOG_DEL  WHERE A.NO_WO = '" & Trim(text(20)) & "' AND A.NO_CHASSIS = '" & Trim(text(17)) & "' AND A.FS_LOG_DEL = '' AND B.NOINV LIKE '" & Trim(text(0)) & "%'"
38        Set rsAddJob = objO.FetchData(sSQL)
          
39        sSQL = "select FS_NM_PERUSH,FS_ALM_PERUSH,'No Telpon : ' + FS_TELP + ', Fax : ' + FS_FAX AS FS_TELP,'Email : ' + FS_EMAIL as FS_EMAIL,'NPWP : ' + FS_NPWP + ' Tanggal : ' as FS_NPWP,'PKP : ' + FS_PKP as FS_PKP,bankname1,bankname2,bankname3,bankno1,bankno2,bankno3 from t_profil"
40        Set rsProfil = objO.FetchData(sSQL)
41        If rsProfil.EOF Then
42            MsgBox "Gagal memuat data profil dealer", vbInformation
43            Exit Sub
44        End If
45        If Not rsRepo.EOF Then
46            rsRepo.MoveFirst
47            nilaiMAT = NulSID(rsRepo!met)
48            isSBE = False
49            no_wo = NulSID(rsRepo!no_wo)
50            tglInv = TglDDMMMMYYYY(rsRepo!InvDate, 1) & " " & rsRepo!INVTIME
51            If NulSID(rsRepo!FS_RC) <> "" Then
52                isSBE = True
53            End If
54            onlySublet = True
55            Do While Not rsRepo.EOF
56                For i = 1 To 8
57                    hlp1 = "REP" & i
58                    hlp2 = "RTE" & i
59                    hlp3 = "OFF" & i
60                    hlp4 = "DISC" & i
                      
61                    If Trim(rsRepo(hlp1)) <> "" Then
62                        With rsCetakNewBilling
63                            .AddNew
64                            !namaperusahaan = NulSID(rsProfil!FS_NM_PERUSH)
65                            !alamatperusahaan = NulSID(rsProfil!FS_ALM_PERUSH)
66                            !teleponperusahaan = NulSID(rsProfil!FS_TELP)
67                            !emailperusahaan = NulSID(rsProfil!FS_EMAIL)
68                            !alamatperusahaanpajak = NulSID("")
69                            !npwpperusahaan = NulSID(rsProfil!FS_NPWP)
70                            !pkpperusahaan = NulSID(rsProfil!FS_PKP)
71                            !noInv = NulSID(rsRepo!noInv)
72                            !kodeLgn = objO.FetchData_String("SELECT idcust FROM [SINGLEIDORI].dbo.bridging WITH (NOLOCK) WHERE kodelgn = '" & NulSID(rsRepo!FS_KD_PLG) & "' AND kodedealer = '" & Right(gDb, 2) & "' ")
73                            !customer = NulSID(rsRepo!customer)
74                            !customeralamat = NulSID(rsRepo!cust_adres)
75                            !NoPolisi = NulSID(rsRepo!NO_POLISI)
76                            !model = NulSID(rsRepo!Fmdl)
77                            !noRangka = NulSID(rsRepo!NO_CHASSIS)
78                            !noMesin = NulSID(rsRepo!No_Engine)
79                            !noOrder = NulSID(rsRepo!no_wo)
80                            !namaSA = objO.FetchData_String("select FS_NM_PETUGAS from T_PETUGAS with (nolock) where FS_KD_PETUGAS = '" & NulSID(rsRepo!FRONT) & "'")
81                            !tipecustomer = NulSID(rsRepo!Status)
82                            !mekanik = NulSID(rsRepo!Mech)
83                            !billing = NulSID(rsRepo!FRONT) ''NulSID(rsRepo!InvDate) ''Format(NulSID(rsRepo!InvDate), "dd-mm-yyyy")
84                            !job = Left(IIf(InStr(NulSID(rsRepo(hlp1)), ".000") > 0, " ", "") & NulSID(rsRepo(hlp1)), 50)
85                            !jobdetail = NulSID("")
86                            !jobRate = k2t(NulSID(rsRepo(hlp2)))
87                            !jobharga = NulSID(rsRepo(hlp3))
88                            !jobdisc = IIf(chCtkDisc.value = False, NulSID(rsRepo(hlp4) & " %"), "") 'NulSID(rsRepo(hlp4) & " %")
89                            !jobsubtotal = NulSID(rsRepo(hlp3) * ((100 - NulSID(rsRepo(hlp4))) / 100))
90                            !dppJasa = Round(NulA(rsRepo!labor) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
91                            !dpppart = Round((NulA(rsRepo!PART_G) + NulA(rsRepo!PART_NG) + NulA(rsRepo!MTRL) + NulA(rsRepo!OIL_G) + NulA(rsRepo!OIL_NG)) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
92                            !dppsublet = Round(NulA(rsRepo!sblt) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
93                            !dpptotal = NulA(!dppJasa + !dpppart + !dppsublet) ''NulA(rsRepo!dpp)
94                            !PPNtotal = NulA(rsRepo!fn_pajak)
95                            !grandTotal = NulA(!dpptotal + !PPNtotal + nilaiMAT)  ''NulA(rsRepo!total)
96                            .Update
97                        End With
98                        onlySublet = False
99                        If isSBE Then
                              Dim nilaisbe As Integer
100                           nilaisbe = CInt(Right(NulSID(rsRepo!FS_RC), 3))
101                           If InStr(NulSID(rsRepo(hlp1)), nilaisbe & ".000") > 0 Then
102                               Set rsRinciSBE = objO.FetchData("select rinciansb as detail from rinciansbe where tipesb = 'R" & Format(nilaisbe, String(3, "0")) & "' order by Status")
103                               If Not rsRinciSBE.EOF Then
104                                   rsRinciSBE.MoveFirst
105                                   Do While Not rsRinciSBE.EOF
106                                       With rsCetakNewBilling
107                                           .AddNew
108                                           !namaperusahaan = NulSID(rsProfil!FS_NM_PERUSH)
109                                           !alamatperusahaan = NulSID(rsProfil!FS_ALM_PERUSH)
110                                           !teleponperusahaan = NulSID(rsProfil!FS_TELP)
111                                           !emailperusahaan = NulSID(rsProfil!FS_EMAIL)
112                                           !alamatperusahaanpajak = NulSID("")
113                                           !npwpperusahaan = NulSID(rsProfil!FS_NPWP)
114                                           !pkpperusahaan = NulSID(rsProfil!FS_PKP)
115                                           !noInv = NulSID(rsRepo!noInv)
116                                           !kodeLgn = objO.FetchData_String("SELECT idcust FROM [SINGLEIDORI].dbo.bridging WITH (NOLOCK) WHERE kodelgn = '" & NulSID(rsRepo!FS_KD_PLG) & "' AND kodedealer = '" & Right(gDb, 2) & "' ")
117                                           !customer = NulSID(rsRepo!customer)
118                                           !customeralamat = NulSID(rsRepo!cust_adres)
119                                           !NoPolisi = NulSID(rsRepo!NO_POLISI)
120                                           !model = NulSID(rsRepo!Fmdl)
121                                           !noRangka = NulSID(rsRepo!NO_CHASSIS)
122                                           !noMesin = NulSID(rsRepo!No_Engine)
123                                           !noOrder = NulSID(rsRepo!no_wo)
124                                           !namaSA = objO.FetchData_String("select FS_NM_PETUGAS from T_PETUGAS with (nolock) where FS_KD_PETUGAS = '" & NulSID(rsRepo!FRONT) & "'")
125                                           !tipecustomer = NulSID(rsRepo!Status)
126                                           !mekanik = NulSID(rsRepo!Mech)
127                                           !billing = NulSID(rsRepo!FRONT) ''NulSID(rsRepo!InvDate) ''Format(NulSID(rsRepo!InvDate), "dd-mm-yyyy")
128                                           !job = Left(" " & NulSID(rsRepo(hlp1)), 50)
129                                           !jobdetail = NulSID(rsRinciSBE!Detail)
130                                           !jobRate = NulSID(rsRepo(hlp2))
131                                           !jobharga = NulSID(rsRepo(hlp3))
132                                           !jobdisc = IIf(chCtkDisc.value = False, NulSID(rsRepo(hlp4) & " %"), "") 'NulSID(rsRepo(hlp4) & " %")
133                                           !jobsubtotal = NulSID(rsRepo(hlp3) * ((100 - NulSID(rsRepo(hlp4))) / 100))
134                                           !dppJasa = Round(NulA(rsRepo!labor) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
135                                           !dpppart = Round((NulA(rsRepo!PART_G) + NulA(rsRepo!PART_NG) + NulA(rsRepo!MTRL) + NulA(rsRepo!OIL_G) + NulA(rsRepo!OIL_NG)) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
136                                           !dppsublet = Round(NulA(rsRepo!sblt) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
137                                           !dpptotal = NulA(!dppJasa + !dpppart + !dppsublet) ''NulA(rsRepo!dpp)
138                                           !PPNtotal = NulA(rsRepo!fn_pajak)
139                                           !grandTotal = NulA(!dpptotal + !PPNtotal + nilaiMAT)  ''NulA(rsRepo!total)
140                                           .Update
141                                       End With
142                                       rsRinciSBE.MoveNext
143                                   Loop
144                               End If
145                           End If
146                       End If
147                   End If
148               Next
                  
149               If onlySublet Then
150                   With rsCetakNewBilling
151                       .AddNew
152                       !namaperusahaan = NulSID(rsProfil!FS_NM_PERUSH)
153                       !alamatperusahaan = NulSID(rsProfil!FS_ALM_PERUSH)
154                       !teleponperusahaan = NulSID(rsProfil!FS_TELP)
155                       !emailperusahaan = NulSID(rsProfil!FS_EMAIL)
156                       !alamatperusahaanpajak = NulSID("")
157                       !npwpperusahaan = NulSID(rsProfil!FS_NPWP)
158                       !pkpperusahaan = NulSID(rsProfil!FS_PKP)
159                       !noInv = NulSID(rsRepo!noInv)
160                       !kodeLgn = objO.FetchData_String("SELECT idcust FROM [SINGLEIDORI].dbo.bridging WITH (NOLOCK) WHERE kodelgn = '" & NulSID(rsRepo!FS_KD_PLG) & "' AND kodedealer = '" & Right(gDb, 2) & "' ")
161                       !customer = NulSID(rsRepo!customer)
162                       !customeralamat = NulSID(rsRepo!cust_adres)
163                       !NoPolisi = NulSID(rsRepo!NO_POLISI)
164                       !model = NulSID(rsRepo!Fmdl)
165                       !noRangka = NulSID(rsRepo!NO_CHASSIS)
166                       !noMesin = NulSID(rsRepo!No_Engine)
167                       !noOrder = NulSID(rsRepo!no_wo)
168                       !namaSA = objO.FetchData_String("select FS_NM_PETUGAS from T_PETUGAS with (nolock) where FS_KD_PETUGAS = '" & NulSID(rsRepo!FRONT) & "'")
169                       !tipecustomer = NulSID(rsRepo!Status)
170                       !mekanik = NulSID(rsRepo!Mech)
171                       !billing = NulSID(rsRepo!FRONT) ''NulSID(rsRepo!InvDate) ''Format(NulSID(rsRepo!InvDate), "dd-mm-yyyy")
172                       !job = NulSID("Order Sublet")
173                       !jobdetail = NulSID("")
174                       !jobRate = 0 'NulSID(rsRepo(hlp2))
175                       !jobharga = 0 'NulSID(rsRepo(hlp3))
176                       !jobdisc = 0 'NulSID(rsRepo(hlp4) & " %")
177                       !jobsubtotal = 0 ' NulSID(rsRepo(hlp3) * ((100 - NulSID(rsRepo(hlp4))) / 100))
178                       !dppJasa = Round(NulA(rsRepo!labor) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
179                       !dpppart = Round((NulA(rsRepo!PART_G) + NulA(rsRepo!PART_NG) + NulA(rsRepo!MTRL) + NulA(rsRepo!OIL_G) + NulA(rsRepo!OIL_NG)) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
180                       !dppsublet = Round(NulA(rsRepo!sblt) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
181                       !dpptotal = NulA(!dppJasa + !dpppart + !dppsublet) ''NulA(rsRepo!dpp)
182                       !PPNtotal = NulA(rsRepo!fn_pajak)
183                       !grandTotal = NulA(!dpptotal + !PPNtotal + nilaiMAT)  ''NulA(rsRepo!total)
184                       .Update
185                   End With
186               End If
187               rsRepo.MoveNext
188           Loop
              
189           If Not rsAddJob.EOF Then
190               rsAddJob.MoveFirst
191               rsRepo.MoveFirst
192               Do While Not rsAddJob.EOF
193                   With rsCetakNewBilling
194                       .AddNew
195                       !namaperusahaan = NulSID(rsProfil!FS_NM_PERUSH)
196                       !alamatperusahaan = NulSID(rsProfil!FS_ALM_PERUSH)
197                       !teleponperusahaan = NulSID(rsProfil!FS_TELP)
198                       !emailperusahaan = NulSID(rsProfil!FS_EMAIL)
199                       !alamatperusahaanpajak = NulSID("")
200                       !npwpperusahaan = NulSID(rsProfil!FS_NPWP)
201                       !pkpperusahaan = NulSID(rsProfil!FS_PKP)
202                       !noInv = NulSID(rsRepo!noInv)
203                       !kodeLgn = objO.FetchData_String("SELECT idcust FROM [SINGLEIDORI].dbo.bridging WITH (NOLOCK) WHERE kodelgn = '" & NulSID(rsRepo!FS_KD_PLG) & "' AND kodedealer = '" & Right(gDb, 2) & "' ")
204                       !customer = NulSID(rsRepo!customer)
205                       !customeralamat = NulSID(rsRepo!cust_adres)
206                       !NoPolisi = NulSID(rsRepo!NO_POLISI)
207                       !model = NulSID(rsRepo!Fmdl)
208                       !noRangka = NulSID(rsRepo!NO_CHASSIS)
209                       !noMesin = NulSID(rsRepo!No_Engine)
210                       !noOrder = NulSID(rsRepo!no_wo)
211                       !namaSA = objO.FetchData_String("select FS_NM_PETUGAS from T_PETUGAS with (nolock) where FS_KD_PETUGAS = '" & NulSID(rsRepo!FRONT) & "'")
212                       !tipecustomer = NulSID(rsRepo!Status)
213                       !mekanik = NulSID(rsRepo!Mech)
214                       !billing = NulSID(rsRepo!FRONT) ''NulSID(rsRepo!InvDate) ''Format(NulSID(rsRepo!InvDate), "dd-mm-yyyy")
215                       !job = Left("Add. " & NulSID(rsAddJob!rep), 50)
216                       !jobdetail = NulSID("")
217                       !jobRate = k2t(NulSID(rsAddJob!rte))
218                       !jobharga = NulSID(rsAddJob!offn)
219                       !jobdisc = 100 'IIf(chCtkDisc.value = False, IIf(NulSID(rsAddJob!freec) = 1, 100, NulSID(rsAddJob!disc)) & " %", "") 'IIf(NulSID(rsAddJob!FREEC) = 1, NulSID(rsAddJob!disc), 0) & " %"
220                       !jobsubtotal = 0 'Round(NulSID(rsAddJob!offn * ((100 - NulSID(IIf(NulSID(rsAddJob!freec) = 1, NulSID(rsAddJob!disc), 0))) / 100)))
221                       !dppJasa = Round(NulA(rsRepo!labor) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
222                       !dpppart = Round((NulA(rsRepo!PART_G) + NulA(rsRepo!PART_NG) + NulA(rsRepo!MTRL) + NulA(rsRepo!OIL_G) + NulA(rsRepo!OIL_NG)) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
223                       !dppsublet = Round(NulA(rsRepo!sblt) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
224                       !dpptotal = NulA(!dppJasa + !dpppart + !dppsublet) ''NulA(rsRepo!dpp)
225                       !PPNtotal = NulA(rsRepo!fn_pajak)
226                       !grandTotal = NulA(!dpptotal + !PPNtotal + nilaiMAT)  ''NulA(rsRepo!total)
227                       .Update
228                   End With
229                   If isSBE Then
                          Dim nilaisbe2 As Integer
230                       nilaisbe2 = CInt(Right(NulSID(rsRepo!FS_RC), 3))
231                       If InStr(NulSID(rsRepo(hlp1)), nilaisbe2 & ".000") > 0 Then
232                           Set rsRinciSBE = objO.FetchData("select rinciansb as detail from rinciansbe where tipesb = 'R" & Format(nilaisbe2, String(3, "0")) & "' order by Status")
233                           If Not rsRinciSBE.EOF Then
234                               rsRinciSBE.MoveFirst
235                               Do While Not rsRinciSBE.EOF
236                                   With rsCetakNewBilling
237                                       .AddNew
238                                       !namaperusahaan = NulSID(rsProfil!FS_NM_PERUSH)
239                                       !alamatperusahaan = NulSID(rsProfil!FS_ALM_PERUSH)
240                                       !teleponperusahaan = NulSID(rsProfil!FS_TELP)
241                                       !emailperusahaan = NulSID(rsProfil!FS_EMAIL)
242                                       !alamatperusahaanpajak = NulSID("")
243                                       !npwpperusahaan = NulSID(rsProfil!FS_NPWP)
244                                       !pkpperusahaan = NulSID(rsProfil!FS_PKP)
245                                       !noInv = NulSID(rsRepo!noInv)
246                                       !kodeLgn = objO.FetchData_String("SELECT idcust FROM [SINGLEIDORI].dbo.bridging WITH (NOLOCK) WHERE kodelgn = '" & NulSID(rsRepo!FS_KD_PLG) & "' AND kodedealer = '" & Right(gDb, 2) & "' ")
247                                       !customer = NulSID(rsRepo!customer)
248                                       !customeralamat = NulSID(rsRepo!cust_adres)
249                                       !NoPolisi = NulSID(rsRepo!NO_POLISI)
250                                       !model = NulSID(rsRepo!Fmdl)
251                                       !noRangka = NulSID(rsRepo!NO_CHASSIS)
252                                       !noMesin = NulSID(rsRepo!No_Engine)
253                                       !noOrder = NulSID(rsRepo!no_wo)
254                                       !namaSA = objO.FetchData_String("select FS_NM_PETUGAS from T_PETUGAS with (nolock) where FS_KD_PETUGAS = '" & NulSID(rsRepo!FRONT) & "'")
255                                       !tipecustomer = NulSID(rsRepo!Status)
256                                       !mekanik = NulSID(rsRepo!Mech)
257                                       !billing = NulSID(rsRepo!FRONT) ''NulSID(rsRepo!InvDate) ''Format(NulSID(rsRepo!InvDate), "dd-mm-yyyy")
258                                       !job = Left(" " & "Add. " & NulSID(rsAddJob!rep), 50)
259                                       !jobdetail = NulSID(rsRinciSBE!Detail)
260                                       !jobRate = NulSID(rsAddJob!rte)
261                                       !jobharga = NulSID(rsAddJob!offn)
262                                       !jobdisc = IIf(chCtkDisc.value = False, IIf(NulSID(rsAddJob!FREEC) = 1, NulSID(rsAddJob!disc), 0) & " %", "") 'IIf(NulSID(rsAddJob!FREEC) = 1, NulSID(rsAddJob!disc), 0) & " %"
263                                       !jobsubtotal = NulSID(rsAddJob!offn * ((100 - NulSID(IIf(NulSID(rsAddJob!FREEC) = 1, NulSID(rsAddJob!disc), 0))) / 100))
264                                       !dppJasa = Round(NulA(rsRepo!labor) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
265                                       !dpppart = Round((NulA(rsRepo!PART_G) + NulA(rsRepo!PART_NG) + NulA(rsRepo!MTRL) + NulA(rsRepo!OIL_G) + NulA(rsRepo!OIL_NG)) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
266                                       !dppsublet = Round(NulA(rsRepo!sblt) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
267                                       !dpptotal = NulA(!dppJasa + !dpppart + !dppsublet) ''NulA(rsRepo!dpp)
268                                       !PPNtotal = NulA(rsRepo!fn_pajak)
269                                       !grandTotal = NulA(!dpptotal + !PPNtotal + nilaiMAT)  ''NulA(rsRepo!total)
270                                       .Update
271                                   End With
272                                   rsRinciSBE.MoveNext
273                               Loop
274                           End If
275                       End If
276                   End If
277                   rsAddJob.MoveNext
278               Loop
279           End If
280           gcat = ""
281           If nRC Then
                  Dim rsHlp As ADODB.Recordset
                  Dim nBulan As String
                  Dim nSbe As String
282               nSbe = ""
283               Set rsHlp = objO.FetchData("SELECT CASE WHEN JT_TGL < JT_BLN THEN JT_TGL ELSE JT_BLN END AS JT_TGL, SBENEXT FROM M_ESTIMASI WITH (NOLOCK) WHERE NORANGKA = '" & Trim(text(17)) & "' AND TGLSVCNOW = '" & Format(text(43), "MM/DD/YYYY") & "'")
284               If Not rsHlp.EOF Then
285                   nBulan = GetNamaBulan(Month(rsHlp("JT_TGL"))) & " " & year(rsHlp("JT_TGL"))
286                   nSbe = Trim(rsHlp("SBENEXT"))
287               End If
288               rsHlp.Close
289               If Trim(nSbe) <> "" Then
                      'gcat = IIf(nHeavy, "Kembali untuk service berkala setelah kelipatan 5.000 KM atau bulan " & nBulan, "Kembali untuk service berkala setelah kelipatan 10.000 KM atau bulan " & nBulan)
                      ''gcat = "Kembali untuk service berkala " & Format(CDbl(nSbe), "#,#0") & " KM atau bulan " & nBulan
290                   gcat = "Kembali untuk service berkala bulan " & nBulan & " atau " & Format(CDbl(nSbe), "#,#0") & " KM "
291                   gcat = gcat + IIf(gcat <> "", " (mana yang tercapai lebih dahulu).", "")
292               Else
293                   Call objO.InsertData("INSERT INTO ESTIMASISALAH (norangka, tglservice, kmsvc, sbesvc, kelipatan, noinv, nowo,customer, alamat,nopolisi, telpon, gsm,cdma) VALUES " & _
                                     "('" & Trim(text(17)) & "','" & Format(text(43), "MM/DD/YYYY") & "'," & CDbl(text(37)) & ",'" & Trim(text(33)) & "','" & IIf(nHeavy, 5000, 10000) & "', " & _
                                     "'" & Trim(text(20)) & "','" & Trim(Trim(text(0))) & "','" & Trim(text(4)) & "','" & Trim(txt5) & "','" & Trim(text(16)) & "','" & Trim(text(7)) & "','" & Trim(text(8)) & "','" & Trim(text(9)) & "')")
294               End If
295           End If
              
296           rpt.txtTerbilang.SetText "Terbilang (Rupiah) " & Trim(Angka2Huruf(rsCetakNewBilling!grandTotal, 1))
              ''20240829 revisi menutup kembali service
297           rpt.txtNextService.SetText "" 'gcat
              
              '' inspection ====================
298           If Log.CR7 = 1 Then
                  Dim rsInspection As ADODB.Recordset
                  Dim total1 As Double
                  Dim total2 As Double
                  Dim total3 As Double
                  Dim total4 As Double
                  
299               sSQL = "select sp1,sp2,sp3,sp4,estn1,estn2,estn3,estn4,estj1,estj2,estj3,estj4,estp1,estp2,estp3,estp4,ests1,ests2,ests3,ests4 from INSPECTION with (nolock) where NoWo = '" & no_wo & "' order by tglentry desc"
300               Set rsInspection = objO.FetchData(sSQL)
301               If Not rsInspection.EOF Then
302                   rsInspection.MoveFirst
303                   rpt.txtSaran1.SetText NulSID(rsInspection!sp1)
304                   rpt.txtSaran2.SetText NulSID(rsInspection!sp2)
305                   rpt.txtSaran3.SetText NulSID(rsInspection!sp3)
306                   rpt.txtSaran4.SetText NulSID(rsInspection!sp4)
                      
307                   If NulSID(rsInspection!sp1) <> "" Then
308                       rpt.txtWaktuSVC1.SetText Format(NulSID(rsInspection!estn1), "dd-mm-yyyy")
309                   Else
310                       rpt.txtWaktuSVC1.SetText ""
311                   End If
312                   If NulSID(rsInspection!sp2) <> "" Then
313                       rpt.txtWaktuSVC2.SetText Format(NulSID(rsInspection!estn2), "dd-mm-yyyy")
314                   Else
315                       rpt.txtWaktuSVC2.SetText ""
316                   End If
317                   If NulSID(rsInspection!sp3) <> "" Then
318                       rpt.txtWaktuSVC3.SetText Format(NulSID(rsInspection!estn3), "dd-mm-yyyy")
319                   Else
320                       rpt.txtWaktuSVC3.SetText ""
321                   End If
322                   If NulSID(rsInspection!sp4) <> "" Then
323                       rpt.txtWaktuSVC4.SetText Format(NulSID(rsInspection!estn4), "dd-mm-yyyy")
324                   Else
325                       rpt.txtWaktuSVC4.SetText ""
326                   End If
                                      
327                   If NulSID(rsInspection!sp1) <> "" Then
328                       rpt.txtEstJasa1.SetText Format(NulSID(rsInspection!estj1), "#,0")
329                   Else
330                       rpt.txtEstJasa1.SetText ""
331                   End If
332                   total1 = total1 + NulSID(rsInspection!estj1)
                      
333                   If NulSID(rsInspection!sp2) <> "" Then
334                       rpt.txtEstJasa2.SetText Format(NulSID(rsInspection!estj2), "#,0")
335                   Else
336                       rpt.txtEstJasa2.SetText ""
337                   End If
338                   total2 = total2 + NulSID(rsInspection!estj2)
                      
339                   If NulSID(rsInspection!sp3) <> "" Then
340                       rpt.txtEstJasa3.SetText Format(NulSID(rsInspection!estj3), "#,0")
341                   Else
342                       rpt.txtEstJasa3.SetText ""
343                   End If
344                   total3 = total3 + NulSID(rsInspection!estj3)
                      
345                   If NulSID(rsInspection!sp4) <> "" Then
346                       rpt.txtEstJasa4.SetText Format(NulSID(rsInspection!estj4), "#,0")
347                   Else
348                       rpt.txtEstJasa4.SetText ""
349                   End If
350                   total4 = total4 + NulSID(rsInspection!estj4)
                                                      
351                   If NulSID(rsInspection!sp1) <> "" Then
352                       rpt.txtEstPart1.SetText Format(NulSID(rsInspection!estp1), "#,0")
353                   Else
354                       rpt.txtEstPart1.SetText ""
355                   End If
356                   total1 = total1 + NulSID(rsInspection!estp1)
                                                      
357                   If NulSID(rsInspection!sp2) <> "" Then
358                       rpt.txtEstPart2.SetText Format(NulSID(rsInspection!estp2), "#,0")
359                   Else
360                       rpt.txtEstPart2.SetText ""
361                   End If
362                   total2 = total2 + NulSID(rsInspection!estp2)
                                                      
363                   If NulSID(rsInspection!sp3) <> "" Then
364                       rpt.txtEstPart3.SetText Format(NulSID(rsInspection!estp3), "#,0")
365                   Else
366                       rpt.txtEstPart3.SetText ""
367                   End If
368                   total3 = total3 + NulSID(rsInspection!estp3)
                                                      
369                   If NulSID(rsInspection!sp4) <> "" Then
370                       rpt.txtEstPart4.SetText Format(NulSID(rsInspection!estp4), "#,0")
371                   Else
372                       rpt.txtEstPart4.SetText ""
373                   End If
374                   total4 = total4 + NulSID(rsInspection!estp4)
                                      
                                                      
375                   If NulSID(rsInspection!sp1) <> "" Then
376                       rpt.txtEstSublet1.SetText Format(NulSID(rsInspection!ests1), "#,0")
377                   Else
378                       rpt.txtEstSublet1.SetText ""
379                   End If
380                   total1 = total1 + NulSID(rsInspection!ests1)
                                                      
381                   If NulSID(rsInspection!sp2) <> "" Then
382                       rpt.txtEstSublet2.SetText Format(NulSID(rsInspection!ests2), "#,0")
383                   Else
384                       rpt.txtEstSublet2.SetText ""
385                   End If
386                   total2 = total2 + NulSID(rsInspection!ests2)
                                                      
387                   If NulSID(rsInspection!sp3) <> "" Then
388                       rpt.txtEstSublet3.SetText Format(NulSID(rsInspection!ests3), "#,0")
389                   Else
390                       rpt.txtEstSublet3.SetText ""
391                   End If
392                   total3 = total3 + NulSID(rsInspection!ests3)
                                                      
393                   If NulSID(rsInspection!sp4) <> "" Then
394                       rpt.txtEstSublet4.SetText Format(NulSID(rsInspection!ests4), "#,0")
395                   Else
396                       rpt.txtEstSublet4.SetText ""
397                   End If
398                   total4 = total4 + NulSID(rsInspection!ests4)
                                      
399                   If NulSID(rsInspection!sp1) <> "" Then
400                       rpt.txtEstTotal1.SetText Format(NulSID(total1), "#,0")
401                   Else
402                       rpt.txtEstTotal1.SetText ""
403                   End If
404                   If NulSID(rsInspection!sp2) <> "" Then
405                       rpt.txtEstTotal2.SetText Format(NulSID(total2), "#,0")
406                   Else
407                       rpt.txtEstTotal2.SetText ""
408                   End If
409                   If NulSID(rsInspection!sp3) <> "" Then
410                       rpt.txtEstTotal3.SetText Format(NulSID(total3), "#,0")
411                   Else
412                       rpt.txtEstTotal3.SetText ""
413                   End If
414                   If NulSID(rsInspection!sp4) <> "" Then
415                       rpt.txtEstTotal4.SetText Format(NulSID(total4), "#,0")
416                   Else
417                       rpt.txtEstTotal4.SetText ""
418                   End If
                      
419               End If
420           End If
              '' akhir inspection ==============
                               
421           rpt.txtBankNama1.SetText NulSID(rsProfil!bankname1)
422           rpt.txtBankNama2.SetText NulSID(rsProfil!bankname2)
423           rpt.txtBankNama3.SetText NulSID(rsProfil!bankname3)
              
424           rpt.txtBankNo1.SetText NulSID(rsProfil!bankno1)
425           rpt.txtBankNo2.SetText NulSID(rsProfil!bankno2)
426           rpt.txtBankNo3.SetText NulSID(rsProfil!bankno3)
              
427           rpt.txtInfoMaterai.SetText ""
428           rpt.txtLunasMaterai.SetText ""
              
429           If Log.fakturMaterai = 1 Then
430               Set rsTemp = objO.FetchData("SELECT * FROM t_invoiceMaterai WITH (NOLOCK) WHERE NOINV = '" & noInv & "'")
431               If Not rsTemp.EOF Then
432                   rpt.txtInfoMaterai.SetText "Ijin Pembubuhan Materai No. " & NulSID(rsTemp!NOIJIN)
433                   rpt.txtLunasMaterai.SetText "BEA MATERAI LUNAS " & NulSID(rsTemp!nilai)
434               End If
435           End If
436           rsCetakNewBilling.MoveFirst
437           rpt.Text57.SetText NulSID(objO.FetchData_String("SELECT FS_KA_BENGKEL FROM T_PROFIL"))
438           rpt.Text8.SetText NulSID(objO.FetchData_String("SELECT FS_KOTA FROM T_PROFIL"))
439           rpt.txtMaterai.SetText IIf(nilaiMAT = 0, "0,00", Format(NulSID(nilaiMAT), "#,###.00"))
440           rpt.txtTglBilling.SetText tglInv
441           rpt.Text41.SetText "PPn " & nilaiPPnInv & " %"
442           rpt.Text55.SetText "PPn " & nilaiPPnInv & " %"
443           rpt.database.SetDataSource rsCetakNewBilling
444           If No = 1 Then
445               With CRViewer1
446                   .ReportSource = rpt
447                   .ViewReport
448                   .Zoom 100
449                   .Top = 0
450                   .Left = asxPowerBanner1.Left + asxPowerBanner1.Width + 20
451                   .Width = Frame2.Width + 100
452                   .Height = asxPowerBanner1.Height
                      
453                   cmdTutup.Top = .Top + 430
454                   cmdTutup.Left = .Left
455                   cmdTutup.Visible = True
456                   .Visible = True
457               End With
458               cmdTutup.Visible = True
459           ElseIf No = 2 Then
460                With CRViewer2
461                   .ReportSource = rpt
462                   .ViewReport
463                   .Zoom 100
464                   .Top = 0
465                   .Left = asxPowerBanner1.Left + asxPowerBanner1.Width + 20
466                   .Width = Frame2.Width + 100
467                   .Height = asxPowerBanner1.Height
                      
468                   cmdTutup2.Top = .Top + 430
469                   cmdTutup2.Left = .Left
470                   cmdTutup2.Visible = True
471                   .Visible = True
472               End With
473               cmdTutup2.Visible = True
474           Else
475                With CRViewer5
476                   .ReportSource = rpt
477                   .ViewReport
478                   .Zoom 100
479                   .Top = 0
480                   .Left = asxPowerBanner1.Left + asxPowerBanner1.Width + 20
481                   .Width = Frame2.Width + 100
482                   .Height = asxPowerBanner1.Height
                      
483                   Command1.Top = .Top + 430
484                   Command1.Left = .Left
485                   Command1.Visible = True
486                   .Visible = True
487               End With
488               DoEvents
489               Command1.Visible = True
490           End If
491           Screen.MousePointer = vbDefault
492           Exit Sub
              
493       Else
494           MsgBox "Data invoice tidak ditemukan", vbInformation
495       End If
End Sub


Public Sub cetakSubletSimplify()
1         Set Report = New rptSublet
          Dim rsRepo As ADODB.Recordset
          
          ''' ASLI NICH !!!
      '    sSQL = "SELECT CUSTOMER, CUST_ADRES, KOTA, A.NO_WO, A.NOINV, B.NO_POLISI, A.NO_CHASSIS, NO_ENGINE, NAMABARANG, HARGA, QTY, KETERANGAN " _
      '         & "FROM HIST_SUBLET A INNER JOIN T_INVOICE B ON A.NO_WO = B.NO_WO AND A.NOINV = B.NOINV " _
      '         & "WHERE A.NOINV = '" & text(0) & "' AND A.NO_WO = '" & text(20) & "' ORDER BY A.NO_WO, NAMABARANG "
          '#########################################################'
          '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..'
          '#########################################################'
2         sSQL = "SELECT DISTINCT CUSTOMER, CUST_ADRES, KOTA, A.NO_WO, A.NOINV, B.NO_POLISI, A.NO_CHASSIS, NO_ENGINE, NAMABARANG, HARGA, QTY, KETERANGAN " _
               & "FROM HIST_SUBLET A WITH(NOLOCK) INNER JOIN T_INVOICE B WITH(NOLOCK) ON A.NO_WO = B.NO_WO AND A.NOINV = B.NOINV  " _
               & "WHERE A.NOINV LIKE '" & text(0) & "%' AND  B.NOINV NOT LIKE '%-B' and B.NOINV NOT LIKE '%-C'  AND A.NO_WO = '" & text(20) & "' ORDER BY A.NO_WO, NAMABARANG "
          
3         Set rsRepo = objO.FetchData(sSQL)
4         If Not rsRepo.EOF Then
5             Report.database.SetDataSource rsRepo
6             Report.txtPT.SetText Trim(gPerush)
7             Report.TxtAlamat.SetText Trim(gAlamat)
8             Report.TxtKota.SetText Trim(gKotaCbg)
9             Report.txtTelp.SetText "Telp. " & Trim(gTelp)
10            Report.TxtFax.SetText ""  '"Fax. "
11            Report.txtKotaTgl.SetText Trim(gKotaCbg) & ", " & TglDDMMMMYYYY(Format(text(1), "DD/MM/YYYY"), 1)
12            Report.txtKabeng.SetText Trim(gSecmanager)
          
13            With CRViewer3
14                .ReportSource = Report
15                .ViewReport
16                .Zoom 100
                  
17                .Top = 0
18                .Left = asxPowerBanner1.Left + asxPowerBanner1.Width + 20
19                .Width = Frame2.Width + 100
20                .Height = TblKeluar.Top + TblKeluar.Height
                  
21                cmdTutup3.Top = .Top + 430
22                cmdTutup3.Left = .Left
23                cmdTutup3.Visible = True
                  
24                cmdSettingPrinter3.Top = cmdTutup3.Top
25                cmdSettingPrinter3.Left = cmdTutup3.Left + cmdTutup3.Width + 50
26                cmdSettingPrinter3.Visible = True
27                .Visible = True
28            End With
29            Screen.MousePointer = vbDefault
30            cmdTutup3.SetFocus
31        End If
End Sub

Public Sub cetakSupplySlipSimplify()
              Dim UkuranKertas As Double
          Dim rsRepo As ADODB.Recordset
          'FrmMain.Timer3.Enabled = False
          Dim SQL As String
          Dim dBaseFK   As String
1         dBaseFK = objO.FetchData_String("SELECT DBASEFK FROM JOBDTA WITH (NOLOCK) WHERE NO_WO = '" & NulSID(text(20)) & "' AND FS_LOG_DEL = '' ")
2         If dBaseFK = "" Then
3           GoTo lanjut
4         ElseIf NulSID(text(47)) = "" Then
5             GoTo lanjut
6         End If
7         Set Report = New Report_SiCetakFakturJual
          
8         If Log.freepart = 1 Then
9             SQL = "select DISTINCT '" & NulSID(text(47)) & "' NoFaktur,'' NoFjDulu,'A' TypeTrn,Armst.NamaLgn, Armst.Alamat1, Armst.Alamat2, Armst.Alamat3,"
10            SQL = SQL & " '' NoPo,a.NO_WO NoSo,b.INVDATE TglJthtmp, b.INVDATE TglFaktur, b.fs_kd_plg Kodelgn, Armst.NamaLgn NamaKirim, Armst.Alamat1 AlamatKirim1, Armst.Alamat2 AlamatKirim2, Armst.Alamat3 AlamatKirim3, round((b.OIL_G+b.OIL_NG+b.PART_G+b.PART_NG)*" & k2t(nilaiDppInv / nilaiTotalInv) & ",0) Subtotal, 0 AS ExtDisJlh, b.MET Meterai, 'T' TypePpn, 0 PPN, 0 OngkosAngkut, round((b.OIL_G+b.OIL_NG+b.PART_G+b.PART_NG)*" & k2t(nilaiDppInv / nilaiTotalInv) & ",0) GrandTotal, 0 UangMuka, round((b.OIL_G+b.OIL_NG+b.PART_G+b.PART_NG)*" & k2t(nilaiDppInv / nilaiTotalInv) & ",0) SisaTotal, '' AS Ketr, ' S U P P L Y   SLIP ' AS KetFaktur, ROW_NUMBER() OVER(order by a.kodeitem) Nomor, a.KodeItem, i.NoRack  NoRack, a.KETERANGAN Keterangan,a.Qty, '' SatuanNs, "
11            SQL = SQL & " isnull((select hargajual from [" & dBaseFK & "]..sisal where NoFaktur = '" & NulSID(text(47)) & "' and kodeitem = a.kodeitem ),round((a.HARGA*" & k2t(nilaiDppInv / nilaiTotalInv) & ")/a.qty,0)) HargaJual, convert(money,round(FN_EXT_DISKON/100,2)) ItemDisPsn, ROUND(a.HARGA*" & k2t(nilaiDppInv / nilaiTotalInv) & ",0) AS Jumlah"
12            SQL = SQL & " from T_PART a with (nolock) "
13            SQL = SQL & " left join t_invoice b with (nolock)  on a.NO_WO = b.NO_WO and a.NOINV = b.NOINV"
14            SQL = SQL & " left join  [" & dBaseFK & "]..Armst Armst with (nolock)  on Armst.kodelgn = b.fs_kd_plg"
15            SQL = SQL & " left join  [" & dBaseFK & "]..inmst i with (nolock)  on  i.KodeItem = a.KODEITEM"
16            SQL = SQL & " where a.fs_log_del = '' and  a.no_wo = '" & NulSID(text(20)) & "' and a.noinv not like '%C' and a.noinv not like '%B'"
              
17            Set rsRepo = objO.FetchData(SQL)
              
18            If Not rsRepo.EOF Then
19                GoTo lanjutReport
20            Else
21                GoTo cekKembali
22            End If
23        End If
cekKembali:
24        SQL = "SELECT DISTINCT Sihdr.NoFaktur, Sihdr.NoFjDulu, Sihdr.TypeTrn, Armst.NamaLgn, Armst.Alamat1, Armst.Alamat2, Armst.Alamat3, Sihdr.NoPo, Sihdr.NoSo, Sihdr.TglJthtmp, Sihdr.TglFaktur, Sihdr.Kodelgn, Sihdr.NamaKirim, Sihdr.AlamatKirim1, Sihdr.AlamatKirim2, Sihdr.AlamatKirim3, Sihdr.Subtotal, [Subtotal]*[ExtDisPersen] AS ExtDisJlh, Sihdr.Meterai, Sihdr.TypePpn, Sihdr.PPN, Sihdr.OngkosAngkut, Sihdr.GrandTotal, Sihdr.UangMuka, Sihdr.SisaTotal, Left(cast([Sihdr].[Keterangan] as varchar),300) AS Ketr,"
25        SQL = SQL + " case when Sihdr.NoSO='' then ' P A R T S  INVOICE' ELSE ' S U P P L Y   SLIP ' END AS KetFaktur,"
26        SQL = SQL + " Sisal.Nomor, Sisal.KodeItem, Inmst.NoRack, Sisal.Keterangan, Sisal.Qty, Sisal.SatuanNs, Sisal.HargaJual, Sisal.ItemDisPsn, "
27        SQL = SQL + "(abs(Sisal.Qty*[HargaJual])-(ItemDisc))*(Sisal.Qty/abs(Sisal.Qty)) AS Jumlah"
28        SQL = SQL + " FROM (([" & dBaseFK & "]..Sihdr Sihdr INNER JOIN [" & dBaseFK & "]..Armst Armst ON Sihdr.KodeLgn = Armst.KodeLgn) INNER JOIN [" & dBaseFK & "]..Sisal Sisal ON Sihdr.NoFaktur = Sisal.NoFaktur) INNER JOIN [" & dBaseFK & "]..Inmst Inmst ON Sisal.KodeItem = Inmst.KodeItem"
29        SQL = SQL + " WHERE (Sihdr.NoFaktur = '" & NulSID(text(47)) & "')"
30        SQL = SQL + " AND (Sihdr.TypeTrn In ('A','B'))"
31        SQL = SQL + " ORDER BY Sihdr.NoFaktur, Sisal.Nomor;"
          
32        Set rsRepo = objO.FetchDataACC(SQL)
lanjutReport:
33        If Not rsRepo.EOF Then
34            Report.database.SetDataSource rsRepo 'DbCon.Execute(SQL)
              
35            Report.txtKodeSOP.SetText "FM-NRM-SVC-01.03"
               
36            Report.TxtNamaPT.SetText profilACC.gNamaProfil
37            Report.TxtAlamatPT.SetText profilACC.AlamatProfil1
38            Report.TxtKotaPT.SetText profilACC.KotaProfil
39            Report.TxtTelpPT.SetText ": " & profilACC.TeleponProfil
40            Report.TxtFax.SetText ": " & profilACC.FaxProfil
41            Report.txtNpwp.SetText ": " & profilACC.npwp
42            Report.TxtPKP.SetText ": " & profilACC.SKPengukuhan
43            Report.TxtTglPengukuhan.SetText ": " & profilACC.TglPengukuhan
               
44            Report.FormulaFields(2).text = "0" ' & IIf(.ChkDisc.value = Checked, 0, 1) & ""
              
45            UkuranKertas = Val(GetSetting("SUPPLYSLIP-SIMPLIFY", "Startup", "PaperSize", 1))
46            If UkuranKertas = 0 Then UkuranKertas = 1
47            Report.PaperSize = UkuranKertas
48            Screen.MousePointer = vbHourglass
49            With CRViewer4
50                 .ReportSource = Report
51                 .ViewReport
52                 .Zoom 100
                   
53                 .Top = 0
54                 .Left = asxPowerBanner1.Left + asxPowerBanner1.Width + 20
55                 .Width = Frame2.Width + 100
56                 .Height = TblKeluar.Top + TblKeluar.Height
                  
57                 cmdTutup4.Top = .Top + 430
58                 cmdTutup4.Left = .Left
59                 cmdTutup4.Visible = True
                  
60                 cmdSettingPrinter4.Top = cmdTutup4.Top
61                 cmdSettingPrinter4.Left = cmdTutup4.Left + cmdTutup4.Width + 50
62                 cmdSettingPrinter4.Visible = True
63                 .Visible = True
64            End With
65        End If
lanjut:
66    Screen.MousePointer = vbDefault

End Sub


Public Sub cetakSupplySlipSimplifyC(noInv As String)
          Dim UkuranKertas As Double
          Dim rsRepo As ADODB.Recordset
          'FrmMain.Timer3.Enabled = False
          Dim SQL As String
          Dim dBaseFK   As String
1         dBaseFK = objO.FetchData_String("SELECT DBASEFK FROM JOBDTA WITH (NOLOCK) WHERE NO_WO = '" & NulSID(text(20)) & "' AND FS_LOG_DEL = '' ")
2         If dBaseFK = "" Then
3           GoTo lanjut
4         ElseIf NulSID(text(47)) = "" Then
5             GoTo lanjut
6         End If
7         Set Report = New Report_SiCetakFakturJual
          
8         SQL = "select '" & NulSID(text(47)) & "' NoFaktur,'' NoFjDulu,'A' TypeTrn,Armst.NamaLgn, Armst.Alamat1, Armst.Alamat2, Armst.Alamat3,"
9         SQL = SQL & " '' NoPo,a.NO_WO NoSo,b.INVDATE TglJthtmp, b.INVDATE TglFaktur, b.fs_kd_plg Kodelgn, Armst.NamaLgn NamaKirim, Armst.Alamat1 AlamatKirim1, Armst.Alamat2 AlamatKirim2, Armst.Alamat3 AlamatKirim3, round((b.OIL_G+b.OIL_NG+b.PART_G+b.PART_NG)*" & k2t(nilaiDppInv / nilaiTotalInv) & ",0) Subtotal, 0 AS ExtDisJlh, b.MET Meterai, 'T' TypePpn, 0 PPN, 0 OngkosAngkut, round((b.OIL_G+b.OIL_NG+b.PART_G+b.PART_NG)*" & k2t(nilaiDppInv / nilaiTotalInv) & ",0) GrandTotal, 0 UangMuka, round((b.OIL_G+b.OIL_NG+b.PART_G+b.PART_NG)*" & k2t(nilaiDppInv / nilaiTotalInv) & ",0) SisaTotal, '' AS Ketr, ' S U P P L Y   SLIP ' AS KetFaktur, ROW_NUMBER() OVER(order by a.kodeitem) Nomor, a.KodeItem, i.NoRack NoRack, a.KETERANGAN Keterangan,a.Qty, '' SatuanNs, "
10        SQL = SQL & " isnull((select hargajual from [" & dBaseFK & "]..sisal where NoFaktur = '" & NulSID(text(47)) & "' and kodeitem = a.kodeitem ),round((a.HARGA*" & k2t(nilaiDppInv / nilaiTotalInv) & ")/a.qty,0))  HargaJual, 0 ItemDisPsn, ROUND(a.HARGA*" & k2t(nilaiDppInv / nilaiTotalInv) & ",0) AS Jumlah"
11        SQL = SQL & " from T_PART a with (nolock) "
12        SQL = SQL & " left join t_invoice b with (nolock)  on a.NO_WO = b.NO_WO and a.NOINV = b.NOINV"
13        SQL = SQL & " left join  [" & dBaseFK & "]..Armst Armst with (nolock)  on Armst.kodelgn = b.fs_kd_plg"
14        SQL = SQL & " left join  [" & dBaseFK & "]..inmst i with (nolock)  on  i.KodeItem = a.KODEITEM"
15        SQL = SQL & " where a.NOINV = '" & noInv & "'"
          
16        Set rsRepo = objO.FetchData(SQL)
          
17        If Not rsRepo.EOF Then
18            Report.database.SetDataSource rsRepo 'DbCon.Execute(SQL)
              
19            Report.txtKodeSOP.SetText "FM-NRM-SVC-01.03"
               
20            Report.TxtNamaPT.SetText profilACC.gNamaProfil
21            Report.TxtAlamatPT.SetText profilACC.AlamatProfil1
22            Report.TxtKotaPT.SetText profilACC.KotaProfil
23            Report.TxtTelpPT.SetText ": " & profilACC.TeleponProfil
24            Report.TxtFax.SetText ": " & profilACC.FaxProfil
25            Report.txtNpwp.SetText ": " & profilACC.npwp
26            Report.TxtPKP.SetText ": " & profilACC.SKPengukuhan
27            Report.TxtTglPengukuhan.SetText ": " & profilACC.TglPengukuhan
               
28            Report.FormulaFields(2).text = "0" ' & IIf(.ChkDisc.value = Checked, 0, 1) & ""
              
29            UkuranKertas = Val(GetSetting("SUPPLYSLIP-SIMPLIFY", "Startup", "PaperSize", 1))
30            If UkuranKertas = 0 Then UkuranKertas = 1
31            Report.PaperSize = UkuranKertas
32            Screen.MousePointer = vbHourglass
33            With CRViewer6
34                 .ReportSource = Report
35                 .ViewReport
36                 .Zoom 100
                   
37                 .Top = 0
38                 .Left = asxPowerBanner1.Left + asxPowerBanner1.Width + 20
39                 .Width = Frame2.Width + 100
40                 .Height = TblKeluar.Top + TblKeluar.Height
                  
41                 cmdTutup6.Top = .Top + 430
42                 cmdTutup6.Left = .Left
43                 cmdTutup6.Visible = True
      '
      '             cmdSettingPrinter4.Top = cmdTutup4.Top
      '             cmdSettingPrinter4.Left = cmdTutup4.Left + cmdTutup4.Width + 50
      '             cmdSettingPrinter4.Visible = True
44                 .Visible = True
45            End With
46        End If
lanjut:
47    Screen.MousePointer = vbDefault

End Sub
                    
'' ========================================================================================

Private Sub loadJobDetailMultiple(noWO As String, noRangka As String)
          Dim rsHlp   As ADODB.Recordset
          Dim rsHlp2  As ADODB.Recordset
          Dim rsHlp3  As ADODB.Recordset
          Dim rsHlp4  As ADODB.Recordset
          
          Dim nilaiSubTotal   As Double
          
          
1         sSQL = "select * from [t_invoicejobdetail] with (nolock) where no_wo = '" & NulSID(text(20)) & "' and no_chassis = '" & NulSID(noRangka) & "' and fs_log_del = ''"
2         Set rsJobDetail = objO.FetchData(sSQL)
3         If Not rsJobDetail.EOF Then
4             rsJobDetail.MoveFirst
5             Do While Not rsJobDetail.EOF
6                 If NulSID(rsJobDetail!id_program) <> "" Then
7                     If objBilling.isDuoInvoice(NulSID(rsJobDetail!id_program)) Then
8                         nDuoInv = True
9                     End If
10                End If
                  rsJobDetail.MoveNext
11            Loop
12            rsJobDetail.MoveFirst
13            If rsJob.State = adStateOpen Then rsJob.Close
14            If rsJob.State = adStateClosed Then doCreateJob
15            Do While Not rsJobDetail.EOF
16                With rsJob
                  
17                    namaProgram = ""
                      
18                   .AddNew
19                   !job = NulSID(rsJobDetail!rep)
20                   !Rate = CDbl(NulSID(rsJobDetail!rte))
                     ''!ACTRT = CDbl(NulSID(rsJobDetail!actrte))
                     '!grp = NulSID(rsJobDetail!grp)
                     '!Mech = NulSID(rsJobDetail!mekanik)
21                   If NulSID(rsJobDetail!isFP) = 1 Then
22                    !hrg = objO.FetchData_Double("select price from [" & dbReplikasiHO & "]..HO_FIXPRICE where id_FP = '" & NulSID(rsJobDetail!id_program) & "' and FS_LOG_DEL = ''")
23                   Else
24                    !hrg = CDbl(rsJobDetail!offn)
25                   End If
26                   !addJob = IIf(NulSID(rsJobDetail!addJob) = 1, True, False)
27                   !disc = CDbl(rsJobDetail!disc)
28                   !DISCRP = Round((!hrg * !Rate) - CDbl(rsJobDetail!Price), 0)
29                   If NulSID(rsJobDetail!isFP) = 1 Then
30                    !STLDISC = !hrg
31                   Else
32                    !STLDISC = Round(CDbl(rsJobDetail!Price), 0)
33                   End If
                    '' !valid = gVerifikator
34                   If NulSID(rsJobDetail!FREEC) = 1 Then
35                       !FREEC = True
36                   Else
37                       !FREEC = False
38                   End If
39                   !idProgram = NulSID(rsJobDetail!id_program)
                     
40                   !isFP = IIf(NulSID(rsJobDetail!isFP) = 1, True, False)
41                   If NulSID(rsJobDetail!isFP) = 1 Then
42                    namaProgram = objO.FetchData_String("select namaFP from [" & dbReplikasiHO & "]..HO_FIXPRICE where id_FP = '" & NulSID(rsJobDetail!id_program) & "' and FS_LOG_DEL = ''")
43                   Else
44                    namaProgram = objO.FetchData_String("select PROGRAM from [" & dbReplikasiHO & "]..HO_PROGRAM where id_Program = '" & NulSID(rsJobDetail!id_program) & "' and FS_LOG_DEL = ''")
45                   End If
46                   !program = namaProgram
47                   If NulSID(rsJobDetail!isFP) = 1 And !FREEC Then
48                    text(22) = text(22) + !hrg
49                   ElseIf nDuoInv = False Then
50                    text(22) = text(22) + !STLDISC
51                   ElseIf !FREEC Then
52                    text(22) = text(22) + !STLDISC
53                   ElseIf !FREEC = False Then
54                    text(22) = text(22) + 0 '!STLDISC
55                   End If
56                   .Update
57                End With
58                subJob = subJob + (CDbl(nHarga) * CDbl(nilaiRate))
59                totRate = totRate + CDbl(nilaiRate)
60                rsJobDetail.MoveNext
61            Loop
62        End If
End Sub


Public Sub cetakSubletSimplifyC()
1         Set Report = New rptSublet
          Dim rsRepo As ADODB.Recordset
          
          ''' ASLI NICH !!!
      '    sSQL = "SELECT CUSTOMER, CUST_ADRES, KOTA, A.NO_WO, A.NOINV, B.NO_POLISI, A.NO_CHASSIS, NO_ENGINE, NAMABARANG, HARGA, QTY, KETERANGAN " _
      '         & "FROM HIST_SUBLET A INNER JOIN T_INVOICE B ON A.NO_WO = B.NO_WO AND A.NOINV = B.NOINV " _
      '         & "WHERE A.NOINV = '" & text(0) & "' AND A.NO_WO = '" & text(20) & "' ORDER BY A.NO_WO, NAMABARANG "
          '#########################################################'
          '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..'
          '#########################################################'
2         sSQL = "SELECT DISTINCT CUSTOMER, CUST_ADRES, KOTA, A.NO_WO, A.NOINV, B.NO_POLISI, A.NO_CHASSIS, NO_ENGINE, NAMABARANG, HARGA, QTY, KETERANGAN " _
               & "FROM HIST_SUBLET A WITH(NOLOCK) INNER JOIN T_INVOICE B WITH(NOLOCK) ON A.NO_WO = B.NO_WO AND A.NOINV = B.NOINV  " _
               & "WHERE A.NOINV LIKE '" & text(0) & "-C' AND A.NO_WO = '" & text(20) & "' ORDER BY A.NO_WO, NAMABARANG "
          
          ''20220609 rev nama alamat dealer penjual diambil dari t_dealer
      '    sSQL = "SELECT DISTINCT ISNULL(C.FS_NM_TAGIH,B.CUSTOMER) AS CUSTOMER, isnull(C.FS_ALAMAT,B.CUST_ADRES) as CUST_ADRES, KOTA, A.NO_WO, A.NOINV, B.NO_POLISI, A.NO_CHASSIS, NO_ENGINE, NAMABARANG, HARGA, QTY, KETERANGAN " _
      '         & "FROM HIST_SUBLET A WITH(NOLOCK) "
      '    sSQL = sSQL & " INNER JOIN T_INVOICE B WITH(NOLOCK) ON A.NO_WO = B.NO_WO AND A.NOINV = B.NOINV  "
      '    sSQL = sSQL & " LEFT JOIN T_DEALER C WITH (NOLOCK) ON B.FS_KD_PLG = C.KDPLG and C.FS_LOG_DEL = ''" _
      '         & "WHERE A.NOINV LIKE '" & Text(0) & "-C' AND A.NO_WO = '" & Text(20) & "' ORDER BY A.NO_WO, NAMABARANG "
          
                
3         Set rsRepo = objO.FetchData(sSQL)
4         If Not rsRepo.EOF Then
5             Report.database.SetDataSource rsRepo
6             Report.txtPT.SetText Trim(gPerush)
7             Report.TxtAlamat.SetText Trim(gAlamat)
8             Report.TxtKota.SetText Trim(gKotaCbg)
9             Report.txtTelp.SetText "Telp. " & Trim(gTelp)
10            Report.TxtFax.SetText ""  '"Fax. "
11            Report.txtKotaTgl.SetText Trim(gKotaCbg) & ", " & TglDDMMMMYYYY(Format(text(1), "DD/MM/YYYY"), 1)
12            Report.txtKabeng.SetText Trim(gSecmanager)
          
13            With CRViewer7
14                .ReportSource = Report
15                .ViewReport
16                .Zoom 100
                  
17                .Top = 0
18                .Left = asxPowerBanner1.Left + asxPowerBanner1.Width + 20
19                .Width = Frame2.Width + 100
20                .Height = TblKeluar.Top + TblKeluar.Height
                  
21                cmdTutup7.Top = .Top + 430
22                cmdTutup7.Left = .Left
23                cmdTutup7.Visible = True
                  
                  'cmdSettingPrinter3.Top = cmdTutup3.Top
                  'cmdSettingPrinter3.Left = cmdTutup3.Left + cmdTutup3.Width + 50
                  'cmdSettingPrinter3.Visible = True
24                .Visible = True
25            End With
26            Screen.MousePointer = vbDefault
27            cmdTutup7.SetFocus
28        End If
End Sub

