VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form FormGRPAdjustment 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8370
   ClientLeft      =   0
   ClientTop       =   4005
   ClientWidth     =   12225
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   Picture         =   "FormGRPAdjustment.frx":0000
   ScaleHeight     =   8370
   ScaleWidth      =   12225
   ShowInTaskbar   =   0   'False
   Begin VB.VScrollBar VScrollFUB 
      Height          =   8350
      LargeChange     =   400
      Left            =   11720
      Max             =   4000
      SmallChange     =   350
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Frame SSFUB 
      BorderStyle     =   0  'None
      Height          =   16000
      Left            =   480
      TabIndex        =   1
      Top             =   -240
      Width           =   11415
      Begin VB.CommandButton cmdTop 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Back To Top"
         Height          =   735
         Left            =   40
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FormGRPAdjustment.frx":F314
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   11640
         Width           =   11220
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "E&xit"
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
         Left            =   10320
         Picture         =   "FormGRPAdjustment.frx":F69F
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   7560
         Width           =   735
      End
      Begin VB.Frame FrameAdj 
         Caption         =   ".:: GRP Adjustment ::."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1290
         Left            =   0
         TabIndex        =   72
         Top             =   7245
         Width           =   7815
         Begin VB.OptionButton OptGRP 
            Caption         =   "Set as GRP"
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
            Left            =   240
            TabIndex        =   76
            Top             =   360
            Width           =   2055
         End
         Begin VB.OptionButton OptSBE 
            Caption         =   "Set as SBE "
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
            Left            =   240
            TabIndex        =   75
            Top             =   720
            Width           =   2055
         End
         Begin VB.ComboBox cboKm 
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
            Left            =   2280
            Style           =   2  'Dropdown List
            TabIndex        =   74
            Top             =   720
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton cmdAdjust 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Adjust"
            Height          =   732
            Left            =   4440
            Picture         =   "FormGRPAdjustment.frx":FE2F
            Style           =   1  'Graphical
            TabIndex        =   73
            Top             =   360
            Width           =   732
         End
         Begin MSComctlLib.ImageList ImageList 
            Left            =   7080
            Top             =   360
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   152
            ImageHeight     =   65
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   2
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormGRPAdjustment.frx":1049C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormGRPAdjustment.frx":120DD
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H00C0C0C0&
            FillColor       =   &H00FFFFFF&
            Height          =   975
            Left            =   4320
            Shape           =   4  'Rounded Rectangle
            Top             =   240
            Width           =   975
         End
         Begin VB.Image ImgView 
            Height          =   975
            Left            =   5400
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.CommandButton CmdXls 
         BackColor       =   &H00FFFFFF&
         Caption         =   "X&LS"
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
         Left            =   9360
         Picture         =   "FormGRPAdjustment.frx":13DF7
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   7560
         Width           =   735
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   2415
         Left            =   0
         TabIndex        =   2
         Top             =   4755
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   4260
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   5
         TabHeight       =   706
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   ".:: Customer ::."
         TabPicture(0)   =   "FormGRPAdjustment.frx":14273
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Line13"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Line12"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Line11"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Line10"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Line9"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label74"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label73"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Label40"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Label45"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Label53"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Label55"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "Label90"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "Label91"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "Label1"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "Label2"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "Label3"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "Label5"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "Label6"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "Label7"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "Label8"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "Label9"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "Label10"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "Label11"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "Label12"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "Label13"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "Line1"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "Line2"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "Label14"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).Control(28)=   "Label42(0)"
         Tab(0).Control(28).Enabled=   0   'False
         Tab(0).Control(29)=   "Label43(0)"
         Tab(0).Control(29).Enabled=   0   'False
         Tab(0).Control(30)=   "Label52(0)"
         Tab(0).Control(30).Enabled=   0   'False
         Tab(0).Control(31)=   "Label54(0)"
         Tab(0).Control(31).Enabled=   0   'False
         Tab(0).Control(32)=   "txtNoEngine"
         Tab(0).Control(32).Enabled=   0   'False
         Tab(0).Control(33)=   "txtTglSvc"
         Tab(0).Control(33).Enabled=   0   'False
         Tab(0).Control(34)=   "txtkmSvc"
         Tab(0).Control(34).Enabled=   0   'False
         Tab(0).Control(35)=   "txtAlamat"
         Tab(0).Control(35).Enabled=   0   'False
         Tab(0).Control(36)=   "txtNama"
         Tab(0).Control(36).Enabled=   0   'False
         Tab(0).Control(37)=   "txtNoPol"
         Tab(0).Control(37).Enabled=   0   'False
         Tab(0).Control(38)=   "txtNoChassis"
         Tab(0).Control(38).Enabled=   0   'False
         Tab(0).Control(39)=   "txtGSM"
         Tab(0).Control(39).Enabled=   0   'False
         Tab(0).Control(40)=   "txtTelp"
         Tab(0).Control(40).Enabled=   0   'False
         Tab(0).Control(41)=   "txtCDMA"
         Tab(0).Control(41).Enabled=   0   'False
         Tab(0).Control(42)=   "txtNoWO(0)"
         Tab(0).Control(42).Enabled=   0   'False
         Tab(0).Control(43)=   "txtNoInv(0)"
         Tab(0).Control(43).Enabled=   0   'False
         Tab(0).ControlCount=   44
         TabCaption(1)   =   ".:: Job Order ::."
         TabPicture(1)   =   "FormGRPAdjustment.frx":1428F
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Line20"
         Tab(1).Control(1)=   "Label23"
         Tab(1).Control(2)=   "Label22"
         Tab(1).Control(3)=   "Label21"
         Tab(1).Control(4)=   "Label20"
         Tab(1).Control(5)=   "Label19"
         Tab(1).Control(6)=   "Label18"
         Tab(1).Control(7)=   "Label17"
         Tab(1).Control(8)=   "Label16"
         Tab(1).Control(9)=   "Label15"
         Tab(1).Control(10)=   "Line19"
         Tab(1).Control(11)=   "Line18"
         Tab(1).Control(12)=   "Line17"
         Tab(1).Control(13)=   "Line16"
         Tab(1).Control(14)=   "Line15"
         Tab(1).Control(15)=   "Line14"
         Tab(1).Control(16)=   "Label42(1)"
         Tab(1).Control(17)=   "Label43(1)"
         Tab(1).Control(18)=   "Label52(1)"
         Tab(1).Control(19)=   "Label54(1)"
         Tab(1).Control(20)=   "tRte(7)"
         Tab(1).Control(21)=   "tRte(6)"
         Tab(1).Control(22)=   "tRte(5)"
         Tab(1).Control(23)=   "tRte(4)"
         Tab(1).Control(24)=   "tRep(7)"
         Tab(1).Control(25)=   "tRep(6)"
         Tab(1).Control(26)=   "tRep(5)"
         Tab(1).Control(27)=   "tRep(4)"
         Tab(1).Control(28)=   "tRte(3)"
         Tab(1).Control(29)=   "tRte(2)"
         Tab(1).Control(30)=   "tRte(1)"
         Tab(1).Control(31)=   "tRte(0)"
         Tab(1).Control(32)=   "tRep(3)"
         Tab(1).Control(33)=   "tRep(2)"
         Tab(1).Control(34)=   "tRep(1)"
         Tab(1).Control(35)=   "tRep(0)"
         Tab(1).Control(36)=   "txtNoWO(1)"
         Tab(1).Control(37)=   "txtNoInv(1)"
         Tab(1).ControlCount=   38
         Begin VB.TextBox txtNoInv 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Index           =   0
            Left            =   6600
            MaxLength       =   30
            TabIndex        =   32
            Top             =   -5
            Width           =   1575
         End
         Begin VB.TextBox txtNoWO 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Index           =   0
            Left            =   9480
            MaxLength       =   30
            TabIndex        =   31
            Top             =   -5
            Width           =   1575
         End
         Begin VB.TextBox txtNoInv 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Index           =   1
            Left            =   -68400
            MaxLength       =   30
            TabIndex        =   30
            Top             =   0
            Width           =   1575
         End
         Begin VB.TextBox txtNoWO 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Index           =   1
            Left            =   -65640
            MaxLength       =   30
            TabIndex        =   29
            Top             =   0
            Width           =   1575
         End
         Begin VB.TextBox txtCDMA 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Left            =   9240
            MaxLength       =   25
            TabIndex        =   28
            Top             =   1800
            Width           =   1815
         End
         Begin VB.TextBox txtTelp 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   330
            Left            =   9240
            MaxLength       =   40
            TabIndex        =   27
            Top             =   1080
            Width           =   1815
         End
         Begin VB.TextBox txtGSM 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Left            =   9240
            MaxLength       =   25
            TabIndex        =   26
            Top             =   1440
            Width           =   1815
         End
         Begin VB.TextBox txtNoChassis 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Left            =   1560
            MaxLength       =   25
            TabIndex        =   25
            Top             =   1680
            Width           =   2655
         End
         Begin VB.TextBox txtNoPol 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Left            =   6360
            MaxLength       =   15
            TabIndex        =   24
            Top             =   2040
            Width           =   1575
         End
         Begin VB.TextBox txtNama 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Left            =   1560
            MaxLength       =   100
            TabIndex        =   23
            Top             =   600
            Width           =   6375
         End
         Begin VB.TextBox txtAlamat 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Left            =   1560
            MaxLength       =   200
            TabIndex        =   22
            Top             =   960
            Width           =   6375
         End
         Begin VB.TextBox txtkmSvc 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Left            =   6360
            MaxLength       =   25
            TabIndex        =   21
            Top             =   1680
            Width           =   1575
         End
         Begin VB.TextBox txtTglSvc 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   330
            Left            =   6360
            MaxLength       =   40
            TabIndex        =   20
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox txtNoEngine 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Left            =   1560
            MaxLength       =   25
            TabIndex        =   19
            Top             =   1320
            Width           =   2655
         End
         Begin VB.TextBox tRep 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
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
            Index           =   0
            Left            =   -74520
            MaxLength       =   100
            TabIndex        =   18
            Top             =   840
            Width           =   4425
         End
         Begin VB.TextBox tRep 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Index           =   1
            Left            =   -74520
            MaxLength       =   100
            TabIndex        =   17
            Top             =   1200
            Width           =   4425
         End
         Begin VB.TextBox tRep 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Index           =   2
            Left            =   -74520
            MaxLength       =   100
            TabIndex        =   16
            Top             =   1560
            Width           =   4425
         End
         Begin VB.TextBox tRep 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Index           =   3
            Left            =   -74520
            MaxLength       =   100
            TabIndex        =   15
            Top             =   1920
            Width           =   4425
         End
         Begin VB.TextBox tRte 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
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
            Index           =   0
            Left            =   -70080
            MaxLength       =   25
            TabIndex        =   14
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox tRte 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Index           =   1
            Left            =   -70080
            MaxLength       =   25
            TabIndex        =   13
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox tRte 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Index           =   2
            Left            =   -70080
            MaxLength       =   25
            TabIndex        =   12
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox tRte 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Index           =   3
            Left            =   -70080
            MaxLength       =   25
            TabIndex        =   11
            Top             =   1920
            Width           =   495
         End
         Begin VB.TextBox tRep 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Index           =   4
            Left            =   -69120
            MaxLength       =   100
            TabIndex        =   10
            Top             =   840
            Width           =   4425
         End
         Begin VB.TextBox tRep 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Index           =   5
            Left            =   -69120
            MaxLength       =   100
            TabIndex        =   9
            Top             =   1200
            Width           =   4425
         End
         Begin VB.TextBox tRep 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Index           =   6
            Left            =   -69120
            MaxLength       =   100
            TabIndex        =   8
            Top             =   1560
            Width           =   4425
         End
         Begin VB.TextBox tRep 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Index           =   7
            Left            =   -69120
            MaxLength       =   100
            TabIndex        =   7
            Top             =   1920
            Width           =   4425
         End
         Begin VB.TextBox tRte 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Index           =   4
            Left            =   -64680
            MaxLength       =   25
            TabIndex        =   6
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox tRte 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Index           =   5
            Left            =   -64680
            MaxLength       =   25
            TabIndex        =   5
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox tRte 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Index           =   6
            Left            =   -64680
            MaxLength       =   25
            TabIndex        =   4
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox tRte 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Index           =   7
            Left            =   -64680
            MaxLength       =   25
            TabIndex        =   3
            Top             =   1920
            Width           =   495
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
            Index           =   0
            Left            =   9360
            TabIndex        =   70
            Top             =   30
            Width           =   75
         End
         Begin VB.Label Label52 
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
            Index           =   0
            Left            =   6480
            TabIndex        =   69
            Top             =   30
            Width           =   75
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Invoice"
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
            Index           =   0
            Left            =   5520
            TabIndex        =   68
            Top             =   30
            Width           =   915
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No WO"
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
            Index           =   0
            Left            =   8640
            TabIndex        =   67
            Top             =   30
            Width           =   645
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
            Index           =   1
            Left            =   -65760
            TabIndex        =   66
            Top             =   30
            Width           =   75
         End
         Begin VB.Label Label52 
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
            Index           =   1
            Left            =   -68520
            TabIndex        =   65
            Top             =   30
            Width           =   75
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Invoice"
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
            Index           =   1
            Left            =   -69480
            TabIndex        =   64
            Top             =   30
            Width           =   915
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No WO"
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
            Index           =   1
            Left            =   -66480
            TabIndex        =   63
            Top             =   30
            Width           =   645
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ".:: TELPON ::."
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
            Left            =   8280
            TabIndex        =   62
            Top             =   675
            Width           =   2775
         End
         Begin VB.Line Line2 
            X1              =   8280
            X2              =   11040
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Line Line1 
            X1              =   8280
            X2              =   11040
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Height          =   195
            Left            =   8280
            TabIndex        =   61
            Top             =   1800
            Width           =   600
         End
         Begin VB.Label Label12 
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
            Left            =   9120
            TabIndex        =   60
            Top             =   1800
            Width           =   75
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PSTN"
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
            Left            =   8280
            TabIndex        =   59
            Top             =   1080
            Width           =   480
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Height          =   195
            Left            =   8280
            TabIndex        =   58
            Top             =   1440
            Width           =   435
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
            Left            =   9120
            TabIndex        =   57
            Top             =   1080
            Width           =   75
         End
         Begin VB.Label Label8 
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
            Left            =   9120
            TabIndex        =   56
            Top             =   1440
            Width           =   75
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Chassis"
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
            Left            =   120
            TabIndex        =   55
            Top             =   1680
            Width           =   1065
         End
         Begin VB.Label Label6 
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
            Left            =   1440
            TabIndex        =   54
            Top             =   1680
            Width           =   75
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
            Left            =   1440
            TabIndex        =   53
            Top             =   960
            Width           =   75
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
            Left            =   1440
            TabIndex        =   52
            Top             =   600
            Width           =   75
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alamat Cust"
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
            Left            =   120
            TabIndex        =   51
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama  Cust"
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
            Left            =   120
            TabIndex        =   50
            Top             =   600
            Width           =   1110
         End
         Begin VB.Label Label91 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nomor Polisi"
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
            Left            =   4920
            TabIndex        =   49
            Top             =   2040
            Width           =   1245
         End
         Begin VB.Label Label90 
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
            Left            =   6240
            TabIndex        =   48
            Top             =   2040
            Width           =   75
         End
         Begin VB.Label Label55 
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
            Left            =   6240
            TabIndex        =   47
            Top             =   1680
            Width           =   75
         End
         Begin VB.Label Label53 
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
            Left            =   6240
            TabIndex        =   46
            Top             =   1320
            Width           =   75
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Km Service"
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
            Left            =   4920
            TabIndex        =   45
            Top             =   1680
            Width           =   1065
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tgl Service"
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
            Left            =   4920
            TabIndex        =   44
            Top             =   1320
            Width           =   1065
         End
         Begin VB.Label Label73 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Engine"
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
            Left            =   120
            TabIndex        =   43
            Top             =   1320
            Width           =   990
         End
         Begin VB.Label Label74 
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
            Left            =   1440
            TabIndex        =   42
            Top             =   1320
            Width           =   75
         End
         Begin VB.Line Line9 
            BorderColor     =   &H000000FF&
            X1              =   2040
            X2              =   11040
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Line Line10 
            BorderColor     =   &H000000FF&
            X1              =   2040
            X2              =   2040
            Y1              =   480
            Y2              =   240
         End
         Begin VB.Line Line11 
            BorderColor     =   &H000000FF&
            X1              =   2040
            X2              =   1800
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Line Line12 
            BorderColor     =   &H000000FF&
            X1              =   360
            X2              =   120
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Line Line13 
            BorderColor     =   &H000000FF&
            X1              =   120
            X2              =   120
            Y1              =   240
            Y2              =   480
         End
         Begin VB.Line Line14 
            BorderColor     =   &H000000FF&
            X1              =   -70680
            X2              =   -63960
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Line Line15 
            BorderColor     =   &H000000FF&
            X1              =   -70680
            X2              =   -70680
            Y1              =   480
            Y2              =   240
         End
         Begin VB.Line Line16 
            BorderColor     =   &H000000FF&
            X1              =   -70680
            X2              =   -70920
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Line Line17 
            BorderColor     =   &H000000FF&
            X1              =   -72480
            X2              =   -72720
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Line Line18 
            BorderColor     =   &H000000FF&
            X1              =   -72720
            X2              =   -72720
            Y1              =   240
            Y2              =   480
         End
         Begin VB.Line Line19 
            BorderColor     =   &H000000FF&
            X1              =   -72720
            X2              =   -75000
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1."
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
            Left            =   -74760
            TabIndex        =   41
            Top             =   840
            Width           =   150
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2."
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
            Left            =   -74760
            TabIndex        =   40
            Top             =   1200
            Width           =   150
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "3."
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
            Left            =   -74760
            TabIndex        =   39
            Top             =   1560
            Width           =   150
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "4."
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
            Left            =   -74760
            TabIndex        =   38
            Top             =   1920
            Width           =   150
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "6."
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
            Left            =   -69360
            TabIndex        =   37
            Top             =   1200
            Width           =   150
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "7."
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
            Left            =   -69360
            TabIndex        =   36
            Top             =   1560
            Width           =   150
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "8."
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
            Left            =   -69360
            TabIndex        =   35
            Top             =   1920
            Width           =   150
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "5."
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
            Left            =   -69360
            TabIndex        =   34
            Top             =   840
            Width           =   150
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   $"FormGRPAdjustment.frx":142AB
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
            Left            =   -74880
            TabIndex        =   33
            Top             =   540
            Width           =   10635
         End
         Begin VB.Line Line20 
            BorderStyle     =   3  'Dot
            X1              =   -74880
            X2              =   -64200
            Y1              =   765
            Y2              =   765
         End
      End
      Begin TrueOleDBGrid70.TDBGrid GridGRP 
         Height          =   3615
         Left            =   0
         TabIndex        =   80
         Top             =   720
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   6376
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "No Invoice"
         Columns(0).DataField=   "NoInv"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "No WO"
         Columns(1).DataField=   "No_Wo"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Tgl Svc"
         Columns(2).DataField=   "Tanggal"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "KM Svc"
         Columns(3).DataField=   "Km"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Nama Customer"
         Columns(4).DataField=   "Customer"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Alamat Customer"
         Columns(5).DataField=   "Cust_adres"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "No Polisi"
         Columns(6).DataField=   "No_Polisi"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "No Engine"
         Columns(7).DataField=   "No_Engine"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "No Chassis"
         Columns(8).DataField=   "No_Chassis"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Telpon"
         Columns(9).DataField=   "Telpon"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "GSM"
         Columns(10).DataField=   "GSM"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "CDMA"
         Columns(11).DataField=   "CDMA"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   12
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=12"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2117"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2011"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2117"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2011"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=2646"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2540"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=2646"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2540"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=5292"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=5186"
         Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(21)=   "Column(5).Width=7064"
         Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=6959"
         Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(25)=   "Column(6).Width=2117"
         Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=2011"
         Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(29)=   "Column(7).Width=3281"
         Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=3175"
         Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(33)=   "Column(8).Width=3889"
         Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=3784"
         Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(37)=   "Column(9).Width=3281"
         Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=3175"
         Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(41)=   "Column(10).Width=3281"
         Splits(0)._ColumnProps(42)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(10)._WidthInPix=3175"
         Splits(0)._ColumnProps(44)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(45)=   "Column(11).Width=3281"
         Splits(0)._ColumnProps(46)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(47)=   "Column(11)._WidthInPix=3175"
         Splits(0)._ColumnProps(48)=   "Column(11).Order=12"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         WrapCellPointer =   -1  'True
         MultipleLines   =   0
         CellTipsWidth   =   0
         MultiSelect     =   2
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Verdana"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=Verdana"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
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
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HD8E9EC&"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=86,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=83,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=84,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=85,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=82,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=79,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=80,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=81,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=78,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=75,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=76,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=77,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=74,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=71,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=72,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=73,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=66,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=62,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=58,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=28,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=25,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=26,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=27,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=54,.parent=13"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=51,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=52,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=53,.parent=17"
         _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=50,.parent=13"
         _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=47,.parent=14"
         _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=48,.parent=15"
         _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=49,.parent=17"
         _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=46,.parent=13"
         _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=43,.parent=14"
         _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=44,.parent=15"
         _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=45,.parent=17"
         _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=32,.parent=13"
         _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=29,.parent=14"
         _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=30,.parent=15"
         _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=31,.parent=17"
         _StyleDefs(84)  =   "Named:id=33:Normal"
         _StyleDefs(85)  =   ":id=33,.parent=0"
         _StyleDefs(86)  =   "Named:id=34:Heading"
         _StyleDefs(87)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(88)  =   ":id=34,.wraptext=-1"
         _StyleDefs(89)  =   "Named:id=35:Footing"
         _StyleDefs(90)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(91)  =   "Named:id=36:Selected"
         _StyleDefs(92)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(93)  =   "Named:id=37:Caption"
         _StyleDefs(94)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(95)  =   "Named:id=38:HighlightRow"
         _StyleDefs(96)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(97)  =   "Named:id=39:EvenRow"
         _StyleDefs(98)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(99)  =   "Named:id=40:OddRow"
         _StyleDefs(100) =   ":id=40,.parent=33"
         _StyleDefs(101) =   "Named:id=41:RecordSelector"
         _StyleDefs(102) =   ":id=41,.parent=34"
         _StyleDefs(103) =   "Named:id=42:FilterBar"
         _StyleDefs(104) =   ":id=42,.parent=33"
      End
      Begin MSComDlg.CommonDialog ComDialogXls 
         Left            =   8760
         Top             =   7800
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame FrameGRP 
         Caption         =   ".:: SELECTED GRP ::."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   2775
         Left            =   0
         TabIndex        =   78
         Top             =   4440
         Width           =   11240
         Begin VB.Label lblMerk 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   2520
            TabIndex        =   79
            Top             =   360
            Width           =   75
         End
         Begin VB.Line Line4 
            BorderColor     =   &H000000FF&
            X1              =   120
            X2              =   11640
            Y1              =   2760
            Y2              =   2760
         End
         Begin VB.Line Line8 
            BorderColor     =   &H000000FF&
            X1              =   120
            X2              =   11280
            Y1              =   240
            Y2              =   240
         End
      End
      Begin TrueOleDBGrid70.TDBGrid GridHist 
         Height          =   2655
         Left            =   0
         TabIndex        =   86
         Top             =   9000
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   4683
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Tanggal"
         Columns(0).DataField=   "tgl"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Km"
         Columns(1).DataField=   "Km"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "SA"
         Columns(2).DataField=   "frt"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Mech"
         Columns(3).DataField=   "Mech"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Kode Dealer"
         Columns(4).DataField=   "KdDealer"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Tipe"
         Columns(5).DataField=   "tp"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Service"
         Columns(6).DataField=   "JobH"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
         Splits(0)._UserFlags=   0
         Splits(0).Locked=   -1  'True
         Splits(0).AllowSizing=   -1  'True
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AllowColSelect=   0   'False
         Splits(0).FetchRowStyle=   -1  'True
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=1773"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1667"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=1773"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1667"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=1773"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=1667"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=2646"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2540"
         Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(21)=   "Column(5).Width=1773"
         Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=1667"
         Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(25)=   "Column(6).Width=5292"
         Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=5186"
         Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
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
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
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
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HD8E9EC&"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=62,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=58,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=54,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=28,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=32,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=29,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=30,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=31,.parent=17"
         _StyleDefs(64)  =   "Named:id=33:Normal"
         _StyleDefs(65)  =   ":id=33,.parent=0"
         _StyleDefs(66)  =   "Named:id=34:Heading"
         _StyleDefs(67)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(68)  =   ":id=34,.wraptext=-1"
         _StyleDefs(69)  =   "Named:id=35:Footing"
         _StyleDefs(70)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(71)  =   "Named:id=36:Selected"
         _StyleDefs(72)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(73)  =   "Named:id=37:Caption"
         _StyleDefs(74)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(75)  =   "Named:id=38:HighlightRow"
         _StyleDefs(76)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(77)  =   "Named:id=39:EvenRow"
         _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(79)  =   "Named:id=40:OddRow"
         _StyleDefs(80)  =   ":id=40,.parent=33"
         _StyleDefs(81)  =   "Named:id=41:RecordSelector"
         _StyleDefs(82)  =   ":id=41,.parent=34"
         _StyleDefs(83)  =   "Named:id=42:FilterBar"
         _StyleDefs(84)  =   ":id=42,.parent=33"
      End
      Begin VB.Label lblJml 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   10080
         TabIndex        =   85
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label70 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Data  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   8640
         TabIndex        =   84
         Top             =   360
         Width           =   1455
      End
      Begin VB.Line Line22 
         BorderColor     =   &H000000FF&
         X1              =   2280
         X2              =   3960
         Y1              =   8760
         Y2              =   8760
      End
      Begin VB.Line Line21 
         BorderColor     =   &H000000FF&
         X1              =   1440
         X2              =   3960
         Y1              =   8880
         Y2              =   8880
      End
      Begin VB.Line Line7 
         BorderColor     =   &H000000FF&
         X1              =   2280
         X2              =   3960
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line6 
         BorderColor     =   &H000000FF&
         X1              =   1440
         X2              =   3960
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line5 
         BorderColor     =   &H000000FF&
         X1              =   240
         X2              =   3960
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   10200
         Shape           =   4  'Rounded Rectangle
         Top             =   7440
         Width           =   975
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   9240
         Shape           =   4  'Rounded Rectangle
         Top             =   7440
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "LIST GRP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   81
         Top             =   315
         Width           =   8655
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         Height          =   495
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   11235
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "HISTORY SERVICE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4320
         TabIndex        =   82
         Top             =   8685
         Width           =   7095
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         Height          =   375
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   8640
         Width           =   11235
      End
   End
End
Attribute VB_Name = "FormGRPAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================
'Motion Captured Data Provided By Acep Suryanto
'==============================================
Option Explicit

Private OnLoad As Boolean

Private objGRP As New ClsEstimasi
Private rsGRP As ADODB.Recordset
Private rsH As ADODB.Recordset

Private col As TrueOleDBGrid70.Column
Private cols As TrueOleDBGrid70.Columns
Private Urut As String
Private rsKm As ADODB.Recordset

Public Property Let OnLoaded(value As Boolean)
   OnLoad = value
End Property

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdTop_Click()
    Me.SSFUB.Top = -200
    Me.VScrollFUB.value = 1
End Sub

Private Sub CmdXls_Click()
1     On Error GoTo Err_Export
          Dim Er As New ErrInfo
2         If rsGRP.RecordCount = 0 Then
3             Er.HeadCap = "Export To Excel !"
4             Er.InfoCap = "List Data GRP Kosong !"
5             Beep
6             Er.show (vbModal)
7             Exit Sub
8             Exit Sub
9         End If
          Dim MyPath As String
          Dim onVal As String
10        MyPath = ""
          Dim f As New FormExplorer
11        f.show (vbModal)
12        MyPath = f.PathSaving
13        If MyPath = "" Then Exit Sub
14        If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
15        If f.FileName = "" Then Exit Sub
16        onVal = MyPath & f.FileName & ".XLS"
          
          ''20200513 REQ PAK KOKOK MINTA DITAMBAHKAN LOG UNTUK USER YG MENDOWNLOAD
          ''-----------------------------------------------------------------------------
17        sSQL = "insert into logDownloadExcelMRA(kodeuser,namamenu,tglExport,tglFilter1,tglFilter2,kriteria,totalrecord,destinasi,kompname)"
18        sSQL = sSQL & " values ('" & gSA & "','FormGRPAdjustment',getdate(),'1900-01-01',getdate(),'','" & rsGRP.RecordCount & "','" & onVal & "','" & ComptName & "') "
19        Call objO.InsertData(sSQL)
          
20         GridGRP.Columns(5).Visible = False
21         GridGRP.ExportToFile onVal, True
22         GridGRP.Columns(5).Visible = True
23                Er.HeadCap = "Export to excel !"
24                Er.InfoCap = "Data sudah di Export ke File Excel." & Chr(13) & _
                               "File : " & onVal
                  
25                Beep
26                Er.show (vbModal)
                  
27         Exit Sub
Err_Export:
28          If err.Number = 32755 Or err.Number = 91 Then GoTo Keluar
29          MsgBox err.Number, "Procedure : " & Me.Name & ".CmdXls_Click"
Keluar:
End Sub

Private Sub Form_Load()
    '' 20200512DEV
    If Log.disableExportMra = 1 Then
        If gTipeUser = "KABENG" Or gTipeUser = "ASSKABENG" Or gTipeUser = "ADMIN" Or gTipeUser = "HO" Then
            'cmdCSV.Enabled = True
            CmdXls.Enabled = True
        Else
            'cmdCSV.Enabled = False
            CmdXls.Enabled = False
        End If
    End If
    Call VScrollFUB_Change
   If OnLoad = False Then Exit Sub
   Call ListGRP_On_Adjust
   Call LastKmSBE
End Sub

Private Sub ListGRP_On_Adjust()
   GridGRP.DataSource = Nothing
   Set rsGRP = objGRP.GetDataGRP
   If rsGRP.RecordCount = 0 Then
      Call BindingNullGRP
      lblJml.Caption = "0"
      Exit Sub
   End If
   lblJml.Caption = rsGRP.RecordCount
   Set Me.GridGRP.DataSource = rsGRP
   rsGRP.MoveFirst
   Call BindToControlGRP(rsGRP)
End Sub

Private Sub BindingNullGRP()
   With Me
      .txtNoInv(0).Text = ""
      .txtNoInv(1).Text = ""
      .txtNoWO(0).Text = ""
      .txtNoWO(1).Text = ""
      .txtTglSvc.Text = ""
      .txtAlamat.Text = ""
      .txtCDMA.Text = ""
      .txtGSM.Text = ""
      .txtkmSvc.Text = ""
      .txtNama.Text = ""
      .txtNoChassis.Text = ""
      .txtNoEngine.Text = ""
      .txtNoPol.Text = ""
      .txtTelp.Text = ""
      .OptGRP.Enabled = False
      .OptSBE.Enabled = False
      .cmdAdjust.Enabled = False
   End With
End Sub

Private Sub BindToControlGRP(rs As ADODB.Recordset)
   With Me
    Set .txtNoInv(0).DataSource = rs
        txtNoInv(0).DataField = rs.Fields(0).Name
    Set .txtNoInv(1).DataSource = rs
        txtNoInv(1).DataField = rs.Fields(0).Name
    Set .txtNoWO(0).DataSource = rs
        txtNoWO(0).DataField = rs.Fields(1).Name
    Set .txtNoWO(1).DataSource = rs
        txtNoWO(1).DataField = rs.Fields(1).Name
    Set .txtTglSvc.DataSource = rs
        txtTglSvc.DataField = rs.Fields(2).Name
    Set .txtkmSvc.DataSource = rs
        txtkmSvc.DataField = rs.Fields(3).Name
    Set .txtNama.DataSource = rs
        txtNama.DataField = rs.Fields(4).Name
    Set .txtAlamat.DataSource = rs
        txtAlamat.DataField = rs.Fields(5).Name
    Set .txtNoPol.DataSource = rs
        txtNoPol.DataField = rs.Fields(6).Name
    Set .txtNoEngine.DataSource = rs
        txtNoEngine.DataField = rs.Fields(7).Name
    Set .txtNoChassis.DataSource = rs
        txtNoChassis.DataField = rs.Fields(8).Name
    Set .txtTelp.DataSource = rs
        txtTelp.DataField = rs.Fields(9).Name
    Set .txtGSM.DataSource = rs
        txtGSM.DataField = rs.Fields(10).Name
    Set .txtCDMA.DataSource = rs
        txtCDMA.DataField = rs.Fields(11).Name
        
            Dim nR As Integer
            For nR = 0 To 7
                Set .tRep(nR).DataSource = rs
                    tRep(nR).DataField = rs.Fields(nR + 12).Name
            Next nR
            For nR = 0 To 7
                Set .tRte(nR).DataSource = rs
                    tRte(nR).DataField = rs.Fields(nR + 20).Name
            Next nR
            
    .OptGRP.Enabled = True
    .OptSBE.Enabled = True
    .cmdAdjust.Enabled = True
    .OptGRP.value = False
    .OptSBE.value = False
    .ImgView.Picture = Nothing
   End With
End Sub

Private Sub LastKmSBE()
   Dim LastKm As Double
   cboKm.Clear
   Set rsKm = objGRP.GetDatakmSBE(txtNoChassis.Text, txtNoWO(0).Text)
   If rsKm.RecordCount = 0 Then
      cboKm.AddItem 1000
      LastKm = 5000
   Else
      rsKm.MoveFirst
      LastKm = Val(rsKm.Fields(0).value)
      LastKm = LastKm + 5000
   End If
   Dim i As Double
   For i = (LastKm - 5000) To 500000 Step 5000
      Me.cboKm.AddItem i
   Next i
End Sub

Private Sub GridGRP_FilterChange()
On Error GoTo err
   Set cols = GridGRP.Columns
   Dim C As Integer
   
   C = GridGRP.col
   GridGRP.HoldFields
   rsGRP.filter = getFilter()
   If rsGRP.EOF = True Then
      For Each col In GridGRP.Columns
         col.FilterText = ""
      Next col
      rsGRP.filter = ""
   End If
   GridGRP.col = C
   GridGRP.EditActive = True
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
         tmp = tmp & col.DataField & " LIKE '" & col.FilterText & "*'"
      End If
   Next col
   getFilter = tmp
   
   Exit Function
err:
   Resume Next
End Function

Private Sub GridGRP_HeadClick(ByVal ColIndex As Integer)
1     On Error GoTo err
2        Urut = IIf(Urut = "DESC", "ASC", "DESC")
3        If ColIndex <= 14 Then rsGRP.Sort = GridGRP.Columns(ColIndex).DataField & " " & Urut
4        GridGRP.SetFocus
5        Sendkeys "{Right " & Val(ColIndex) & "}"
6     Exit Sub
err:
7        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure getFilter of Form FormGRPAdjustment")
End Sub

Private Sub GridGRP_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo Keluar
   If rsGRP.RecordCount = 0 Then Exit Sub
   Call BindToControlGRP(rsGRP)
   Set rsH = objGRP.FetchAllHISTORY(Trim(txtNoChassis.Text))
   If rsH.RecordCount <> 0 Then Set GridHist.DataSource = rsH
Keluar:
End Sub

Private Sub OptGRP_Click()
   Me.ImgView.Picture = Nothing
   If OptGRP.value = True Then
      Me.ImgView.Picture = Me.ImageList.ListImages(1).Picture
      Me.cboKm.Visible = False
      Me.OptSBE.Caption = "Set as SBE    "
   End If
End Sub
Private Sub OptSBE_Click()
   Me.ImgView.Picture = Nothing
   If OptSBE.value = True Then
      Me.ImgView.Picture = Me.ImageList.ListImages(2).Picture
      Me.cboKm.Visible = True
      Me.OptSBE.Caption = "Set as SBE    KM :"
   End If
End Sub


Private Sub cmdAdjust_Click()
   If OptGRP.value = True Then
      If MsgBox("Set data To GRP ?", vbQuestion + vbYesNo) = vbYes Then
         'Delete Item GRP di M_ToSBE
         Call objGRP.DeleteItemGRP_On_ToSBE(txtNoChassis.Text, txtNoWO(0).Text)
      End If
   End If
   If OptSBE.value = True Then
      If MsgBox("Set data To SBE : " & Me.cboKm.Text & " ?", vbQuestion + vbYesNo) = vbYes Then
         'Delete Item GRP di M_ToSBE
         Call objGRP.DeleteItemGRP_On_ToSBE(txtNoChassis.Text, txtNoWO(0).Text)
         'Update Item History
         Dim xKm As String
         If Len(Mid(cboKm.Text, 1, Len(cboKm.Text) - 3)) = 2 Then
            xKm = "R0" & Mid(cboKm.Text, 1, Len(cboKm.Text) - 3)
         ElseIf Len(Mid(cboKm.Text, 1, Len(cboKm.Text) - 3)) = 3 Then
            xKm = "R" & Mid(cboKm.Text, 1, Len(cboKm.Text) - 3)
         End If
         'Update History
         Call objGRP.AdjustGRP_ItemHistory(txtNoChassis.Text, txtNoWO(0).Text, xKm)
      End If
   End If
   'Refresh Data
   Call ListGRP_On_Adjust
   Call LastKmSBE
End Sub

Private Sub tRep_Change(Index As Integer)
    If tRep(Index).Text <> "" Then tRep(Index).BackColor = &HFFFFFF
    If tRep(Index).Text = "" Then tRep(Index).BackColor = &HE0E0E0
End Sub
Private Sub tRte_Change(Index As Integer)
    If tRte(Index).Text <> "" Then tRte(Index).BackColor = &HFFFFFF
    If tRte(Index).Text = "" Then tRte(Index).BackColor = &HE0E0E0
End Sub

Private Sub VScrollFUB_Change()
    Me.SSFUB.Top = -(VScrollFUB.value + 200)
End Sub
Private Sub VScrollFUB_Scroll()
    Me.SSFUB.Top = -(VScrollFUB.value + 200)
End Sub
