VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmLoyalitiVerifikasiCustomer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Verifikasi Customer"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11190
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   11190
   Begin VB.Frame Frame1 
      Height          =   2910
      Left            =   3225
      TabIndex        =   6
      Top             =   1800
      Visible         =   0   'False
      Width           =   7020
      Begin VB.Frame Frame2 
         Caption         =   "DATA VERIFIKASI CUSTOMER"
         Height          =   1650
         Left            =   60
         TabIndex        =   8
         Top             =   165
         Width           =   6915
         Begin MSMask.MaskEdBox Text 
            Height          =   390
            Index           =   0
            Left            =   1665
            TabIndex        =   9
            Top             =   255
            Width           =   5160
            _ExtentX        =   9102
            _ExtentY        =   688
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   30
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   12
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
            DataField       =   "c"
            Height          =   390
            Index           =   2
            Left            =   1665
            TabIndex        =   10
            Top             =   1110
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   688
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "9-9-9-9-9"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   390
            Index           =   1
            Left            =   1665
            TabIndex        =   14
            Top             =   690
            Width           =   5160
            _ExtentX        =   9102
            _ExtentY        =   688
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   30
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   12
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
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "No HP"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   930
            TabIndex        =   15
            Top             =   750
            Width           =   675
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Verifikasi"
            DataField       =   "c"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   45
            TabIndex        =   13
            Top             =   1125
            Width           =   1560
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Email"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1020
            TabIndex        =   11
            Top             =   315
            Width           =   585
         End
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tutup"
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
         Left            =   6090
         Picture         =   "frmLoyalitiVerifikasiCustomer.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1950
         Width           =   735
      End
      Begin VB.CommandButton cmdSearch2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cek Data"
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
         Left            =   5010
         Picture         =   "frmLoyalitiVerifikasiCustomer.frx":0790
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1950
         Width           =   990
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   4905
         Shape           =   4  'Rounded Rectangle
         Top             =   1860
         Width           =   2040
      End
   End
   Begin VB.CommandButton CmdRefreshLap 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Refresh"
      Height          =   735
      Left            =   9525
      Picture         =   "frmLoyalitiVerifikasiCustomer.frx":0DDF
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5820
      Width           =   735
   End
   Begin VB.CommandButton cmdProcess 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pr&ocess"
      Height          =   732
      Left            =   8745
      Picture         =   "frmLoyalitiVerifikasiCustomer.frx":1443
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5820
      Width           =   732
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
      Left            =   10290
      Picture         =   "frmLoyalitiVerifikasiCustomer.frx":1AB0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5835
      Width           =   735
   End
   Begin TrueOleDBGrid70.TDBGrid gridCustomer 
      Height          =   4920
      Left            =   540
      TabIndex        =   16
      Top             =   660
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   8678
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
      Columns.Count   =   5
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).FetchRowStyle=   -1  'True
      Splits(0).DividerColor=   14933984
      Splits(0).FilterBar=   -1  'True
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
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "DAFTAR VALIDASI CUSTOMER"
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   12577002
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=54,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=32,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
      _StyleDefs(56)  =   "Named:id=33:Normal"
      _StyleDefs(57)  =   ":id=33,.parent=0"
      _StyleDefs(58)  =   "Named:id=34:Heading"
      _StyleDefs(59)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(60)  =   ":id=34,.wraptext=-1"
      _StyleDefs(61)  =   "Named:id=35:Footing"
      _StyleDefs(62)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(63)  =   "Named:id=36:Selected"
      _StyleDefs(64)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(65)  =   "Named:id=37:Caption"
      _StyleDefs(66)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(67)  =   "Named:id=38:HighlightRow"
      _StyleDefs(68)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(69)  =   "Named:id=39:EvenRow"
      _StyleDefs(70)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(71)  =   "Named:id=40:OddRow"
      _StyleDefs(72)  =   ":id=40,.parent=33"
      _StyleDefs(73)  =   "Named:id=41:RecordSelector"
      _StyleDefs(74)  =   ":id=41,.parent=34"
      _StyleDefs(75)  =   "Named:id=42:FilterBar"
      _StyleDefs(76)  =   ":id=42,.parent=33"
   End
   Begin VB.Label lblTotal 
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
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   2625
      TabIndex        =   4
      Top             =   5820
      Width           =   540
   End
   Begin VB.Label Label2 
      Caption         =   "TOTAL DATA = "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   675
      TabIndex        =   3
      Top             =   5865
      Width           =   1905
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   8625
      Shape           =   4  'Rounded Rectangle
      Top             =   5700
      Width           =   2475
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   ".:: VERIFIKASI CUSTOMER ::."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   555
      TabIndex        =   0
      Top             =   195
      Width           =   10500
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      FillColor       =   &H0080C0FF&
      FillStyle       =   0  'Solid
      Height          =   525
      Left            =   510
      Shape           =   4  'Rounded Rectangle
      Top             =   90
      Width           =   10605
   End
   Begin VB.Image Image1 
      Height          =   8325
      Left            =   0
      Picture         =   "frmLoyalitiVerifikasiCustomer.frx":2240
      Top             =   60
      Width           =   450
   End
End
Attribute VB_Name = "frmLoyalitiVerifikasiCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmLoyalitiVerifikasiCustomer
' DateTime  : 20/10/2017 14:31
' Author    : Rudi Kristanto
' Purpose   :
'---------------------------------------------------------------------------------------

Private rsTemp As ADODB.Recordset

Private Sub createHeader()
    With gridCustomer
        .Columns(0).Caption = "Email"
        .Columns(0).DataField = "alamatemail"
        .Columns(0).Width = 3000
        .Columns(0).Locked = True
        
        .Columns(1).Caption = "Email"
        .Columns(1).DataField = "alamatemail"
        .Columns(1).Width = 2000
        .Columns(1).Locked = True
        .Columns(1).Visible = False
        
        .Columns(2).Caption = "Nama Customer"
        .Columns(2).DataField = "Nama"
        .Columns(2).Width = 4000
        .Columns(2).Locked = True
        
        .Columns(3).Caption = "Telp"
        .Columns(3).DataField = "Telp"
        .Columns(3).Width = 2500
        .Columns(3).Locked = True
        
        .Columns(4).Caption = "KodeAktivasi"
        .Columns(4).DataField = "KodeAktivasi"
        .Columns(4).Width = 1500
        .Columns(4).Visible = False
    End With
End Sub

Private Sub refreshDatagrid()
1        On Error GoTo refreshDatagrid_Error

2         sSQL = "select AlamatEmail,Nama,Telp,KodeAktivasi from RPLNASAPP..MasterUser where [VerifikasiEmail] = 1 and Verifikasi = 0"
3         Set rsTemp = objO.FetchData(sSQL)
4         Call createHeader
5         gridCustomer.DataSource = rsTemp
6         gridCustomer.Refresh
7         lblTotal.Caption = rsTemp.RecordCount

8        On Error GoTo 0
9        Exit Sub

refreshDatagrid_Error:

10        MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure refreshDatagrid of Form frmLoyalitiVerifikasiCustomer"
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdProcess_Click()
    If MsgBox("Apakah yakin validasi data customer " & gridCustomer.Columns(2).value & " ?", vbYesNo) = vbYes Then
        text(0) = Trim(gridCustomer.Columns(0).value)
        text(1) = Trim(gridCustomer.Columns(3).value)
        'Text(2).SetFocus
        
        Frame1.Visible = True
    End If
End Sub

Private Sub CmdRefreshLap_Click()
    Call refreshDatagrid
End Sub

Private Sub cmdSearch2_Click()
    If Trim(Replace(text(2), "-", "")) = Trim(gridCustomer.Columns(4).value) Then
'        Call objO.InsertData("UPDATE RPLNASAPP..MASTERUSER SET verifikasiCbg = '1' WHERE alamatemail = '" & Trim(gridCustomer.Columns(0).value) & "' ")
'        Frame1.Visible = False
'        Call refreshDatagrid
        'Dim f As frmLoyalityUnitCustomer
        frmLoyalityUnitCustomer.emailCustomer = Trim(gridCustomer.Columns(0).value)
        frmLoyalityUnitCustomer.Left = 0
        frmLoyalityUnitCustomer.Top = 750
        frmLoyalityUnitCustomer.Show vbModal
        
        Frame1.Visible = False
        Call refreshDatagrid
        
'        frmLoyalityUnitCustomer.emailCustomer = Trim(gridCustomer.Columns(0).value)
'        frmLoyalityUnitCustomer.Left = 0
'        frmLoyalityUnitCustomer.Top = 0
'        frmLoyalityUnitCustomer.Show
        
    Else
        MsgBox "Kode verifikasi tidak cocok", vbInformation
    End If
End Sub

Private Sub Command1_Click()
    Frame1.Visible = False
    Call refreshDatagrid
End Sub

Private Sub Form_Load()
    Call refreshDatagrid
End Sub
