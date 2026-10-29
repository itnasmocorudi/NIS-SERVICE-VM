VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{EE757A1F-B0AC-40BC-9E72-B8651740F53E}#1.0#0"; "ARProgBar.ocx"
Begin VB.Form frmRepoBulananByModel 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   13845
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4200
      Top             =   1080
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   13575
      Begin VB.Label lblJudul 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   ".:: LAPORAN BULANAN BY MODEL ::."
         BeginProperty Font 
            Name            =   "Trajan Pro"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   200
         Width           =   13335
      End
   End
   Begin ARProgBarCtrl.ARProgressBar progKeterangan 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   6000
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   661
      Value           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   9340
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
      Columns.Count   =   71
      Splits(0)._UserFlags=   0
      Splits(0).AllowFocus=   0   'False
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   953
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AllowColSelect=   0   'False
      Splits(0).AllowRowSelect=   0   'False
      Splits(0).FetchRowStyle=   -1  'True
      Splits(0).DividerColor=   15790320
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=71"
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
      Splits(0)._ColumnProps(37)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(41)=   "Column(10).Width=2725"
      Splits(0)._ColumnProps(42)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(10)._WidthInPix=2646"
      Splits(0)._ColumnProps(44)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(45)=   "Column(11).Width=2725"
      Splits(0)._ColumnProps(46)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(47)=   "Column(11)._WidthInPix=2646"
      Splits(0)._ColumnProps(48)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(49)=   "Column(12).Width=2725"
      Splits(0)._ColumnProps(50)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(12)._WidthInPix=2646"
      Splits(0)._ColumnProps(52)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(53)=   "Column(13).Width=2725"
      Splits(0)._ColumnProps(54)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(55)=   "Column(13)._WidthInPix=2646"
      Splits(0)._ColumnProps(56)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(57)=   "Column(14).Width=2725"
      Splits(0)._ColumnProps(58)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(59)=   "Column(14)._WidthInPix=2646"
      Splits(0)._ColumnProps(60)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(61)=   "Column(15).Width=2725"
      Splits(0)._ColumnProps(62)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(63)=   "Column(15)._WidthInPix=2646"
      Splits(0)._ColumnProps(64)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(65)=   "Column(16).Width=2725"
      Splits(0)._ColumnProps(66)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(67)=   "Column(16)._WidthInPix=2646"
      Splits(0)._ColumnProps(68)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(69)=   "Column(17).Width=2725"
      Splits(0)._ColumnProps(70)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(71)=   "Column(17)._WidthInPix=2646"
      Splits(0)._ColumnProps(72)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(73)=   "Column(18).Width=2725"
      Splits(0)._ColumnProps(74)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(75)=   "Column(18)._WidthInPix=2646"
      Splits(0)._ColumnProps(76)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(77)=   "Column(19).Width=2725"
      Splits(0)._ColumnProps(78)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(79)=   "Column(19)._WidthInPix=2646"
      Splits(0)._ColumnProps(80)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(81)=   "Column(20).Width=2725"
      Splits(0)._ColumnProps(82)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(83)=   "Column(20)._WidthInPix=2646"
      Splits(0)._ColumnProps(84)=   "Column(20).Order=21"
      Splits(0)._ColumnProps(85)=   "Column(21).Width=2725"
      Splits(0)._ColumnProps(86)=   "Column(21).DividerColor=0"
      Splits(0)._ColumnProps(87)=   "Column(21)._WidthInPix=2646"
      Splits(0)._ColumnProps(88)=   "Column(21).Order=22"
      Splits(0)._ColumnProps(89)=   "Column(22).Width=2725"
      Splits(0)._ColumnProps(90)=   "Column(22).DividerColor=0"
      Splits(0)._ColumnProps(91)=   "Column(22)._WidthInPix=2646"
      Splits(0)._ColumnProps(92)=   "Column(22).Order=23"
      Splits(0)._ColumnProps(93)=   "Column(23).Width=2725"
      Splits(0)._ColumnProps(94)=   "Column(23).DividerColor=0"
      Splits(0)._ColumnProps(95)=   "Column(23)._WidthInPix=2646"
      Splits(0)._ColumnProps(96)=   "Column(23).Order=24"
      Splits(0)._ColumnProps(97)=   "Column(24).Width=2725"
      Splits(0)._ColumnProps(98)=   "Column(24).DividerColor=0"
      Splits(0)._ColumnProps(99)=   "Column(24)._WidthInPix=2646"
      Splits(0)._ColumnProps(100)=   "Column(24).Order=25"
      Splits(0)._ColumnProps(101)=   "Column(25).Width=2725"
      Splits(0)._ColumnProps(102)=   "Column(25).DividerColor=0"
      Splits(0)._ColumnProps(103)=   "Column(25)._WidthInPix=2646"
      Splits(0)._ColumnProps(104)=   "Column(25).Order=26"
      Splits(0)._ColumnProps(105)=   "Column(26).Width=2725"
      Splits(0)._ColumnProps(106)=   "Column(26).DividerColor=0"
      Splits(0)._ColumnProps(107)=   "Column(26)._WidthInPix=2646"
      Splits(0)._ColumnProps(108)=   "Column(26).Order=27"
      Splits(0)._ColumnProps(109)=   "Column(27).Width=2725"
      Splits(0)._ColumnProps(110)=   "Column(27).DividerColor=0"
      Splits(0)._ColumnProps(111)=   "Column(27)._WidthInPix=2646"
      Splits(0)._ColumnProps(112)=   "Column(27).Order=28"
      Splits(0)._ColumnProps(113)=   "Column(28).Width=2725"
      Splits(0)._ColumnProps(114)=   "Column(28).DividerColor=0"
      Splits(0)._ColumnProps(115)=   "Column(28)._WidthInPix=2646"
      Splits(0)._ColumnProps(116)=   "Column(28).Order=29"
      Splits(0)._ColumnProps(117)=   "Column(29).Width=2725"
      Splits(0)._ColumnProps(118)=   "Column(29).DividerColor=0"
      Splits(0)._ColumnProps(119)=   "Column(29)._WidthInPix=2646"
      Splits(0)._ColumnProps(120)=   "Column(29).Order=30"
      Splits(0)._ColumnProps(121)=   "Column(30).Width=2725"
      Splits(0)._ColumnProps(122)=   "Column(30).DividerColor=0"
      Splits(0)._ColumnProps(123)=   "Column(30)._WidthInPix=2646"
      Splits(0)._ColumnProps(124)=   "Column(30).Order=31"
      Splits(0)._ColumnProps(125)=   "Column(31).Width=2725"
      Splits(0)._ColumnProps(126)=   "Column(31).DividerColor=0"
      Splits(0)._ColumnProps(127)=   "Column(31)._WidthInPix=2646"
      Splits(0)._ColumnProps(128)=   "Column(31).Order=32"
      Splits(0)._ColumnProps(129)=   "Column(32).Width=2725"
      Splits(0)._ColumnProps(130)=   "Column(32).DividerColor=0"
      Splits(0)._ColumnProps(131)=   "Column(32)._WidthInPix=2646"
      Splits(0)._ColumnProps(132)=   "Column(32).Order=33"
      Splits(0)._ColumnProps(133)=   "Column(33).Width=2725"
      Splits(0)._ColumnProps(134)=   "Column(33).DividerColor=0"
      Splits(0)._ColumnProps(135)=   "Column(33)._WidthInPix=2646"
      Splits(0)._ColumnProps(136)=   "Column(33).Order=34"
      Splits(0)._ColumnProps(137)=   "Column(34).Width=2725"
      Splits(0)._ColumnProps(138)=   "Column(34).DividerColor=0"
      Splits(0)._ColumnProps(139)=   "Column(34)._WidthInPix=2646"
      Splits(0)._ColumnProps(140)=   "Column(34).Order=35"
      Splits(0)._ColumnProps(141)=   "Column(35).Width=2725"
      Splits(0)._ColumnProps(142)=   "Column(35).DividerColor=0"
      Splits(0)._ColumnProps(143)=   "Column(35)._WidthInPix=2646"
      Splits(0)._ColumnProps(144)=   "Column(35).Order=36"
      Splits(0)._ColumnProps(145)=   "Column(36).Width=2725"
      Splits(0)._ColumnProps(146)=   "Column(36).DividerColor=0"
      Splits(0)._ColumnProps(147)=   "Column(36)._WidthInPix=2646"
      Splits(0)._ColumnProps(148)=   "Column(36).Order=37"
      Splits(0)._ColumnProps(149)=   "Column(37).Width=2725"
      Splits(0)._ColumnProps(150)=   "Column(37).DividerColor=0"
      Splits(0)._ColumnProps(151)=   "Column(37)._WidthInPix=2646"
      Splits(0)._ColumnProps(152)=   "Column(37).Order=38"
      Splits(0)._ColumnProps(153)=   "Column(38).Width=2725"
      Splits(0)._ColumnProps(154)=   "Column(38).DividerColor=0"
      Splits(0)._ColumnProps(155)=   "Column(38)._WidthInPix=2646"
      Splits(0)._ColumnProps(156)=   "Column(38).Order=39"
      Splits(0)._ColumnProps(157)=   "Column(39).Width=2725"
      Splits(0)._ColumnProps(158)=   "Column(39).DividerColor=0"
      Splits(0)._ColumnProps(159)=   "Column(39)._WidthInPix=2646"
      Splits(0)._ColumnProps(160)=   "Column(39).Order=40"
      Splits(0)._ColumnProps(161)=   "Column(40).Width=2725"
      Splits(0)._ColumnProps(162)=   "Column(40).DividerColor=0"
      Splits(0)._ColumnProps(163)=   "Column(40)._WidthInPix=2646"
      Splits(0)._ColumnProps(164)=   "Column(40).Order=41"
      Splits(0)._ColumnProps(165)=   "Column(41).Width=2725"
      Splits(0)._ColumnProps(166)=   "Column(41).DividerColor=0"
      Splits(0)._ColumnProps(167)=   "Column(41)._WidthInPix=2646"
      Splits(0)._ColumnProps(168)=   "Column(41).Order=42"
      Splits(0)._ColumnProps(169)=   "Column(42).Width=2725"
      Splits(0)._ColumnProps(170)=   "Column(42).DividerColor=0"
      Splits(0)._ColumnProps(171)=   "Column(42)._WidthInPix=2646"
      Splits(0)._ColumnProps(172)=   "Column(42).Order=43"
      Splits(0)._ColumnProps(173)=   "Column(43).Width=2725"
      Splits(0)._ColumnProps(174)=   "Column(43).DividerColor=0"
      Splits(0)._ColumnProps(175)=   "Column(43)._WidthInPix=2646"
      Splits(0)._ColumnProps(176)=   "Column(43).Order=44"
      Splits(0)._ColumnProps(177)=   "Column(44).Width=2725"
      Splits(0)._ColumnProps(178)=   "Column(44).DividerColor=0"
      Splits(0)._ColumnProps(179)=   "Column(44)._WidthInPix=2646"
      Splits(0)._ColumnProps(180)=   "Column(44).Order=45"
      Splits(0)._ColumnProps(181)=   "Column(45).Width=2725"
      Splits(0)._ColumnProps(182)=   "Column(45).DividerColor=0"
      Splits(0)._ColumnProps(183)=   "Column(45)._WidthInPix=2646"
      Splits(0)._ColumnProps(184)=   "Column(45).Order=46"
      Splits(0)._ColumnProps(185)=   "Column(46).Width=2725"
      Splits(0)._ColumnProps(186)=   "Column(46).DividerColor=0"
      Splits(0)._ColumnProps(187)=   "Column(46)._WidthInPix=2646"
      Splits(0)._ColumnProps(188)=   "Column(46).Order=47"
      Splits(0)._ColumnProps(189)=   "Column(47).Width=2725"
      Splits(0)._ColumnProps(190)=   "Column(47).DividerColor=0"
      Splits(0)._ColumnProps(191)=   "Column(47)._WidthInPix=2646"
      Splits(0)._ColumnProps(192)=   "Column(47).Order=48"
      Splits(0)._ColumnProps(193)=   "Column(48).Width=2725"
      Splits(0)._ColumnProps(194)=   "Column(48).DividerColor=0"
      Splits(0)._ColumnProps(195)=   "Column(48)._WidthInPix=2646"
      Splits(0)._ColumnProps(196)=   "Column(48).Order=49"
      Splits(0)._ColumnProps(197)=   "Column(49).Width=2725"
      Splits(0)._ColumnProps(198)=   "Column(49).DividerColor=0"
      Splits(0)._ColumnProps(199)=   "Column(49)._WidthInPix=2646"
      Splits(0)._ColumnProps(200)=   "Column(49).Order=50"
      Splits(0)._ColumnProps(201)=   "Column(50).Width=2725"
      Splits(0)._ColumnProps(202)=   "Column(50).DividerColor=0"
      Splits(0)._ColumnProps(203)=   "Column(50)._WidthInPix=2646"
      Splits(0)._ColumnProps(204)=   "Column(50).Order=51"
      Splits(0)._ColumnProps(205)=   "Column(51).Width=2725"
      Splits(0)._ColumnProps(206)=   "Column(51).DividerColor=0"
      Splits(0)._ColumnProps(207)=   "Column(51)._WidthInPix=2646"
      Splits(0)._ColumnProps(208)=   "Column(51).Order=52"
      Splits(0)._ColumnProps(209)=   "Column(52).Width=2725"
      Splits(0)._ColumnProps(210)=   "Column(52).DividerColor=0"
      Splits(0)._ColumnProps(211)=   "Column(52)._WidthInPix=2646"
      Splits(0)._ColumnProps(212)=   "Column(52).Order=53"
      Splits(0)._ColumnProps(213)=   "Column(53).Width=2725"
      Splits(0)._ColumnProps(214)=   "Column(53).DividerColor=0"
      Splits(0)._ColumnProps(215)=   "Column(53)._WidthInPix=2646"
      Splits(0)._ColumnProps(216)=   "Column(53).Order=54"
      Splits(0)._ColumnProps(217)=   "Column(54).Width=2725"
      Splits(0)._ColumnProps(218)=   "Column(54).DividerColor=0"
      Splits(0)._ColumnProps(219)=   "Column(54)._WidthInPix=2646"
      Splits(0)._ColumnProps(220)=   "Column(54).Order=55"
      Splits(0)._ColumnProps(221)=   "Column(55).Width=2725"
      Splits(0)._ColumnProps(222)=   "Column(55).DividerColor=0"
      Splits(0)._ColumnProps(223)=   "Column(55)._WidthInPix=2646"
      Splits(0)._ColumnProps(224)=   "Column(55).Order=56"
      Splits(0)._ColumnProps(225)=   "Column(56).Width=2725"
      Splits(0)._ColumnProps(226)=   "Column(56).DividerColor=0"
      Splits(0)._ColumnProps(227)=   "Column(56)._WidthInPix=2646"
      Splits(0)._ColumnProps(228)=   "Column(56).Order=57"
      Splits(0)._ColumnProps(229)=   "Column(57).Width=2725"
      Splits(0)._ColumnProps(230)=   "Column(57).DividerColor=0"
      Splits(0)._ColumnProps(231)=   "Column(57)._WidthInPix=2646"
      Splits(0)._ColumnProps(232)=   "Column(57).Order=58"
      Splits(0)._ColumnProps(233)=   "Column(58).Width=2725"
      Splits(0)._ColumnProps(234)=   "Column(58).DividerColor=0"
      Splits(0)._ColumnProps(235)=   "Column(58)._WidthInPix=2646"
      Splits(0)._ColumnProps(236)=   "Column(58).Order=59"
      Splits(0)._ColumnProps(237)=   "Column(59).Width=2725"
      Splits(0)._ColumnProps(238)=   "Column(59).DividerColor=0"
      Splits(0)._ColumnProps(239)=   "Column(59)._WidthInPix=2646"
      Splits(0)._ColumnProps(240)=   "Column(59).Order=60"
      Splits(0)._ColumnProps(241)=   "Column(60).Width=2725"
      Splits(0)._ColumnProps(242)=   "Column(60).DividerColor=0"
      Splits(0)._ColumnProps(243)=   "Column(60)._WidthInPix=2646"
      Splits(0)._ColumnProps(244)=   "Column(60).Order=61"
      Splits(0)._ColumnProps(245)=   "Column(61).Width=2725"
      Splits(0)._ColumnProps(246)=   "Column(61).DividerColor=0"
      Splits(0)._ColumnProps(247)=   "Column(61)._WidthInPix=2646"
      Splits(0)._ColumnProps(248)=   "Column(61).Order=62"
      Splits(0)._ColumnProps(249)=   "Column(62).Width=2725"
      Splits(0)._ColumnProps(250)=   "Column(62).DividerColor=0"
      Splits(0)._ColumnProps(251)=   "Column(62)._WidthInPix=2646"
      Splits(0)._ColumnProps(252)=   "Column(62).Order=63"
      Splits(0)._ColumnProps(253)=   "Column(63).Width=2725"
      Splits(0)._ColumnProps(254)=   "Column(63).DividerColor=0"
      Splits(0)._ColumnProps(255)=   "Column(63)._WidthInPix=2646"
      Splits(0)._ColumnProps(256)=   "Column(63).Order=64"
      Splits(0)._ColumnProps(257)=   "Column(64).Width=2725"
      Splits(0)._ColumnProps(258)=   "Column(64).DividerColor=0"
      Splits(0)._ColumnProps(259)=   "Column(64)._WidthInPix=2646"
      Splits(0)._ColumnProps(260)=   "Column(64).Order=65"
      Splits(0)._ColumnProps(261)=   "Column(65).Width=2725"
      Splits(0)._ColumnProps(262)=   "Column(65).DividerColor=0"
      Splits(0)._ColumnProps(263)=   "Column(65)._WidthInPix=2646"
      Splits(0)._ColumnProps(264)=   "Column(65).Order=66"
      Splits(0)._ColumnProps(265)=   "Column(66).Width=2725"
      Splits(0)._ColumnProps(266)=   "Column(66).DividerColor=0"
      Splits(0)._ColumnProps(267)=   "Column(66)._WidthInPix=2646"
      Splits(0)._ColumnProps(268)=   "Column(66).Order=67"
      Splits(0)._ColumnProps(269)=   "Column(67).Width=2725"
      Splits(0)._ColumnProps(270)=   "Column(67).DividerColor=0"
      Splits(0)._ColumnProps(271)=   "Column(67)._WidthInPix=2646"
      Splits(0)._ColumnProps(272)=   "Column(67).Order=68"
      Splits(0)._ColumnProps(273)=   "Column(68).Width=2725"
      Splits(0)._ColumnProps(274)=   "Column(68).DividerColor=0"
      Splits(0)._ColumnProps(275)=   "Column(68)._WidthInPix=2646"
      Splits(0)._ColumnProps(276)=   "Column(68).Order=69"
      Splits(0)._ColumnProps(277)=   "Column(69).Width=2725"
      Splits(0)._ColumnProps(278)=   "Column(69).DividerColor=0"
      Splits(0)._ColumnProps(279)=   "Column(69)._WidthInPix=2646"
      Splits(0)._ColumnProps(280)=   "Column(69).Order=70"
      Splits(0)._ColumnProps(281)=   "Column(70).Width=2725"
      Splits(0)._ColumnProps(282)=   "Column(70).DividerColor=0"
      Splits(0)._ColumnProps(283)=   "Column(70)._WidthInPix=2646"
      Splits(0)._ColumnProps(284)=   "Column(70).Order=71"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      ColumnHeaders   =   0   'False
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      AllowArrows     =   0   'False
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   15790320
      RowDividerColor =   15790320
      RowSubDividerColor=   15790320
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=174,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=171,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=172,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=173,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=170,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=167,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=168,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=169,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=166,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=163,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=164,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=165,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=162,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=159,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=160,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=161,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=158,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=155,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=156,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=157,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=154,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=151,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=152,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=153,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=150,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=147,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=148,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=149,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=146,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=143,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=144,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=145,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=142,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=139,.parent=14"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=140,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=141,.parent=17"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=138,.parent=13"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=135,.parent=14"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=136,.parent=15"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=137,.parent=17"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=134,.parent=13"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=131,.parent=14"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=132,.parent=15"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=133,.parent=17"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=130,.parent=13"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=127,.parent=14"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=128,.parent=15"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=129,.parent=17"
      _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=126,.parent=13"
      _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=123,.parent=14"
      _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=124,.parent=15"
      _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=125,.parent=17"
      _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=122,.parent=13"
      _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=119,.parent=14"
      _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=120,.parent=15"
      _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=121,.parent=17"
      _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=118,.parent=13"
      _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=115,.parent=14"
      _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=116,.parent=15"
      _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=117,.parent=17"
      _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=114,.parent=13"
      _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=111,.parent=14"
      _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=112,.parent=15"
      _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=113,.parent=17"
      _StyleDefs(100) =   "Splits(0).Columns(16).Style:id=110,.parent=13"
      _StyleDefs(101) =   "Splits(0).Columns(16).HeadingStyle:id=107,.parent=14"
      _StyleDefs(102) =   "Splits(0).Columns(16).FooterStyle:id=108,.parent=15"
      _StyleDefs(103) =   "Splits(0).Columns(16).EditorStyle:id=109,.parent=17"
      _StyleDefs(104) =   "Splits(0).Columns(17).Style:id=106,.parent=13"
      _StyleDefs(105) =   "Splits(0).Columns(17).HeadingStyle:id=103,.parent=14"
      _StyleDefs(106) =   "Splits(0).Columns(17).FooterStyle:id=104,.parent=15"
      _StyleDefs(107) =   "Splits(0).Columns(17).EditorStyle:id=105,.parent=17"
      _StyleDefs(108) =   "Splits(0).Columns(18).Style:id=102,.parent=13"
      _StyleDefs(109) =   "Splits(0).Columns(18).HeadingStyle:id=99,.parent=14"
      _StyleDefs(110) =   "Splits(0).Columns(18).FooterStyle:id=100,.parent=15"
      _StyleDefs(111) =   "Splits(0).Columns(18).EditorStyle:id=101,.parent=17"
      _StyleDefs(112) =   "Splits(0).Columns(19).Style:id=98,.parent=13"
      _StyleDefs(113) =   "Splits(0).Columns(19).HeadingStyle:id=95,.parent=14"
      _StyleDefs(114) =   "Splits(0).Columns(19).FooterStyle:id=96,.parent=15"
      _StyleDefs(115) =   "Splits(0).Columns(19).EditorStyle:id=97,.parent=17"
      _StyleDefs(116) =   "Splits(0).Columns(20).Style:id=94,.parent=13"
      _StyleDefs(117) =   "Splits(0).Columns(20).HeadingStyle:id=91,.parent=14"
      _StyleDefs(118) =   "Splits(0).Columns(20).FooterStyle:id=92,.parent=15"
      _StyleDefs(119) =   "Splits(0).Columns(20).EditorStyle:id=93,.parent=17"
      _StyleDefs(120) =   "Splits(0).Columns(21).Style:id=90,.parent=13"
      _StyleDefs(121) =   "Splits(0).Columns(21).HeadingStyle:id=87,.parent=14"
      _StyleDefs(122) =   "Splits(0).Columns(21).FooterStyle:id=88,.parent=15"
      _StyleDefs(123) =   "Splits(0).Columns(21).EditorStyle:id=89,.parent=17"
      _StyleDefs(124) =   "Splits(0).Columns(22).Style:id=86,.parent=13"
      _StyleDefs(125) =   "Splits(0).Columns(22).HeadingStyle:id=83,.parent=14"
      _StyleDefs(126) =   "Splits(0).Columns(22).FooterStyle:id=84,.parent=15"
      _StyleDefs(127) =   "Splits(0).Columns(22).EditorStyle:id=85,.parent=17"
      _StyleDefs(128) =   "Splits(0).Columns(23).Style:id=82,.parent=13"
      _StyleDefs(129) =   "Splits(0).Columns(23).HeadingStyle:id=79,.parent=14"
      _StyleDefs(130) =   "Splits(0).Columns(23).FooterStyle:id=80,.parent=15"
      _StyleDefs(131) =   "Splits(0).Columns(23).EditorStyle:id=81,.parent=17"
      _StyleDefs(132) =   "Splits(0).Columns(24).Style:id=78,.parent=13"
      _StyleDefs(133) =   "Splits(0).Columns(24).HeadingStyle:id=75,.parent=14"
      _StyleDefs(134) =   "Splits(0).Columns(24).FooterStyle:id=76,.parent=15"
      _StyleDefs(135) =   "Splits(0).Columns(24).EditorStyle:id=77,.parent=17"
      _StyleDefs(136) =   "Splits(0).Columns(25).Style:id=74,.parent=13"
      _StyleDefs(137) =   "Splits(0).Columns(25).HeadingStyle:id=71,.parent=14"
      _StyleDefs(138) =   "Splits(0).Columns(25).FooterStyle:id=72,.parent=15"
      _StyleDefs(139) =   "Splits(0).Columns(25).EditorStyle:id=73,.parent=17"
      _StyleDefs(140) =   "Splits(0).Columns(26).Style:id=70,.parent=13"
      _StyleDefs(141) =   "Splits(0).Columns(26).HeadingStyle:id=67,.parent=14"
      _StyleDefs(142) =   "Splits(0).Columns(26).FooterStyle:id=68,.parent=15"
      _StyleDefs(143) =   "Splits(0).Columns(26).EditorStyle:id=69,.parent=17"
      _StyleDefs(144) =   "Splits(0).Columns(27).Style:id=66,.parent=13"
      _StyleDefs(145) =   "Splits(0).Columns(27).HeadingStyle:id=63,.parent=14"
      _StyleDefs(146) =   "Splits(0).Columns(27).FooterStyle:id=64,.parent=15"
      _StyleDefs(147) =   "Splits(0).Columns(27).EditorStyle:id=65,.parent=17"
      _StyleDefs(148) =   "Splits(0).Columns(28).Style:id=62,.parent=13"
      _StyleDefs(149) =   "Splits(0).Columns(28).HeadingStyle:id=59,.parent=14"
      _StyleDefs(150) =   "Splits(0).Columns(28).FooterStyle:id=60,.parent=15"
      _StyleDefs(151) =   "Splits(0).Columns(28).EditorStyle:id=61,.parent=17"
      _StyleDefs(152) =   "Splits(0).Columns(29).Style:id=58,.parent=13"
      _StyleDefs(153) =   "Splits(0).Columns(29).HeadingStyle:id=55,.parent=14"
      _StyleDefs(154) =   "Splits(0).Columns(29).FooterStyle:id=56,.parent=15"
      _StyleDefs(155) =   "Splits(0).Columns(29).EditorStyle:id=57,.parent=17"
      _StyleDefs(156) =   "Splits(0).Columns(30).Style:id=54,.parent=13"
      _StyleDefs(157) =   "Splits(0).Columns(30).HeadingStyle:id=51,.parent=14"
      _StyleDefs(158) =   "Splits(0).Columns(30).FooterStyle:id=52,.parent=15"
      _StyleDefs(159) =   "Splits(0).Columns(30).EditorStyle:id=53,.parent=17"
      _StyleDefs(160) =   "Splits(0).Columns(31).Style:id=50,.parent=13"
      _StyleDefs(161) =   "Splits(0).Columns(31).HeadingStyle:id=47,.parent=14"
      _StyleDefs(162) =   "Splits(0).Columns(31).FooterStyle:id=48,.parent=15"
      _StyleDefs(163) =   "Splits(0).Columns(31).EditorStyle:id=49,.parent=17"
      _StyleDefs(164) =   "Splits(0).Columns(32).Style:id=46,.parent=13"
      _StyleDefs(165) =   "Splits(0).Columns(32).HeadingStyle:id=43,.parent=14"
      _StyleDefs(166) =   "Splits(0).Columns(32).FooterStyle:id=44,.parent=15"
      _StyleDefs(167) =   "Splits(0).Columns(32).EditorStyle:id=45,.parent=17"
      _StyleDefs(168) =   "Splits(0).Columns(33).Style:id=28,.parent=13"
      _StyleDefs(169) =   "Splits(0).Columns(33).HeadingStyle:id=25,.parent=14"
      _StyleDefs(170) =   "Splits(0).Columns(33).FooterStyle:id=26,.parent=15"
      _StyleDefs(171) =   "Splits(0).Columns(33).EditorStyle:id=27,.parent=17"
      _StyleDefs(172) =   "Splits(0).Columns(34).Style:id=32,.parent=13"
      _StyleDefs(173) =   "Splits(0).Columns(34).HeadingStyle:id=29,.parent=14"
      _StyleDefs(174) =   "Splits(0).Columns(34).FooterStyle:id=30,.parent=15"
      _StyleDefs(175) =   "Splits(0).Columns(34).EditorStyle:id=31,.parent=17"
      _StyleDefs(176) =   "Splits(0).Columns(35).Style:id=194,.parent=13"
      _StyleDefs(177) =   "Splits(0).Columns(35).HeadingStyle:id=191,.parent=14"
      _StyleDefs(178) =   "Splits(0).Columns(35).FooterStyle:id=192,.parent=15"
      _StyleDefs(179) =   "Splits(0).Columns(35).EditorStyle:id=193,.parent=17"
      _StyleDefs(180) =   "Splits(0).Columns(36).Style:id=198,.parent=13"
      _StyleDefs(181) =   "Splits(0).Columns(36).HeadingStyle:id=195,.parent=14"
      _StyleDefs(182) =   "Splits(0).Columns(36).FooterStyle:id=196,.parent=15"
      _StyleDefs(183) =   "Splits(0).Columns(36).EditorStyle:id=197,.parent=17"
      _StyleDefs(184) =   "Splits(0).Columns(37).Style:id=202,.parent=13"
      _StyleDefs(185) =   "Splits(0).Columns(37).HeadingStyle:id=199,.parent=14"
      _StyleDefs(186) =   "Splits(0).Columns(37).FooterStyle:id=200,.parent=15"
      _StyleDefs(187) =   "Splits(0).Columns(37).EditorStyle:id=201,.parent=17"
      _StyleDefs(188) =   "Splits(0).Columns(38).Style:id=206,.parent=13"
      _StyleDefs(189) =   "Splits(0).Columns(38).HeadingStyle:id=203,.parent=14"
      _StyleDefs(190) =   "Splits(0).Columns(38).FooterStyle:id=204,.parent=15"
      _StyleDefs(191) =   "Splits(0).Columns(38).EditorStyle:id=205,.parent=17"
      _StyleDefs(192) =   "Splits(0).Columns(39).Style:id=210,.parent=13"
      _StyleDefs(193) =   "Splits(0).Columns(39).HeadingStyle:id=207,.parent=14"
      _StyleDefs(194) =   "Splits(0).Columns(39).FooterStyle:id=208,.parent=15"
      _StyleDefs(195) =   "Splits(0).Columns(39).EditorStyle:id=209,.parent=17"
      _StyleDefs(196) =   "Splits(0).Columns(40).Style:id=214,.parent=13"
      _StyleDefs(197) =   "Splits(0).Columns(40).HeadingStyle:id=211,.parent=14"
      _StyleDefs(198) =   "Splits(0).Columns(40).FooterStyle:id=212,.parent=15"
      _StyleDefs(199) =   "Splits(0).Columns(40).EditorStyle:id=213,.parent=17"
      _StyleDefs(200) =   "Splits(0).Columns(41).Style:id=218,.parent=13"
      _StyleDefs(201) =   "Splits(0).Columns(41).HeadingStyle:id=215,.parent=14"
      _StyleDefs(202) =   "Splits(0).Columns(41).FooterStyle:id=216,.parent=15"
      _StyleDefs(203) =   "Splits(0).Columns(41).EditorStyle:id=217,.parent=17"
      _StyleDefs(204) =   "Splits(0).Columns(42).Style:id=222,.parent=13"
      _StyleDefs(205) =   "Splits(0).Columns(42).HeadingStyle:id=219,.parent=14"
      _StyleDefs(206) =   "Splits(0).Columns(42).FooterStyle:id=220,.parent=15"
      _StyleDefs(207) =   "Splits(0).Columns(42).EditorStyle:id=221,.parent=17"
      _StyleDefs(208) =   "Splits(0).Columns(43).Style:id=226,.parent=13"
      _StyleDefs(209) =   "Splits(0).Columns(43).HeadingStyle:id=223,.parent=14"
      _StyleDefs(210) =   "Splits(0).Columns(43).FooterStyle:id=224,.parent=15"
      _StyleDefs(211) =   "Splits(0).Columns(43).EditorStyle:id=225,.parent=17"
      _StyleDefs(212) =   "Splits(0).Columns(44).Style:id=230,.parent=13"
      _StyleDefs(213) =   "Splits(0).Columns(44).HeadingStyle:id=227,.parent=14"
      _StyleDefs(214) =   "Splits(0).Columns(44).FooterStyle:id=228,.parent=15"
      _StyleDefs(215) =   "Splits(0).Columns(44).EditorStyle:id=229,.parent=17"
      _StyleDefs(216) =   "Splits(0).Columns(45).Style:id=234,.parent=13"
      _StyleDefs(217) =   "Splits(0).Columns(45).HeadingStyle:id=231,.parent=14"
      _StyleDefs(218) =   "Splits(0).Columns(45).FooterStyle:id=232,.parent=15"
      _StyleDefs(219) =   "Splits(0).Columns(45).EditorStyle:id=233,.parent=17"
      _StyleDefs(220) =   "Splits(0).Columns(46).Style:id=238,.parent=13"
      _StyleDefs(221) =   "Splits(0).Columns(46).HeadingStyle:id=235,.parent=14"
      _StyleDefs(222) =   "Splits(0).Columns(46).FooterStyle:id=236,.parent=15"
      _StyleDefs(223) =   "Splits(0).Columns(46).EditorStyle:id=237,.parent=17"
      _StyleDefs(224) =   "Splits(0).Columns(47).Style:id=242,.parent=13"
      _StyleDefs(225) =   "Splits(0).Columns(47).HeadingStyle:id=239,.parent=14"
      _StyleDefs(226) =   "Splits(0).Columns(47).FooterStyle:id=240,.parent=15"
      _StyleDefs(227) =   "Splits(0).Columns(47).EditorStyle:id=241,.parent=17"
      _StyleDefs(228) =   "Splits(0).Columns(48).Style:id=246,.parent=13"
      _StyleDefs(229) =   "Splits(0).Columns(48).HeadingStyle:id=243,.parent=14"
      _StyleDefs(230) =   "Splits(0).Columns(48).FooterStyle:id=244,.parent=15"
      _StyleDefs(231) =   "Splits(0).Columns(48).EditorStyle:id=245,.parent=17"
      _StyleDefs(232) =   "Splits(0).Columns(49).Style:id=250,.parent=13"
      _StyleDefs(233) =   "Splits(0).Columns(49).HeadingStyle:id=247,.parent=14"
      _StyleDefs(234) =   "Splits(0).Columns(49).FooterStyle:id=248,.parent=15"
      _StyleDefs(235) =   "Splits(0).Columns(49).EditorStyle:id=249,.parent=17"
      _StyleDefs(236) =   "Splits(0).Columns(50).Style:id=254,.parent=13"
      _StyleDefs(237) =   "Splits(0).Columns(50).HeadingStyle:id=251,.parent=14"
      _StyleDefs(238) =   "Splits(0).Columns(50).FooterStyle:id=252,.parent=15"
      _StyleDefs(239) =   "Splits(0).Columns(50).EditorStyle:id=253,.parent=17"
      _StyleDefs(240) =   "Splits(0).Columns(51).Style:id=258,.parent=13"
      _StyleDefs(241) =   "Splits(0).Columns(51).HeadingStyle:id=255,.parent=14"
      _StyleDefs(242) =   "Splits(0).Columns(51).FooterStyle:id=256,.parent=15"
      _StyleDefs(243) =   "Splits(0).Columns(51).EditorStyle:id=257,.parent=17"
      _StyleDefs(244) =   "Splits(0).Columns(52).Style:id=262,.parent=13"
      _StyleDefs(245) =   "Splits(0).Columns(52).HeadingStyle:id=259,.parent=14"
      _StyleDefs(246) =   "Splits(0).Columns(52).FooterStyle:id=260,.parent=15"
      _StyleDefs(247) =   "Splits(0).Columns(52).EditorStyle:id=261,.parent=17"
      _StyleDefs(248) =   "Splits(0).Columns(53).Style:id=266,.parent=13"
      _StyleDefs(249) =   "Splits(0).Columns(53).HeadingStyle:id=263,.parent=14"
      _StyleDefs(250) =   "Splits(0).Columns(53).FooterStyle:id=264,.parent=15"
      _StyleDefs(251) =   "Splits(0).Columns(53).EditorStyle:id=265,.parent=17"
      _StyleDefs(252) =   "Splits(0).Columns(54).Style:id=270,.parent=13"
      _StyleDefs(253) =   "Splits(0).Columns(54).HeadingStyle:id=267,.parent=14"
      _StyleDefs(254) =   "Splits(0).Columns(54).FooterStyle:id=268,.parent=15"
      _StyleDefs(255) =   "Splits(0).Columns(54).EditorStyle:id=269,.parent=17"
      _StyleDefs(256) =   "Splits(0).Columns(55).Style:id=274,.parent=13"
      _StyleDefs(257) =   "Splits(0).Columns(55).HeadingStyle:id=271,.parent=14"
      _StyleDefs(258) =   "Splits(0).Columns(55).FooterStyle:id=272,.parent=15"
      _StyleDefs(259) =   "Splits(0).Columns(55).EditorStyle:id=273,.parent=17"
      _StyleDefs(260) =   "Splits(0).Columns(56).Style:id=278,.parent=13"
      _StyleDefs(261) =   "Splits(0).Columns(56).HeadingStyle:id=275,.parent=14"
      _StyleDefs(262) =   "Splits(0).Columns(56).FooterStyle:id=276,.parent=15"
      _StyleDefs(263) =   "Splits(0).Columns(56).EditorStyle:id=277,.parent=17"
      _StyleDefs(264) =   "Splits(0).Columns(57).Style:id=282,.parent=13"
      _StyleDefs(265) =   "Splits(0).Columns(57).HeadingStyle:id=279,.parent=14"
      _StyleDefs(266) =   "Splits(0).Columns(57).FooterStyle:id=280,.parent=15"
      _StyleDefs(267) =   "Splits(0).Columns(57).EditorStyle:id=281,.parent=17"
      _StyleDefs(268) =   "Splits(0).Columns(58).Style:id=286,.parent=13"
      _StyleDefs(269) =   "Splits(0).Columns(58).HeadingStyle:id=283,.parent=14"
      _StyleDefs(270) =   "Splits(0).Columns(58).FooterStyle:id=284,.parent=15"
      _StyleDefs(271) =   "Splits(0).Columns(58).EditorStyle:id=285,.parent=17"
      _StyleDefs(272) =   "Splits(0).Columns(59).Style:id=290,.parent=13"
      _StyleDefs(273) =   "Splits(0).Columns(59).HeadingStyle:id=287,.parent=14"
      _StyleDefs(274) =   "Splits(0).Columns(59).FooterStyle:id=288,.parent=15"
      _StyleDefs(275) =   "Splits(0).Columns(59).EditorStyle:id=289,.parent=17"
      _StyleDefs(276) =   "Splits(0).Columns(60).Style:id=294,.parent=13"
      _StyleDefs(277) =   "Splits(0).Columns(60).HeadingStyle:id=291,.parent=14"
      _StyleDefs(278) =   "Splits(0).Columns(60).FooterStyle:id=292,.parent=15"
      _StyleDefs(279) =   "Splits(0).Columns(60).EditorStyle:id=293,.parent=17"
      _StyleDefs(280) =   "Splits(0).Columns(61).Style:id=298,.parent=13"
      _StyleDefs(281) =   "Splits(0).Columns(61).HeadingStyle:id=295,.parent=14"
      _StyleDefs(282) =   "Splits(0).Columns(61).FooterStyle:id=296,.parent=15"
      _StyleDefs(283) =   "Splits(0).Columns(61).EditorStyle:id=297,.parent=17"
      _StyleDefs(284) =   "Splits(0).Columns(62).Style:id=302,.parent=13"
      _StyleDefs(285) =   "Splits(0).Columns(62).HeadingStyle:id=299,.parent=14"
      _StyleDefs(286) =   "Splits(0).Columns(62).FooterStyle:id=300,.parent=15"
      _StyleDefs(287) =   "Splits(0).Columns(62).EditorStyle:id=301,.parent=17"
      _StyleDefs(288) =   "Splits(0).Columns(63).Style:id=306,.parent=13"
      _StyleDefs(289) =   "Splits(0).Columns(63).HeadingStyle:id=303,.parent=14"
      _StyleDefs(290) =   "Splits(0).Columns(63).FooterStyle:id=304,.parent=15"
      _StyleDefs(291) =   "Splits(0).Columns(63).EditorStyle:id=305,.parent=17"
      _StyleDefs(292) =   "Splits(0).Columns(64).Style:id=310,.parent=13"
      _StyleDefs(293) =   "Splits(0).Columns(64).HeadingStyle:id=307,.parent=14"
      _StyleDefs(294) =   "Splits(0).Columns(64).FooterStyle:id=308,.parent=15"
      _StyleDefs(295) =   "Splits(0).Columns(64).EditorStyle:id=309,.parent=17"
      _StyleDefs(296) =   "Splits(0).Columns(65).Style:id=314,.parent=13"
      _StyleDefs(297) =   "Splits(0).Columns(65).HeadingStyle:id=311,.parent=14"
      _StyleDefs(298) =   "Splits(0).Columns(65).FooterStyle:id=312,.parent=15"
      _StyleDefs(299) =   "Splits(0).Columns(65).EditorStyle:id=313,.parent=17"
      _StyleDefs(300) =   "Splits(0).Columns(66).Style:id=318,.parent=13"
      _StyleDefs(301) =   "Splits(0).Columns(66).HeadingStyle:id=315,.parent=14"
      _StyleDefs(302) =   "Splits(0).Columns(66).FooterStyle:id=316,.parent=15"
      _StyleDefs(303) =   "Splits(0).Columns(66).EditorStyle:id=317,.parent=17"
      _StyleDefs(304) =   "Splits(0).Columns(67).Style:id=178,.parent=13"
      _StyleDefs(305) =   "Splits(0).Columns(67).HeadingStyle:id=175,.parent=14"
      _StyleDefs(306) =   "Splits(0).Columns(67).FooterStyle:id=176,.parent=15"
      _StyleDefs(307) =   "Splits(0).Columns(67).EditorStyle:id=177,.parent=17"
      _StyleDefs(308) =   "Splits(0).Columns(68).Style:id=182,.parent=13"
      _StyleDefs(309) =   "Splits(0).Columns(68).HeadingStyle:id=179,.parent=14"
      _StyleDefs(310) =   "Splits(0).Columns(68).FooterStyle:id=180,.parent=15"
      _StyleDefs(311) =   "Splits(0).Columns(68).EditorStyle:id=181,.parent=17"
      _StyleDefs(312) =   "Splits(0).Columns(69).Style:id=186,.parent=13"
      _StyleDefs(313) =   "Splits(0).Columns(69).HeadingStyle:id=183,.parent=14"
      _StyleDefs(314) =   "Splits(0).Columns(69).FooterStyle:id=184,.parent=15"
      _StyleDefs(315) =   "Splits(0).Columns(69).EditorStyle:id=185,.parent=17"
      _StyleDefs(316) =   "Splits(0).Columns(70).Style:id=190,.parent=13"
      _StyleDefs(317) =   "Splits(0).Columns(70).HeadingStyle:id=187,.parent=14"
      _StyleDefs(318) =   "Splits(0).Columns(70).FooterStyle:id=188,.parent=15"
      _StyleDefs(319) =   "Splits(0).Columns(70).EditorStyle:id=189,.parent=17"
      _StyleDefs(320) =   "Named:id=33:Normal"
      _StyleDefs(321) =   ":id=33,.parent=0"
      _StyleDefs(322) =   "Named:id=34:Heading"
      _StyleDefs(323) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(324) =   ":id=34,.wraptext=-1,.bold=-1,.fontsize=975,.italic=0,.underline=0"
      _StyleDefs(325) =   ":id=34,.strikethrough=0,.charset=0"
      _StyleDefs(326) =   ":id=34,.fontname=MS Sans Serif"
      _StyleDefs(327) =   "Named:id=35:Footing"
      _StyleDefs(328) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(329) =   "Named:id=36:Selected"
      _StyleDefs(330) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(331) =   "Named:id=37:Caption"
      _StyleDefs(332) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(333) =   "Named:id=38:HighlightRow"
      _StyleDefs(334) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(335) =   "Named:id=39:EvenRow"
      _StyleDefs(336) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(337) =   "Named:id=40:OddRow"
      _StyleDefs(338) =   ":id=40,.parent=33"
      _StyleDefs(339) =   "Named:id=41:RecordSelector"
      _StyleDefs(340) =   ":id=41,.parent=34"
      _StyleDefs(341) =   "Named:id=42:FilterBar"
      _StyleDefs(342) =   ":id=42,.parent=33"
      _StyleDefs(343) =   "Named:id=319:NormalBold"
      _StyleDefs(344) =   ":id=319,.parent=33,.bold=-1,.fontsize=975,.italic=0,.underline=0"
      _StyleDefs(345) =   ":id=319,.strikethrough=0,.charset=0"
      _StyleDefs(346) =   ":id=319,.fontname=MS Sans Serif"
   End
   Begin VB.CommandButton CmdXls 
      BackColor       =   &H00FFFFFF&
      Caption         =   "X&LS"
      Enabled         =   0   'False
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
      Left            =   120
      Picture         =   "frmRepoBulananByModel.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6000
      Width           =   735
   End
   Begin VB.Label lblKeterangan 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   6480
      Visible         =   0   'False
      Width           =   6615
   End
End
Attribute VB_Name = "frmRepoBulananByModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsTemp     As New ADODB.Recordset
Dim sSQL       As String
Dim rsReport   As New ADODB.Recordset

Public Bulan   As Integer
Public Tahun   As Integer


Public Sub createTableReport(total As Integer)
    If rsReport.State Then rsReport.Close
    Set rsReport = Nothing
    With rsReport
        .Fields.Append "URUT", adInteger
        .Fields.Append "KETERANGAN", adVarChar, 200
        .Fields.Append "BOLD", adBoolean
        For i = 1 To total
            .Fields.Append "Unit" & i, adVarChar, 100
        Next
        .Fields.Append "HEADER", adBoolean
        .Open
    End With
    
End Sub

Private Sub HeaderTable(total As Integer)
With TDBGrid1
        Dim i As Integer
        Dim kolom As Integer
        
         i = 0
         Do While i <= 70
             .Splits(0).Columns(i).Visible = False
             i = i + 1
         Loop
        
        .Columns(0).Width = 500
        .Columns(0).Caption = "URUT"
        .Columns(0).DataField = "URUT"
        .Splits(0).Columns(0).Locked = True
        .Splits(0).Columns(0).Visible = True
        
        .Columns(1).Width = 3500
        .Columns(1).Caption = "KETERANGAN"
        .Columns(1).DataField = "KETERANGAN"
        .Splits(0).Columns(1).Locked = True
        .Splits(0).Columns(1).Visible = True
        
        i = 2
        kolom = 1
        Do While i <= (total + 2)
            .Columns(i).Width = 2500
            .Columns(i).Caption = "UNIT" & kolom
            .Columns(i).DataField = "Unit" & kolom
            .Splits(0).Columns(i).Visible = True
            i = i + 1
            kolom = kolom + 1
         Loop
 
    End With
End Sub

Private Sub prosesPerhitungan()
    progKeterangan.Visible = True
    lblKeterangan.Visible = True
    lblKeterangan.Caption = "Pengambilan Data ..."
    progKeterangan.Min = 1
    progKeterangan.Max = 10
    Call objO.DeleteData("Delete FROM BatchBulananByModel Where Bulan = '" & Bulan & "' and Tahun = '" & Tahun & "'")
                
    progKeterangan.value = 2
    DoEvents
                
   '6 Agustus 2012
    Call objO.InsertData("Insert Into BatchBulananByModel exec repobulanan3A_ByModel_FS_GRP_MDL " & Bulan & "," & Tahun & "")
    progKeterangan.value = 4
    DoEvents
    
    Call objO.InsertData("Insert Into BatchBulananByModel exec repobulanan3B_ByModel_FS_GRP_MDL " & Bulan & "," & Tahun & "")
    progKeterangan.value = 6
    DoEvents
    
    Call objO.InsertData("Insert Into BatchBulananByModel exec repobulanan3C_ByModelFS_GRP_MDL " & Bulan & "," & Tahun & "")
    progKeterangan.value = 8
    DoEvents
    
    Call objO.InsertData("Insert Into BatchBulananByModel exec repobulanan3D_ByModel_FS_GRP_MDL " & Bulan & "," & Tahun & "")
    progKeterangan.value = 9
    DoEvents
    
    Call objO.InsertData("Insert Into BatchBulananByModel exec repobulanan3E_ByModel_FS_GRP_MDL " & Bulan & "," & Tahun & "")
    progKeterangan.value = 10
    DoEvents
    
    progKeterangan.Visible = False
    lblKeterangan.Visible = False
    Timer1.Enabled = False
End Sub

Private Sub CmdXls_Click()
1        On Error GoTo Err_Export
          Dim Er As New ErrInfo
2         If rsReport.RecordCount = 0 Then
3             Er.HeadCap = "Export To Excel !"
4             Er.InfoCap = "List Data Reminder Service Kosong !"
5             Beep
6             Er.Show (vbModal)
7             Exit Sub
8         End If
          Dim MyPath As String
          Dim onVal As String
9         MyPath = ""
          Dim f As New FormExplorer
10        f.Show (vbModal)
11        MyPath = f.PathSaving
12        If MyPath = "" Then Exit Sub
13        If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
14        If f.FileName = "" Then Exit Sub
15        onVal = MyPath & f.FileName & ".XLS"
          
16         TDBGrid1.ExportToFile onVal, True
17                Er.HeadCap = "Export to excel !"
18                Er.InfoCap = "Data sudah di Export ke File Excel." & Chr(13) & _
                               "File : " & onVal
                  
19                Beep
20                Er.Show (vbModal)
21         Exit Sub
Err_Export:
22          If err.Number = 32755 Or err.Number = 91 Then GoTo Keluar
23          MsgBox err.Number, "Procedure : " & Me.Name & ".CmdXls_Click"
Keluar:

End Sub

Private Sub Form_Load()
   Timer1.Enabled = True
End Sub

Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
   Dim rs As New ADODB.Recordset
    Set rs = rsReport.Clone
    If Bookmark > 1 Then
        rs.Bookmark = Bookmark
    End If
    If rs("header") = True Then
      RowStyle = "Heading"
    ElseIf rs!Bold = True Then
      RowStyle = "NormalBold"
    Else
      RowStyle = "Normal"
    End If
End Sub

Private Sub Timer1_Timer()
   lblJudul.Caption = ".:: LAPORAN BULANAN BY MODEL " & GetNamaBulan(Bulan, 1) & " - " & Tahun & " ::."
      
   If Log.RPLHistori = 0 Then
        Set rsTemp = objO.FetchData("select distinct fs_grp_mdl from model with (nolock) where FS_LOG_DEL = '' ORDER BY FS_GRP_MDL ")
   Else
        Set rsTemp = objO.FetchData("select distinct fs_grp_mdl from [" & dbReplikasiHO & "]..model with (nolock) where FS_LOG_DEL = '' ORDER BY FS_GRP_MDL ")
   End If
   If Not rsTemp.EOF Then
      Dim i As Integer
      Dim help As String
      i = 1
      Dim totalKolom As Integer
      
      Call prosesPerhitungan
      
      totalKolom = rsTemp.RecordCount '' mengambil total kolom
      
      Call createTableReport(rsTemp.RecordCount)
      Call HeaderTable(rsTemp.RecordCount)
      With rsReport
         .AddNew
         !Urut = 1
         !keterangan = "KETERANGAN"
         !Bold = False
         !Header = True
         
         Do While Not rsTemp.EOF
            help = "Unit" & i
            .Fields(help) = rsTemp!fs_grp_mdl
            i = i + 1
            rsTemp.MoveNext
         Loop
         .Update
      End With
      
      Call prosesRekapData(totalKolom)
      CmdXls.Enabled = True
   End If
   TDBGrid1.DataSource = rsReport
   TDBGrid1.FetchRowStyle = True
End Sub

Private Sub prosesRekapData(totalKolom As Integer)
   Dim nTotSBE(100)           As Double
   Dim nTotGR(100)            As Double
   Dim totCPUS(100)           As Double
   Dim nonCPUS(100)           As Double
   Dim totUE(100)             As Double
   Dim labor(100)             As Double
   
   Set rsTemp = objO.FetchData("Select Keterangan, SUM(JUMLAH) AS JUMLAH, Bulan, TAHUN,case isnull(KODE_KATEGORY,'Others') when '' then 'Others' when 'Others' then 'Others' else isnull(KODE_KATEGORY,'Others') end as KODE_KATEGORY from BatchBulananBymodel  Where Bulan = '" & Bulan & "' and Tahun = '" & Tahun & "' group by BULAN, Tahun, KETERANGAN,KODE_KATEGORY")
   If Not rsTemp.EOF Then
      Dim i                As Integer
      Dim nUrut            As Integer
      Dim rsPencarian      As New ADODB.Recordset
      Dim rsHeader         As New ADODB.Recordset
      Dim kolom            As Integer
      Dim help             As String
      Dim nilaiKolom       As Double
      
      Set rsHeader = rsReport.Clone
      progKeterangan.Visible = True
      progKeterangan.Min = 1
      progKeterangan.Max = 30
      
      lblKeterangan.Visible = True
      lblKeterangan.Caption = "Proses Rekap Data.."
      i = 0
      Do While i <= 30
         With rsReport
            .AddNew
            Select Case i
               Case 1
                   nUrut = nI
                   Set rsPencarian = rsTemp.Clone
                   nKeterangan = "10K saja "
                   rsPencarian.filter = "KETERANGAN = '10'"
                   kolom = 1
                     !Urut = 1
                     !keterangan = nKeterangan
                     !Bold = False
                     !Header = False
                     For kolom = 1 To totalKolom
                        help = "Unit" & kolom
                        nilaiKolom = 0
                        nTotSBE(kolom) = 0
                        If Not rsPencarian.EOF Then
                           rsPencarian.MoveFirst
                           Do While Not rsPencarian.EOF
                              If Trim(rsPencarian!kode_kategory) = rsHeader.Fields(help) Then
                                 nilaiKolom = nilaiKolom + CDbl(rsPencarian("JUMLAH"))
                              End If
                              rsPencarian.MoveNext
                           Loop
                           rsPencarian.MoveFirst
                        End If
                        .Fields(help) = nilaiKolom
                        nTotSBE(kolom) = nTotSBE(kolom) + CDbl(nilaiKolom)
                        
                     Next
                   progKeterangan.value = 2
                   DoEvents
               Case 2
                   nUrut = nI
                   Set rsPencarian = rsTemp.Clone
                   nKeterangan = "20K saja "
                   rsPencarian.filter = "KETERANGAN = '20'"
                   kolom = 1
                     !Urut = 2
                     !keterangan = nKeterangan
                     !Bold = False
                     !Header = False
                     For kolom = 1 To totalKolom
                        help = "Unit" & kolom
                        nilaiKolom = 0
                        'nTotSBE(kolom) = 0
                        If Not rsPencarian.EOF Then
                           rsPencarian.MoveFirst
                           Do While Not rsPencarian.EOF
                              If Trim(rsPencarian!kode_kategory) = rsHeader.Fields(help) Then
                                 nilaiKolom = nilaiKolom + CDbl(rsPencarian("JUMLAH"))
                              End If
                              rsPencarian.MoveNext
                           Loop
                           rsPencarian.MoveFirst
                        End If
                        .Fields(help) = nilaiKolom
                        nTotSBE(kolom) = nTotSBE(kolom) + CDbl(nilaiKolom)
                        
                     Next
                   progKeterangan.value = 3
                   DoEvents
                   
               Case 3
                   nUrut = nI
                   Set rsPencarian = rsTemp.Clone
                   nKeterangan = "30K saja "
                   rsPencarian.filter = "KETERANGAN = '30'"
                   kolom = 1
                     !Urut = 3
                     !keterangan = nKeterangan
                     !Bold = False
                     !Header = False
                     For kolom = 1 To totalKolom
                        help = "Unit" & kolom
                        nilaiKolom = 0
                        'nTotSBE(kolom) = 0
                        If Not rsPencarian.EOF Then
                           rsPencarian.MoveFirst
                           Do While Not rsPencarian.EOF
                              If Trim(rsPencarian!kode_kategory) = rsHeader.Fields(help) Then
                                 nilaiKolom = nilaiKolom + CDbl(rsPencarian("JUMLAH"))
                              End If
                              rsPencarian.MoveNext
                           Loop
                           rsPencarian.MoveFirst
                        End If
                        .Fields(help) = nilaiKolom
                        nTotSBE(kolom) = nTotSBE(kolom) + CDbl(nilaiKolom)
                        
                     Next
                   progKeterangan.value = 4
                   DoEvents
               Case 4
                   nUrut = nI
                   Set rsPencarian = rsTemp.Clone
                   nKeterangan = "40K saja "
                   rsPencarian.filter = "KETERANGAN = '40'"
                   kolom = 1
                     !Urut = 4
                     !keterangan = nKeterangan
                     !Bold = False
                     !Header = False
                     For kolom = 1 To totalKolom
                        help = "Unit" & kolom
                        nilaiKolom = 0
                        'nTotSBE(kolom) = 0
                        If Not rsPencarian.EOF Then
                           rsPencarian.MoveFirst
                           Do While Not rsPencarian.EOF
                              If Trim(rsPencarian!kode_kategory) = rsHeader.Fields(help) Then
                                 nilaiKolom = nilaiKolom + CDbl(rsPencarian("JUMLAH"))
                              End If
                              rsPencarian.MoveNext
                           Loop
                           rsPencarian.MoveFirst
                        End If
                        .Fields(help) = nilaiKolom
                        nTotSBE(kolom) = nTotSBE(kolom) + CDbl(nilaiKolom)
                        
                     Next
                   progKeterangan.value = 4
                   DoEvents
               Case 5
                   nUrut = nI
                   Set rsPencarian = rsTemp.Clone
                   nKeterangan = "50K saja "
                   rsPencarian.filter = "KETERANGAN = '50'"
                   kolom = 1
                     !Urut = 5
                     !keterangan = nKeterangan
                     !Bold = False
                     !Header = False
                     For kolom = 1 To totalKolom
                        help = "Unit" & kolom
                        nilaiKolom = 0
                        'nTotSBE(kolom) = 0
                        If Not rsPencarian.EOF Then
                           rsPencarian.MoveFirst
                           Do While Not rsPencarian.EOF
                              If Trim(rsPencarian!kode_kategory) = rsHeader.Fields(help) Then
                                 nilaiKolom = nilaiKolom + CDbl(rsPencarian("JUMLAH"))
                              End If
                              rsPencarian.MoveNext
                           Loop
                           rsPencarian.MoveFirst
                        End If
                        .Fields(help) = nilaiKolom
                        nTotSBE(kolom) = nTotSBE(kolom) + CDbl(nilaiKolom)
                        
                     Next
                   progKeterangan.value = 5
                   DoEvents
               Case 6
                   nUrut = nI
                   Set rsPencarian = rsTemp.Clone
                   nKeterangan = "60K saja "
                   rsPencarian.filter = "KETERANGAN = '60'"
                   kolom = 1
                     !Urut = 6
                     !keterangan = nKeterangan
                     !Bold = False
                     !Header = False
                     For kolom = 1 To totalKolom
                        help = "Unit" & kolom
                        nilaiKolom = 0
                        'nTotSBE(kolom) = 0
                        If Not rsPencarian.EOF Then
                           rsPencarian.MoveFirst
                           Do While Not rsPencarian.EOF
                              If Trim(rsPencarian!kode_kategory) = rsHeader.Fields(help) Then
                                 nilaiKolom = nilaiKolom + CDbl(rsPencarian("JUMLAH"))
                              End If
                              rsPencarian.MoveNext
                           Loop
                           rsPencarian.MoveFirst
                        End If
                        .Fields(help) = nilaiKolom
                        nTotSBE(kolom) = nTotSBE(kolom) + CDbl(nilaiKolom)
                        
                     Next
                   progKeterangan.value = 6
                   DoEvents
               Case 7
                   nUrut = nI
                   Set rsPencarian = rsTemp.Clone
                   nKeterangan = "70K saja "
                   rsPencarian.filter = "KETERANGAN = '70'"
                   kolom = 1
                     !Urut = 7
                     !keterangan = nKeterangan
                     !Bold = False
                     !Header = False
                     For kolom = 1 To totalKolom
                        help = "Unit" & kolom
                        nilaiKolom = 0
                        'nTotSBE(kolom) = 0
                        If Not rsPencarian.EOF Then
                           rsPencarian.MoveFirst
                           Do While Not rsPencarian.EOF
                              If Trim(rsPencarian!kode_kategory) = rsHeader.Fields(help) Then
                                 nilaiKolom = nilaiKolom + CDbl(rsPencarian("JUMLAH"))
                              End If
                              rsPencarian.MoveNext
                           Loop
                           rsPencarian.MoveFirst
                        End If
                        .Fields(help) = nilaiKolom
                        nTotSBE(kolom) = nTotSBE(kolom) + CDbl(nilaiKolom)
                        
                     Next
                   progKeterangan.value = 7
                   DoEvents
               Case 8
                   nUrut = nI
                   Set rsPencarian = rsTemp.Clone
                   nKeterangan = "80K saja "
                   rsPencarian.filter = "KETERANGAN = '80'"
                   kolom = 1
                     !Urut = 8
                     !keterangan = nKeterangan
                     !Bold = False
                     !Header = False
                     For kolom = 1 To totalKolom
                        help = "Unit" & kolom
                        nilaiKolom = 0
                        'nTotSBE(kolom) = 0
                        If Not rsPencarian.EOF Then
                           rsPencarian.MoveFirst
                           Do While Not rsPencarian.EOF
                              If Trim(rsPencarian!kode_kategory) = rsHeader.Fields(help) Then
                                 nilaiKolom = nilaiKolom + CDbl(rsPencarian("JUMLAH"))
                              End If
                              rsPencarian.MoveNext
                           Loop
                           rsPencarian.MoveFirst
                        End If
                        .Fields(help) = nilaiKolom
                        nTotSBE(kolom) = nTotSBE(kolom) + CDbl(nilaiKolom)
                        
                     Next
                   progKeterangan.value = 8
                   DoEvents
               Case 8
                   nUrut = nI
                   Set rsPencarian = rsTemp.Clone
                   nKeterangan = "80K saja "
                   rsPencarian.filter = "KETERANGAN = '80'"
                   kolom = 1
                     !Urut = 8
                     !keterangan = nKeterangan
                     !Bold = False
                     !Header = False
                     For kolom = 1 To totalKolom
                        help = "Unit" & kolom
                        nilaiKolom = 0
                        'nTotSBE(kolom) = 0
                        If Not rsPencarian.EOF Then
                           rsPencarian.MoveFirst
                           Do While Not rsPencarian.EOF
                              If Trim(rsPencarian!kode_kategory) = rsHeader.Fields(help) Then
                                 nilaiKolom = nilaiKolom + CDbl(rsPencarian("JUMLAH"))
                              End If
                              rsPencarian.MoveNext
                           Loop
                           rsPencarian.MoveFirst
                        End If
                        .Fields(help) = nilaiKolom
                        nTotSBE(kolom) = nTotSBE(kolom) + CDbl(nilaiKolom)
                        
                     Next
                   progKeterangan.value = 8
                   DoEvents
               Case 9
                   nUrut = nI
                   Set rsPencarian = rsTemp.Clone
                   nKeterangan = "90K saja "
                   rsPencarian.filter = "KETERANGAN = '90'"
                   kolom = 1
                     !Urut = 9
                     !keterangan = nKeterangan
                     !Bold = False
                     !Header = False
                     For kolom = 1 To totalKolom
                        help = "Unit" & kolom
                        nilaiKolom = 0
                        'nTotSBE(kolom) = 0
                        If Not rsPencarian.EOF Then
                           rsPencarian.MoveFirst
                           Do While Not rsPencarian.EOF
                              If Trim(rsPencarian!kode_kategory) = rsHeader.Fields(help) Then
                                 nilaiKolom = nilaiKolom + CDbl(rsPencarian("JUMLAH"))
                              End If
                              rsPencarian.MoveNext
                           Loop
                           rsPencarian.MoveFirst
                        End If
                        .Fields(help) = nilaiKolom
                        nTotSBE(kolom) = nTotSBE(kolom) + CDbl(nilaiKolom)
                        
                     Next
                   progKeterangan.value = 9
                   DoEvents
               Case 10
                   nUrut = nI
                   Set rsPencarian = rsTemp.Clone
                   nKeterangan = "100K saja "
                   rsPencarian.filter = "KETERANGAN = '100'"
                   kolom = 1
                     !Urut = 10
                     !keterangan = nKeterangan
                     !Bold = False
                     !Header = False
                     For kolom = 1 To totalKolom
                        help = "Unit" & kolom
                        nilaiKolom = 0
                        'nTotSBE(kolom) = 0
                        If Not rsPencarian.EOF Then
                           rsPencarian.MoveFirst
                           Do While Not rsPencarian.EOF
                              If Trim(rsPencarian!kode_kategory) = rsHeader.Fields(help) Then
                                 nilaiKolom = nilaiKolom + CDbl(rsPencarian("JUMLAH"))
                              End If
                              rsPencarian.MoveNext
                           Loop
                           rsPencarian.MoveFirst
                        End If
                        .Fields(help) = nilaiKolom
                        nTotSBE(kolom) = nTotSBE(kolom) + CDbl(nilaiKolom)
                        
                     Next
                   progKeterangan.value = 10
                   DoEvents
               Case 10
                   nUrut = nI
                   Set rsPencarian = rsTemp.Clone
                   nKeterangan = "100K saja "
                   rsPencarian.filter = "KETERANGAN = '100'"
                   kolom = 1
                     !Urut = 10
                     !keterangan = nKeterangan
                     !Bold = False
                     !Header = False
                     For kolom = 1 To totalKolom
                        help = "Unit" & kolom
                        nilaiKolom = 0
                        'nTotSBE(kolom) = 0
                        If Not rsPencarian.EOF Then
                           rsPencarian.MoveFirst
                           Do While Not rsPencarian.EOF
                              If Trim(rsPencarian!kode_kategory) = rsHeader.Fields(help) Then
                                 nilaiKolom = nilaiKolom + CDbl(rsPencarian("JUMLAH"))
                              End If
                              rsPencarian.MoveNext
                           Loop
                           rsPencarian.MoveFirst
                        End If
                        .Fields(help) = nilaiKolom
                        nTotSBE(kolom) = nTotSBE(kolom) + CDbl(nilaiKolom)
                        
                     Next
                   progKeterangan.value = 10
                   DoEvents
               Case 11
                   nUrut = nI
                   Set rsPencarian = rsTemp.Clone
                   nKeterangan = "Kelipatan 5K dibawah 100K "
                   rsPencarian.filter = "KETERANGAN = 'Kelipatan 5K'"
                   kolom = 1
                     !Urut = 11
                     !keterangan = nKeterangan
                     !Bold = False
                     !Header = False
                     For kolom = 1 To totalKolom
                        help = "Unit" & kolom
                        nilaiKolom = 0
                        'nTotSBE(kolom) = 0
                        If Not rsPencarian.EOF Then
                           rsPencarian.MoveFirst
                           Do While Not rsPencarian.EOF
                              If Trim(rsPencarian!kode_kategory) = rsHeader.Fields(help) Then
                                 nilaiKolom = nilaiKolom + CDbl(rsPencarian("JUMLAH"))
                              End If
                              rsPencarian.MoveNext
                           Loop
                           rsPencarian.MoveFirst
                        End If
                        .Fields(help) = nilaiKolom
                        nTotSBE(kolom) = nTotSBE(kolom) + CDbl(nilaiKolom)
                        
                     Next
                   progKeterangan.value = 11
                   DoEvents
               Case 12
                   nUrut = nI
                   Set rsPencarian = rsTemp.Clone
                   nKeterangan = "SBE lainnya diatas 100K "
                   rsPencarian.filter = "KETERANGAN = 'Lainnya Diatas 100K'"
                   kolom = 1
                     !Urut = 12
                     !keterangan = nKeterangan
                     !Bold = False
                     !Header = False
                     For kolom = 1 To totalKolom
                        help = "Unit" & kolom
                        nilaiKolom = 0
                        'nTotSBE(kolom) = 0
                        If Not rsPencarian.EOF Then
                           rsPencarian.MoveFirst
                           Do While Not rsPencarian.EOF
                              If Trim(rsPencarian!kode_kategory) = rsHeader.Fields(help) Then
                                 nilaiKolom = nilaiKolom + CDbl(rsPencarian("JUMLAH"))
                              End If
                              rsPencarian.MoveNext
                           Loop
                           rsPencarian.MoveFirst
                        End If
                        .Fields(help) = nilaiKolom
                        nTotSBE(kolom) = nTotSBE(kolom) + CDbl(nilaiKolom)
                        
                     Next
                   progKeterangan.value = 12
                   DoEvents
               Case 13
                   nUrut = nI
                   Set rsPencarian = rsTemp.Clone
                   nKeterangan = "  TOTAL SERVIS BERKALA EKSTERNAL "
                  ' rsPencarian.filter = "KETERANGAN = 'Lainnya Diatas 100K'"
                   kolom = 1
                     !Urut = 13
                     !keterangan = nKeterangan
                     !Bold = True
                     !Header = False
                     For kolom = 1 To totalKolom
                        help = "Unit" & kolom
                        nilaiKolom = 0
                        'nTotSBE(kolom) = 0
'                        If Not rsPencarian.EOF Then
'                           rsPencarian.MoveFirst
'                           Do While Not rsPencarian.EOF
'                              If Trim(rsPencarian!kode_kategory) = rsHeader.Fields(help) Then
'                                 nilaiKolom = nilaiKolom + CDbl(rsPencarian("JUMLAH"))
'                              End If
'                              rsPencarian.MoveNext
'                           Loop
'                           rsPencarian.MoveFirst
'                        End If
                        .Fields(help) = nTotSBE(kolom) 'nilaiKolom
'                        nTotSBE(kolom) = nTotSBE(kolom) + CDbl(nilaiKolom)
                        
                     Next
                   progKeterangan.value = 13
                   DoEvents
               Case 14
                   nUrut = nI
                   Set rsPencarian = rsTemp.Clone
                   nKeterangan = "Unit Ganti Oli Mesin "
                   rsPencarian.filter = "KETERANGAN = 'Unit Ganti Oli Mesin'"
                   kolom = 1
                     !Urut = 14
                     !keterangan = nKeterangan
                     !Bold = False
                     !Header = False
                     For kolom = 1 To totalKolom
                        help = "Unit" & kolom
                        nilaiKolom = 0
                        nTotGR(kolom) = 0
                        If Not rsPencarian.EOF Then
                           rsPencarian.MoveFirst
                           Do While Not rsPencarian.EOF
                              If Trim(rsPencarian!kode_kategory) = rsHeader.Fields(help) Then
                                 nilaiKolom = nilaiKolom + CDbl(rsPencarian("JUMLAH"))
                              End If
                              rsPencarian.MoveNext
                           Loop
                           rsPencarian.MoveFirst
                        End If
                        .Fields(help) = nilaiKolom
                        nTotGR(kolom) = nTotGR(kolom) + CDbl(nilaiKolom)
                        
                     Next
                   progKeterangan.value = 14
                   DoEvents
               Case 15
                   nUrut = nI
                   Set rsPencarian = rsTemp.Clone
                   nKeterangan = "General Repair Lainnya "
                   rsPencarian.filter = "KETERANGAN = 'Lainnya General Repair1'"
                   kolom = 1
                     !Urut = 15
                     !keterangan = nKeterangan
                     !Bold = False
                     !Header = False
                     For kolom = 1 To totalKolom
                        help = "Unit" & kolom
                        nilaiKolom = 0
                       ' nTotGR(kolom) = 0
                        If Not rsPencarian.EOF Then
                           rsPencarian.MoveFirst
                           Do While Not rsPencarian.EOF
                              If Trim(rsPencarian!kode_kategory) = rsHeader.Fields(help) Then
                                 nilaiKolom = nilaiKolom + CDbl(rsPencarian("JUMLAH"))
                              End If
                              rsPencarian.MoveNext
                           Loop
                           rsPencarian.MoveFirst
                        End If
                        .Fields(help) = nilaiKolom
                        nTotGR(kolom) = nTotGR(kolom) + CDbl(nilaiKolom)
                        
                     Next
                   progKeterangan.value = 15
                   DoEvents
                Case 16
                   nUrut = nI
                   Set rsPencarian = rsTemp.Clone
                   nKeterangan = "  SUBTOTAL GENERAL REPAIR "
                  ' rsPencarian.filter = "KETERANGAN = 'Lainnya General Repair1'"
                   kolom = 1
                     !Urut = 16
                     !keterangan = nKeterangan
                     !Bold = True
                     !Header = False
                     For kolom = 1 To totalKolom
                        help = "Unit" & kolom
                        nilaiKolom = 0
                       ' nTotGR(kolom) = 0
'                        If Not rsPencarian.EOF Then
'                           rsPencarian.MoveFirst
'                           Do While Not rsPencarian.EOF
'                              If Trim(rsPencarian!kode_kategory) = rsHeader.Fields(help) Then
'                                 nilaiKolom = nilaiKolom + CDbl(rsPencarian("JUMLAH"))
'                              End If
'                              rsPencarian.MoveNext
'                           Loop
'                           rsPencarian.MoveFirst
'                        End If
                        .Fields(help) = nTotGR(kolom) ' nilaiKolom
'                        nTotGR(kolom) = nTotGR(kolom) + CDbl(nilaiKolom)
                        
                     Next
                   progKeterangan.value = 16
                   DoEvents
               Case 17
                   nUrut = nI
                   Set rsPencarian = rsTemp.Clone
                   nKeterangan = "  TOTAL CPUS "
                  ' rsPencarian.filter = "KETERANGAN = 'Lainnya General Repair1'"
                   kolom = 1
                     !Urut = 17
                     !keterangan = nKeterangan
                     !Bold = True
                     !Header = False
                     For kolom = 1 To totalKolom
                        help = "Unit" & kolom
                        nilaiKolom = 0
                        totCPUS(kolom) = 0
                       ' nTotGR(kolom) = 0
'                        If Not rsPencarian.EOF Then
'                           rsPencarian.MoveFirst
'                           Do While Not rsPencarian.EOF
'                              If Trim(rsPencarian!kode_kategory) = rsHeader.Fields(help) Then
'                                 nilaiKolom = nilaiKolom + CDbl(rsPencarian("JUMLAH"))
'                              End If
'                              rsPencarian.MoveNext
'                           Loop
'                           rsPencarian.MoveFirst
'                        End If
                        .Fields(help) = CDbl(nTotGR(kolom)) + CDbl(nTotSBE(kolom)) ' nilaiKolom
                        totCPUS(kolom) = CDbl(totCPUS(kolom)) + CDbl(nTotGR(kolom)) + CDbl(nTotSBE(kolom))
'                        nTotGR(kolom) = nTotGR(kolom) + CDbl(nilaiKolom)
                        
                     Next
                   progKeterangan.value = 17
                   DoEvents
               Case 18
                   nUrut = nI
                   Set rsPencarian = rsTemp.Clone
                   nKeterangan = "1K Penjualan Outlet Sama "
                   rsPencarian.filter = "KETERANGAN = 'SBI 1 dealer'"
                   kolom = 1
                     !Urut = 18
                     !keterangan = nKeterangan
                     !Bold = False
                     !Header = False
                     For kolom = 1 To totalKolom
                        help = "Unit" & kolom
                        nilaiKolom = 0
                       ' nTotGR(kolom) = 0
                        If Not rsPencarian.EOF Then
                           rsPencarian.MoveFirst
                           Do While Not rsPencarian.EOF
                              If Trim(rsPencarian!kode_kategory) = rsHeader.Fields(help) Then
                                 nilaiKolom = nilaiKolom + CDbl(rsPencarian("JUMLAH"))
                              End If
                              rsPencarian.MoveNext
                           Loop
                           rsPencarian.MoveFirst
                        End If
                        .Fields(help) = nilaiKolom
'                        nTotGR(kolom) = nTotGR(kolom) + CDbl(nilaiKolom)
                        
                     Next
                   progKeterangan.value = 18
                   DoEvents
               
               Case 19
                   nUrut = nI
                   Set rsPencarian = rsTemp.Clone
                   nKeterangan = "1K Total "
                   rsPencarian.filter = "KETERANGAN = 'SBI beda dealer'"
                   kolom = 1
                     !Urut = 19
                     !keterangan = nKeterangan
                     !Bold = False
                     !Header = False
                     For kolom = 1 To totalKolom
                        help = "Unit" & kolom
                        nilaiKolom = 0
                       ' nTotGR(kolom) = 0
                        If Not rsPencarian.EOF Then
                           rsPencarian.MoveFirst
                           Do While Not rsPencarian.EOF
                              If Trim(rsPencarian!kode_kategory) = rsHeader.Fields(help) Then
                                 nilaiKolom = nilaiKolom + CDbl(rsPencarian("JUMLAH"))
                              End If
                              rsPencarian.MoveNext
                           Loop
                           rsPencarian.MoveFirst
                        End If
                        .Fields(help) = nilaiKolom
'                        nTotGR(kolom) = nTotGR(kolom) + CDbl(nilaiKolom)
                        nonCPUS(kolom) = nonCPUS(kolom) + nilaiKolom
                        
                     Next
                   progKeterangan.value = 19
                   DoEvents
               
               Case 20
                   nUrut = nI
                   Set rsPencarian = rsTemp.Clone
                   nKeterangan = "Warranty "
                   rsPencarian.filter = "KETERANGAN = 'nilai WARRANTY'"
                   kolom = 1
                     !Urut = 20
                     !keterangan = nKeterangan
                     !Bold = False
                     !Header = False
                     For kolom = 1 To totalKolom
                        help = "Unit" & kolom
                        nilaiKolom = 0
                       ' nTotGR(kolom) = 0
                        If Not rsPencarian.EOF Then
                           rsPencarian.MoveFirst
                           Do While Not rsPencarian.EOF
                              If Trim(rsPencarian!kode_kategory) = rsHeader.Fields(help) Then
                                 nilaiKolom = nilaiKolom + CDbl(rsPencarian("JUMLAH"))
                              End If
                              rsPencarian.MoveNext
                           Loop
                           rsPencarian.MoveFirst
                        End If
                        .Fields(help) = nilaiKolom
'                        nTotGR(kolom) = nTotGR(kolom) + CDbl(nilaiKolom)
                        nonCPUS(kolom) = nonCPUS(kolom) + nilaiKolom
                        
                     Next
                   progKeterangan.value = 20
                   DoEvents
               Case 21
                   nUrut = nI
                   Set rsPencarian = rsTemp.Clone
                   nKeterangan = "Return Job "
                   rsPencarian.filter = "KETERANGAN = 'Return Job'"
                   kolom = 1
                     !Urut = 21
                     !keterangan = nKeterangan
                     !Bold = False
                     !Header = False
                     For kolom = 1 To totalKolom
                        help = "Unit" & kolom
                        nilaiKolom = 0
                       ' nTotGR(kolom) = 0
                        If Not rsPencarian.EOF Then
                           rsPencarian.MoveFirst
                           Do While Not rsPencarian.EOF
                              If Trim(rsPencarian!kode_kategory) = rsHeader.Fields(help) Then
                                 nilaiKolom = nilaiKolom + CDbl(rsPencarian("JUMLAH"))
                              End If
                              rsPencarian.MoveNext
                           Loop
                           rsPencarian.MoveFirst
                        End If
                        .Fields(help) = nilaiKolom
'                        nTotGR(kolom) = nTotGR(kolom) + CDbl(nilaiKolom)
                        nonCPUS(kolom) = nonCPUS(kolom) + nilaiKolom
                        
                     Next
                   progKeterangan.value = 21
                   DoEvents
               Case 22
                   nUrut = nI
                   Set rsPencarian = rsTemp.Clone
                   nKeterangan = "Pre Delivery Servis "
                   rsPencarian.filter = "KETERANGAN = 'PDS'"
                   kolom = 1
                     !Urut = 22
                     !keterangan = nKeterangan
                     !Bold = False
                     !Header = False
                     For kolom = 1 To totalKolom
                        help = "Unit" & kolom
                        nilaiKolom = 0
                       ' nTotGR(kolom) = 0
                        If Not rsPencarian.EOF Then
                           rsPencarian.MoveFirst
                           Do While Not rsPencarian.EOF
                              If Trim(rsPencarian!kode_kategory) = rsHeader.Fields(help) Then
                                 nilaiKolom = nilaiKolom + CDbl(rsPencarian("JUMLAH"))
                              End If
                              rsPencarian.MoveNext
                           Loop
                           rsPencarian.MoveFirst
                        End If
                        .Fields(help) = nilaiKolom
'                        nTotGR(kolom) = nTotGR(kolom) + CDbl(nilaiKolom)
                        nonCPUS(kolom) = nonCPUS(kolom) + nilaiKolom
                        
                     Next
                   progKeterangan.value = 22
                   DoEvents
               Case 23
                   nUrut = nI
                   Set rsPencarian = rsTemp.Clone
                   nKeterangan = "  SUBTOTAL NON CPUS "
                   'rsPencarian.filter = "KETERANGAN = 'PDS'"
                   kolom = 1
                     !Urut = 23
                     !keterangan = nKeterangan
                     !Bold = True
                     !Header = False
                     For kolom = 1 To totalKolom
                        help = "Unit" & kolom
                        nilaiKolom = 0
                       ' nTotGR(kolom) = 0
'                        If Not rsPencarian.EOF Then
'                           rsPencarian.MoveFirst
'                           Do While Not rsPencarian.EOF
'                              If Trim(rsPencarian!kode_kategory) = rsHeader.Fields(help) Then
'                                 nilaiKolom = nilaiKolom + CDbl(rsPencarian("JUMLAH"))
'                              End If
'                              rsPencarian.MoveNext
'                           Loop
'                           rsPencarian.MoveFirst
'                        End If
                        .Fields(help) = nonCPUS(kolom)
'                        nTotGR(kolom) = nTotGR(kolom) + CDbl(nilaiKolom)
'                        nonCPUS(kolom) = nonCPUS(kolom) + nilaiKolom
                        
                     Next
                   progKeterangan.value = 23
                   DoEvents
               Case 24
                   nUrut = nI
                   Set rsPencarian = rsTemp.Clone
                   nKeterangan = "  JUMLAH TOTAL UNIT "
                   'rsPencarian.filter = "KETERANGAN = 'PDS'"
                   kolom = 1
                     !Urut = 24
                     !keterangan = nKeterangan
                     !Bold = True
                     !Header = False
                     For kolom = 1 To totalKolom
                        help = "Unit" & kolom
                        nilaiKolom = 0
                       ' nTotGR(kolom) = 0
'                        If Not rsPencarian.EOF Then
'                           rsPencarian.MoveFirst
'                           Do While Not rsPencarian.EOF
'                              If Trim(rsPencarian!kode_kategory) = rsHeader.Fields(help) Then
'                                 nilaiKolom = nilaiKolom + CDbl(rsPencarian("JUMLAH"))
'                              End If
'                              rsPencarian.MoveNext
'                           Loop
'                           rsPencarian.MoveFirst
'                        End If
                        .Fields(help) = CDbl(nonCPUS(kolom)) + CDbl(totCPUS(kolom))
                        totUE(kolom) = CDbl(totUE(kolom)) + CDbl(nonCPUS(kolom)) + CDbl(totCPUS(kolom))
'                        nTotGR(kolom) = nTotGR(kolom) + CDbl(nilaiKolom)
'                        nonCPUS(kolom) = nonCPUS(kolom) + nilaiKolom
                        
                     Next
                   progKeterangan.value = 24
                   DoEvents
               Case 25
                   nUrut = nI
                   Set rsPencarian = rsTemp.Clone
                   nKeterangan = "  REVENUE BY MODEL [Rupiah aktual yang diterima bengkel setelah harga diskon] """
                   'rsPencarian.filter = "KETERANGAN = 'PDS'"
                   kolom = 1
                     !Urut = 25
                     !keterangan = nKeterangan
                     !Bold = True
                     !Header = False
                     For kolom = 1 To totalKolom
                        help = "Unit" & kolom
                        nilaiKolom = 0
                       ' nTotGR(kolom) = 0
'                        If Not rsPencarian.EOF Then
'                           rsPencarian.MoveFirst
'                           Do While Not rsPencarian.EOF
'                              If Trim(rsPencarian!kode_kategory) = rsHeader.Fields(help) Then
'                                 nilaiKolom = nilaiKolom + CDbl(rsPencarian("JUMLAH"))
'                              End If
'                              rsPencarian.MoveNext
'                           Loop
'                           rsPencarian.MoveFirst
'                        End If
                        .Fields(help) = 0 'CDbl(nonCPUS(kolom)) + CDbl(totCPUS(kolom))
                        'totUE(kolom) = CDbl(totUE(kolom)) + CDbl(nonCPUS(kolom)) + CDbl(totCPUS(kolom))
'                        nTotGR(kolom) = nTotGR(kolom) + CDbl(nilaiKolom)
'                        nonCPUS(kolom) = nonCPUS(kolom) + nilaiKolom
                        
                     Next
                   progKeterangan.value = 25
                   DoEvents
               Case 26
                   nUrut = nI
                   Set rsPencarian = rsTemp.Clone
                   nKeterangan = "Labor "
                   rsPencarian.filter = "KETERANGAN = 'Revenue Jasa Total'"
                   kolom = 1
                     !Urut = 26
                     !keterangan = nKeterangan
                     !Bold = False
                     !Header = False
                     For kolom = 1 To totalKolom
                        help = "Unit" & kolom
                        nilaiKolom = 0
                       ' nTotGR(kolom) = 0
                        If Not rsPencarian.EOF Then
                           rsPencarian.MoveFirst
                           Do While Not rsPencarian.EOF
                              If Trim(rsPencarian!kode_kategory) = rsHeader.Fields(help) Then
                                 nilaiKolom = nilaiKolom + CDbl(rsPencarian("JUMLAH"))
                              End If
                              rsPencarian.MoveNext
                           Loop
                           rsPencarian.MoveFirst
                        End If
                        .Fields(help) = nilaiKolom
                        labor(kolom) = labor(kolom) + nilaiKolom
                        
                     Next
                   progKeterangan.value = 26
                   DoEvents
               Case 27
                   nUrut = nI
                   Set rsPencarian = rsTemp.Clone
                   nKeterangan = "Part Toyota & Oli TGMO (CPUS + Non CPUS) "
                   rsPencarian.filter = "KETERANGAN = 'Revenue Part dan Oli TGMO Total'"
                   kolom = 1
                     !Urut = 27
                     !keterangan = nKeterangan
                     !Bold = False
                     !Header = False
                     For kolom = 1 To totalKolom
                        help = "Unit" & kolom
                        nilaiKolom = 0
                       ' nTotGR(kolom) = 0
                        If Not rsPencarian.EOF Then
                           rsPencarian.MoveFirst
                           Do While Not rsPencarian.EOF
                              If Trim(rsPencarian!kode_kategory) = rsHeader.Fields(help) Then
                                 nilaiKolom = nilaiKolom + CDbl(rsPencarian("JUMLAH"))
                              End If
                              rsPencarian.MoveNext
                           Loop
                           rsPencarian.MoveFirst
                        End If
                        .Fields(help) = nilaiKolom
                        labor(kolom) = labor(kolom) + nilaiKolom
                        
                     Next
                   progKeterangan.value = 27
                   DoEvents
               Case 28
                   nUrut = nI
                   Set rsPencarian = rsTemp.Clone
                   nKeterangan = "Oli / Bahan Non Toyota (CPUS + Non CPUS) "
                   rsPencarian.filter = "KETERANGAN = 'Revenue Oli dan Bahan Non Toyota Total'"
                   kolom = 1
                     !Urut = 28
                     !keterangan = nKeterangan
                     !Bold = False
                     !Header = False
                     For kolom = 1 To totalKolom
                        help = "Unit" & kolom
                        nilaiKolom = 0
                       ' nTotGR(kolom) = 0
                        If Not rsPencarian.EOF Then
                           rsPencarian.MoveFirst
                           Do While Not rsPencarian.EOF
                              If Trim(rsPencarian!kode_kategory) = rsHeader.Fields(help) Then
                                 nilaiKolom = nilaiKolom + CDbl(rsPencarian("JUMLAH"))
                              End If
                              rsPencarian.MoveNext
                           Loop
                           rsPencarian.MoveFirst
                        End If
                        .Fields(help) = nilaiKolom
                        labor(kolom) = labor(kolom) + nilaiKolom
                        
                     Next
                   progKeterangan.value = 28
                   DoEvents
               Case 29
                   nUrut = nI
                   Set rsPencarian = rsTemp.Clone
                   nKeterangan = "Pekerjaan Luar (CPUS + Non CPUS) "
                   rsPencarian.filter = "KETERANGAN = 'Revenue Pekerjaan Luar Total'"
                   kolom = 1
                     !Urut = 29
                     !keterangan = nKeterangan
                     !Bold = False
                     !Header = False
                     For kolom = 1 To totalKolom
                        help = "Unit" & kolom
                        nilaiKolom = 0
                       ' nTotGR(kolom) = 0
                        If Not rsPencarian.EOF Then
                           rsPencarian.MoveFirst
                           Do While Not rsPencarian.EOF
                              If Trim(rsPencarian!kode_kategory) = rsHeader.Fields(help) Then
                                 nilaiKolom = nilaiKolom + CDbl(rsPencarian("JUMLAH"))
                              End If
                              rsPencarian.MoveNext
                           Loop
                           rsPencarian.MoveFirst
                        End If
                        .Fields(help) = nilaiKolom
                        labor(kolom) = labor(kolom) + nilaiKolom
                        
                     Next
                   progKeterangan.value = 29
                   DoEvents
               Case 30
                   nUrut = nI
                   Set rsPencarian = rsTemp.Clone
                   nKeterangan = "  JUMLAH TOTAL UNIT "
                   'rsPencarian.filter = "KETERANGAN = 'PDS'"
                   kolom = 1
                     !Urut = 30
                     !keterangan = nKeterangan
                     !Bold = True
                     !Header = False
                     For kolom = 1 To totalKolom
                        help = "Unit" & kolom
                        nilaiKolom = 0
                       ' nTotGR(kolom) = 0
'                        If Not rsPencarian.EOF Then
'                           rsPencarian.MoveFirst
'                           Do While Not rsPencarian.EOF
'                              If Trim(rsPencarian!kode_kategory) = rsHeader.Fields(help) Then
'                                 nilaiKolom = nilaiKolom + CDbl(rsPencarian("JUMLAH"))
'                              End If
'                              rsPencarian.MoveNext
'                           Loop
'                           rsPencarian.MoveFirst
'                        End If
                        .Fields(help) = CDbl(labor(kolom))
                        ''totUE(kolom) = CDbl(totUE(kolom)) + CDbl(nonCPUS(kolom)) + CDbl(totCPUS(kolom))
'                        nTotGR(kolom) = nTotGR(kolom) + CDbl(nilaiKolom)
'                        nonCPUS(kolom) = nonCPUS(kolom) + nilaiKolom
                        
                     Next
                   progKeterangan.value = 30
                   DoEvents
            End Select
            .Update
         End With
         i = i + 1
      Loop
   End If
  lblKeterangan.Visible = False
  progKeterangan.Visible = False
  
End Sub
