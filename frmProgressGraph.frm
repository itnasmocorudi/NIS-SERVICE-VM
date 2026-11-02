VERSION 5.00
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Begin VB.Form frmProgressGraph 
   Caption         =   "Display Activity Mekanik"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7995
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton TblKeluar 
      Caption         =   "&Kembali ke Menu"
      Height          =   390
      Left            =   9915
      TabIndex        =   0
      Top             =   7125
      Width           =   1890
   End
   Begin TrueDBGrid70.TDBGrid TData1 
      Height          =   7065
      Left            =   360
      TabIndex        =   1
      Top             =   0
      Width           =   11445
      _ExtentX        =   20188
      _ExtentY        =   12462
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
      Columns.Count   =   69
      Splits(0)._UserFlags=   0
      Splits(0).SizeMode=   2
      Splits(0).Size  =   2
      Splits(0).Size.vt=   2
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0).ScrollBars=   0
      Splits(0).AllowColSelect=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=69"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3995"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3916"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=423"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=344"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=1773"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1693"
      Splits(0)._ColumnProps(12)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(14)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(17)=   "Column(3).Visible=0"
      Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(19)=   "Column(4).Width=1879"
      Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=1799"
      Splits(0)._ColumnProps(22)=   "Column(4).Visible=0"
      Splits(0)._ColumnProps(23)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(24)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(25)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(27)=   "Column(5).Visible=0"
      Splits(0)._ColumnProps(28)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(29)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(30)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(32)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(33)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(34)=   "Column(7).Width=2725"
      Splits(0)._ColumnProps(35)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(36)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(37)=   "Column(7).Visible=0"
      Splits(0)._ColumnProps(38)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(39)=   "Column(8).Width=2725"
      Splits(0)._ColumnProps(40)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(41)=   "Column(8)._WidthInPix=2646"
      Splits(0)._ColumnProps(42)=   "Column(8).Visible=0"
      Splits(0)._ColumnProps(43)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(44)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(45)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(47)=   "Column(9).Visible=0"
      Splits(0)._ColumnProps(48)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(49)=   "Column(10).Width=2725"
      Splits(0)._ColumnProps(50)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(10)._WidthInPix=2646"
      Splits(0)._ColumnProps(52)=   "Column(10).Visible=0"
      Splits(0)._ColumnProps(53)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(54)=   "Column(11).Width=847"
      Splits(0)._ColumnProps(55)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(56)=   "Column(11)._WidthInPix=767"
      Splits(0)._ColumnProps(57)=   "Column(11).Visible=0"
      Splits(0)._ColumnProps(58)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(59)=   "Column(12).Width=2725"
      Splits(0)._ColumnProps(60)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(61)=   "Column(12)._WidthInPix=2646"
      Splits(0)._ColumnProps(62)=   "Column(12).Visible=0"
      Splits(0)._ColumnProps(63)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(64)=   "Column(13).Width=2725"
      Splits(0)._ColumnProps(65)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(66)=   "Column(13)._WidthInPix=2646"
      Splits(0)._ColumnProps(67)=   "Column(13).Visible=0"
      Splits(0)._ColumnProps(68)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(69)=   "Column(14).Width=2725"
      Splits(0)._ColumnProps(70)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(71)=   "Column(14)._WidthInPix=2646"
      Splits(0)._ColumnProps(72)=   "Column(14).Visible=0"
      Splits(0)._ColumnProps(73)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(74)=   "Column(15).Width=2725"
      Splits(0)._ColumnProps(75)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(76)=   "Column(15)._WidthInPix=2646"
      Splits(0)._ColumnProps(77)=   "Column(15).Visible=0"
      Splits(0)._ColumnProps(78)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(79)=   "Column(16).Width=2725"
      Splits(0)._ColumnProps(80)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(81)=   "Column(16)._WidthInPix=2646"
      Splits(0)._ColumnProps(82)=   "Column(16).Visible=0"
      Splits(0)._ColumnProps(83)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(84)=   "Column(17).Width=2725"
      Splits(0)._ColumnProps(85)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(86)=   "Column(17)._WidthInPix=2646"
      Splits(0)._ColumnProps(87)=   "Column(17).Visible=0"
      Splits(0)._ColumnProps(88)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(89)=   "Column(18).Width=2725"
      Splits(0)._ColumnProps(90)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(91)=   "Column(18)._WidthInPix=2646"
      Splits(0)._ColumnProps(92)=   "Column(18).Visible=0"
      Splits(0)._ColumnProps(93)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(94)=   "Column(19).Width=2725"
      Splits(0)._ColumnProps(95)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(96)=   "Column(19)._WidthInPix=2646"
      Splits(0)._ColumnProps(97)=   "Column(19).Visible=0"
      Splits(0)._ColumnProps(98)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(99)=   "Column(20).Width=2725"
      Splits(0)._ColumnProps(100)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(101)=   "Column(20)._WidthInPix=2646"
      Splits(0)._ColumnProps(102)=   "Column(20).Visible=0"
      Splits(0)._ColumnProps(103)=   "Column(20).Order=21"
      Splits(0)._ColumnProps(104)=   "Column(21).Width=2725"
      Splits(0)._ColumnProps(105)=   "Column(21).DividerColor=0"
      Splits(0)._ColumnProps(106)=   "Column(21)._WidthInPix=2646"
      Splits(0)._ColumnProps(107)=   "Column(21).Visible=0"
      Splits(0)._ColumnProps(108)=   "Column(21).Order=22"
      Splits(0)._ColumnProps(109)=   "Column(22).Width=2725"
      Splits(0)._ColumnProps(110)=   "Column(22).DividerColor=0"
      Splits(0)._ColumnProps(111)=   "Column(22)._WidthInPix=2646"
      Splits(0)._ColumnProps(112)=   "Column(22).Visible=0"
      Splits(0)._ColumnProps(113)=   "Column(22).Order=23"
      Splits(0)._ColumnProps(114)=   "Column(23).Width=2725"
      Splits(0)._ColumnProps(115)=   "Column(23).DividerColor=0"
      Splits(0)._ColumnProps(116)=   "Column(23)._WidthInPix=2646"
      Splits(0)._ColumnProps(117)=   "Column(23).Visible=0"
      Splits(0)._ColumnProps(118)=   "Column(23).Order=24"
      Splits(0)._ColumnProps(119)=   "Column(24).Width=2725"
      Splits(0)._ColumnProps(120)=   "Column(24).DividerColor=0"
      Splits(0)._ColumnProps(121)=   "Column(24)._WidthInPix=2646"
      Splits(0)._ColumnProps(122)=   "Column(24).Visible=0"
      Splits(0)._ColumnProps(123)=   "Column(24).Order=25"
      Splits(0)._ColumnProps(124)=   "Column(25).Width=2725"
      Splits(0)._ColumnProps(125)=   "Column(25).DividerColor=0"
      Splits(0)._ColumnProps(126)=   "Column(25)._WidthInPix=2646"
      Splits(0)._ColumnProps(127)=   "Column(25).Visible=0"
      Splits(0)._ColumnProps(128)=   "Column(25).Order=26"
      Splits(0)._ColumnProps(129)=   "Column(26).Width=2725"
      Splits(0)._ColumnProps(130)=   "Column(26).DividerColor=0"
      Splits(0)._ColumnProps(131)=   "Column(26)._WidthInPix=2646"
      Splits(0)._ColumnProps(132)=   "Column(26).Visible=0"
      Splits(0)._ColumnProps(133)=   "Column(26).Order=27"
      Splits(0)._ColumnProps(134)=   "Column(27).Width=2725"
      Splits(0)._ColumnProps(135)=   "Column(27).DividerColor=0"
      Splits(0)._ColumnProps(136)=   "Column(27)._WidthInPix=2646"
      Splits(0)._ColumnProps(137)=   "Column(27).Visible=0"
      Splits(0)._ColumnProps(138)=   "Column(27).Order=28"
      Splits(0)._ColumnProps(139)=   "Column(28).Width=2725"
      Splits(0)._ColumnProps(140)=   "Column(28).DividerColor=0"
      Splits(0)._ColumnProps(141)=   "Column(28)._WidthInPix=2646"
      Splits(0)._ColumnProps(142)=   "Column(28).Visible=0"
      Splits(0)._ColumnProps(143)=   "Column(28).Order=29"
      Splits(0)._ColumnProps(144)=   "Column(29).Width=2725"
      Splits(0)._ColumnProps(145)=   "Column(29).DividerColor=0"
      Splits(0)._ColumnProps(146)=   "Column(29)._WidthInPix=2646"
      Splits(0)._ColumnProps(147)=   "Column(29).Visible=0"
      Splits(0)._ColumnProps(148)=   "Column(29).Order=30"
      Splits(0)._ColumnProps(149)=   "Column(30).Width=2725"
      Splits(0)._ColumnProps(150)=   "Column(30).DividerColor=0"
      Splits(0)._ColumnProps(151)=   "Column(30)._WidthInPix=2646"
      Splits(0)._ColumnProps(152)=   "Column(30).Visible=0"
      Splits(0)._ColumnProps(153)=   "Column(30).Order=31"
      Splits(0)._ColumnProps(154)=   "Column(31).Width=2725"
      Splits(0)._ColumnProps(155)=   "Column(31).DividerColor=0"
      Splits(0)._ColumnProps(156)=   "Column(31)._WidthInPix=2646"
      Splits(0)._ColumnProps(157)=   "Column(31).Visible=0"
      Splits(0)._ColumnProps(158)=   "Column(31).Order=32"
      Splits(0)._ColumnProps(159)=   "Column(32).Width=2725"
      Splits(0)._ColumnProps(160)=   "Column(32).DividerColor=0"
      Splits(0)._ColumnProps(161)=   "Column(32)._WidthInPix=2646"
      Splits(0)._ColumnProps(162)=   "Column(32).Visible=0"
      Splits(0)._ColumnProps(163)=   "Column(32).Order=33"
      Splits(0)._ColumnProps(164)=   "Column(33).Width=2725"
      Splits(0)._ColumnProps(165)=   "Column(33).DividerColor=0"
      Splits(0)._ColumnProps(166)=   "Column(33)._WidthInPix=2646"
      Splits(0)._ColumnProps(167)=   "Column(33).Visible=0"
      Splits(0)._ColumnProps(168)=   "Column(33).Order=34"
      Splits(0)._ColumnProps(169)=   "Column(34).Width=2725"
      Splits(0)._ColumnProps(170)=   "Column(34).DividerColor=0"
      Splits(0)._ColumnProps(171)=   "Column(34)._WidthInPix=2646"
      Splits(0)._ColumnProps(172)=   "Column(34).Visible=0"
      Splits(0)._ColumnProps(173)=   "Column(34).Order=35"
      Splits(0)._ColumnProps(174)=   "Column(35).Width=2725"
      Splits(0)._ColumnProps(175)=   "Column(35).DividerColor=0"
      Splits(0)._ColumnProps(176)=   "Column(35)._WidthInPix=2646"
      Splits(0)._ColumnProps(177)=   "Column(35).Visible=0"
      Splits(0)._ColumnProps(178)=   "Column(35).Order=36"
      Splits(0)._ColumnProps(179)=   "Column(36).Width=159"
      Splits(0)._ColumnProps(180)=   "Column(36).DividerColor=0"
      Splits(0)._ColumnProps(181)=   "Column(36)._WidthInPix=79"
      Splits(0)._ColumnProps(182)=   "Column(36).Visible=0"
      Splits(0)._ColumnProps(183)=   "Column(36).Order=37"
      Splits(0)._ColumnProps(184)=   "Column(37).Width=2725"
      Splits(0)._ColumnProps(185)=   "Column(37).DividerColor=0"
      Splits(0)._ColumnProps(186)=   "Column(37)._WidthInPix=2646"
      Splits(0)._ColumnProps(187)=   "Column(37).Visible=0"
      Splits(0)._ColumnProps(188)=   "Column(37).Order=38"
      Splits(0)._ColumnProps(189)=   "Column(38).Width=2725"
      Splits(0)._ColumnProps(190)=   "Column(38).DividerColor=0"
      Splits(0)._ColumnProps(191)=   "Column(38)._WidthInPix=2646"
      Splits(0)._ColumnProps(192)=   "Column(38).Visible=0"
      Splits(0)._ColumnProps(193)=   "Column(38).Order=39"
      Splits(0)._ColumnProps(194)=   "Column(39).Width=2725"
      Splits(0)._ColumnProps(195)=   "Column(39).DividerColor=0"
      Splits(0)._ColumnProps(196)=   "Column(39)._WidthInPix=2646"
      Splits(0)._ColumnProps(197)=   "Column(39).Visible=0"
      Splits(0)._ColumnProps(198)=   "Column(39).Order=40"
      Splits(0)._ColumnProps(199)=   "Column(40).Width=2725"
      Splits(0)._ColumnProps(200)=   "Column(40).DividerColor=0"
      Splits(0)._ColumnProps(201)=   "Column(40)._WidthInPix=2646"
      Splits(0)._ColumnProps(202)=   "Column(40).Visible=0"
      Splits(0)._ColumnProps(203)=   "Column(40).Order=41"
      Splits(0)._ColumnProps(204)=   "Column(41).Width=2725"
      Splits(0)._ColumnProps(205)=   "Column(41).DividerColor=0"
      Splits(0)._ColumnProps(206)=   "Column(41)._WidthInPix=2646"
      Splits(0)._ColumnProps(207)=   "Column(41).Visible=0"
      Splits(0)._ColumnProps(208)=   "Column(41).Order=42"
      Splits(0)._ColumnProps(209)=   "Column(42).Width=2725"
      Splits(0)._ColumnProps(210)=   "Column(42).DividerColor=0"
      Splits(0)._ColumnProps(211)=   "Column(42)._WidthInPix=2646"
      Splits(0)._ColumnProps(212)=   "Column(42).Visible=0"
      Splits(0)._ColumnProps(213)=   "Column(42).Order=43"
      Splits(0)._ColumnProps(214)=   "Column(43).Width=2725"
      Splits(0)._ColumnProps(215)=   "Column(43).DividerColor=0"
      Splits(0)._ColumnProps(216)=   "Column(43)._WidthInPix=2646"
      Splits(0)._ColumnProps(217)=   "Column(43).Visible=0"
      Splits(0)._ColumnProps(218)=   "Column(43).Order=44"
      Splits(0)._ColumnProps(219)=   "Column(44).Width=2725"
      Splits(0)._ColumnProps(220)=   "Column(44).DividerColor=0"
      Splits(0)._ColumnProps(221)=   "Column(44)._WidthInPix=2646"
      Splits(0)._ColumnProps(222)=   "Column(44).Visible=0"
      Splits(0)._ColumnProps(223)=   "Column(44).Order=45"
      Splits(0)._ColumnProps(224)=   "Column(45).Width=2725"
      Splits(0)._ColumnProps(225)=   "Column(45).DividerColor=0"
      Splits(0)._ColumnProps(226)=   "Column(45)._WidthInPix=2646"
      Splits(0)._ColumnProps(227)=   "Column(45).Visible=0"
      Splits(0)._ColumnProps(228)=   "Column(45).Order=46"
      Splits(0)._ColumnProps(229)=   "Column(46).Width=2725"
      Splits(0)._ColumnProps(230)=   "Column(46).DividerColor=0"
      Splits(0)._ColumnProps(231)=   "Column(46)._WidthInPix=2646"
      Splits(0)._ColumnProps(232)=   "Column(46).Visible=0"
      Splits(0)._ColumnProps(233)=   "Column(46).Order=47"
      Splits(0)._ColumnProps(234)=   "Column(47).Width=2725"
      Splits(0)._ColumnProps(235)=   "Column(47).DividerColor=0"
      Splits(0)._ColumnProps(236)=   "Column(47)._WidthInPix=2646"
      Splits(0)._ColumnProps(237)=   "Column(47).Visible=0"
      Splits(0)._ColumnProps(238)=   "Column(47).Order=48"
      Splits(0)._ColumnProps(239)=   "Column(48).Width=2725"
      Splits(0)._ColumnProps(240)=   "Column(48).DividerColor=0"
      Splits(0)._ColumnProps(241)=   "Column(48)._WidthInPix=2646"
      Splits(0)._ColumnProps(242)=   "Column(48).Visible=0"
      Splits(0)._ColumnProps(243)=   "Column(48).Order=49"
      Splits(0)._ColumnProps(244)=   "Column(49).Width=2725"
      Splits(0)._ColumnProps(245)=   "Column(49).DividerColor=0"
      Splits(0)._ColumnProps(246)=   "Column(49)._WidthInPix=2646"
      Splits(0)._ColumnProps(247)=   "Column(49).Visible=0"
      Splits(0)._ColumnProps(248)=   "Column(49).Order=50"
      Splits(0)._ColumnProps(249)=   "Column(50).Width=2725"
      Splits(0)._ColumnProps(250)=   "Column(50).DividerColor=0"
      Splits(0)._ColumnProps(251)=   "Column(50)._WidthInPix=2646"
      Splits(0)._ColumnProps(252)=   "Column(50).Visible=0"
      Splits(0)._ColumnProps(253)=   "Column(50).Order=51"
      Splits(0)._ColumnProps(254)=   "Column(51).Width=2725"
      Splits(0)._ColumnProps(255)=   "Column(51).DividerColor=0"
      Splits(0)._ColumnProps(256)=   "Column(51)._WidthInPix=2646"
      Splits(0)._ColumnProps(257)=   "Column(51).Visible=0"
      Splits(0)._ColumnProps(258)=   "Column(51).Order=52"
      Splits(0)._ColumnProps(259)=   "Column(52).Width=2725"
      Splits(0)._ColumnProps(260)=   "Column(52).DividerColor=0"
      Splits(0)._ColumnProps(261)=   "Column(52)._WidthInPix=2646"
      Splits(0)._ColumnProps(262)=   "Column(52).Visible=0"
      Splits(0)._ColumnProps(263)=   "Column(52).Order=53"
      Splits(0)._ColumnProps(264)=   "Column(53).Width=2725"
      Splits(0)._ColumnProps(265)=   "Column(53).DividerColor=0"
      Splits(0)._ColumnProps(266)=   "Column(53)._WidthInPix=2646"
      Splits(0)._ColumnProps(267)=   "Column(53).Visible=0"
      Splits(0)._ColumnProps(268)=   "Column(53).Order=54"
      Splits(0)._ColumnProps(269)=   "Column(54).Width=2725"
      Splits(0)._ColumnProps(270)=   "Column(54).DividerColor=0"
      Splits(0)._ColumnProps(271)=   "Column(54)._WidthInPix=2646"
      Splits(0)._ColumnProps(272)=   "Column(54).Visible=0"
      Splits(0)._ColumnProps(273)=   "Column(54).Order=55"
      Splits(0)._ColumnProps(274)=   "Column(55).Width=2725"
      Splits(0)._ColumnProps(275)=   "Column(55).DividerColor=0"
      Splits(0)._ColumnProps(276)=   "Column(55)._WidthInPix=2646"
      Splits(0)._ColumnProps(277)=   "Column(55).Visible=0"
      Splits(0)._ColumnProps(278)=   "Column(55).Order=56"
      Splits(0)._ColumnProps(279)=   "Column(56).Width=185"
      Splits(0)._ColumnProps(280)=   "Column(56).DividerColor=0"
      Splits(0)._ColumnProps(281)=   "Column(56)._WidthInPix=106"
      Splits(0)._ColumnProps(282)=   "Column(56).Visible=0"
      Splits(0)._ColumnProps(283)=   "Column(56).Order=57"
      Splits(0)._ColumnProps(284)=   "Column(57).Width=397"
      Splits(0)._ColumnProps(285)=   "Column(57).DividerColor=0"
      Splits(0)._ColumnProps(286)=   "Column(57)._WidthInPix=318"
      Splits(0)._ColumnProps(287)=   "Column(57).Visible=0"
      Splits(0)._ColumnProps(288)=   "Column(57).Order=58"
      Splits(0)._ColumnProps(289)=   "Column(58).Width=2725"
      Splits(0)._ColumnProps(290)=   "Column(58).DividerColor=0"
      Splits(0)._ColumnProps(291)=   "Column(58)._WidthInPix=2646"
      Splits(0)._ColumnProps(292)=   "Column(58).Visible=0"
      Splits(0)._ColumnProps(293)=   "Column(58).Order=59"
      Splits(0)._ColumnProps(294)=   "Column(59).Width=2725"
      Splits(0)._ColumnProps(295)=   "Column(59).DividerColor=0"
      Splits(0)._ColumnProps(296)=   "Column(59)._WidthInPix=2646"
      Splits(0)._ColumnProps(297)=   "Column(59).Visible=0"
      Splits(0)._ColumnProps(298)=   "Column(59).Order=60"
      Splits(0)._ColumnProps(299)=   "Column(60).Width=2725"
      Splits(0)._ColumnProps(300)=   "Column(60).DividerColor=0"
      Splits(0)._ColumnProps(301)=   "Column(60)._WidthInPix=2646"
      Splits(0)._ColumnProps(302)=   "Column(60).Visible=0"
      Splits(0)._ColumnProps(303)=   "Column(60).Order=61"
      Splits(0)._ColumnProps(304)=   "Column(61).Width=159"
      Splits(0)._ColumnProps(305)=   "Column(61).DividerColor=0"
      Splits(0)._ColumnProps(306)=   "Column(61)._WidthInPix=79"
      Splits(0)._ColumnProps(307)=   "Column(61).Visible=0"
      Splits(0)._ColumnProps(308)=   "Column(61).Order=62"
      Splits(0)._ColumnProps(309)=   "Column(62).Width=2725"
      Splits(0)._ColumnProps(310)=   "Column(62).DividerColor=0"
      Splits(0)._ColumnProps(311)=   "Column(62)._WidthInPix=2646"
      Splits(0)._ColumnProps(312)=   "Column(62).Visible=0"
      Splits(0)._ColumnProps(313)=   "Column(62).Order=63"
      Splits(0)._ColumnProps(314)=   "Column(63).Width=2725"
      Splits(0)._ColumnProps(315)=   "Column(63).DividerColor=0"
      Splits(0)._ColumnProps(316)=   "Column(63)._WidthInPix=2646"
      Splits(0)._ColumnProps(317)=   "Column(63).Visible=0"
      Splits(0)._ColumnProps(318)=   "Column(63).Order=64"
      Splits(0)._ColumnProps(319)=   "Column(64).Width=2725"
      Splits(0)._ColumnProps(320)=   "Column(64).DividerColor=0"
      Splits(0)._ColumnProps(321)=   "Column(64)._WidthInPix=2646"
      Splits(0)._ColumnProps(322)=   "Column(64).Visible=0"
      Splits(0)._ColumnProps(323)=   "Column(64).Order=65"
      Splits(0)._ColumnProps(324)=   "Column(65).Width=2725"
      Splits(0)._ColumnProps(325)=   "Column(65).DividerColor=0"
      Splits(0)._ColumnProps(326)=   "Column(65)._WidthInPix=2646"
      Splits(0)._ColumnProps(327)=   "Column(65).Visible=0"
      Splits(0)._ColumnProps(328)=   "Column(65).Order=66"
      Splits(0)._ColumnProps(329)=   "Column(66).Width=2725"
      Splits(0)._ColumnProps(330)=   "Column(66).DividerColor=0"
      Splits(0)._ColumnProps(331)=   "Column(66)._WidthInPix=2646"
      Splits(0)._ColumnProps(332)=   "Column(66).Visible=0"
      Splits(0)._ColumnProps(333)=   "Column(66).Order=67"
      Splits(0)._ColumnProps(334)=   "Column(67).Width=2725"
      Splits(0)._ColumnProps(335)=   "Column(67).DividerColor=0"
      Splits(0)._ColumnProps(336)=   "Column(67)._WidthInPix=2646"
      Splits(0)._ColumnProps(337)=   "Column(67).Visible=0"
      Splits(0)._ColumnProps(338)=   "Column(67).Order=68"
      Splits(0)._ColumnProps(339)=   "Column(68).Width=2196"
      Splits(0)._ColumnProps(340)=   "Column(68).DividerColor=0"
      Splits(0)._ColumnProps(341)=   "Column(68)._WidthInPix=2117"
      Splits(0)._ColumnProps(342)=   "Column(68).Order=69"
      Splits(1)._UserFlags=   0
      Splits(1).AllowRowSizing=   0   'False
      Splits(1).RecordSelectors=   0   'False
      Splits(1).RecordSelectorWidth=   503
      Splits(1).ScrollBars=   3
      Splits(1).DividerColor=   12632256
      Splits(1).SpringMode=   0   'False
      Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(1)._ColumnProps(0)=   "Columns.Count=69"
      Splits(1)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(1)._ColumnProps(4)=   "Column(0).Visible=0"
      Splits(1)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(1)._ColumnProps(6)=   "Column(1).Width=1561"
      Splits(1)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(1)._ColumnProps(8)=   "Column(1)._WidthInPix=1482"
      Splits(1)._ColumnProps(9)=   "Column(1).Visible=0"
      Splits(1)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(1)._ColumnProps(11)=   "Column(2).Width=2725"
      Splits(1)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(1)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
      Splits(1)._ColumnProps(14)=   "Column(2).Order=3"
      Splits(1)._ColumnProps(15)=   "Column(3).Width=2725"
      Splits(1)._ColumnProps(16)=   "Column(3).DividerColor=0"
      Splits(1)._ColumnProps(17)=   "Column(3)._WidthInPix=2646"
      Splits(1)._ColumnProps(18)=   "Column(3).Order=4"
      Splits(1)._ColumnProps(19)=   "Column(4).Width=2725"
      Splits(1)._ColumnProps(20)=   "Column(4).DividerColor=0"
      Splits(1)._ColumnProps(21)=   "Column(4)._WidthInPix=2646"
      Splits(1)._ColumnProps(22)=   "Column(4).Order=5"
      Splits(1)._ColumnProps(23)=   "Column(5).Width=2725"
      Splits(1)._ColumnProps(24)=   "Column(5).DividerColor=0"
      Splits(1)._ColumnProps(25)=   "Column(5)._WidthInPix=2646"
      Splits(1)._ColumnProps(26)=   "Column(5).Order=6"
      Splits(1)._ColumnProps(27)=   "Column(6).Width=2725"
      Splits(1)._ColumnProps(28)=   "Column(6).DividerColor=0"
      Splits(1)._ColumnProps(29)=   "Column(6)._WidthInPix=2646"
      Splits(1)._ColumnProps(30)=   "Column(6).Order=7"
      Splits(1)._ColumnProps(31)=   "Column(7).Width=2725"
      Splits(1)._ColumnProps(32)=   "Column(7).DividerColor=0"
      Splits(1)._ColumnProps(33)=   "Column(7)._WidthInPix=2646"
      Splits(1)._ColumnProps(34)=   "Column(7).Order=8"
      Splits(1)._ColumnProps(35)=   "Column(8).Width=2725"
      Splits(1)._ColumnProps(36)=   "Column(8).DividerColor=0"
      Splits(1)._ColumnProps(37)=   "Column(8)._WidthInPix=2646"
      Splits(1)._ColumnProps(38)=   "Column(8).Order=9"
      Splits(1)._ColumnProps(39)=   "Column(9).Width=2725"
      Splits(1)._ColumnProps(40)=   "Column(9).DividerColor=0"
      Splits(1)._ColumnProps(41)=   "Column(9)._WidthInPix=2646"
      Splits(1)._ColumnProps(42)=   "Column(9).Order=10"
      Splits(1)._ColumnProps(43)=   "Column(10).Width=2725"
      Splits(1)._ColumnProps(44)=   "Column(10).DividerColor=0"
      Splits(1)._ColumnProps(45)=   "Column(10)._WidthInPix=2646"
      Splits(1)._ColumnProps(46)=   "Column(10).Order=11"
      Splits(1)._ColumnProps(47)=   "Column(11).Width=2725"
      Splits(1)._ColumnProps(48)=   "Column(11).DividerColor=0"
      Splits(1)._ColumnProps(49)=   "Column(11)._WidthInPix=2646"
      Splits(1)._ColumnProps(50)=   "Column(11).Order=12"
      Splits(1)._ColumnProps(51)=   "Column(12).Width=2725"
      Splits(1)._ColumnProps(52)=   "Column(12).DividerColor=0"
      Splits(1)._ColumnProps(53)=   "Column(12)._WidthInPix=2646"
      Splits(1)._ColumnProps(54)=   "Column(12).Order=13"
      Splits(1)._ColumnProps(55)=   "Column(13).Width=2725"
      Splits(1)._ColumnProps(56)=   "Column(13).DividerColor=0"
      Splits(1)._ColumnProps(57)=   "Column(13)._WidthInPix=2646"
      Splits(1)._ColumnProps(58)=   "Column(13).Order=14"
      Splits(1)._ColumnProps(59)=   "Column(14).Width=2725"
      Splits(1)._ColumnProps(60)=   "Column(14).DividerColor=0"
      Splits(1)._ColumnProps(61)=   "Column(14)._WidthInPix=2646"
      Splits(1)._ColumnProps(62)=   "Column(14).Order=15"
      Splits(1)._ColumnProps(63)=   "Column(15).Width=2725"
      Splits(1)._ColumnProps(64)=   "Column(15).DividerColor=0"
      Splits(1)._ColumnProps(65)=   "Column(15)._WidthInPix=2646"
      Splits(1)._ColumnProps(66)=   "Column(15).Order=16"
      Splits(1)._ColumnProps(67)=   "Column(16).Width=2725"
      Splits(1)._ColumnProps(68)=   "Column(16).DividerColor=0"
      Splits(1)._ColumnProps(69)=   "Column(16)._WidthInPix=2646"
      Splits(1)._ColumnProps(70)=   "Column(16).Order=17"
      Splits(1)._ColumnProps(71)=   "Column(17).Width=2725"
      Splits(1)._ColumnProps(72)=   "Column(17).DividerColor=0"
      Splits(1)._ColumnProps(73)=   "Column(17)._WidthInPix=2646"
      Splits(1)._ColumnProps(74)=   "Column(17).Order=18"
      Splits(1)._ColumnProps(75)=   "Column(18).Width=2725"
      Splits(1)._ColumnProps(76)=   "Column(18).DividerColor=0"
      Splits(1)._ColumnProps(77)=   "Column(18)._WidthInPix=2646"
      Splits(1)._ColumnProps(78)=   "Column(18).Order=19"
      Splits(1)._ColumnProps(79)=   "Column(19).Width=2725"
      Splits(1)._ColumnProps(80)=   "Column(19).DividerColor=0"
      Splits(1)._ColumnProps(81)=   "Column(19)._WidthInPix=2646"
      Splits(1)._ColumnProps(82)=   "Column(19).Order=20"
      Splits(1)._ColumnProps(83)=   "Column(20).Width=2725"
      Splits(1)._ColumnProps(84)=   "Column(20).DividerColor=0"
      Splits(1)._ColumnProps(85)=   "Column(20)._WidthInPix=2646"
      Splits(1)._ColumnProps(86)=   "Column(20).Order=21"
      Splits(1)._ColumnProps(87)=   "Column(21).Width=2725"
      Splits(1)._ColumnProps(88)=   "Column(21).DividerColor=0"
      Splits(1)._ColumnProps(89)=   "Column(21)._WidthInPix=2646"
      Splits(1)._ColumnProps(90)=   "Column(21).Order=22"
      Splits(1)._ColumnProps(91)=   "Column(22).Width=2725"
      Splits(1)._ColumnProps(92)=   "Column(22).DividerColor=0"
      Splits(1)._ColumnProps(93)=   "Column(22)._WidthInPix=2646"
      Splits(1)._ColumnProps(94)=   "Column(22).Order=23"
      Splits(1)._ColumnProps(95)=   "Column(23).Width=2725"
      Splits(1)._ColumnProps(96)=   "Column(23).DividerColor=0"
      Splits(1)._ColumnProps(97)=   "Column(23)._WidthInPix=2646"
      Splits(1)._ColumnProps(98)=   "Column(23).Order=24"
      Splits(1)._ColumnProps(99)=   "Column(24).Width=2725"
      Splits(1)._ColumnProps(100)=   "Column(24).DividerColor=0"
      Splits(1)._ColumnProps(101)=   "Column(24)._WidthInPix=2646"
      Splits(1)._ColumnProps(102)=   "Column(24).Order=25"
      Splits(1)._ColumnProps(103)=   "Column(25).Width=2725"
      Splits(1)._ColumnProps(104)=   "Column(25).DividerColor=0"
      Splits(1)._ColumnProps(105)=   "Column(25)._WidthInPix=2646"
      Splits(1)._ColumnProps(106)=   "Column(25).Order=26"
      Splits(1)._ColumnProps(107)=   "Column(26).Width=2725"
      Splits(1)._ColumnProps(108)=   "Column(26).DividerColor=0"
      Splits(1)._ColumnProps(109)=   "Column(26)._WidthInPix=2646"
      Splits(1)._ColumnProps(110)=   "Column(26).Order=27"
      Splits(1)._ColumnProps(111)=   "Column(27).Width=2725"
      Splits(1)._ColumnProps(112)=   "Column(27).DividerColor=0"
      Splits(1)._ColumnProps(113)=   "Column(27)._WidthInPix=2646"
      Splits(1)._ColumnProps(114)=   "Column(27).Order=28"
      Splits(1)._ColumnProps(115)=   "Column(28).Width=2725"
      Splits(1)._ColumnProps(116)=   "Column(28).DividerColor=0"
      Splits(1)._ColumnProps(117)=   "Column(28)._WidthInPix=2646"
      Splits(1)._ColumnProps(118)=   "Column(28).Order=29"
      Splits(1)._ColumnProps(119)=   "Column(29).Width=2725"
      Splits(1)._ColumnProps(120)=   "Column(29).DividerColor=0"
      Splits(1)._ColumnProps(121)=   "Column(29)._WidthInPix=2646"
      Splits(1)._ColumnProps(122)=   "Column(29).Order=30"
      Splits(1)._ColumnProps(123)=   "Column(30).Width=2725"
      Splits(1)._ColumnProps(124)=   "Column(30).DividerColor=0"
      Splits(1)._ColumnProps(125)=   "Column(30)._WidthInPix=2646"
      Splits(1)._ColumnProps(126)=   "Column(30).Order=31"
      Splits(1)._ColumnProps(127)=   "Column(31).Width=2725"
      Splits(1)._ColumnProps(128)=   "Column(31).DividerColor=0"
      Splits(1)._ColumnProps(129)=   "Column(31)._WidthInPix=2646"
      Splits(1)._ColumnProps(130)=   "Column(31).Order=32"
      Splits(1)._ColumnProps(131)=   "Column(32).Width=2725"
      Splits(1)._ColumnProps(132)=   "Column(32).DividerColor=0"
      Splits(1)._ColumnProps(133)=   "Column(32)._WidthInPix=2646"
      Splits(1)._ColumnProps(134)=   "Column(32).Order=33"
      Splits(1)._ColumnProps(135)=   "Column(33).Width=2725"
      Splits(1)._ColumnProps(136)=   "Column(33).DividerColor=0"
      Splits(1)._ColumnProps(137)=   "Column(33)._WidthInPix=2646"
      Splits(1)._ColumnProps(138)=   "Column(33).Order=34"
      Splits(1)._ColumnProps(139)=   "Column(34).Width=2725"
      Splits(1)._ColumnProps(140)=   "Column(34).DividerColor=0"
      Splits(1)._ColumnProps(141)=   "Column(34)._WidthInPix=2646"
      Splits(1)._ColumnProps(142)=   "Column(34).Order=35"
      Splits(1)._ColumnProps(143)=   "Column(35).Width=2725"
      Splits(1)._ColumnProps(144)=   "Column(35).DividerColor=0"
      Splits(1)._ColumnProps(145)=   "Column(35)._WidthInPix=2646"
      Splits(1)._ColumnProps(146)=   "Column(35).Order=36"
      Splits(1)._ColumnProps(147)=   "Column(36).Width=2725"
      Splits(1)._ColumnProps(148)=   "Column(36).DividerColor=0"
      Splits(1)._ColumnProps(149)=   "Column(36)._WidthInPix=2646"
      Splits(1)._ColumnProps(150)=   "Column(36).Order=37"
      Splits(1)._ColumnProps(151)=   "Column(37).Width=2725"
      Splits(1)._ColumnProps(152)=   "Column(37).DividerColor=0"
      Splits(1)._ColumnProps(153)=   "Column(37)._WidthInPix=2646"
      Splits(1)._ColumnProps(154)=   "Column(37).Order=38"
      Splits(1)._ColumnProps(155)=   "Column(38).Width=2725"
      Splits(1)._ColumnProps(156)=   "Column(38).DividerColor=0"
      Splits(1)._ColumnProps(157)=   "Column(38)._WidthInPix=2646"
      Splits(1)._ColumnProps(158)=   "Column(38).Order=39"
      Splits(1)._ColumnProps(159)=   "Column(39).Width=2725"
      Splits(1)._ColumnProps(160)=   "Column(39).DividerColor=0"
      Splits(1)._ColumnProps(161)=   "Column(39)._WidthInPix=2646"
      Splits(1)._ColumnProps(162)=   "Column(39).Order=40"
      Splits(1)._ColumnProps(163)=   "Column(40).Width=2725"
      Splits(1)._ColumnProps(164)=   "Column(40).DividerColor=0"
      Splits(1)._ColumnProps(165)=   "Column(40)._WidthInPix=2646"
      Splits(1)._ColumnProps(166)=   "Column(40).Order=41"
      Splits(1)._ColumnProps(167)=   "Column(41).Width=2725"
      Splits(1)._ColumnProps(168)=   "Column(41).DividerColor=0"
      Splits(1)._ColumnProps(169)=   "Column(41)._WidthInPix=2646"
      Splits(1)._ColumnProps(170)=   "Column(41).Order=42"
      Splits(1)._ColumnProps(171)=   "Column(42).Width=2725"
      Splits(1)._ColumnProps(172)=   "Column(42).DividerColor=0"
      Splits(1)._ColumnProps(173)=   "Column(42)._WidthInPix=2646"
      Splits(1)._ColumnProps(174)=   "Column(42).Order=43"
      Splits(1)._ColumnProps(175)=   "Column(43).Width=2725"
      Splits(1)._ColumnProps(176)=   "Column(43).DividerColor=0"
      Splits(1)._ColumnProps(177)=   "Column(43)._WidthInPix=2646"
      Splits(1)._ColumnProps(178)=   "Column(43).Order=44"
      Splits(1)._ColumnProps(179)=   "Column(44).Width=2725"
      Splits(1)._ColumnProps(180)=   "Column(44).DividerColor=0"
      Splits(1)._ColumnProps(181)=   "Column(44)._WidthInPix=2646"
      Splits(1)._ColumnProps(182)=   "Column(44).Order=45"
      Splits(1)._ColumnProps(183)=   "Column(45).Width=2725"
      Splits(1)._ColumnProps(184)=   "Column(45).DividerColor=0"
      Splits(1)._ColumnProps(185)=   "Column(45)._WidthInPix=2646"
      Splits(1)._ColumnProps(186)=   "Column(45).Order=46"
      Splits(1)._ColumnProps(187)=   "Column(46).Width=2725"
      Splits(1)._ColumnProps(188)=   "Column(46).DividerColor=0"
      Splits(1)._ColumnProps(189)=   "Column(46)._WidthInPix=2646"
      Splits(1)._ColumnProps(190)=   "Column(46).Order=47"
      Splits(1)._ColumnProps(191)=   "Column(47).Width=2725"
      Splits(1)._ColumnProps(192)=   "Column(47).DividerColor=0"
      Splits(1)._ColumnProps(193)=   "Column(47)._WidthInPix=2646"
      Splits(1)._ColumnProps(194)=   "Column(47).Order=48"
      Splits(1)._ColumnProps(195)=   "Column(48).Width=2725"
      Splits(1)._ColumnProps(196)=   "Column(48).DividerColor=0"
      Splits(1)._ColumnProps(197)=   "Column(48)._WidthInPix=2646"
      Splits(1)._ColumnProps(198)=   "Column(48).Order=49"
      Splits(1)._ColumnProps(199)=   "Column(49).Width=2725"
      Splits(1)._ColumnProps(200)=   "Column(49).DividerColor=0"
      Splits(1)._ColumnProps(201)=   "Column(49)._WidthInPix=2646"
      Splits(1)._ColumnProps(202)=   "Column(49).Order=50"
      Splits(1)._ColumnProps(203)=   "Column(50).Width=2725"
      Splits(1)._ColumnProps(204)=   "Column(50).DividerColor=0"
      Splits(1)._ColumnProps(205)=   "Column(50)._WidthInPix=2646"
      Splits(1)._ColumnProps(206)=   "Column(50).Order=51"
      Splits(1)._ColumnProps(207)=   "Column(51).Width=2725"
      Splits(1)._ColumnProps(208)=   "Column(51).DividerColor=0"
      Splits(1)._ColumnProps(209)=   "Column(51)._WidthInPix=2646"
      Splits(1)._ColumnProps(210)=   "Column(51).Order=52"
      Splits(1)._ColumnProps(211)=   "Column(52).Width=2725"
      Splits(1)._ColumnProps(212)=   "Column(52).DividerColor=0"
      Splits(1)._ColumnProps(213)=   "Column(52)._WidthInPix=2646"
      Splits(1)._ColumnProps(214)=   "Column(52).Order=53"
      Splits(1)._ColumnProps(215)=   "Column(53).Width=2725"
      Splits(1)._ColumnProps(216)=   "Column(53).DividerColor=0"
      Splits(1)._ColumnProps(217)=   "Column(53)._WidthInPix=2646"
      Splits(1)._ColumnProps(218)=   "Column(53).Order=54"
      Splits(1)._ColumnProps(219)=   "Column(54).Width=2725"
      Splits(1)._ColumnProps(220)=   "Column(54).DividerColor=0"
      Splits(1)._ColumnProps(221)=   "Column(54)._WidthInPix=2646"
      Splits(1)._ColumnProps(222)=   "Column(54).Order=55"
      Splits(1)._ColumnProps(223)=   "Column(55).Width=2725"
      Splits(1)._ColumnProps(224)=   "Column(55).DividerColor=0"
      Splits(1)._ColumnProps(225)=   "Column(55)._WidthInPix=2646"
      Splits(1)._ColumnProps(226)=   "Column(55).Order=56"
      Splits(1)._ColumnProps(227)=   "Column(56).Width=2725"
      Splits(1)._ColumnProps(228)=   "Column(56).DividerColor=0"
      Splits(1)._ColumnProps(229)=   "Column(56)._WidthInPix=2646"
      Splits(1)._ColumnProps(230)=   "Column(56).Order=57"
      Splits(1)._ColumnProps(231)=   "Column(57).Width=2725"
      Splits(1)._ColumnProps(232)=   "Column(57).DividerColor=0"
      Splits(1)._ColumnProps(233)=   "Column(57)._WidthInPix=2646"
      Splits(1)._ColumnProps(234)=   "Column(57).Order=58"
      Splits(1)._ColumnProps(235)=   "Column(58).Width=2725"
      Splits(1)._ColumnProps(236)=   "Column(58).DividerColor=0"
      Splits(1)._ColumnProps(237)=   "Column(58)._WidthInPix=2646"
      Splits(1)._ColumnProps(238)=   "Column(58).Order=59"
      Splits(1)._ColumnProps(239)=   "Column(59).Width=2725"
      Splits(1)._ColumnProps(240)=   "Column(59).DividerColor=0"
      Splits(1)._ColumnProps(241)=   "Column(59)._WidthInPix=2646"
      Splits(1)._ColumnProps(242)=   "Column(59).Order=60"
      Splits(1)._ColumnProps(243)=   "Column(60).Width=2725"
      Splits(1)._ColumnProps(244)=   "Column(60).DividerColor=0"
      Splits(1)._ColumnProps(245)=   "Column(60)._WidthInPix=2646"
      Splits(1)._ColumnProps(246)=   "Column(60).Order=61"
      Splits(1)._ColumnProps(247)=   "Column(61).Width=2725"
      Splits(1)._ColumnProps(248)=   "Column(61).DividerColor=0"
      Splits(1)._ColumnProps(249)=   "Column(61)._WidthInPix=2646"
      Splits(1)._ColumnProps(250)=   "Column(61).Order=62"
      Splits(1)._ColumnProps(251)=   "Column(62).Width=2725"
      Splits(1)._ColumnProps(252)=   "Column(62).DividerColor=0"
      Splits(1)._ColumnProps(253)=   "Column(62)._WidthInPix=2646"
      Splits(1)._ColumnProps(254)=   "Column(62).Order=63"
      Splits(1)._ColumnProps(255)=   "Column(63).Width=2725"
      Splits(1)._ColumnProps(256)=   "Column(63).DividerColor=0"
      Splits(1)._ColumnProps(257)=   "Column(63)._WidthInPix=2646"
      Splits(1)._ColumnProps(258)=   "Column(63).Order=64"
      Splits(1)._ColumnProps(259)=   "Column(64).Width=2725"
      Splits(1)._ColumnProps(260)=   "Column(64).DividerColor=0"
      Splits(1)._ColumnProps(261)=   "Column(64)._WidthInPix=2646"
      Splits(1)._ColumnProps(262)=   "Column(64).Order=65"
      Splits(1)._ColumnProps(263)=   "Column(65).Width=2725"
      Splits(1)._ColumnProps(264)=   "Column(65).DividerColor=0"
      Splits(1)._ColumnProps(265)=   "Column(65)._WidthInPix=2646"
      Splits(1)._ColumnProps(266)=   "Column(65).Order=66"
      Splits(1)._ColumnProps(267)=   "Column(66).Width=2725"
      Splits(1)._ColumnProps(268)=   "Column(66).DividerColor=0"
      Splits(1)._ColumnProps(269)=   "Column(66)._WidthInPix=2646"
      Splits(1)._ColumnProps(270)=   "Column(66).Order=67"
      Splits(1)._ColumnProps(271)=   "Column(67).Width=2725"
      Splits(1)._ColumnProps(272)=   "Column(67).DividerColor=0"
      Splits(1)._ColumnProps(273)=   "Column(67)._WidthInPix=2646"
      Splits(1)._ColumnProps(274)=   "Column(67).Order=68"
      Splits(1)._ColumnProps(275)=   "Column(68).Width=2196"
      Splits(1)._ColumnProps(276)=   "Column(68).DividerColor=0"
      Splits(1)._ColumnProps(277)=   "Column(68)._WidthInPix=2117"
      Splits(1)._ColumnProps(278)=   "Column(68).Order=69"
      Splits.Count    =   2
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Times New Roman"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Times New Roman"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      AllowArrows     =   0   'False
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   12632256
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
      DirectionAfterEnter=   1
      MaxRows         =   250000
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=12,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Times New Roman"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&H404040&"
      _StyleDefs(11)  =   ":id=2,.fgcolor=&HFFFFFF&"
      _StyleDefs(12)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(13)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(14)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&H800080&"
      _StyleDefs(15)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(16)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H800080&"
      _StyleDefs(17)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(18)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40,.bgcolor=&H40FF00&"
      _StyleDefs(19)  =   ":id=10,.borderColor=&H80000018&"
      _StyleDefs(20)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41,.locked=0"
      _StyleDefs(21)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(22)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(23)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(24)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(25)  =   ":id=14,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(26)  =   ":id=14,.fontname=Times New Roman"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=58,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=78,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=75,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=76,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=77,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=86,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=83,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=84,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=85,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=94,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=91,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=92,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=93,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=102,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=99,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=100,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=101,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=110,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=107,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=108,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=109,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=118,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=115,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=116,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=117,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=126,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=123,.parent=14"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=124,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=125,.parent=17"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=134,.parent=13"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=131,.parent=14"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=132,.parent=15"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=133,.parent=17"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=142,.parent=13"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=139,.parent=14"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=140,.parent=15"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=141,.parent=17"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=150,.parent=13"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=147,.parent=14"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=148,.parent=15"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=149,.parent=17"
      _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=158,.parent=13"
      _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=155,.parent=14"
      _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=156,.parent=15"
      _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=157,.parent=17"
      _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=166,.parent=13"
      _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=163,.parent=14"
      _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=164,.parent=15"
      _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=165,.parent=17"
      _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=174,.parent=13"
      _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=171,.parent=14"
      _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=172,.parent=15"
      _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=173,.parent=17"
      _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=182,.parent=13"
      _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=179,.parent=14"
      _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=180,.parent=15"
      _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=181,.parent=17"
      _StyleDefs(100) =   "Splits(0).Columns(16).Style:id=190,.parent=13"
      _StyleDefs(101) =   "Splits(0).Columns(16).HeadingStyle:id=187,.parent=14"
      _StyleDefs(102) =   "Splits(0).Columns(16).FooterStyle:id=188,.parent=15"
      _StyleDefs(103) =   "Splits(0).Columns(16).EditorStyle:id=189,.parent=17"
      _StyleDefs(104) =   "Splits(0).Columns(17).Style:id=198,.parent=13"
      _StyleDefs(105) =   "Splits(0).Columns(17).HeadingStyle:id=195,.parent=14"
      _StyleDefs(106) =   "Splits(0).Columns(17).FooterStyle:id=196,.parent=15"
      _StyleDefs(107) =   "Splits(0).Columns(17).EditorStyle:id=197,.parent=17"
      _StyleDefs(108) =   "Splits(0).Columns(18).Style:id=206,.parent=13"
      _StyleDefs(109) =   "Splits(0).Columns(18).HeadingStyle:id=203,.parent=14"
      _StyleDefs(110) =   "Splits(0).Columns(18).FooterStyle:id=204,.parent=15"
      _StyleDefs(111) =   "Splits(0).Columns(18).EditorStyle:id=205,.parent=17"
      _StyleDefs(112) =   "Splits(0).Columns(19).Style:id=214,.parent=13"
      _StyleDefs(113) =   "Splits(0).Columns(19).HeadingStyle:id=211,.parent=14"
      _StyleDefs(114) =   "Splits(0).Columns(19).FooterStyle:id=212,.parent=15"
      _StyleDefs(115) =   "Splits(0).Columns(19).EditorStyle:id=213,.parent=17"
      _StyleDefs(116) =   "Splits(0).Columns(20).Style:id=222,.parent=13"
      _StyleDefs(117) =   "Splits(0).Columns(20).HeadingStyle:id=219,.parent=14"
      _StyleDefs(118) =   "Splits(0).Columns(20).FooterStyle:id=220,.parent=15"
      _StyleDefs(119) =   "Splits(0).Columns(20).EditorStyle:id=221,.parent=17"
      _StyleDefs(120) =   "Splits(0).Columns(21).Style:id=230,.parent=13"
      _StyleDefs(121) =   "Splits(0).Columns(21).HeadingStyle:id=227,.parent=14"
      _StyleDefs(122) =   "Splits(0).Columns(21).FooterStyle:id=228,.parent=15"
      _StyleDefs(123) =   "Splits(0).Columns(21).EditorStyle:id=229,.parent=17"
      _StyleDefs(124) =   "Splits(0).Columns(22).Style:id=238,.parent=13"
      _StyleDefs(125) =   "Splits(0).Columns(22).HeadingStyle:id=235,.parent=14"
      _StyleDefs(126) =   "Splits(0).Columns(22).FooterStyle:id=236,.parent=15"
      _StyleDefs(127) =   "Splits(0).Columns(22).EditorStyle:id=237,.parent=17"
      _StyleDefs(128) =   "Splits(0).Columns(23).Style:id=246,.parent=13"
      _StyleDefs(129) =   "Splits(0).Columns(23).HeadingStyle:id=243,.parent=14"
      _StyleDefs(130) =   "Splits(0).Columns(23).FooterStyle:id=244,.parent=15"
      _StyleDefs(131) =   "Splits(0).Columns(23).EditorStyle:id=245,.parent=17"
      _StyleDefs(132) =   "Splits(0).Columns(24).Style:id=254,.parent=13"
      _StyleDefs(133) =   "Splits(0).Columns(24).HeadingStyle:id=251,.parent=14"
      _StyleDefs(134) =   "Splits(0).Columns(24).FooterStyle:id=252,.parent=15"
      _StyleDefs(135) =   "Splits(0).Columns(24).EditorStyle:id=253,.parent=17"
      _StyleDefs(136) =   "Splits(0).Columns(25).Style:id=262,.parent=13"
      _StyleDefs(137) =   "Splits(0).Columns(25).HeadingStyle:id=259,.parent=14"
      _StyleDefs(138) =   "Splits(0).Columns(25).FooterStyle:id=260,.parent=15"
      _StyleDefs(139) =   "Splits(0).Columns(25).EditorStyle:id=261,.parent=17"
      _StyleDefs(140) =   "Splits(0).Columns(26).Style:id=270,.parent=13"
      _StyleDefs(141) =   "Splits(0).Columns(26).HeadingStyle:id=267,.parent=14"
      _StyleDefs(142) =   "Splits(0).Columns(26).FooterStyle:id=268,.parent=15"
      _StyleDefs(143) =   "Splits(0).Columns(26).EditorStyle:id=269,.parent=17"
      _StyleDefs(144) =   "Splits(0).Columns(27).Style:id=278,.parent=13"
      _StyleDefs(145) =   "Splits(0).Columns(27).HeadingStyle:id=275,.parent=14"
      _StyleDefs(146) =   "Splits(0).Columns(27).FooterStyle:id=276,.parent=15"
      _StyleDefs(147) =   "Splits(0).Columns(27).EditorStyle:id=277,.parent=17"
      _StyleDefs(148) =   "Splits(0).Columns(28).Style:id=286,.parent=13"
      _StyleDefs(149) =   "Splits(0).Columns(28).HeadingStyle:id=283,.parent=14"
      _StyleDefs(150) =   "Splits(0).Columns(28).FooterStyle:id=284,.parent=15"
      _StyleDefs(151) =   "Splits(0).Columns(28).EditorStyle:id=285,.parent=17"
      _StyleDefs(152) =   "Splits(0).Columns(29).Style:id=294,.parent=13"
      _StyleDefs(153) =   "Splits(0).Columns(29).HeadingStyle:id=291,.parent=14"
      _StyleDefs(154) =   "Splits(0).Columns(29).FooterStyle:id=292,.parent=15"
      _StyleDefs(155) =   "Splits(0).Columns(29).EditorStyle:id=293,.parent=17"
      _StyleDefs(156) =   "Splits(0).Columns(30).Style:id=302,.parent=13"
      _StyleDefs(157) =   "Splits(0).Columns(30).HeadingStyle:id=299,.parent=14"
      _StyleDefs(158) =   "Splits(0).Columns(30).FooterStyle:id=300,.parent=15"
      _StyleDefs(159) =   "Splits(0).Columns(30).EditorStyle:id=301,.parent=17"
      _StyleDefs(160) =   "Splits(0).Columns(31).Style:id=322,.parent=13"
      _StyleDefs(161) =   "Splits(0).Columns(31).HeadingStyle:id=319,.parent=14"
      _StyleDefs(162) =   "Splits(0).Columns(31).FooterStyle:id=320,.parent=15"
      _StyleDefs(163) =   "Splits(0).Columns(31).EditorStyle:id=321,.parent=17"
      _StyleDefs(164) =   "Splits(0).Columns(32).Style:id=330,.parent=13"
      _StyleDefs(165) =   "Splits(0).Columns(32).HeadingStyle:id=327,.parent=14"
      _StyleDefs(166) =   "Splits(0).Columns(32).FooterStyle:id=328,.parent=15"
      _StyleDefs(167) =   "Splits(0).Columns(32).EditorStyle:id=329,.parent=17"
      _StyleDefs(168) =   "Splits(0).Columns(33).Style:id=338,.parent=13"
      _StyleDefs(169) =   "Splits(0).Columns(33).HeadingStyle:id=335,.parent=14"
      _StyleDefs(170) =   "Splits(0).Columns(33).FooterStyle:id=336,.parent=15"
      _StyleDefs(171) =   "Splits(0).Columns(33).EditorStyle:id=337,.parent=17"
      _StyleDefs(172) =   "Splits(0).Columns(34).Style:id=346,.parent=13"
      _StyleDefs(173) =   "Splits(0).Columns(34).HeadingStyle:id=343,.parent=14"
      _StyleDefs(174) =   "Splits(0).Columns(34).FooterStyle:id=344,.parent=15"
      _StyleDefs(175) =   "Splits(0).Columns(34).EditorStyle:id=345,.parent=17"
      _StyleDefs(176) =   "Splits(0).Columns(35).Style:id=354,.parent=13"
      _StyleDefs(177) =   "Splits(0).Columns(35).HeadingStyle:id=351,.parent=14"
      _StyleDefs(178) =   "Splits(0).Columns(35).FooterStyle:id=352,.parent=15"
      _StyleDefs(179) =   "Splits(0).Columns(35).EditorStyle:id=353,.parent=17"
      _StyleDefs(180) =   "Splits(0).Columns(36).Style:id=362,.parent=13"
      _StyleDefs(181) =   "Splits(0).Columns(36).HeadingStyle:id=359,.parent=14"
      _StyleDefs(182) =   "Splits(0).Columns(36).FooterStyle:id=360,.parent=15"
      _StyleDefs(183) =   "Splits(0).Columns(36).EditorStyle:id=361,.parent=17"
      _StyleDefs(184) =   "Splits(0).Columns(37).Style:id=370,.parent=13"
      _StyleDefs(185) =   "Splits(0).Columns(37).HeadingStyle:id=367,.parent=14"
      _StyleDefs(186) =   "Splits(0).Columns(37).FooterStyle:id=368,.parent=15"
      _StyleDefs(187) =   "Splits(0).Columns(37).EditorStyle:id=369,.parent=17"
      _StyleDefs(188) =   "Splits(0).Columns(38).Style:id=378,.parent=13"
      _StyleDefs(189) =   "Splits(0).Columns(38).HeadingStyle:id=375,.parent=14"
      _StyleDefs(190) =   "Splits(0).Columns(38).FooterStyle:id=376,.parent=15"
      _StyleDefs(191) =   "Splits(0).Columns(38).EditorStyle:id=377,.parent=17"
      _StyleDefs(192) =   "Splits(0).Columns(39).Style:id=386,.parent=13"
      _StyleDefs(193) =   "Splits(0).Columns(39).HeadingStyle:id=383,.parent=14"
      _StyleDefs(194) =   "Splits(0).Columns(39).FooterStyle:id=384,.parent=15"
      _StyleDefs(195) =   "Splits(0).Columns(39).EditorStyle:id=385,.parent=17"
      _StyleDefs(196) =   "Splits(0).Columns(40).Style:id=394,.parent=13"
      _StyleDefs(197) =   "Splits(0).Columns(40).HeadingStyle:id=391,.parent=14"
      _StyleDefs(198) =   "Splits(0).Columns(40).FooterStyle:id=392,.parent=15"
      _StyleDefs(199) =   "Splits(0).Columns(40).EditorStyle:id=393,.parent=17"
      _StyleDefs(200) =   "Splits(0).Columns(41).Style:id=402,.parent=13"
      _StyleDefs(201) =   "Splits(0).Columns(41).HeadingStyle:id=399,.parent=14"
      _StyleDefs(202) =   "Splits(0).Columns(41).FooterStyle:id=400,.parent=15"
      _StyleDefs(203) =   "Splits(0).Columns(41).EditorStyle:id=401,.parent=17"
      _StyleDefs(204) =   "Splits(0).Columns(42).Style:id=410,.parent=13"
      _StyleDefs(205) =   "Splits(0).Columns(42).HeadingStyle:id=407,.parent=14"
      _StyleDefs(206) =   "Splits(0).Columns(42).FooterStyle:id=408,.parent=15"
      _StyleDefs(207) =   "Splits(0).Columns(42).EditorStyle:id=409,.parent=17"
      _StyleDefs(208) =   "Splits(0).Columns(43).Style:id=418,.parent=13"
      _StyleDefs(209) =   "Splits(0).Columns(43).HeadingStyle:id=415,.parent=14"
      _StyleDefs(210) =   "Splits(0).Columns(43).FooterStyle:id=416,.parent=15"
      _StyleDefs(211) =   "Splits(0).Columns(43).EditorStyle:id=417,.parent=17"
      _StyleDefs(212) =   "Splits(0).Columns(44).Style:id=426,.parent=13"
      _StyleDefs(213) =   "Splits(0).Columns(44).HeadingStyle:id=423,.parent=14"
      _StyleDefs(214) =   "Splits(0).Columns(44).FooterStyle:id=424,.parent=15"
      _StyleDefs(215) =   "Splits(0).Columns(44).EditorStyle:id=425,.parent=17"
      _StyleDefs(216) =   "Splits(0).Columns(45).Style:id=434,.parent=13"
      _StyleDefs(217) =   "Splits(0).Columns(45).HeadingStyle:id=431,.parent=14"
      _StyleDefs(218) =   "Splits(0).Columns(45).FooterStyle:id=432,.parent=15"
      _StyleDefs(219) =   "Splits(0).Columns(45).EditorStyle:id=433,.parent=17"
      _StyleDefs(220) =   "Splits(0).Columns(46).Style:id=442,.parent=13"
      _StyleDefs(221) =   "Splits(0).Columns(46).HeadingStyle:id=439,.parent=14"
      _StyleDefs(222) =   "Splits(0).Columns(46).FooterStyle:id=440,.parent=15"
      _StyleDefs(223) =   "Splits(0).Columns(46).EditorStyle:id=441,.parent=17"
      _StyleDefs(224) =   "Splits(0).Columns(47).Style:id=450,.parent=13"
      _StyleDefs(225) =   "Splits(0).Columns(47).HeadingStyle:id=447,.parent=14"
      _StyleDefs(226) =   "Splits(0).Columns(47).FooterStyle:id=448,.parent=15"
      _StyleDefs(227) =   "Splits(0).Columns(47).EditorStyle:id=449,.parent=17"
      _StyleDefs(228) =   "Splits(0).Columns(48).Style:id=458,.parent=13"
      _StyleDefs(229) =   "Splits(0).Columns(48).HeadingStyle:id=455,.parent=14"
      _StyleDefs(230) =   "Splits(0).Columns(48).FooterStyle:id=456,.parent=15"
      _StyleDefs(231) =   "Splits(0).Columns(48).EditorStyle:id=457,.parent=17"
      _StyleDefs(232) =   "Splits(0).Columns(49).Style:id=466,.parent=13"
      _StyleDefs(233) =   "Splits(0).Columns(49).HeadingStyle:id=463,.parent=14"
      _StyleDefs(234) =   "Splits(0).Columns(49).FooterStyle:id=464,.parent=15"
      _StyleDefs(235) =   "Splits(0).Columns(49).EditorStyle:id=465,.parent=17"
      _StyleDefs(236) =   "Splits(0).Columns(50).Style:id=474,.parent=13"
      _StyleDefs(237) =   "Splits(0).Columns(50).HeadingStyle:id=471,.parent=14"
      _StyleDefs(238) =   "Splits(0).Columns(50).FooterStyle:id=472,.parent=15"
      _StyleDefs(239) =   "Splits(0).Columns(50).EditorStyle:id=473,.parent=17"
      _StyleDefs(240) =   "Splits(0).Columns(51).Style:id=482,.parent=13"
      _StyleDefs(241) =   "Splits(0).Columns(51).HeadingStyle:id=479,.parent=14"
      _StyleDefs(242) =   "Splits(0).Columns(51).FooterStyle:id=480,.parent=15"
      _StyleDefs(243) =   "Splits(0).Columns(51).EditorStyle:id=481,.parent=17"
      _StyleDefs(244) =   "Splits(0).Columns(52).Style:id=490,.parent=13"
      _StyleDefs(245) =   "Splits(0).Columns(52).HeadingStyle:id=487,.parent=14"
      _StyleDefs(246) =   "Splits(0).Columns(52).FooterStyle:id=488,.parent=15"
      _StyleDefs(247) =   "Splits(0).Columns(52).EditorStyle:id=489,.parent=17"
      _StyleDefs(248) =   "Splits(0).Columns(53).Style:id=498,.parent=13"
      _StyleDefs(249) =   "Splits(0).Columns(53).HeadingStyle:id=495,.parent=14"
      _StyleDefs(250) =   "Splits(0).Columns(53).FooterStyle:id=496,.parent=15"
      _StyleDefs(251) =   "Splits(0).Columns(53).EditorStyle:id=497,.parent=17"
      _StyleDefs(252) =   "Splits(0).Columns(54).Style:id=506,.parent=13"
      _StyleDefs(253) =   "Splits(0).Columns(54).HeadingStyle:id=503,.parent=14"
      _StyleDefs(254) =   "Splits(0).Columns(54).FooterStyle:id=504,.parent=15"
      _StyleDefs(255) =   "Splits(0).Columns(54).EditorStyle:id=505,.parent=17"
      _StyleDefs(256) =   "Splits(0).Columns(55).Style:id=514,.parent=13"
      _StyleDefs(257) =   "Splits(0).Columns(55).HeadingStyle:id=511,.parent=14"
      _StyleDefs(258) =   "Splits(0).Columns(55).FooterStyle:id=512,.parent=15"
      _StyleDefs(259) =   "Splits(0).Columns(55).EditorStyle:id=513,.parent=17"
      _StyleDefs(260) =   "Splits(0).Columns(56).Style:id=522,.parent=13"
      _StyleDefs(261) =   "Splits(0).Columns(56).HeadingStyle:id=519,.parent=14"
      _StyleDefs(262) =   "Splits(0).Columns(56).FooterStyle:id=520,.parent=15"
      _StyleDefs(263) =   "Splits(0).Columns(56).EditorStyle:id=521,.parent=17"
      _StyleDefs(264) =   "Splits(0).Columns(57).Style:id=530,.parent=13"
      _StyleDefs(265) =   "Splits(0).Columns(57).HeadingStyle:id=527,.parent=14"
      _StyleDefs(266) =   "Splits(0).Columns(57).FooterStyle:id=528,.parent=15"
      _StyleDefs(267) =   "Splits(0).Columns(57).EditorStyle:id=529,.parent=17"
      _StyleDefs(268) =   "Splits(0).Columns(58).Style:id=538,.parent=13"
      _StyleDefs(269) =   "Splits(0).Columns(58).HeadingStyle:id=535,.parent=14"
      _StyleDefs(270) =   "Splits(0).Columns(58).FooterStyle:id=536,.parent=15"
      _StyleDefs(271) =   "Splits(0).Columns(58).EditorStyle:id=537,.parent=17"
      _StyleDefs(272) =   "Splits(0).Columns(59).Style:id=546,.parent=13"
      _StyleDefs(273) =   "Splits(0).Columns(59).HeadingStyle:id=543,.parent=14"
      _StyleDefs(274) =   "Splits(0).Columns(59).FooterStyle:id=544,.parent=15"
      _StyleDefs(275) =   "Splits(0).Columns(59).EditorStyle:id=545,.parent=17"
      _StyleDefs(276) =   "Splits(0).Columns(60).Style:id=554,.parent=13"
      _StyleDefs(277) =   "Splits(0).Columns(60).HeadingStyle:id=551,.parent=14"
      _StyleDefs(278) =   "Splits(0).Columns(60).FooterStyle:id=552,.parent=15"
      _StyleDefs(279) =   "Splits(0).Columns(60).EditorStyle:id=553,.parent=17"
      _StyleDefs(280) =   "Splits(0).Columns(61).Style:id=562,.parent=13"
      _StyleDefs(281) =   "Splits(0).Columns(61).HeadingStyle:id=559,.parent=14"
      _StyleDefs(282) =   "Splits(0).Columns(61).FooterStyle:id=560,.parent=15"
      _StyleDefs(283) =   "Splits(0).Columns(61).EditorStyle:id=561,.parent=17"
      _StyleDefs(284) =   "Splits(0).Columns(62).Style:id=570,.parent=13"
      _StyleDefs(285) =   "Splits(0).Columns(62).HeadingStyle:id=567,.parent=14"
      _StyleDefs(286) =   "Splits(0).Columns(62).FooterStyle:id=568,.parent=15"
      _StyleDefs(287) =   "Splits(0).Columns(62).EditorStyle:id=569,.parent=17"
      _StyleDefs(288) =   "Splits(0).Columns(63).Style:id=578,.parent=13"
      _StyleDefs(289) =   "Splits(0).Columns(63).HeadingStyle:id=575,.parent=14"
      _StyleDefs(290) =   "Splits(0).Columns(63).FooterStyle:id=576,.parent=15"
      _StyleDefs(291) =   "Splits(0).Columns(63).EditorStyle:id=577,.parent=17"
      _StyleDefs(292) =   "Splits(0).Columns(64).Style:id=586,.parent=13"
      _StyleDefs(293) =   "Splits(0).Columns(64).HeadingStyle:id=583,.parent=14"
      _StyleDefs(294) =   "Splits(0).Columns(64).FooterStyle:id=584,.parent=15"
      _StyleDefs(295) =   "Splits(0).Columns(64).EditorStyle:id=585,.parent=17"
      _StyleDefs(296) =   "Splits(0).Columns(65).Style:id=594,.parent=13"
      _StyleDefs(297) =   "Splits(0).Columns(65).HeadingStyle:id=591,.parent=14"
      _StyleDefs(298) =   "Splits(0).Columns(65).FooterStyle:id=592,.parent=15"
      _StyleDefs(299) =   "Splits(0).Columns(65).EditorStyle:id=593,.parent=17"
      _StyleDefs(300) =   "Splits(0).Columns(66).Style:id=602,.parent=13"
      _StyleDefs(301) =   "Splits(0).Columns(66).HeadingStyle:id=599,.parent=14"
      _StyleDefs(302) =   "Splits(0).Columns(66).FooterStyle:id=600,.parent=15"
      _StyleDefs(303) =   "Splits(0).Columns(66).EditorStyle:id=601,.parent=17"
      _StyleDefs(304) =   "Splits(0).Columns(67).Style:id=610,.parent=13"
      _StyleDefs(305) =   "Splits(0).Columns(67).HeadingStyle:id=607,.parent=14"
      _StyleDefs(306) =   "Splits(0).Columns(67).FooterStyle:id=608,.parent=15"
      _StyleDefs(307) =   "Splits(0).Columns(67).EditorStyle:id=609,.parent=17"
      _StyleDefs(308) =   "Splits(0).Columns(68).Style:id=28,.parent=13"
      _StyleDefs(309) =   "Splits(0).Columns(68).HeadingStyle:id=25,.parent=14"
      _StyleDefs(310) =   "Splits(0).Columns(68).FooterStyle:id=26,.parent=15"
      _StyleDefs(311) =   "Splits(0).Columns(68).EditorStyle:id=27,.parent=17"
      _StyleDefs(312) =   "Splits(1).Style:id=51,.parent=1"
      _StyleDefs(313) =   "Splits(1).CaptionStyle:id=64,.parent=4"
      _StyleDefs(314) =   "Splits(1).HeadingStyle:id=52,.parent=2,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(315) =   ":id=52,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(316) =   ":id=52,.fontname=Times New Roman"
      _StyleDefs(317) =   "Splits(1).FooterStyle:id=53,.parent=3"
      _StyleDefs(318) =   "Splits(1).InactiveStyle:id=54,.parent=5"
      _StyleDefs(319) =   "Splits(1).SelectedStyle:id=60,.parent=6"
      _StyleDefs(320) =   "Splits(1).EditorStyle:id=59,.parent=7"
      _StyleDefs(321) =   "Splits(1).HighlightRowStyle:id=61,.parent=8"
      _StyleDefs(322) =   "Splits(1).EvenRowStyle:id=62,.parent=9"
      _StyleDefs(323) =   "Splits(1).OddRowStyle:id=63,.parent=10"
      _StyleDefs(324) =   "Splits(1).RecordSelectorStyle:id=65,.parent=11"
      _StyleDefs(325) =   "Splits(1).FilterBarStyle:id=66,.parent=12"
      _StyleDefs(326) =   "Splits(1).Columns(0).Style:id=70,.parent=51"
      _StyleDefs(327) =   "Splits(1).Columns(0).HeadingStyle:id=67,.parent=52"
      _StyleDefs(328) =   "Splits(1).Columns(0).FooterStyle:id=68,.parent=53"
      _StyleDefs(329) =   "Splits(1).Columns(0).EditorStyle:id=69,.parent=59"
      _StyleDefs(330) =   "Splits(1).Columns(1).Style:id=74,.parent=51"
      _StyleDefs(331) =   "Splits(1).Columns(1).HeadingStyle:id=71,.parent=52"
      _StyleDefs(332) =   "Splits(1).Columns(1).FooterStyle:id=72,.parent=53"
      _StyleDefs(333) =   "Splits(1).Columns(1).EditorStyle:id=73,.parent=59"
      _StyleDefs(334) =   "Splits(1).Columns(2).Style:id=82,.parent=51"
      _StyleDefs(335) =   "Splits(1).Columns(2).HeadingStyle:id=79,.parent=52"
      _StyleDefs(336) =   "Splits(1).Columns(2).FooterStyle:id=80,.parent=53"
      _StyleDefs(337) =   "Splits(1).Columns(2).EditorStyle:id=81,.parent=59"
      _StyleDefs(338) =   "Splits(1).Columns(3).Style:id=90,.parent=51"
      _StyleDefs(339) =   "Splits(1).Columns(3).HeadingStyle:id=87,.parent=52"
      _StyleDefs(340) =   "Splits(1).Columns(3).FooterStyle:id=88,.parent=53"
      _StyleDefs(341) =   "Splits(1).Columns(3).EditorStyle:id=89,.parent=59"
      _StyleDefs(342) =   "Splits(1).Columns(4).Style:id=98,.parent=51"
      _StyleDefs(343) =   "Splits(1).Columns(4).HeadingStyle:id=95,.parent=52"
      _StyleDefs(344) =   "Splits(1).Columns(4).FooterStyle:id=96,.parent=53"
      _StyleDefs(345) =   "Splits(1).Columns(4).EditorStyle:id=97,.parent=59"
      _StyleDefs(346) =   "Splits(1).Columns(5).Style:id=106,.parent=51"
      _StyleDefs(347) =   "Splits(1).Columns(5).HeadingStyle:id=103,.parent=52"
      _StyleDefs(348) =   "Splits(1).Columns(5).FooterStyle:id=104,.parent=53"
      _StyleDefs(349) =   "Splits(1).Columns(5).EditorStyle:id=105,.parent=59"
      _StyleDefs(350) =   "Splits(1).Columns(6).Style:id=114,.parent=51"
      _StyleDefs(351) =   "Splits(1).Columns(6).HeadingStyle:id=111,.parent=52"
      _StyleDefs(352) =   "Splits(1).Columns(6).FooterStyle:id=112,.parent=53"
      _StyleDefs(353) =   "Splits(1).Columns(6).EditorStyle:id=113,.parent=59"
      _StyleDefs(354) =   "Splits(1).Columns(7).Style:id=122,.parent=51"
      _StyleDefs(355) =   "Splits(1).Columns(7).HeadingStyle:id=119,.parent=52"
      _StyleDefs(356) =   "Splits(1).Columns(7).FooterStyle:id=120,.parent=53"
      _StyleDefs(357) =   "Splits(1).Columns(7).EditorStyle:id=121,.parent=59"
      _StyleDefs(358) =   "Splits(1).Columns(8).Style:id=130,.parent=51"
      _StyleDefs(359) =   "Splits(1).Columns(8).HeadingStyle:id=127,.parent=52"
      _StyleDefs(360) =   "Splits(1).Columns(8).FooterStyle:id=128,.parent=53"
      _StyleDefs(361) =   "Splits(1).Columns(8).EditorStyle:id=129,.parent=59"
      _StyleDefs(362) =   "Splits(1).Columns(9).Style:id=138,.parent=51"
      _StyleDefs(363) =   "Splits(1).Columns(9).HeadingStyle:id=135,.parent=52"
      _StyleDefs(364) =   "Splits(1).Columns(9).FooterStyle:id=136,.parent=53"
      _StyleDefs(365) =   "Splits(1).Columns(9).EditorStyle:id=137,.parent=59"
      _StyleDefs(366) =   "Splits(1).Columns(10).Style:id=146,.parent=51"
      _StyleDefs(367) =   "Splits(1).Columns(10).HeadingStyle:id=143,.parent=52"
      _StyleDefs(368) =   "Splits(1).Columns(10).FooterStyle:id=144,.parent=53"
      _StyleDefs(369) =   "Splits(1).Columns(10).EditorStyle:id=145,.parent=59"
      _StyleDefs(370) =   "Splits(1).Columns(11).Style:id=154,.parent=51"
      _StyleDefs(371) =   "Splits(1).Columns(11).HeadingStyle:id=151,.parent=52"
      _StyleDefs(372) =   "Splits(1).Columns(11).FooterStyle:id=152,.parent=53"
      _StyleDefs(373) =   "Splits(1).Columns(11).EditorStyle:id=153,.parent=59"
      _StyleDefs(374) =   "Splits(1).Columns(12).Style:id=162,.parent=51"
      _StyleDefs(375) =   "Splits(1).Columns(12).HeadingStyle:id=159,.parent=52"
      _StyleDefs(376) =   "Splits(1).Columns(12).FooterStyle:id=160,.parent=53"
      _StyleDefs(377) =   "Splits(1).Columns(12).EditorStyle:id=161,.parent=59"
      _StyleDefs(378) =   "Splits(1).Columns(13).Style:id=170,.parent=51"
      _StyleDefs(379) =   "Splits(1).Columns(13).HeadingStyle:id=167,.parent=52"
      _StyleDefs(380) =   "Splits(1).Columns(13).FooterStyle:id=168,.parent=53"
      _StyleDefs(381) =   "Splits(1).Columns(13).EditorStyle:id=169,.parent=59"
      _StyleDefs(382) =   "Splits(1).Columns(14).Style:id=178,.parent=51"
      _StyleDefs(383) =   "Splits(1).Columns(14).HeadingStyle:id=175,.parent=52"
      _StyleDefs(384) =   "Splits(1).Columns(14).FooterStyle:id=176,.parent=53"
      _StyleDefs(385) =   "Splits(1).Columns(14).EditorStyle:id=177,.parent=59"
      _StyleDefs(386) =   "Splits(1).Columns(15).Style:id=186,.parent=51"
      _StyleDefs(387) =   "Splits(1).Columns(15).HeadingStyle:id=183,.parent=52"
      _StyleDefs(388) =   "Splits(1).Columns(15).FooterStyle:id=184,.parent=53"
      _StyleDefs(389) =   "Splits(1).Columns(15).EditorStyle:id=185,.parent=59"
      _StyleDefs(390) =   "Splits(1).Columns(16).Style:id=194,.parent=51"
      _StyleDefs(391) =   "Splits(1).Columns(16).HeadingStyle:id=191,.parent=52"
      _StyleDefs(392) =   "Splits(1).Columns(16).FooterStyle:id=192,.parent=53"
      _StyleDefs(393) =   "Splits(1).Columns(16).EditorStyle:id=193,.parent=59"
      _StyleDefs(394) =   "Splits(1).Columns(17).Style:id=202,.parent=51"
      _StyleDefs(395) =   "Splits(1).Columns(17).HeadingStyle:id=199,.parent=52"
      _StyleDefs(396) =   "Splits(1).Columns(17).FooterStyle:id=200,.parent=53"
      _StyleDefs(397) =   "Splits(1).Columns(17).EditorStyle:id=201,.parent=59"
      _StyleDefs(398) =   "Splits(1).Columns(18).Style:id=210,.parent=51"
      _StyleDefs(399) =   "Splits(1).Columns(18).HeadingStyle:id=207,.parent=52"
      _StyleDefs(400) =   "Splits(1).Columns(18).FooterStyle:id=208,.parent=53"
      _StyleDefs(401) =   "Splits(1).Columns(18).EditorStyle:id=209,.parent=59"
      _StyleDefs(402) =   "Splits(1).Columns(19).Style:id=218,.parent=51"
      _StyleDefs(403) =   "Splits(1).Columns(19).HeadingStyle:id=215,.parent=52"
      _StyleDefs(404) =   "Splits(1).Columns(19).FooterStyle:id=216,.parent=53"
      _StyleDefs(405) =   "Splits(1).Columns(19).EditorStyle:id=217,.parent=59"
      _StyleDefs(406) =   "Splits(1).Columns(20).Style:id=226,.parent=51"
      _StyleDefs(407) =   "Splits(1).Columns(20).HeadingStyle:id=223,.parent=52"
      _StyleDefs(408) =   "Splits(1).Columns(20).FooterStyle:id=224,.parent=53"
      _StyleDefs(409) =   "Splits(1).Columns(20).EditorStyle:id=225,.parent=59"
      _StyleDefs(410) =   "Splits(1).Columns(21).Style:id=234,.parent=51"
      _StyleDefs(411) =   "Splits(1).Columns(21).HeadingStyle:id=231,.parent=52"
      _StyleDefs(412) =   "Splits(1).Columns(21).FooterStyle:id=232,.parent=53"
      _StyleDefs(413) =   "Splits(1).Columns(21).EditorStyle:id=233,.parent=59"
      _StyleDefs(414) =   "Splits(1).Columns(22).Style:id=242,.parent=51"
      _StyleDefs(415) =   "Splits(1).Columns(22).HeadingStyle:id=239,.parent=52"
      _StyleDefs(416) =   "Splits(1).Columns(22).FooterStyle:id=240,.parent=53"
      _StyleDefs(417) =   "Splits(1).Columns(22).EditorStyle:id=241,.parent=59"
      _StyleDefs(418) =   "Splits(1).Columns(23).Style:id=250,.parent=51"
      _StyleDefs(419) =   "Splits(1).Columns(23).HeadingStyle:id=247,.parent=52"
      _StyleDefs(420) =   "Splits(1).Columns(23).FooterStyle:id=248,.parent=53"
      _StyleDefs(421) =   "Splits(1).Columns(23).EditorStyle:id=249,.parent=59"
      _StyleDefs(422) =   "Splits(1).Columns(24).Style:id=258,.parent=51"
      _StyleDefs(423) =   "Splits(1).Columns(24).HeadingStyle:id=255,.parent=52"
      _StyleDefs(424) =   "Splits(1).Columns(24).FooterStyle:id=256,.parent=53"
      _StyleDefs(425) =   "Splits(1).Columns(24).EditorStyle:id=257,.parent=59"
      _StyleDefs(426) =   "Splits(1).Columns(25).Style:id=266,.parent=51"
      _StyleDefs(427) =   "Splits(1).Columns(25).HeadingStyle:id=263,.parent=52"
      _StyleDefs(428) =   "Splits(1).Columns(25).FooterStyle:id=264,.parent=53"
      _StyleDefs(429) =   "Splits(1).Columns(25).EditorStyle:id=265,.parent=59"
      _StyleDefs(430) =   "Splits(1).Columns(26).Style:id=274,.parent=51"
      _StyleDefs(431) =   "Splits(1).Columns(26).HeadingStyle:id=271,.parent=52"
      _StyleDefs(432) =   "Splits(1).Columns(26).FooterStyle:id=272,.parent=53"
      _StyleDefs(433) =   "Splits(1).Columns(26).EditorStyle:id=273,.parent=59"
      _StyleDefs(434) =   "Splits(1).Columns(27).Style:id=282,.parent=51"
      _StyleDefs(435) =   "Splits(1).Columns(27).HeadingStyle:id=279,.parent=52"
      _StyleDefs(436) =   "Splits(1).Columns(27).FooterStyle:id=280,.parent=53"
      _StyleDefs(437) =   "Splits(1).Columns(27).EditorStyle:id=281,.parent=59"
      _StyleDefs(438) =   "Splits(1).Columns(28).Style:id=290,.parent=51"
      _StyleDefs(439) =   "Splits(1).Columns(28).HeadingStyle:id=287,.parent=52"
      _StyleDefs(440) =   "Splits(1).Columns(28).FooterStyle:id=288,.parent=53"
      _StyleDefs(441) =   "Splits(1).Columns(28).EditorStyle:id=289,.parent=59"
      _StyleDefs(442) =   "Splits(1).Columns(29).Style:id=298,.parent=51"
      _StyleDefs(443) =   "Splits(1).Columns(29).HeadingStyle:id=295,.parent=52"
      _StyleDefs(444) =   "Splits(1).Columns(29).FooterStyle:id=296,.parent=53"
      _StyleDefs(445) =   "Splits(1).Columns(29).EditorStyle:id=297,.parent=59"
      _StyleDefs(446) =   "Splits(1).Columns(30).Style:id=318,.parent=51"
      _StyleDefs(447) =   "Splits(1).Columns(30).HeadingStyle:id=315,.parent=52"
      _StyleDefs(448) =   "Splits(1).Columns(30).FooterStyle:id=316,.parent=53"
      _StyleDefs(449) =   "Splits(1).Columns(30).EditorStyle:id=317,.parent=59"
      _StyleDefs(450) =   "Splits(1).Columns(31).Style:id=326,.parent=51"
      _StyleDefs(451) =   "Splits(1).Columns(31).HeadingStyle:id=323,.parent=52"
      _StyleDefs(452) =   "Splits(1).Columns(31).FooterStyle:id=324,.parent=53"
      _StyleDefs(453) =   "Splits(1).Columns(31).EditorStyle:id=325,.parent=59"
      _StyleDefs(454) =   "Splits(1).Columns(32).Style:id=334,.parent=51"
      _StyleDefs(455) =   "Splits(1).Columns(32).HeadingStyle:id=331,.parent=52"
      _StyleDefs(456) =   "Splits(1).Columns(32).FooterStyle:id=332,.parent=53"
      _StyleDefs(457) =   "Splits(1).Columns(32).EditorStyle:id=333,.parent=59"
      _StyleDefs(458) =   "Splits(1).Columns(33).Style:id=342,.parent=51"
      _StyleDefs(459) =   "Splits(1).Columns(33).HeadingStyle:id=339,.parent=52"
      _StyleDefs(460) =   "Splits(1).Columns(33).FooterStyle:id=340,.parent=53"
      _StyleDefs(461) =   "Splits(1).Columns(33).EditorStyle:id=341,.parent=59"
      _StyleDefs(462) =   "Splits(1).Columns(34).Style:id=350,.parent=51"
      _StyleDefs(463) =   "Splits(1).Columns(34).HeadingStyle:id=347,.parent=52"
      _StyleDefs(464) =   "Splits(1).Columns(34).FooterStyle:id=348,.parent=53"
      _StyleDefs(465) =   "Splits(1).Columns(34).EditorStyle:id=349,.parent=59"
      _StyleDefs(466) =   "Splits(1).Columns(35).Style:id=358,.parent=51"
      _StyleDefs(467) =   "Splits(1).Columns(35).HeadingStyle:id=355,.parent=52"
      _StyleDefs(468) =   "Splits(1).Columns(35).FooterStyle:id=356,.parent=53"
      _StyleDefs(469) =   "Splits(1).Columns(35).EditorStyle:id=357,.parent=59"
      _StyleDefs(470) =   "Splits(1).Columns(36).Style:id=366,.parent=51"
      _StyleDefs(471) =   "Splits(1).Columns(36).HeadingStyle:id=363,.parent=52"
      _StyleDefs(472) =   "Splits(1).Columns(36).FooterStyle:id=364,.parent=53"
      _StyleDefs(473) =   "Splits(1).Columns(36).EditorStyle:id=365,.parent=59"
      _StyleDefs(474) =   "Splits(1).Columns(37).Style:id=374,.parent=51"
      _StyleDefs(475) =   "Splits(1).Columns(37).HeadingStyle:id=371,.parent=52"
      _StyleDefs(476) =   "Splits(1).Columns(37).FooterStyle:id=372,.parent=53"
      _StyleDefs(477) =   "Splits(1).Columns(37).EditorStyle:id=373,.parent=59"
      _StyleDefs(478) =   "Splits(1).Columns(38).Style:id=382,.parent=51"
      _StyleDefs(479) =   "Splits(1).Columns(38).HeadingStyle:id=379,.parent=52"
      _StyleDefs(480) =   "Splits(1).Columns(38).FooterStyle:id=380,.parent=53"
      _StyleDefs(481) =   "Splits(1).Columns(38).EditorStyle:id=381,.parent=59"
      _StyleDefs(482) =   "Splits(1).Columns(39).Style:id=390,.parent=51"
      _StyleDefs(483) =   "Splits(1).Columns(39).HeadingStyle:id=387,.parent=52"
      _StyleDefs(484) =   "Splits(1).Columns(39).FooterStyle:id=388,.parent=53"
      _StyleDefs(485) =   "Splits(1).Columns(39).EditorStyle:id=389,.parent=59"
      _StyleDefs(486) =   "Splits(1).Columns(40).Style:id=398,.parent=51"
      _StyleDefs(487) =   "Splits(1).Columns(40).HeadingStyle:id=395,.parent=52"
      _StyleDefs(488) =   "Splits(1).Columns(40).FooterStyle:id=396,.parent=53"
      _StyleDefs(489) =   "Splits(1).Columns(40).EditorStyle:id=397,.parent=59"
      _StyleDefs(490) =   "Splits(1).Columns(41).Style:id=406,.parent=51"
      _StyleDefs(491) =   "Splits(1).Columns(41).HeadingStyle:id=403,.parent=52"
      _StyleDefs(492) =   "Splits(1).Columns(41).FooterStyle:id=404,.parent=53"
      _StyleDefs(493) =   "Splits(1).Columns(41).EditorStyle:id=405,.parent=59"
      _StyleDefs(494) =   "Splits(1).Columns(42).Style:id=414,.parent=51"
      _StyleDefs(495) =   "Splits(1).Columns(42).HeadingStyle:id=411,.parent=52"
      _StyleDefs(496) =   "Splits(1).Columns(42).FooterStyle:id=412,.parent=53"
      _StyleDefs(497) =   "Splits(1).Columns(42).EditorStyle:id=413,.parent=59"
      _StyleDefs(498) =   "Splits(1).Columns(43).Style:id=422,.parent=51"
      _StyleDefs(499) =   "Splits(1).Columns(43).HeadingStyle:id=419,.parent=52"
      _StyleDefs(500) =   "Splits(1).Columns(43).FooterStyle:id=420,.parent=53"
      _StyleDefs(501) =   "Splits(1).Columns(43).EditorStyle:id=421,.parent=59"
      _StyleDefs(502) =   "Splits(1).Columns(44).Style:id=430,.parent=51"
      _StyleDefs(503) =   "Splits(1).Columns(44).HeadingStyle:id=427,.parent=52"
      _StyleDefs(504) =   "Splits(1).Columns(44).FooterStyle:id=428,.parent=53"
      _StyleDefs(505) =   "Splits(1).Columns(44).EditorStyle:id=429,.parent=59"
      _StyleDefs(506) =   "Splits(1).Columns(45).Style:id=438,.parent=51"
      _StyleDefs(507) =   "Splits(1).Columns(45).HeadingStyle:id=435,.parent=52"
      _StyleDefs(508) =   "Splits(1).Columns(45).FooterStyle:id=436,.parent=53"
      _StyleDefs(509) =   "Splits(1).Columns(45).EditorStyle:id=437,.parent=59"
      _StyleDefs(510) =   "Splits(1).Columns(46).Style:id=446,.parent=51"
      _StyleDefs(511) =   "Splits(1).Columns(46).HeadingStyle:id=443,.parent=52"
      _StyleDefs(512) =   "Splits(1).Columns(46).FooterStyle:id=444,.parent=53"
      _StyleDefs(513) =   "Splits(1).Columns(46).EditorStyle:id=445,.parent=59"
      _StyleDefs(514) =   "Splits(1).Columns(47).Style:id=454,.parent=51"
      _StyleDefs(515) =   "Splits(1).Columns(47).HeadingStyle:id=451,.parent=52"
      _StyleDefs(516) =   "Splits(1).Columns(47).FooterStyle:id=452,.parent=53"
      _StyleDefs(517) =   "Splits(1).Columns(47).EditorStyle:id=453,.parent=59"
      _StyleDefs(518) =   "Splits(1).Columns(48).Style:id=462,.parent=51"
      _StyleDefs(519) =   "Splits(1).Columns(48).HeadingStyle:id=459,.parent=52"
      _StyleDefs(520) =   "Splits(1).Columns(48).FooterStyle:id=460,.parent=53"
      _StyleDefs(521) =   "Splits(1).Columns(48).EditorStyle:id=461,.parent=59"
      _StyleDefs(522) =   "Splits(1).Columns(49).Style:id=470,.parent=51"
      _StyleDefs(523) =   "Splits(1).Columns(49).HeadingStyle:id=467,.parent=52"
      _StyleDefs(524) =   "Splits(1).Columns(49).FooterStyle:id=468,.parent=53"
      _StyleDefs(525) =   "Splits(1).Columns(49).EditorStyle:id=469,.parent=59"
      _StyleDefs(526) =   "Splits(1).Columns(50).Style:id=478,.parent=51"
      _StyleDefs(527) =   "Splits(1).Columns(50).HeadingStyle:id=475,.parent=52"
      _StyleDefs(528) =   "Splits(1).Columns(50).FooterStyle:id=476,.parent=53"
      _StyleDefs(529) =   "Splits(1).Columns(50).EditorStyle:id=477,.parent=59"
      _StyleDefs(530) =   "Splits(1).Columns(51).Style:id=486,.parent=51"
      _StyleDefs(531) =   "Splits(1).Columns(51).HeadingStyle:id=483,.parent=52"
      _StyleDefs(532) =   "Splits(1).Columns(51).FooterStyle:id=484,.parent=53"
      _StyleDefs(533) =   "Splits(1).Columns(51).EditorStyle:id=485,.parent=59"
      _StyleDefs(534) =   "Splits(1).Columns(52).Style:id=494,.parent=51"
      _StyleDefs(535) =   "Splits(1).Columns(52).HeadingStyle:id=491,.parent=52"
      _StyleDefs(536) =   "Splits(1).Columns(52).FooterStyle:id=492,.parent=53"
      _StyleDefs(537) =   "Splits(1).Columns(52).EditorStyle:id=493,.parent=59"
      _StyleDefs(538) =   "Splits(1).Columns(53).Style:id=502,.parent=51"
      _StyleDefs(539) =   "Splits(1).Columns(53).HeadingStyle:id=499,.parent=52"
      _StyleDefs(540) =   "Splits(1).Columns(53).FooterStyle:id=500,.parent=53"
      _StyleDefs(541) =   "Splits(1).Columns(53).EditorStyle:id=501,.parent=59"
      _StyleDefs(542) =   "Splits(1).Columns(54).Style:id=510,.parent=51"
      _StyleDefs(543) =   "Splits(1).Columns(54).HeadingStyle:id=507,.parent=52"
      _StyleDefs(544) =   "Splits(1).Columns(54).FooterStyle:id=508,.parent=53"
      _StyleDefs(545) =   "Splits(1).Columns(54).EditorStyle:id=509,.parent=59"
      _StyleDefs(546) =   "Splits(1).Columns(55).Style:id=518,.parent=51"
      _StyleDefs(547) =   "Splits(1).Columns(55).HeadingStyle:id=515,.parent=52"
      _StyleDefs(548) =   "Splits(1).Columns(55).FooterStyle:id=516,.parent=53"
      _StyleDefs(549) =   "Splits(1).Columns(55).EditorStyle:id=517,.parent=59"
      _StyleDefs(550) =   "Splits(1).Columns(56).Style:id=526,.parent=51"
      _StyleDefs(551) =   "Splits(1).Columns(56).HeadingStyle:id=523,.parent=52"
      _StyleDefs(552) =   "Splits(1).Columns(56).FooterStyle:id=524,.parent=53"
      _StyleDefs(553) =   "Splits(1).Columns(56).EditorStyle:id=525,.parent=59"
      _StyleDefs(554) =   "Splits(1).Columns(57).Style:id=534,.parent=51"
      _StyleDefs(555) =   "Splits(1).Columns(57).HeadingStyle:id=531,.parent=52"
      _StyleDefs(556) =   "Splits(1).Columns(57).FooterStyle:id=532,.parent=53"
      _StyleDefs(557) =   "Splits(1).Columns(57).EditorStyle:id=533,.parent=59"
      _StyleDefs(558) =   "Splits(1).Columns(58).Style:id=542,.parent=51"
      _StyleDefs(559) =   "Splits(1).Columns(58).HeadingStyle:id=539,.parent=52"
      _StyleDefs(560) =   "Splits(1).Columns(58).FooterStyle:id=540,.parent=53"
      _StyleDefs(561) =   "Splits(1).Columns(58).EditorStyle:id=541,.parent=59"
      _StyleDefs(562) =   "Splits(1).Columns(59).Style:id=550,.parent=51"
      _StyleDefs(563) =   "Splits(1).Columns(59).HeadingStyle:id=547,.parent=52"
      _StyleDefs(564) =   "Splits(1).Columns(59).FooterStyle:id=548,.parent=53"
      _StyleDefs(565) =   "Splits(1).Columns(59).EditorStyle:id=549,.parent=59"
      _StyleDefs(566) =   "Splits(1).Columns(60).Style:id=558,.parent=51"
      _StyleDefs(567) =   "Splits(1).Columns(60).HeadingStyle:id=555,.parent=52"
      _StyleDefs(568) =   "Splits(1).Columns(60).FooterStyle:id=556,.parent=53"
      _StyleDefs(569) =   "Splits(1).Columns(60).EditorStyle:id=557,.parent=59"
      _StyleDefs(570) =   "Splits(1).Columns(61).Style:id=566,.parent=51"
      _StyleDefs(571) =   "Splits(1).Columns(61).HeadingStyle:id=563,.parent=52"
      _StyleDefs(572) =   "Splits(1).Columns(61).FooterStyle:id=564,.parent=53"
      _StyleDefs(573) =   "Splits(1).Columns(61).EditorStyle:id=565,.parent=59"
      _StyleDefs(574) =   "Splits(1).Columns(62).Style:id=574,.parent=51"
      _StyleDefs(575) =   "Splits(1).Columns(62).HeadingStyle:id=571,.parent=52"
      _StyleDefs(576) =   "Splits(1).Columns(62).FooterStyle:id=572,.parent=53"
      _StyleDefs(577) =   "Splits(1).Columns(62).EditorStyle:id=573,.parent=59"
      _StyleDefs(578) =   "Splits(1).Columns(63).Style:id=582,.parent=51"
      _StyleDefs(579) =   "Splits(1).Columns(63).HeadingStyle:id=579,.parent=52"
      _StyleDefs(580) =   "Splits(1).Columns(63).FooterStyle:id=580,.parent=53"
      _StyleDefs(581) =   "Splits(1).Columns(63).EditorStyle:id=581,.parent=59"
      _StyleDefs(582) =   "Splits(1).Columns(64).Style:id=590,.parent=51"
      _StyleDefs(583) =   "Splits(1).Columns(64).HeadingStyle:id=587,.parent=52"
      _StyleDefs(584) =   "Splits(1).Columns(64).FooterStyle:id=588,.parent=53"
      _StyleDefs(585) =   "Splits(1).Columns(64).EditorStyle:id=589,.parent=59"
      _StyleDefs(586) =   "Splits(1).Columns(65).Style:id=598,.parent=51"
      _StyleDefs(587) =   "Splits(1).Columns(65).HeadingStyle:id=595,.parent=52"
      _StyleDefs(588) =   "Splits(1).Columns(65).FooterStyle:id=596,.parent=53"
      _StyleDefs(589) =   "Splits(1).Columns(65).EditorStyle:id=597,.parent=59"
      _StyleDefs(590) =   "Splits(1).Columns(66).Style:id=606,.parent=51"
      _StyleDefs(591) =   "Splits(1).Columns(66).HeadingStyle:id=603,.parent=52"
      _StyleDefs(592) =   "Splits(1).Columns(66).FooterStyle:id=604,.parent=53"
      _StyleDefs(593) =   "Splits(1).Columns(66).EditorStyle:id=605,.parent=59"
      _StyleDefs(594) =   "Splits(1).Columns(67).Style:id=614,.parent=51"
      _StyleDefs(595) =   "Splits(1).Columns(67).HeadingStyle:id=611,.parent=52"
      _StyleDefs(596) =   "Splits(1).Columns(67).FooterStyle:id=612,.parent=53"
      _StyleDefs(597) =   "Splits(1).Columns(67).EditorStyle:id=613,.parent=59"
      _StyleDefs(598) =   "Splits(1).Columns(68).Style:id=46,.parent=51"
      _StyleDefs(599) =   "Splits(1).Columns(68).HeadingStyle:id=43,.parent=52"
      _StyleDefs(600) =   "Splits(1).Columns(68).FooterStyle:id=44,.parent=53"
      _StyleDefs(601) =   "Splits(1).Columns(68).EditorStyle:id=45,.parent=59"
      _StyleDefs(602) =   "Named:id=33:Normal"
      _StyleDefs(603) =   ":id=33,.parent=0"
      _StyleDefs(604) =   "Named:id=34:Heading"
      _StyleDefs(605) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(606) =   ":id=34,.wraptext=-1"
      _StyleDefs(607) =   "Named:id=35:Footing"
      _StyleDefs(608) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(609) =   "Named:id=36:Selected"
      _StyleDefs(610) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(611) =   "Named:id=37:Caption"
      _StyleDefs(612) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(613) =   "Named:id=38:HighlightRow"
      _StyleDefs(614) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(615) =   "Named:id=39:EvenRow"
      _StyleDefs(616) =   ":id=39,.parent=33,.bgcolor=&H6DB9DA&"
      _StyleDefs(617) =   "Named:id=40:OddRow"
      _StyleDefs(618) =   ":id=40,.parent=33,.bgcolor=&H82C4DF&"
      _StyleDefs(619) =   "Named:id=41:RecordSelector"
      _StyleDefs(620) =   ":id=41,.parent=34,.wraptext=-1"
      _StyleDefs(621) =   "Named:id=42:FilterBar"
      _StyleDefs(622) =   ":id=42,.parent=33"
   End
   Begin VB.Frame Frame1 
      Height          =   525
      Left            =   405
      TabIndex        =   2
      Top             =   6990
      Width           =   7335
      Begin VB.Label Label4 
         BackColor       =   &H00008000&
         Caption         =   "-------------------"
         Height          =   270
         Left            =   3420
         TabIndex        =   6
         Top             =   165
         Width           =   870
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Posisi Waktu Aktif"
         Height          =   300
         Left            =   4425
         TabIndex        =   5
         Top             =   225
         Width           =   2190
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Lama Aktivitas Mekanik"
         Height          =   300
         Left            =   1140
         TabIndex        =   4
         Top             =   210
         Width           =   2190
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         Caption         =   "Label1"
         Height          =   270
         Left            =   135
         TabIndex        =   3
         Top             =   165
         Width           =   870
      End
   End
   Begin AIFCmp1.asxPowerBanner asxPowerBanner1 
      Height          =   7560
      Left            =   0
      Top             =   0
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   13335
      StartColor      =   4210752
      EndColor        =   12632064
      FormatString    =   "Display Activity Mekanik"
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
Attribute VB_Name = "frmProgressGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PosSkr As Integer
Dim jam As New XArrayDB
Dim MenitNya As Integer
Dim menit As Integer
Dim JamNya As Integer
Dim MaxBaris As Integer
Const MaxKolom = 68
Const NilaiAwal = 2
Const jamAwal = 7

Private Sub Waktu(JmlKolom As Integer)
    Dim i As Integer
    TData1.Splits(0).Columns(0).Caption = "NAMA MEKANIK"
    TData1.Splits(0).Columns(0).HeadAlignment = dbgCenter
    TData1.Splits(0).Columns(0).Alignment = dbgRight
    
    For i = NilaiAwal To JmlKolom
        Select Case i
            Case NilaiAwal
                Judul = "7"
            Case NilaiAwal + 6
                Judul = "8"
            Case NilaiAwal + 12
                Judul = "9"
            Case NilaiAwal + 18
                Judul = "10"
            Case NilaiAwal + 24
                Judul = "11"
            Case NilaiAwal + 30
                Judul = "12"
            Case NilaiAwal + 36
                Judul = "13"
            Case NilaiAwal + 42
                Judul = "14"
            Case NilaiAwal + 48
                Judul = "15"
            Case NilaiAwal + 54
                Judul = "16"
            Case NilaiAwal + 60
                Judul = "17"
            Case NilaiAwal + 66
                Judul = "18"
            Case Else
                Judul = "."
        End Select
        
        TData1.Splits(1).Columns(i).Caption = Judul
        TData1.Splits(1).Columns(i).Width = 250
        TData1.Splits(1).Columns(i).Alignment = dbgCenter
    Next
End Sub


Private Sub Form_Load()
    Dim i As Integer
    Call SettingGrid
    
    jam.Clear
    'Lebar Kolom dan barisnya
    sSQL = "select count(*) as JML_BRS from MEKANIK"
    oRs.Open sSQL, oConn, adOpenStatic
    If Not oRs.BOF Then
        MaxBaris = oRs("JML_BRS") + 1
    End If
    oRs.Close
    MaxBaris = MaxBaris * 2
    jam.ReDim 0, MaxBaris, 0, MaxKolom
    PosSkr = PosisiJam(MaxBaris)
'    PosSkr = PosisiJam(MaxKolom)
    i = 0
    sSQL = "select NAME from MEKANIK order by [NAME]"
    oRs.Open sSQL, oConn, adOpenStatic
    If Not oRs.BOF Then
        Do While Not oRs.EOF
            If i Mod 2 = 0 Then
                jam(i, 0) = Trim(oRs("NAME"))
                jam(i, 1) = "F"
                oRs.MoveNext
            Else
                jam(i, 0) = ""
                jam(i, 1) = "A"
            End If
            i = i + 1
        Loop
        jam(i, 0) = ""
        jam(i, 1) = "A"
    End If
    oRs.Close
    
    Call Waktu(MaxKolom)
    Call PosisiMekanik(PosSkr, MaxBaris)
    
    TData1.Array = jam
    TData1.ReBind
End Sub
            
'Call BuatGaris(JamAw, MenitAw, JamAk, MenitAk, Posisi, PosisiSekarang)
            
Private Sub BuatGaris(ByVal JamAw As Integer, ByVal MnAw As Integer, ByVal JamAk As Integer, ByVal MnAk As Integer, ByVal Pos As Integer, ByVal PosWaktu As Integer, Sls As Boolean)
    On Error GoTo Err
    
    Dim i As Integer
    Dim PosisiKotak As Integer
    Dim PanjangKotak As Integer
    Dim TempPanjang As Integer
    Dim MENIT1_TOTAL As Integer
    Dim MENIT2_TOTAL As Integer
    
    If JamAw >= 7 Or JamAw <= 19 And JamAw >= jamAwal Then
        PosisiKotak = ((JamAw - jamAwal) * 6) + Round(MnAw / 10) + 1
'        PosisiKotak = ((JamAw - JamAwal) * 6) + Int(MnAw / 10) + 1
        
        If JamAk < 7 Or JamAk >= 19 Then
            PanjangKotak = PosWaktu - PosisiKotak  '- 1 '- NilaiAwal '- JamAwal
        Else
            MENIT1_TOTAL = JamAk * 60 + MnAk
            MENIT2_TOTAL = JamAw * 60 + MnAw
            
            TempPanjang = Int(((MENIT1_TOTAL - MENIT2_TOTAL) / 10))
            PanjangKotak = TempPanjang
        End If
        
        If Sls = False Then
            PanjangKotak = PosSkr - PosisiKotak - 1
        End If
        
        For i = 1 To PanjangKotak
            If (PosisiKotak + i) > PosSkr Then
                Exit For
            End If
            jam(Pos, PosisiKotak + i) = "#"
        Next
    End If
    
    Exit Sub
Err:
    MsgBox "Garis Rate terlalu panjang !", vbOKOnly, "Error"
End Sub

Private Sub BuatGarisFlate(ByVal JamAw As Integer, ByVal MenitAw As Integer, ByVal Panjang As Integer, ByVal Pos As Integer, ByVal PosWaktu As Integer)
    On Error GoTo Err
    
    Dim i As Integer
    Dim PosisiKotak As Integer
    Dim PanjangKotak As Integer
    Dim TempPanjang As Integer
    
    If JamAw >= 7 Or JamAw <= 19 And JamAw >= jamAwal Then
        PosisiKotak = ((JamAw - jamAwal) * 6) + Round(MenitAw / 10) + 1

        PanjangKotak = Panjang
            
        For i = 1 To PanjangKotak
            jam(Pos, PosisiKotak + i) = "-"
        Next
    End If
    
    Exit Sub
Err:
    MsgBox "Garis Flate Rate terlalu panjang !", vbOKOnly, "Error"
End Sub

Private Sub PosisiMekanik(PosisiSekarang As Integer, MaxBaris As Integer)
    Dim JamAk As Integer
    Dim MenitAk As Integer
    Dim JamAw As Integer
    Dim MenitAw As Integer
    Dim Posisi As Integer
    Dim Keterangan As String
    Dim Panjang As Integer
    Dim Selesai As Boolean
    
    sSQL = "select RTE1+RTE2+RTE3+RTE4+RTE5+RTE6+RTE7+RTE8 as RATE, " _
         & " TIME_M, TIME_S, [NAME], NO_WO, NO_POLISI from jobdta as A, MEKANIK as B " _
         & " where A.deliv<>'BTL' and A.[DATE]='" & TglYYYYMMDD(gTanggal) & "' and A.MECH=B.MECH " _
         & " and TIME_M <> '' order by TIME_M"
    
    oRs.Open sSQL, oConn, adOpenStatic
    If Not oRs.BOF Then
        Do While Not oRs.EOF

            For i = 0 To MaxBaris
                If Trim(jam(i, 0)) = Trim(oRs("NAME")) Then
                    Posisi = i + 1
                    Exit For
                End If
            Next i
            JamAw = Val(Left(Trim(oRs("TIME_M")), 2))
            MenitAw = Val(Right(Trim(oRs("TIME_M")), 2))
            
            JamAk = Val(Left(Trim(oRs("TIME_S")), 2))
            MenitAk = Val(Right(Trim(oRs("TIME_S")), 2))
            Keterangan = ""
            
            If JamAk = 0 Then
                Selesai = False
            Else
                Selesai = True
            End If
            If JamAk = 0 Or JamAk > JamNya Or (JamAk = JamNya And MenitAk > menit) Then
'                JamAk = Val(Left(Trim(oRs("TIME_S")), 2))
'                MenitAk = Val(Left(Trim(oRs("TIME_S")), 2))
                JamAk = JamNya
                MenitAk = menit
                Keterangan = Trim(oRs("NO_WO")) + " - " + Trim(oRs("NO_POLISI"))
            End If
            
            Panjang = oRs("RATE") * 6
            jam(Posisi, 0) = Keterangan
            Call BuatGaris(JamAw, MenitAw, JamAk, MenitAk, Posisi, PosisiSekarang, Selesai)
            Call BuatGarisFlate(JamAw, MenitAw, Panjang, Posisi - 1, PosisiSekarang)
            oRs.MoveNext
        Loop
    End If
    oRs.Close
End Sub

Private Function PosisiJam(MaxBaris As Integer) As Integer
    Dim i As Integer
    Dim TempMenit As Double
    
    JamNya = Format(Time, "HH")
    MenitNya = Int(Format(Time, "NN") / 10)
    menit = Format(Time, "NN")
    PosisiJam = NilaiAwal + ((JamNya - jamAwal) * 6) + MenitNya
    
    If JamNya < 0 Then
        PosisiJam = NilaiAwal
    End If
    If (JamNya >= 19 And MenitNya > 0) Or JamNya > 19 Then
        PosisiJam = MaxKolom
    End If
    
    If PosisiJam < 0 Then
        PosisiJam = NilaiAwal
    End If
    
    For i = 0 To MaxBaris
        jam(i, PosisiJam) = "|"
    Next
End Function

Private Sub SettingGrid()
    'Pengaturan Grid
    Dim i As Integer
    Dim NoItem As New TrueDBGrid70.Style
    Dim ProgressItem As New TrueDBGrid70.Style
    Dim ProgressFlate As New TrueDBGrid70.Style
    
    NoItem.BackColor = &H8000&
    ProgressItem.BackColor = vbBlack
    TData1.AddRegexCellStyle dbgAllCells, NoItem, "^|"
    
    TData1.AddRegexCellStyle dbgAllCells, ProgressItem, "^#"
    
    ProgressFlate.BackColor = vbRed
    TData1.AddRegexCellStyle dbgAllCells, ProgressFlate, "^-"
    
    TData1.EvenRowStyle.BackColor = &HC0FFC0   ' &H80FFFF
    TData1.OddRowStyle.BackColor = &HFFC0C0  ' &HC0C000     ' &HC0FFFF
    TData1.AlternatingRowStyle = True
    MaxBaris = 0
    For i = 0 To MaxKolom
        TData1.Splits(0).Columns(i).AllowSizing = False
        TData1.Splits(1).Columns(i).AllowSizing = False
    Next
    For i = NilaiAwal To MaxKolom
        TData1.Splits(0).Columns(i).Visible = False
    Next i
End Sub

Private Sub TblKeluar_Click()
    Unload Me
End Sub
