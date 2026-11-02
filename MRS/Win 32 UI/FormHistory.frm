VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FormHistory 
   BackColor       =   &H00BFE8EA&
   BorderStyle     =   0  'None
   Caption         =   "FormHistory"
   ClientHeight    =   4410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11250
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TrueOleDBGrid70.TDBGrid Grid 
      Height          =   2940
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   5186
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Tanggal"
      Columns(0).DataField=   "Tanggal"
      Columns(0).NumberFormat=   "Short Date"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "No Invoice"
      Columns(1).DataField=   "NoINV"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Total Invoice"
      Columns(2).DataField=   "TotInv"
      Columns(2).NumberFormat=   "Currency"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "No WO"
      Columns(3).DataField=   "NoWO"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "KM"
      Columns(4).DataField=   "KM"
      Columns(4).NumberFormat=   "General Number"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "SA"
      Columns(5).DataField=   "SA"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "MECH"
      Columns(6).DataField=   "MECH"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Job Order"
      Columns(7).DataField=   "JobOrder"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Pesan Khusus"
      Columns(8).DataField=   "Pesan"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Saran"
      Columns(9).DataField=   "Saran"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Catatan"
      Columns(10).DataField=   "Catat"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "Permintaan"
      Columns(11).DataField=   "Minta"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "Keluhan"
      Columns(12).DataField=   "Keluh"
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "Dealer Service"
      Columns(13).DataField=   "Dealer"
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "Hasil FI"
      Columns(14).DataField=   "HasilFI"
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "Hasil Test Drive"
      Columns(15).DataField=   "HasilTD"
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "Next Job ( CR 7 )"
      Columns(16).DataField=   "NextJob"
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "Komplain User"
      Columns(17).DataField=   "Komplain"
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "FIR"
      Columns(18).DataField=   "fir"
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   19
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=19"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2461"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2355"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2461"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2355"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=2646"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2540"
      Splits(0)._ColumnProps(12)=   "Column(2)._ColStyle=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(14)=   "Column(3).Width=2302"
      Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=2196"
      Splits(0)._ColumnProps(17)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(18)=   "Column(4).Width=2302"
      Splits(0)._ColumnProps(19)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(4)._WidthInPix=2196"
      Splits(0)._ColumnProps(21)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(22)=   "Column(5).Width=1244"
      Splits(0)._ColumnProps(23)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(5)._WidthInPix=1138"
      Splits(0)._ColumnProps(25)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(26)=   "Column(6).Width=1244"
      Splits(0)._ColumnProps(27)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(6)._WidthInPix=1138"
      Splits(0)._ColumnProps(29)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(30)=   "Column(7).Width=6165"
      Splits(0)._ColumnProps(31)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(32)=   "Column(7)._WidthInPix=6059"
      Splits(0)._ColumnProps(33)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(34)=   "Column(8).Width=4419"
      Splits(0)._ColumnProps(35)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(36)=   "Column(8)._WidthInPix=4313"
      Splits(0)._ColumnProps(37)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(38)=   "Column(9).Width=4419"
      Splits(0)._ColumnProps(39)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(9)._WidthInPix=4313"
      Splits(0)._ColumnProps(41)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(42)=   "Column(10).Width=4419"
      Splits(0)._ColumnProps(43)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(44)=   "Column(10)._WidthInPix=4313"
      Splits(0)._ColumnProps(45)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(46)=   "Column(11).Width=4419"
      Splits(0)._ColumnProps(47)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(48)=   "Column(11)._WidthInPix=4313"
      Splits(0)._ColumnProps(49)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(50)=   "Column(12).Width=4419"
      Splits(0)._ColumnProps(51)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(52)=   "Column(12)._WidthInPix=4313"
      Splits(0)._ColumnProps(53)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(54)=   "Column(13).Width=4419"
      Splits(0)._ColumnProps(55)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(56)=   "Column(13)._WidthInPix=4313"
      Splits(0)._ColumnProps(57)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(58)=   "Column(14).Width=3281"
      Splits(0)._ColumnProps(59)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(60)=   "Column(14)._WidthInPix=3175"
      Splits(0)._ColumnProps(61)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(62)=   "Column(15).Width=3281"
      Splits(0)._ColumnProps(63)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(64)=   "Column(15)._WidthInPix=3175"
      Splits(0)._ColumnProps(65)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(66)=   "Column(16).Width=3281"
      Splits(0)._ColumnProps(67)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(68)=   "Column(16)._WidthInPix=3175"
      Splits(0)._ColumnProps(69)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(70)=   "Column(17).Width=3281"
      Splits(0)._ColumnProps(71)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(72)=   "Column(17)._WidthInPix=3175"
      Splits(0)._ColumnProps(73)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(74)=   "Column(18).Width=3281"
      Splits(0)._ColumnProps(75)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(76)=   "Column(18)._WidthInPix=3175"
      Splits(0)._ColumnProps(77)=   "Column(18).Order=19"
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
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HC1EDFF&"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=90,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=87,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=88,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=89,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=86,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=83,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=84,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=85,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=82,.parent=13,.alignment=1"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=79,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=80,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=81,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=78,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=75,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=76,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=77,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=74,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=71,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=72,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=73,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=70,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=66,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=63,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=64,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=65,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=62,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=58,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=55,.parent=14"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=56,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=57,.parent=17"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=54,.parent=13"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=51,.parent=14"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=52,.parent=15"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=53,.parent=17"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=50,.parent=13"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=47,.parent=14"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=48,.parent=15"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=49,.parent=17"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=46,.parent=13"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=43,.parent=14"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=44,.parent=15"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=45,.parent=17"
      _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=28,.parent=13"
      _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=25,.parent=14"
      _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=26,.parent=15"
      _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=27,.parent=17"
      _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=32,.parent=13"
      _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=29,.parent=14"
      _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=30,.parent=15"
      _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=31,.parent=17"
      _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=94,.parent=13"
      _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=91,.parent=14"
      _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=92,.parent=15"
      _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=93,.parent=17"
      _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=98,.parent=13"
      _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=95,.parent=14"
      _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=96,.parent=15"
      _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=97,.parent=17"
      _StyleDefs(100) =   "Splits(0).Columns(16).Style:id=102,.parent=13"
      _StyleDefs(101) =   "Splits(0).Columns(16).HeadingStyle:id=99,.parent=14"
      _StyleDefs(102) =   "Splits(0).Columns(16).FooterStyle:id=100,.parent=15"
      _StyleDefs(103) =   "Splits(0).Columns(16).EditorStyle:id=101,.parent=17"
      _StyleDefs(104) =   "Splits(0).Columns(17).Style:id=106,.parent=13"
      _StyleDefs(105) =   "Splits(0).Columns(17).HeadingStyle:id=103,.parent=14"
      _StyleDefs(106) =   "Splits(0).Columns(17).FooterStyle:id=104,.parent=15"
      _StyleDefs(107) =   "Splits(0).Columns(17).EditorStyle:id=105,.parent=17"
      _StyleDefs(108) =   "Splits(0).Columns(18).Style:id=110,.parent=13"
      _StyleDefs(109) =   "Splits(0).Columns(18).HeadingStyle:id=107,.parent=14"
      _StyleDefs(110) =   "Splits(0).Columns(18).FooterStyle:id=108,.parent=15"
      _StyleDefs(111) =   "Splits(0).Columns(18).EditorStyle:id=109,.parent=17"
      _StyleDefs(112) =   "Named:id=33:Normal"
      _StyleDefs(113) =   ":id=33,.parent=0"
      _StyleDefs(114) =   "Named:id=34:Heading"
      _StyleDefs(115) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(116) =   ":id=34,.wraptext=-1"
      _StyleDefs(117) =   "Named:id=35:Footing"
      _StyleDefs(118) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(119) =   "Named:id=36:Selected"
      _StyleDefs(120) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(121) =   "Named:id=37:Caption"
      _StyleDefs(122) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(123) =   "Named:id=38:HighlightRow"
      _StyleDefs(124) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(125) =   "Named:id=39:EvenRow"
      _StyleDefs(126) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(127) =   "Named:id=40:OddRow"
      _StyleDefs(128) =   ":id=40,.parent=33"
      _StyleDefs(129) =   "Named:id=41:RecordSelector"
      _StyleDefs(130) =   ":id=41,.parent=34"
      _StyleDefs(131) =   "Named:id=42:FilterBar"
      _StyleDefs(132) =   ":id=42,.parent=33"
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
      Left            =   10200
      Picture         =   "FormHistory.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   735
   End
   Begin TrueOleDBGrid70.TDBGrid GridHist 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   11235
      _ExtentX        =   19817
      _ExtentY        =   5106
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
      Columns(2).DataField=   "front"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Mech"
      Columns(3).DataField=   "Mech"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Nama Dealer"
      Columns(4).DataField=   "FS_KD_DEALER"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Tipe"
      Columns(5).DataField=   "tipe"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Service"
      Columns(6).DataField=   "Job"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Pesan"
      Columns(7).DataField=   "pesan"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Saran"
      Columns(8).DataField=   "saran"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Catatan"
      Columns(9).DataField=   "catatan"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   10
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=10"
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
      Splits(0)._ColumnProps(25)=   "Column(6).Width=8811"
      Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=8705"
      Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(29)=   "Column(7).Width=3519"
      Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=3413"
      Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(33)=   "Column(8).Width=3519"
      Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=3413"
      Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(37)=   "Column(9).Width=3519"
      Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=3413"
      Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
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
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
      _StyleDefs(76)  =   "Named:id=33:Normal"
      _StyleDefs(77)  =   ":id=33,.parent=0"
      _StyleDefs(78)  =   "Named:id=34:Heading"
      _StyleDefs(79)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(80)  =   ":id=34,.wraptext=-1"
      _StyleDefs(81)  =   "Named:id=35:Footing"
      _StyleDefs(82)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(83)  =   "Named:id=36:Selected"
      _StyleDefs(84)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(85)  =   "Named:id=37:Caption"
      _StyleDefs(86)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(87)  =   "Named:id=38:HighlightRow"
      _StyleDefs(88)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(89)  =   "Named:id=39:EvenRow"
      _StyleDefs(90)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(91)  =   "Named:id=40:OddRow"
      _StyleDefs(92)  =   ":id=40,.parent=33"
      _StyleDefs(93)  =   "Named:id=41:RecordSelector"
      _StyleDefs(94)  =   ":id=41,.parent=34"
      _StyleDefs(95)  =   "Named:id=42:FilterBar"
      _StyleDefs(96)  =   ":id=42,.parent=33"
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   11160
      Y1              =   3300
      Y2              =   3300
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   10080
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Width           =   975
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
      TabIndex        =   1
      Top             =   45
      Width           =   7095
   End
   Begin VB.Line Line21 
      BorderColor     =   &H000000FF&
      X1              =   1440
      X2              =   3960
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line22 
      BorderColor     =   &H000000FF&
      X1              =   2280
      X2              =   3960
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   375
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   11235
   End
End
Attribute VB_Name = "FormHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================
'Motion Captured Data Provided By Acep Suryanto
'==============================================
Option Explicit



Private xMe As Boolean

Private noRangka As String
Private objCu As New ClsEstimasi
Private rsCu As ADODB.Recordset
Private NoSPK As String
Private isHO As Boolean

Public Property Let isMe(value As Boolean)
   xMe = value
End Property

Public Property Let isHOService(value As Boolean)
    isHO = value
End Property

Public Property Let NomorRangka(value As String)
   noRangka = value
End Property

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
          Dim strSql As String
          Dim keteranganHide As String
          
          ' GENERATE HISTORY
1         If isHO Then
2             grid(0).Visible = False
3             GridHist.Visible = True
              'jika history dari ho
4             Label24.Caption = "HISTORY SERVICE ONLINE HO"
5             Set rsCu = objCu.panggilHistoryHO(noRangka)
6             If rsCu.RecordCount = 0 Then Exit Sub
7             Set Me.GridHist.DataSource = rsCu
8         Else
9             grid(0).Visible = True
10            GridHist.Visible = False
11            Label24.Caption = "HISTORY SERVICE"
              'ganti panggil history
              'Call objCu.panggilHistory(noRangka)
              
              '' 20180515REV02'INSPECTION
              
      '        strSql = "select isnull(komplain,'') komplain,IsNull(J.CATAT,'') as CATAT, D.DATE as Tanggal,D.NOINV,D.TOTAL as TotInv, D.NO_WO as NOWO, D.KM,D.FRONT as SA ,D.MECH ," _
      '             & " REPLACE(REPLACE(REPLACE(D.REP1 + '; ' + D.REP2 + '; ' + D.REP3 + '; ' + D.REP4 + '; ' + D.REP5 + '; ' + D.REP6 + '; ' " _
      '             & " + D.REP7,'; ;',''),'; ;',''),'; ;','') as JobOrder, D.REMARK1 as Pesan, D.SUGGEST as Saran, " _
      '             & " D.REQ1 + '; ' + D.REQ2 + '; ' + D.REQ3 + '; ' + D.REQ4 + '; ' + D.REQ5 as Minta,D.FS_NM_DEALER as Dealer, " _
      '             & " REPLACE(REPLACE(REPLACE(D.CMPL1 + '; ' + D.CMPL2 + '; ' + D.CMPL3 + '; ' + D.CMPL4 + '; ' + D.CMPL5 ,'; ;',''),'; ;',''),'; ;','') as Keluh,isnull(i.HasilFI,'') HasilFI,isnull(i.HasilTestDrive,'') as HasilTD, " _
      '             & " isnull(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(i.SNJ1 + '; ' + i.SNJ2 + '; ' + i.SNJ3 + '; ' + i.SNJ4 + '; ' + i.SNJ5 + '; ' + i.SNJ6 + '; ' + i.SNJ7 ,'; ;',''),'; ;',''),'; ;',''),'; ;',''),'; ;',''),'') as NextJob,FU.KELUHAN1 + ' ' + FU.KELUHAN2 + ' ' + FU.KELUHAN3 AS FIR from " _
      '             & " T_INVOICE D WITH (NOLOCK) LEFT OUTER JOIN DTLFU FU WITH (NOLOCK) ON FU.NO_WO = D.NO_WO left join complain c with (nolock) on c.no_wo = d.no_wo  Left Join Inspection i WITH (NOLOCK) on i.nowo = d.no_wo Left Join JobDTA J WITH (NOLOCK) On D.No_WO = J.NO_WO where D.NO_CHASSIS ='" & noRangka & "' AND D.FS_LOG_DEL = '' order by D.DATE DESC, D.NOINV DESC, D.KM DESC"
              
12            If Log.hideService = 1 Then
13                keteranganHide = " AND D.KM > 10 "
14            End If
              
15            strSql = "select isnull(komplain,'') komplain,IsNull(J.CATAT,'') as CATAT, D.DATE as Tanggal,D.NOINV,D.TOTAL as TotInv, D.NO_WO as NOWO, D.KM,D.FRONT as SA ,D.MECH ," _
                   & " REPLACE(REPLACE(REPLACE(D.REP1 + '; ' + D.REP2 + '; ' + D.REP3 + '; ' + D.REP4 + '; ' + D.REP5 + '; ' + D.REP6 + '; ' " _
                   & " + D.REP7,'; ;',''),'; ;',''),'; ;','') as JobOrder, D.REMARK1 as Pesan, D.SUGGEST as Saran, " _
                   & " D.REQ1 + '; ' + D.REQ2 + '; ' + D.REQ3 + '; ' + D.REQ4 + '; ' + D.REQ5 as Minta,D.FS_NM_DEALER as Dealer, " _
                   & " REPLACE(REPLACE(REPLACE(D.CMPL1 + '; ' + D.CMPL2 + '; ' + D.CMPL3 + '; ' + D.CMPL4 + '; ' + D.CMPL5 ,'; ;',''),'; ;',''),'; ;','') as Keluh,isnull(i.HasilFI,'') HasilFI,isnull(i.HasilTestDrive,'') as HasilTD, " _
                   & " isnull(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(i.SNJ1 + '; ' + i.SNJ2 + '; ' + i.SNJ3 + '; ' + i.SNJ4 + '; ' + i.SNJ5 + '; ' + i.SNJ6 + '; ' + i.SNJ7 ,'; ;',''),'; ;',''),'; ;',''),'; ;',''),'; ;',''),'') as NextJob,FU.KELUHAN1 + ' ' + FU.KELUHAN2 + ' ' + FU.KELUHAN3 AS FIR from " _
                   & " T_INVOICE D WITH (NOLOCK) LEFT OUTER JOIN DTLFU FU WITH (NOLOCK) ON FU.NO_WO = D.NO_WO left join complain c with (nolock) on c.no_wo = d.no_wo  Left Join Inspection i WITH (NOLOCK) on i.nowo = d.no_wo and i.TglEntry = (select top 1 TglEntry from INSPECTION where NoWo = d.NO_WO order by TglEntry desc ) Left Join JobDTA J WITH (NOLOCK) On D.No_WO = J.NO_WO where D.NO_CHASSIS ='" & noRangka & "' AND D.FS_LOG_DEL = '' " & keteranganHide & " order by D.DATE DESC, D.NOINV DESC, D.KM DESC"
                   
                   
16            Set rsCu = objO.FetchData(strSql)
              'Set rsCu = objCu.FetchAllHISTORY(noRangka)
17            If rsCu.RecordCount = 0 Then Exit Sub
18            Set Me.grid(0).DataSource = rsCu
19        End If
End Sub

