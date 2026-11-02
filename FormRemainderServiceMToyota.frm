VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{EE757A1F-B0AC-40BC-9E72-B8651740F53E}#1.0#0"; "ARProgBar.OCX"
Begin VB.Form FormRemainderServiceMToyota 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8835
   ClientLeft      =   0
   ClientTop       =   4005
   ClientWidth     =   12375
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FormRemainderServiceMToyota.frx":0000
   ScaleHeight     =   8835
   ScaleWidth      =   12375
   ShowInTaskbar   =   0   'False
   Begin VB.VScrollBar VScrollFUB 
      Height          =   8775
      LargeChange     =   300
      Left            =   11880
      Max             =   5200
      SmallChange     =   350
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   8775
      Left            =   480
      ScaleHeight     =   8775
      ScaleWidth      =   11655
      TabIndex        =   1
      Top             =   0
      Width           =   11655
      Begin TabDlg.SSTab SSFUB 
         Height          =   16005
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   28231
         _Version        =   393216
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   520
         TabCaption(0)   =   "  "
         TabPicture(0)   =   "FormRemainderServiceMToyota.frx":EC25
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Shape6"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lblPart"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Shape2"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label4"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Line1"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label2"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Line2"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Line3"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Line4"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "ImgView"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "Line6"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "Line7"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "Shape3"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "lblKet"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "Line18"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "Line19"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "Line20"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "Line16"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "Line21"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "Line22"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "Shape1"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "lblJml"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "Label48"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "Shape4"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "lblSave"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "Line23"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "Label59"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).Control(28)=   "Label57"
         Tab(0).Control(28).Enabled=   0   'False
         Tab(0).Control(29)=   "Shape9"
         Tab(0).Control(29).Enabled=   0   'False
         Tab(0).Control(30)=   "Line9"
         Tab(0).Control(30).Enabled=   0   'False
         Tab(0).Control(31)=   "Line25"
         Tab(0).Control(31).Enabled=   0   'False
         Tab(0).Control(32)=   "Shape8"
         Tab(0).Control(32).Enabled=   0   'False
         Tab(0).Control(33)=   "lbLen"
         Tab(0).Control(33).Enabled=   0   'False
         Tab(0).Control(34)=   "Label39"
         Tab(0).Control(34).Enabled=   0   'False
         Tab(0).Control(35)=   "Label49"
         Tab(0).Control(35).Enabled=   0   'False
         Tab(0).Control(36)=   "Line24"
         Tab(0).Control(36).Enabled=   0   'False
         Tab(0).Control(37)=   "Label56"
         Tab(0).Control(37).Enabled=   0   'False
         Tab(0).Control(38)=   "Label50"
         Tab(0).Control(38).Enabled=   0   'False
         Tab(0).Control(39)=   "Picture2"
         Tab(0).Control(39).Enabled=   0   'False
         Tab(0).Control(40)=   "DTSend"
         Tab(0).Control(40).Enabled=   0   'False
         Tab(0).Control(41)=   "DtRemain0"
         Tab(0).Control(41).Enabled=   0   'False
         Tab(0).Control(42)=   "ComDialogXls"
         Tab(0).Control(42).Enabled=   0   'False
         Tab(0).Control(43)=   "DtRemain"
         Tab(0).Control(43).Enabled=   0   'False
         Tab(0).Control(44)=   "ImageList"
         Tab(0).Control(44).Enabled=   0   'False
         Tab(0).Control(45)=   "cboType"
         Tab(0).Control(45).Enabled=   0   'False
         Tab(0).Control(46)=   "cmdProcess"
         Tab(0).Control(46).Enabled=   0   'False
         Tab(0).Control(47)=   "FrameGRP"
         Tab(0).Control(47).Enabled=   0   'False
         Tab(0).Control(48)=   "Frame1"
         Tab(0).Control(48).Enabled=   0   'False
         Tab(0).Control(49)=   "Frame9"
         Tab(0).Control(49).Enabled=   0   'False
         Tab(0).Control(50)=   "Frame10"
         Tab(0).Control(50).Enabled=   0   'False
         Tab(0).Control(51)=   "Frame4"
         Tab(0).Control(51).Enabled=   0   'False
         Tab(0).Control(52)=   "cmdTop"
         Tab(0).Control(52).Enabled=   0   'False
         Tab(0).Control(53)=   "cmdExit"
         Tab(0).Control(53).Enabled=   0   'False
         Tab(0).Control(54)=   "CmdXls"
         Tab(0).Control(54).Enabled=   0   'False
         Tab(0).Control(55)=   "cmdSave"
         Tab(0).Control(55).Enabled=   0   'False
         Tab(0).Control(56)=   "fot"
         Tab(0).Control(56).Enabled=   0   'False
         Tab(0).Control(57)=   "txtPesan"
         Tab(0).Control(57).Enabled=   0   'False
         Tab(0).Control(58)=   "cmdTemplate"
         Tab(0).Control(58).Enabled=   0   'False
         Tab(0).Control(59)=   "chGSM"
         Tab(0).Control(59).Enabled=   0   'False
         Tab(0).Control(60)=   "chCDMA"
         Tab(0).Control(60).Enabled=   0   'False
         Tab(0).Control(61)=   "cmdSMS"
         Tab(0).Control(61).Enabled=   0   'False
         Tab(0).Control(62)=   "GridRemainder"
         Tab(0).Control(62).Enabled=   0   'False
         Tab(0).Control(63)=   "cboKm"
         Tab(0).Control(63).Enabled=   0   'False
         Tab(0).Control(64)=   "MAPIMessages1"
         Tab(0).Control(64).Enabled=   0   'False
         Tab(0).Control(65)=   "MAPISession1"
         Tab(0).Control(65).Enabled=   0   'False
         Tab(0).Control(66)=   "progBar"
         Tab(0).Control(66).Enabled=   0   'False
         Tab(0).ControlCount=   67
         Begin ARProgBarCtrl.ARProgressBar progBar 
            Height          =   375
            Left            =   120
            TabIndex        =   136
            Top             =   1080
            Visible         =   0   'False
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   661
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
         Begin MSMAPI.MAPISession MAPISession1 
            Left            =   360
            Top             =   240
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DownloadMail    =   -1  'True
            LogonUI         =   -1  'True
            NewSession      =   0   'False
         End
         Begin MSMAPI.MAPIMessages MAPIMessages1 
            Left            =   0
            Top             =   120
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            AddressEditFieldCount=   1
            AddressModifiable=   0   'False
            AddressResolveUI=   0   'False
            FetchSorted     =   0   'False
            FetchUnreadOnly =   0   'False
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
            Left            =   5205
            Style           =   2  'Dropdown List
            TabIndex        =   132
            Top             =   960
            Visible         =   0   'False
            Width           =   1065
         End
         Begin TrueOleDBGrid70.TDBGrid GridRemainder 
            Height          =   2475
            Left            =   120
            TabIndex        =   39
            Top             =   1875
            Width           =   4950
            _ExtentX        =   8731
            _ExtentY        =   4366
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "outlet_ID"
            Columns(0).DataField=   "outlet_ID"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "one_account_ID"
            Columns(1).DataField=   "one_account_ID"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "dealer_customer_ID"
            Columns(2).DataField=   "dealer_customer_ID"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "email"
            Columns(3).DataField=   "email"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "preferred_contact_channel"
            Columns(4).DataField=   "preferred_contact_channel"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "reminder_ID"
            Columns(5).DataField=   "reminder_ID"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "activity"
            Columns(6).DataField=   "activity"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "activity_plan_scheduled_date"
            Columns(7).DataField=   "activity_plan_scheduled_date"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "auto_reminder_status"
            Columns(8).DataField=   "auto_reminder_status"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "priority_call"
            Columns(9).DataField=   "priority_call"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "extended_warranty_status"
            Columns(10).DataField=   "extended_warranty_status"
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(11)._VlistStyle=   0
            Columns(11)._MaxComboItems=   5
            Columns(11).Caption=   "customer_habit"
            Columns(11).DataField=   "customer_habit"
            Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(12)._VlistStyle=   0
            Columns(12)._MaxComboItems=   5
            Columns(12).Caption=   "last_habit"
            Columns(12).DataField=   "last_habit"
            Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(13)._VlistStyle=   0
            Columns(13)._MaxComboItems=   5
            Columns(13).Caption=   "next_service_status"
            Columns(13).DataField=   "next_service_status"
            Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(14)._VlistStyle=   0
            Columns(14)._MaxComboItems=   5
            Columns(14).Caption=   "last_service_date"
            Columns(14).DataField=   "last_service_date"
            Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(15)._VlistStyle=   0
            Columns(15)._MaxComboItems=   5
            Columns(15).Caption=   "next_service_date"
            Columns(15).DataField=   "next_service_date"
            Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(16)._VlistStyle=   0
            Columns(16)._MaxComboItems=   5
            Columns(16).Caption=   "ncs_status"
            Columns(16).DataField=   "ncs_status"
            Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(17)._VlistStyle=   0
            Columns(17)._MaxComboItems=   5
            Columns(17).Caption=   "katashiki_suffix"
            Columns(17).DataField=   "katashiki_suffix"
            Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(18)._VlistStyle=   0
            Columns(18)._MaxComboItems=   5
            Columns(18).Caption=   "vin"
            Columns(18).DataField=   "vin"
            Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(19)._VlistStyle=   0
            Columns(19)._MaxComboItems=   5
            Columns(19).Caption=   "model"
            Columns(19).DataField=   "model"
            Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(20)._VlistStyle=   0
            Columns(20)._MaxComboItems=   5
            Columns(20).Caption=   "variant"
            Columns(20).DataField=   "variant"
            Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(21)._VlistStyle=   0
            Columns(21)._MaxComboItems=   5
            Columns(21).Caption=   "FullName"
            Columns(21).DataField=   "FullName"
            Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(22)._VlistStyle=   0
            Columns(22)._MaxComboItems=   5
            Columns(22).Caption=   "police_number"
            Columns(22).DataField=   "police_number"
            Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(23)._VlistStyle=   0
            Columns(23)._MaxComboItems=   5
            Columns(23).Caption=   "phone_number"
            Columns(23).DataField=   "phone_number"
            Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(24)._VlistStyle=   0
            Columns(24)._MaxComboItems=   5
            Columns(24).Caption=   "Color"
            Columns(24).DataField=   "Color"
            Columns(24)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(25)._VlistStyle=   0
            Columns(25)._MaxComboItems=   5
            Columns(25).Caption=   "program_tab"
            Columns(25).DataField=   "program_tab"
            Columns(25)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(26)._VlistStyle=   0
            Columns(26)._MaxComboItems=   5
            Columns(26).Caption=   "next_service_stage"
            Columns(26).DataField=   "next_service_stage"
            Columns(26)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(27)._VlistStyle=   0
            Columns(27)._MaxComboItems=   5
            Columns(27).Caption=   "reminder_message"
            Columns(27).DataField=   "reminder_message"
            Columns(27)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   28
            Splits(0)._UserFlags=   0
            Splits(0).AnchorRightColumn=   -1  'True
            Splits(0).AllowSizing=   -1  'True
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).AllowColSelect=   0   'False
            Splits(0).AllowColMove=   -1  'True
            Splits(0).FetchRowStyle=   -1  'True
            Splits(0).DividerColor=   14215660
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=28"
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
            Splits(0)._ColumnProps(45)=   "Column(11).Width=3281"
            Splits(0)._ColumnProps(46)=   "Column(11).DividerColor=0"
            Splits(0)._ColumnProps(47)=   "Column(11)._WidthInPix=3175"
            Splits(0)._ColumnProps(48)=   "Column(11).Order=12"
            Splits(0)._ColumnProps(49)=   "Column(12).Width=3281"
            Splits(0)._ColumnProps(50)=   "Column(12).DividerColor=0"
            Splits(0)._ColumnProps(51)=   "Column(12)._WidthInPix=3175"
            Splits(0)._ColumnProps(52)=   "Column(12).Order=13"
            Splits(0)._ColumnProps(53)=   "Column(13).Width=3281"
            Splits(0)._ColumnProps(54)=   "Column(13).DividerColor=0"
            Splits(0)._ColumnProps(55)=   "Column(13)._WidthInPix=3175"
            Splits(0)._ColumnProps(56)=   "Column(13).Order=14"
            Splits(0)._ColumnProps(57)=   "Column(14).Width=3281"
            Splits(0)._ColumnProps(58)=   "Column(14).DividerColor=0"
            Splits(0)._ColumnProps(59)=   "Column(14)._WidthInPix=3175"
            Splits(0)._ColumnProps(60)=   "Column(14).Order=15"
            Splits(0)._ColumnProps(61)=   "Column(15).Width=3281"
            Splits(0)._ColumnProps(62)=   "Column(15).DividerColor=0"
            Splits(0)._ColumnProps(63)=   "Column(15)._WidthInPix=3175"
            Splits(0)._ColumnProps(64)=   "Column(15).Order=16"
            Splits(0)._ColumnProps(65)=   "Column(16).Width=3281"
            Splits(0)._ColumnProps(66)=   "Column(16).DividerColor=0"
            Splits(0)._ColumnProps(67)=   "Column(16)._WidthInPix=3175"
            Splits(0)._ColumnProps(68)=   "Column(16).Order=17"
            Splits(0)._ColumnProps(69)=   "Column(17).Width=3281"
            Splits(0)._ColumnProps(70)=   "Column(17).DividerColor=0"
            Splits(0)._ColumnProps(71)=   "Column(17)._WidthInPix=3175"
            Splits(0)._ColumnProps(72)=   "Column(17).Order=18"
            Splits(0)._ColumnProps(73)=   "Column(18).Width=3281"
            Splits(0)._ColumnProps(74)=   "Column(18).DividerColor=0"
            Splits(0)._ColumnProps(75)=   "Column(18)._WidthInPix=3175"
            Splits(0)._ColumnProps(76)=   "Column(18).Order=19"
            Splits(0)._ColumnProps(77)=   "Column(19).Width=3281"
            Splits(0)._ColumnProps(78)=   "Column(19).DividerColor=0"
            Splits(0)._ColumnProps(79)=   "Column(19)._WidthInPix=3175"
            Splits(0)._ColumnProps(80)=   "Column(19).Order=20"
            Splits(0)._ColumnProps(81)=   "Column(20).Width=3281"
            Splits(0)._ColumnProps(82)=   "Column(20).DividerColor=0"
            Splits(0)._ColumnProps(83)=   "Column(20)._WidthInPix=3175"
            Splits(0)._ColumnProps(84)=   "Column(20).Order=21"
            Splits(0)._ColumnProps(85)=   "Column(21).Width=3281"
            Splits(0)._ColumnProps(86)=   "Column(21).DividerColor=0"
            Splits(0)._ColumnProps(87)=   "Column(21)._WidthInPix=3175"
            Splits(0)._ColumnProps(88)=   "Column(21).Order=22"
            Splits(0)._ColumnProps(89)=   "Column(22).Width=3281"
            Splits(0)._ColumnProps(90)=   "Column(22).DividerColor=0"
            Splits(0)._ColumnProps(91)=   "Column(22)._WidthInPix=3175"
            Splits(0)._ColumnProps(92)=   "Column(22).Order=23"
            Splits(0)._ColumnProps(93)=   "Column(23).Width=3281"
            Splits(0)._ColumnProps(94)=   "Column(23).DividerColor=0"
            Splits(0)._ColumnProps(95)=   "Column(23)._WidthInPix=3175"
            Splits(0)._ColumnProps(96)=   "Column(23).Order=24"
            Splits(0)._ColumnProps(97)=   "Column(24).Width=3281"
            Splits(0)._ColumnProps(98)=   "Column(24).DividerColor=0"
            Splits(0)._ColumnProps(99)=   "Column(24)._WidthInPix=3175"
            Splits(0)._ColumnProps(100)=   "Column(24).Order=25"
            Splits(0)._ColumnProps(101)=   "Column(25).Width=3281"
            Splits(0)._ColumnProps(102)=   "Column(25).DividerColor=0"
            Splits(0)._ColumnProps(103)=   "Column(25)._WidthInPix=3175"
            Splits(0)._ColumnProps(104)=   "Column(25).Order=26"
            Splits(0)._ColumnProps(105)=   "Column(26).Width=3281"
            Splits(0)._ColumnProps(106)=   "Column(26).DividerColor=0"
            Splits(0)._ColumnProps(107)=   "Column(26)._WidthInPix=3175"
            Splits(0)._ColumnProps(108)=   "Column(26).Order=27"
            Splits(0)._ColumnProps(109)=   "Column(27).Width=3281"
            Splits(0)._ColumnProps(110)=   "Column(27).DividerColor=0"
            Splits(0)._ColumnProps(111)=   "Column(27)._WidthInPix=3175"
            Splits(0)._ColumnProps(112)=   "Column(27).Order=28"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            Appearance      =   0
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=155,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=152,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=153,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=154,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=151,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=148,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=149,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=150,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=147,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=144,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=145,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=146,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=143,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=140,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=141,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=142,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=139,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=136,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=137,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=138,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=135,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=132,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=133,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=134,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=131,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=128,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=129,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=130,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=127,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=124,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=125,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=126,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=123,.parent=13"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=120,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=121,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=122,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=119,.parent=13"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=116,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=117,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=118,.parent=17"
            _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=114,.parent=13"
            _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=111,.parent=14"
            _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=112,.parent=15"
            _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=113,.parent=17"
            _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=110,.parent=13"
            _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=107,.parent=14"
            _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=108,.parent=15"
            _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=109,.parent=17"
            _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=106,.parent=13"
            _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=103,.parent=14"
            _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=104,.parent=15"
            _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=105,.parent=17"
            _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=102,.parent=13"
            _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=99,.parent=14"
            _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=100,.parent=15"
            _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=101,.parent=17"
            _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=98,.parent=13"
            _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=95,.parent=14"
            _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=96,.parent=15"
            _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=97,.parent=17"
            _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=94,.parent=13"
            _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=91,.parent=14"
            _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=92,.parent=15"
            _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=93,.parent=17"
            _StyleDefs(100) =   "Splits(0).Columns(16).Style:id=90,.parent=13"
            _StyleDefs(101) =   "Splits(0).Columns(16).HeadingStyle:id=87,.parent=14"
            _StyleDefs(102) =   "Splits(0).Columns(16).FooterStyle:id=88,.parent=15"
            _StyleDefs(103) =   "Splits(0).Columns(16).EditorStyle:id=89,.parent=17"
            _StyleDefs(104) =   "Splits(0).Columns(17).Style:id=86,.parent=13"
            _StyleDefs(105) =   "Splits(0).Columns(17).HeadingStyle:id=83,.parent=14"
            _StyleDefs(106) =   "Splits(0).Columns(17).FooterStyle:id=84,.parent=15"
            _StyleDefs(107) =   "Splits(0).Columns(17).EditorStyle:id=85,.parent=17"
            _StyleDefs(108) =   "Splits(0).Columns(18).Style:id=82,.parent=13"
            _StyleDefs(109) =   "Splits(0).Columns(18).HeadingStyle:id=79,.parent=14"
            _StyleDefs(110) =   "Splits(0).Columns(18).FooterStyle:id=80,.parent=15"
            _StyleDefs(111) =   "Splits(0).Columns(18).EditorStyle:id=81,.parent=17"
            _StyleDefs(112) =   "Splits(0).Columns(19).Style:id=78,.parent=13"
            _StyleDefs(113) =   "Splits(0).Columns(19).HeadingStyle:id=75,.parent=14"
            _StyleDefs(114) =   "Splits(0).Columns(19).FooterStyle:id=76,.parent=15"
            _StyleDefs(115) =   "Splits(0).Columns(19).EditorStyle:id=77,.parent=17"
            _StyleDefs(116) =   "Splits(0).Columns(20).Style:id=74,.parent=13"
            _StyleDefs(117) =   "Splits(0).Columns(20).HeadingStyle:id=71,.parent=14"
            _StyleDefs(118) =   "Splits(0).Columns(20).FooterStyle:id=72,.parent=15"
            _StyleDefs(119) =   "Splits(0).Columns(20).EditorStyle:id=73,.parent=17"
            _StyleDefs(120) =   "Splits(0).Columns(21).Style:id=70,.parent=13"
            _StyleDefs(121) =   "Splits(0).Columns(21).HeadingStyle:id=67,.parent=14"
            _StyleDefs(122) =   "Splits(0).Columns(21).FooterStyle:id=68,.parent=15"
            _StyleDefs(123) =   "Splits(0).Columns(21).EditorStyle:id=69,.parent=17"
            _StyleDefs(124) =   "Splits(0).Columns(22).Style:id=66,.parent=13"
            _StyleDefs(125) =   "Splits(0).Columns(22).HeadingStyle:id=63,.parent=14"
            _StyleDefs(126) =   "Splits(0).Columns(22).FooterStyle:id=64,.parent=15"
            _StyleDefs(127) =   "Splits(0).Columns(22).EditorStyle:id=65,.parent=17"
            _StyleDefs(128) =   "Splits(0).Columns(23).Style:id=62,.parent=13"
            _StyleDefs(129) =   "Splits(0).Columns(23).HeadingStyle:id=59,.parent=14"
            _StyleDefs(130) =   "Splits(0).Columns(23).FooterStyle:id=60,.parent=15"
            _StyleDefs(131) =   "Splits(0).Columns(23).EditorStyle:id=61,.parent=17"
            _StyleDefs(132) =   "Splits(0).Columns(24).Style:id=58,.parent=13"
            _StyleDefs(133) =   "Splits(0).Columns(24).HeadingStyle:id=55,.parent=14"
            _StyleDefs(134) =   "Splits(0).Columns(24).FooterStyle:id=56,.parent=15"
            _StyleDefs(135) =   "Splits(0).Columns(24).EditorStyle:id=57,.parent=17"
            _StyleDefs(136) =   "Splits(0).Columns(25).Style:id=54,.parent=13"
            _StyleDefs(137) =   "Splits(0).Columns(25).HeadingStyle:id=51,.parent=14"
            _StyleDefs(138) =   "Splits(0).Columns(25).FooterStyle:id=52,.parent=15"
            _StyleDefs(139) =   "Splits(0).Columns(25).EditorStyle:id=53,.parent=17"
            _StyleDefs(140) =   "Splits(0).Columns(26).Style:id=50,.parent=13"
            _StyleDefs(141) =   "Splits(0).Columns(26).HeadingStyle:id=47,.parent=14"
            _StyleDefs(142) =   "Splits(0).Columns(26).FooterStyle:id=48,.parent=15"
            _StyleDefs(143) =   "Splits(0).Columns(26).EditorStyle:id=49,.parent=17"
            _StyleDefs(144) =   "Splits(0).Columns(27).Style:id=46,.parent=13"
            _StyleDefs(145) =   "Splits(0).Columns(27).HeadingStyle:id=43,.parent=14"
            _StyleDefs(146) =   "Splits(0).Columns(27).FooterStyle:id=44,.parent=15"
            _StyleDefs(147) =   "Splits(0).Columns(27).EditorStyle:id=45,.parent=17"
            _StyleDefs(148) =   "Named:id=33:Normal"
            _StyleDefs(149) =   ":id=33,.parent=0"
            _StyleDefs(150) =   "Named:id=34:Heading"
            _StyleDefs(151) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(152) =   ":id=34,.wraptext=-1"
            _StyleDefs(153) =   "Named:id=35:Footing"
            _StyleDefs(154) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(155) =   "Named:id=36:Selected"
            _StyleDefs(156) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(157) =   "Named:id=37:Caption"
            _StyleDefs(158) =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(159) =   "Named:id=38:HighlightRow"
            _StyleDefs(160) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(161) =   "Named:id=39:EvenRow"
            _StyleDefs(162) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(163) =   "Named:id=40:OddRow"
            _StyleDefs(164) =   ":id=40,.parent=33"
            _StyleDefs(165) =   "Named:id=41:RecordSelector"
            _StyleDefs(166) =   ":id=41,.parent=34"
            _StyleDefs(167) =   "Named:id=42:FilterBar"
            _StyleDefs(168) =   ":id=42,.parent=33"
            _StyleDefs(169) =   "Named:id=115:adaWO"
            _StyleDefs(170) =   ":id=115,.parent=33,.bgcolor=&HDDF3C7&,.fgcolor=&H4080FF&,.bold=-1,.fontsize=825"
            _StyleDefs(171) =   ":id=115,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(172) =   ":id=115,.fontname=MS Sans Serif"
         End
         Begin VB.CommandButton cmdSMS 
            BackColor       =   &H00FFFFFF&
            Caption         =   "SEND"
            Height          =   735
            Left            =   10380
            Picture         =   "FormRemainderServiceMToyota.frx":EC41
            Style           =   1  'Graphical
            TabIndex        =   124
            Top             =   3180
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CheckBox chCDMA 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6000
            TabIndex        =   121
            Top             =   3600
            Width           =   2775
         End
         Begin VB.CheckBox chGSM 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6000
            TabIndex        =   120
            Top             =   3240
            Width           =   2535
         End
         Begin VB.CommandButton cmdTemplate 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Template"
            Height          =   732
            Left            =   10380
            Picture         =   "FormRemainderServiceMToyota.frx":EFC3
            Style           =   1  'Graphical
            TabIndex        =   112
            Top             =   1980
            Width           =   855
         End
         Begin VB.TextBox txtPesan 
            Appearance      =   0  'Flat
            Height          =   855
            Left            =   5280
            MaxLength       =   1000
            MultiLine       =   -1  'True
            TabIndex        =   111
            Top             =   1920
            Width           =   4935
         End
         Begin VB.CheckBox fot 
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
            Left            =   5280
            TabIndex        =   110
            Top             =   2820
            Width           =   3255
         End
         Begin VB.CommandButton cmdSave 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Save"
            Height          =   735
            Left            =   9480
            Picture         =   "FormRemainderServiceMToyota.frx":F618
            Style           =   1  'Graphical
            TabIndex        =   109
            Top             =   3180
            Width           =   735
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
            Picture         =   "FormRemainderServiceMToyota.frx":FCE4
            Style           =   1  'Graphical
            TabIndex        =   106
            Top             =   480
            Width           =   735
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
            Picture         =   "FormRemainderServiceMToyota.frx":10160
            Style           =   1  'Graphical
            TabIndex        =   92
            Top             =   480
            Width           =   735
         End
         Begin VB.CommandButton cmdTop 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Back To Top"
            Height          =   735
            Left            =   120
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FormRemainderServiceMToyota.frx":108F0
            Style           =   1  'Graphical
            TabIndex        =   91
            Top             =   12945
            Width           =   11175
         End
         Begin VB.Frame Frame4 
            Caption         =   " FAKTUR DATA STNK "
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
            Height          =   2055
            Left            =   120
            TabIndex        =   66
            Top             =   10440
            Width           =   11175
            Begin VB.TextBox txtAgama 
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
               Left            =   7680
               MaxLength       =   15
               TabIndex        =   73
               Top             =   720
               Width           =   1670
            End
            Begin VB.TextBox txtHoby 
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
               Left            =   7680
               MaxLength       =   15
               TabIndex        =   72
               Top             =   1440
               Width           =   3255
            End
            Begin VB.TextBox txtPekerjaan 
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
               Left            =   7680
               MaxLength       =   15
               TabIndex        =   71
               Top             =   1080
               Width           =   3255
            End
            Begin VB.TextBox txtTempatLahir 
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
               Left            =   7680
               MaxLength       =   15
               TabIndex        =   70
               Top             =   360
               Width           =   1670
            End
            Begin VB.TextBox txtKotaS 
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
               Left            =   1680
               MaxLength       =   30
               TabIndex        =   69
               Top             =   1440
               Width           =   1815
            End
            Begin VB.TextBox txtAlamatS 
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
               Height          =   690
               Left            =   1680
               MaxLength       =   200
               MultiLine       =   -1  'True
               TabIndex        =   68
               Top             =   720
               Width           =   4455
            End
            Begin VB.TextBox txtNamaS 
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
               Left            =   1680
               MaxLength       =   50
               TabIndex        =   67
               Top             =   360
               Width           =   4455
            End
            Begin TDBDate6Ctl.TDBDate DTLahir 
               Height          =   330
               Left            =   9360
               TabIndex        =   74
               Top             =   360
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calendar        =   "FormRemainderServiceMToyota.frx":10C7B
               Caption         =   "FormRemainderServiceMToyota.frx":10DA7
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FormRemainderServiceMToyota.frx":10E13
               Keys            =   "FormRemainderServiceMToyota.frx":10E31
               Spin            =   "FormRemainderServiceMToyota.frx":10E8F
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
               Enabled         =   0
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
            Begin VB.Label Label68 
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
               Left            =   7560
               TabIndex        =   89
               Top             =   1080
               Width           =   75
            End
            Begin VB.Label Label67 
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
               Left            =   7560
               TabIndex        =   88
               Top             =   1440
               Width           =   75
            End
            Begin VB.Label Label66 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Hoby"
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
               Left            =   6480
               TabIndex        =   87
               Top             =   1440
               Width           =   495
            End
            Begin VB.Label Label65 
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
               Left            =   7560
               TabIndex        =   86
               Top             =   360
               Width           =   75
            End
            Begin VB.Label Label64 
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
               Left            =   7560
               TabIndex        =   85
               Top             =   720
               Width           =   75
            End
            Begin VB.Label Label63 
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
               Left            =   8280
               TabIndex        =   84
               Top             =   1200
               Width           =   75
            End
            Begin VB.Label Label62 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "TTL"
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
               Left            =   6480
               TabIndex        =   83
               Top             =   360
               Width           =   345
            End
            Begin VB.Label Label61 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Agama"
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
               Left            =   6480
               TabIndex        =   82
               Top             =   705
               Width           =   645
            End
            Begin VB.Label Label58 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Pekerjaan"
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
               Left            =   6480
               TabIndex        =   81
               Top             =   1080
               Width           =   930
            End
            Begin VB.Label Label41 
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
               TabIndex        =   80
               Top             =   1440
               Width           =   75
            End
            Begin VB.Label Label38 
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
               TabIndex        =   79
               Top             =   720
               Width           =   75
            End
            Begin VB.Label Label37 
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
               TabIndex        =   78
               Top             =   360
               Width           =   75
            End
            Begin VB.Label Label36 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Kota"
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
               Left            =   240
               TabIndex        =   77
               Top             =   1440
               Width           =   465
            End
            Begin VB.Label Label35 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Alamat "
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
               Left            =   240
               TabIndex        =   76
               Top             =   720
               Width           =   750
            End
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nama STNK"
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
               Left            =   240
               TabIndex        =   75
               Top             =   360
               Width           =   1080
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   " CR 2 "
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
            Height          =   1215
            Left            =   7800
            TabIndex        =   59
            Top             =   9240
            Width           =   3495
            Begin VB.TextBox txtKMCR2 
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
               Left            =   1680
               MaxLength       =   15
               TabIndex        =   60
               Top             =   720
               Width           =   1575
            End
            Begin TDBDate6Ctl.TDBDate DTCR2 
               Height          =   330
               Left            =   1680
               TabIndex        =   61
               Top             =   360
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calendar        =   "FormRemainderServiceMToyota.frx":10EB7
               Caption         =   "FormRemainderServiceMToyota.frx":10FE3
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FormRemainderServiceMToyota.frx":1104F
               Keys            =   "FormRemainderServiceMToyota.frx":1106D
               Spin            =   "FormRemainderServiceMToyota.frx":110CB
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
               Enabled         =   0
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
            Begin VB.Label Label33 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tgl Validasi"
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
               Left            =   240
               TabIndex        =   65
               Top             =   360
               Width           =   1110
            End
            Begin VB.Label Label32 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Angka KM"
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
               Left            =   240
               TabIndex        =   64
               Top             =   720
               Width           =   960
            End
            Begin VB.Label Label31 
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
               TabIndex        =   63
               Top             =   375
               Width           =   75
            End
            Begin VB.Label Label30 
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
               TabIndex        =   62
               Top             =   735
               Width           =   75
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   " CR 1 "
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
            Height          =   1215
            Left            =   4320
            TabIndex        =   52
            Top             =   9240
            Width           =   3495
            Begin VB.TextBox txtKMCR1 
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
               Left            =   1680
               MaxLength       =   15
               TabIndex        =   53
               Top             =   720
               Width           =   1575
            End
            Begin TDBDate6Ctl.TDBDate DTCR1 
               Height          =   330
               Left            =   1680
               TabIndex        =   54
               Top             =   360
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calendar        =   "FormRemainderServiceMToyota.frx":110F3
               Caption         =   "FormRemainderServiceMToyota.frx":1121F
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FormRemainderServiceMToyota.frx":1128B
               Keys            =   "FormRemainderServiceMToyota.frx":112A9
               Spin            =   "FormRemainderServiceMToyota.frx":11307
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
               Enabled         =   0
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
            Begin VB.Label Label29 
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
               TabIndex        =   58
               Top             =   735
               Width           =   75
            End
            Begin VB.Label Label28 
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
               TabIndex        =   57
               Top             =   375
               Width           =   75
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Angka KM"
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
               Left            =   240
               TabIndex        =   56
               Top             =   720
               Width           =   960
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tgl Validasi"
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
               Left            =   240
               TabIndex        =   55
               Top             =   360
               Width           =   1110
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   " BPK "
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
            Height          =   1215
            Left            =   120
            TabIndex        =   45
            Top             =   9240
            Width           =   4215
            Begin TDBDate6Ctl.TDBDate DtDO 
               Height          =   330
               Left            =   1680
               TabIndex        =   46
               Top             =   360
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calendar        =   "FormRemainderServiceMToyota.frx":1132F
               Caption         =   "FormRemainderServiceMToyota.frx":1145B
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FormRemainderServiceMToyota.frx":114C7
               Keys            =   "FormRemainderServiceMToyota.frx":114E5
               Spin            =   "FormRemainderServiceMToyota.frx":11543
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
               Enabled         =   0
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
            Begin TDBDate6Ctl.TDBDate DtDEC 
               Height          =   330
               Left            =   1680
               TabIndex        =   47
               Top             =   720
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calendar        =   "FormRemainderServiceMToyota.frx":1156B
               Caption         =   "FormRemainderServiceMToyota.frx":11697
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FormRemainderServiceMToyota.frx":11703
               Keys            =   "FormRemainderServiceMToyota.frx":11721
               Spin            =   "FormRemainderServiceMToyota.frx":1177F
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
               Enabled         =   0
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
            Begin VB.Label Label25 
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
               TabIndex        =   51
               Top             =   735
               Width           =   75
            End
            Begin VB.Label Label24 
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
               TabIndex        =   50
               Top             =   375
               Width           =   75
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tgl DO / BPK"
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
               Left            =   240
               TabIndex        =   49
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tgl DEC"
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
               Left            =   240
               TabIndex        =   48
               Top             =   720
               Width           =   765
            End
         End
         Begin VB.Frame FrameGRP 
            Caption         =   ".:: SELECTED ITEM SERVICE ::."
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
            Height          =   4455
            Left            =   120
            TabIndex        =   5
            Top             =   4320
            Width           =   11205
            Begin VB.TextBox txtColor 
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
               Left            =   1800
               MaxLength       =   25
               TabIndex        =   133
               Top             =   2640
               Width           =   2655
            End
            Begin VB.CommandButton cmdHistory 
               BackColor       =   &H00FFFFFF&
               Caption         =   "&History"
               Height          =   735
               Left            =   7080
               Picture         =   "FormRemainderServiceMToyota.frx":117A7
               Style           =   1  'Graphical
               TabIndex        =   129
               Top             =   360
               Width           =   855
            End
            Begin VB.CheckBox onD 
               BackColor       =   &H00FFFFFF&
               Caption         =   "DM"
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
               Left            =   10200
               TabIndex        =   128
               Top             =   720
               Value           =   1  'Checked
               Width           =   615
            End
            Begin VB.CheckBox onS 
               BackColor       =   &H00FFFFFF&
               Caption         =   "SMS"
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
               Left            =   8280
               TabIndex        =   126
               Top             =   720
               Value           =   1  'Checked
               Width           =   975
            End
            Begin VB.CheckBox onC 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Call"
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
               Left            =   9240
               TabIndex        =   125
               Top             =   720
               Value           =   1  'Checked
               Width           =   975
            End
            Begin VB.TextBox txtNextSvcStage 
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
               Left            =   9720
               MaxLength       =   40
               TabIndex        =   107
               Top             =   1560
               Width           =   1335
            End
            Begin VB.TextBox txtProgram 
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
               Left            =   5760
               MaxLength       =   25
               TabIndex        =   99
               Top             =   1560
               Width           =   2175
            End
            Begin VB.TextBox txtTipe 
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
               Left            =   1800
               MaxLength       =   25
               TabIndex        =   96
               Top             =   2280
               Width           =   2655
            End
            Begin VB.CheckBox chCall 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Call By Phone"
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
               Left            =   6480
               TabIndex        =   95
               Top             =   4035
               Width           =   1695
            End
            Begin VB.CheckBox chSMS 
               BackColor       =   &H00FFFFFF&
               Caption         =   "SMS"
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
               Left            =   6480
               TabIndex        =   94
               Top             =   3795
               Width           =   1695
            End
            Begin VB.CommandButton cmdCall 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Call"
               Height          =   735
               Left            =   5160
               Picture         =   "FormRemainderServiceMToyota.frx":11E14
               Style           =   1  'Graphical
               TabIndex        =   16
               Top             =   3540
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.CommandButton cmdDM 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Create DM"
               Height          =   735
               Left            =   5160
               Picture         =   "FormRemainderServiceMToyota.frx":12134
               Style           =   1  'Graphical
               TabIndex        =   15
               Top             =   3540
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.TextBox txtDateLast 
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
               Left            =   5400
               MaxLength       =   40
               TabIndex        =   14
               Top             =   360
               Width           =   1575
            End
            Begin VB.TextBox txtKetModel 
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
               Left            =   1800
               MaxLength       =   25
               TabIndex        =   13
               Top             =   1920
               Width           =   2655
            End
            Begin VB.TextBox txtDateNext 
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
               Left            =   5400
               MaxLength       =   40
               TabIndex        =   12
               Top             =   720
               Width           =   1575
            End
            Begin VB.TextBox txtReminderID 
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
               Left            =   1800
               MaxLength       =   30
               TabIndex        =   11
               Top             =   360
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
               Left            =   1800
               MaxLength       =   100
               TabIndex        =   10
               Top             =   1200
               Width           =   6135
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
               Left            =   1800
               MaxLength       =   15
               TabIndex        =   9
               Top             =   1560
               Width           =   2655
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
               Left            =   1800
               MaxLength       =   25
               TabIndex        =   8
               Top             =   3000
               Width           =   2655
            End
            Begin VB.TextBox txtReminder 
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
               Height          =   1050
               Left            =   5760
               MaxLength       =   25
               TabIndex        =   7
               Top             =   2280
               Width           =   5295
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
               Left            =   5760
               MaxLength       =   40
               TabIndex        =   6
               Top             =   1920
               Width           =   2175
            End
            Begin VB.CheckBox chDM 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Direct Mail"
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
               Left            =   6480
               TabIndex        =   93
               Top             =   3555
               Width           =   1695
            End
            Begin VB.Label lblFOP 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   375
               Left            =   8160
               TabIndex        =   137
               Top             =   2880
               Width           =   2895
            End
            Begin VB.Label Label70 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Warna"
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
               Left            =   240
               TabIndex        =   135
               Top             =   2640
               Width           =   570
            End
            Begin VB.Label Label69 
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
               Left            =   1680
               TabIndex        =   134
               Top             =   2640
               Width           =   75
            End
            Begin VB.Line Line27 
               BorderColor     =   &H00C0C0C0&
               X1              =   8280
               X2              =   11040
               Y1              =   1000
               Y2              =   1000
            End
            Begin VB.Line Line26 
               BorderColor     =   &H00C0C0C0&
               X1              =   8280
               X2              =   11040
               Y1              =   650
               Y2              =   650
            End
            Begin VB.Label Label51 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "View List Reminder :"
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
               Left            =   8280
               TabIndex        =   127
               Top             =   400
               Width           =   2100
            End
            Begin VB.Line Line8 
               BorderColor     =   &H000000FF&
               X1              =   120
               X2              =   11160
               Y1              =   4320
               Y2              =   4320
            End
            Begin VB.Label nCall 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ".  .  .  .  .  .  .  .  ."
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
               Left            =   8400
               TabIndex        =   104
               Top             =   4065
               Width           =   1500
            End
            Begin VB.Label nSMS 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ".  .  .  .  .  .  .  .  ."
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
               Left            =   8400
               TabIndex        =   103
               Top             =   3810
               Width           =   1500
            End
            Begin VB.Label nDM 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ".  .  .  .  .  .  .  .  ."
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
               Left            =   8400
               TabIndex        =   102
               Top             =   3555
               Width           =   1500
            End
            Begin VB.Label Label47 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Program"
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
               Left            =   4680
               TabIndex        =   101
               Top             =   1560
               Width           =   750
            End
            Begin VB.Label Label46 
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
               Left            =   5640
               TabIndex        =   100
               Top             =   1560
               Width           =   75
            End
            Begin VB.Label Label44 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tipe"
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
               Left            =   240
               TabIndex        =   98
               Top             =   2280
               Width           =   360
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
               Left            =   1680
               TabIndex        =   97
               Top             =   2280
               Width           =   75
            End
            Begin VB.Shape Shape7 
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00808080&
               Height          =   840
               Left            =   6360
               Top             =   3480
               Width           =   4695
            End
            Begin VB.Shape Shape5 
               BorderColor     =   &H00808080&
               FillColor       =   &H00FFFFFF&
               Height          =   840
               Left            =   5040
               Shape           =   4  'Rounded Rectangle
               Top             =   3460
               Width           =   1215
            End
            Begin VB.Line Line17 
               BorderColor     =   &H000000FF&
               X1              =   120
               X2              =   11160
               Y1              =   3400
               Y2              =   3400
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Date Last Service"
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
               Left            =   3720
               TabIndex        =   37
               Top             =   360
               Width           =   1500
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Date Next Service"
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
               Left            =   3720
               TabIndex        =   36
               Top             =   720
               Width           =   1515
            End
            Begin VB.Label Label19 
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
               Left            =   5280
               TabIndex        =   35
               Top             =   360
               Width           =   75
            End
            Begin VB.Label Label18 
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
               Left            =   5280
               TabIndex        =   34
               Top             =   720
               Width           =   75
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
               Left            =   1680
               TabIndex        =   33
               Top             =   1920
               Width           =   75
            End
            Begin VB.Label Label73 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Model"
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
               Left            =   240
               TabIndex        =   32
               Top             =   1920
               Width           =   510
            End
            Begin VB.Label Label43 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Reminder ID"
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
               Left            =   240
               TabIndex        =   31
               Top             =   345
               Width           =   1035
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
               Left            =   1680
               TabIndex        =   30
               Top             =   360
               Width           =   75
            End
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
               TabIndex        =   29
               Top             =   360
               Width           =   75
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
               Left            =   1680
               TabIndex        =   28
               Top             =   1560
               Width           =   75
            End
            Begin VB.Label Label91 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "No Polisi"
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
               Left            =   240
               TabIndex        =   27
               Top             =   1560
               Width           =   855
            End
            Begin VB.Label Label17 
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
               Left            =   240
               TabIndex        =   26
               Top             =   1200
               Width           =   1110
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Next Service Stage"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   8040
               TabIndex        =   25
               Top             =   1560
               Width           =   1500
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label15 
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
               Left            =   1680
               TabIndex        =   24
               Top             =   1200
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
               Left            =   9600
               TabIndex        =   23
               Top             =   1560
               Width           =   75
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
               Left            =   1680
               TabIndex        =   22
               Top             =   3000
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
               Left            =   240
               TabIndex        =   21
               Top             =   3000
               Width           =   1065
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
               Left            =   5640
               TabIndex        =   20
               Top             =   2400
               Width           =   75
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
               Left            =   5640
               TabIndex        =   19
               Top             =   1920
               Width           =   75
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Reminder Message"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Left            =   4680
               TabIndex        =   18
               Top             =   2280
               Width           =   930
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Telpon"
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
               Left            =   4680
               TabIndex        =   17
               Top             =   1920
               Width           =   570
            End
            Begin VB.Line Line13 
               BorderColor     =   &H000000FF&
               X1              =   120
               X2              =   11160
               Y1              =   1125
               Y2              =   1125
            End
            Begin VB.Line Line12 
               BorderColor     =   &H000000FF&
               X1              =   120
               X2              =   11160
               Y1              =   3360
               Y2              =   3360
            End
            Begin VB.Line Line11 
               BorderColor     =   &H000000FF&
               X1              =   120
               X2              =   11160
               Y1              =   240
               Y2              =   240
            End
            Begin VB.Shape Shape10 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00FFFFFF&
               Height          =   690
               Left            =   8160
               Top             =   360
               Width           =   2940
            End
            Begin VB.Label Label60 
               Caption         =   "Label60"
               Height          =   495
               Left            =   9840
               TabIndex        =   130
               Top             =   3480
               Width           =   975
            End
         End
         Begin VB.CommandButton cmdProcess 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Pr&ocess"
            Height          =   732
            Left            =   5280
            Picture         =   "FormRemainderServiceMToyota.frx":12450
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   240
            Width           =   732
         End
         Begin VB.ComboBox cboType 
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
            Left            =   3240
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   975
            Width           =   1920
         End
         Begin MSComctlLib.ImageList ImageList 
            Left            =   7680
            Top             =   360
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   145
            ImageHeight     =   78
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   3
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormRemainderServiceMToyota.frx":12ABD
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormRemainderServiceMToyota.frx":13FD4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormRemainderServiceMToyota.frx":15CCF
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin TDBDate6Ctl.TDBDate DtRemain 
            Height          =   315
            Left            =   3720
            TabIndex        =   38
            Top             =   270
            Width           =   1350
            _Version        =   65536
            _ExtentX        =   2381
            _ExtentY        =   556
            Calendar        =   "FormRemainderServiceMToyota.frx":173BF
            Caption         =   "FormRemainderServiceMToyota.frx":174EB
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "FormRemainderServiceMToyota.frx":17557
            Keys            =   "FormRemainderServiceMToyota.frx":17575
            Spin            =   "FormRemainderServiceMToyota.frx":175D3
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
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
         Begin MSComDlg.CommonDialog ComDialogXls 
            Left            =   8760
            Top             =   840
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin TDBDate6Ctl.TDBDate DtRemain0 
            Height          =   315
            Left            =   2040
            TabIndex        =   108
            Top             =   270
            Width           =   1350
            _Version        =   65536
            _ExtentX        =   2381
            _ExtentY        =   556
            Calendar        =   "FormRemainderServiceMToyota.frx":175FB
            Caption         =   "FormRemainderServiceMToyota.frx":17727
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "FormRemainderServiceMToyota.frx":17793
            Keys            =   "FormRemainderServiceMToyota.frx":177B1
            Spin            =   "FormRemainderServiceMToyota.frx":1780F
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
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
         Begin TDBDate6Ctl.TDBDate DTSend 
            Height          =   330
            Left            =   7320
            TabIndex        =   113
            Top             =   3990
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   582
            Calendar        =   "FormRemainderServiceMToyota.frx":17837
            Caption         =   "FormRemainderServiceMToyota.frx":17963
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "FormRemainderServiceMToyota.frx":179CF
            Keys            =   "FormRemainderServiceMToyota.frx":179ED
            Spin            =   "FormRemainderServiceMToyota.frx":17A4B
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
            Enabled         =   0
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
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   8400
            ScaleHeight     =   315
            ScaleWidth      =   2715
            TabIndex        =   131
            Top             =   7080
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.Label Label50 
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
            Left            =   5880
            TabIndex        =   123
            Top             =   3720
            Width           =   75
         End
         Begin VB.Label Label56 
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
            Left            =   5880
            TabIndex        =   122
            Top             =   3360
            Width           =   75
         End
         Begin VB.Line Line24 
            BorderColor     =   &H000000FF&
            X1              =   5160
            X2              =   5160
            Y1              =   1920
            Y2              =   4320
         End
         Begin VB.Label Label49 
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
            Left            =   5280
            TabIndex        =   119
            Top             =   3600
            Width           =   1080
         End
         Begin VB.Label Label39 
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
            Left            =   5280
            TabIndex        =   118
            Top             =   3240
            Width           =   1035
         End
         Begin VB.Label lbLen 
            Alignment       =   1  'Right Justify
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
            Left            =   8880
            TabIndex        =   117
            Top             =   2820
            Width           =   1335
         End
         Begin VB.Shape Shape8 
            BorderColor     =   &H00C0C0C0&
            FillColor       =   &H00808080&
            Height          =   855
            Left            =   10320
            Top             =   1920
            Width           =   975
         End
         Begin VB.Line Line25 
            BorderColor     =   &H00808080&
            X1              =   12720
            X2              =   17640
            Y1              =   4320
            Y2              =   4320
         End
         Begin VB.Line Line9 
            BorderColor     =   &H00808080&
            X1              =   5280
            X2              =   10200
            Y1              =   3960
            Y2              =   3960
         End
         Begin VB.Shape Shape9 
            BorderColor     =   &H00C0C0C0&
            FillColor       =   &H00808080&
            Height          =   855
            Left            =   10320
            Top             =   3120
            Width           =   975
         End
         Begin VB.Label Label57 
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
            Left            =   6960
            TabIndex        =   116
            Top             =   4035
            Width           =   75
         End
         Begin VB.Label Label59 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Waktu Kirim SMS"
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
            TabIndex        =   115
            Top             =   4035
            Width           =   1485
         End
         Begin VB.Line Line23 
            BorderColor     =   &H00808080&
            X1              =   5280
            X2              =   11280
            Y1              =   4335
            Y2              =   4335
         End
         Begin VB.Label lblSave 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Save Item : 0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   9480
            TabIndex        =   114
            Top             =   4020
            Width           =   1725
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00C0C0C0&
            FillColor       =   &H00FFFFFF&
            Height          =   975
            Left            =   9240
            Shape           =   4  'Rounded Rectangle
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label48 
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
            Left            =   9120
            TabIndex        =   105
            Top             =   1500
            Width           =   1455
         End
         Begin VB.Label lblJml 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   10200
            TabIndex        =   43
            Top             =   1500
            Width           =   975
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00C0C0C0&
            FillColor       =   &H00FFFFFF&
            Height          =   975
            Left            =   10200
            Shape           =   4  'Rounded Rectangle
            Top             =   360
            Width           =   975
         End
         Begin VB.Line Line22 
            BorderColor     =   &H000000FF&
            X1              =   120
            X2              =   9000
            Y1              =   12600
            Y2              =   12600
         End
         Begin VB.Line Line21 
            BorderColor     =   &H000000FF&
            X1              =   1200
            X2              =   10200
            Y1              =   12705
            Y2              =   12705
         End
         Begin VB.Line Line16 
            BorderColor     =   &H000000FF&
            X1              =   2280
            X2              =   11280
            Y1              =   12810
            Y2              =   12810
         End
         Begin VB.Line Line20 
            BorderColor     =   &H000000FF&
            X1              =   720
            X2              =   3960
            Y1              =   9165
            Y2              =   9165
         End
         Begin VB.Line Line19 
            BorderColor     =   &H000000FF&
            X1              =   1440
            X2              =   3960
            Y1              =   9060
            Y2              =   9060
         End
         Begin VB.Line Line18 
            BorderColor     =   &H000000FF&
            X1              =   2160
            X2              =   3960
            Y1              =   8970
            Y2              =   8970
         End
         Begin VB.Label lblKet 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   42
            Top             =   650
            Width           =   4695
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H00C0C0C0&
            FillColor       =   &H00FFFFFF&
            Height          =   975
            Left            =   5175
            Shape           =   4  'Rounded Rectangle
            Top             =   120
            Width           =   975
         End
         Begin VB.Line Line7 
            BorderColor     =   &H000000FF&
            X1              =   2415
            X2              =   4095
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Line Line6 
            BorderColor     =   &H000000FF&
            X1              =   1440
            X2              =   4095
            Y1              =   1680
            Y2              =   1680
         End
         Begin VB.Image ImgView 
            Height          =   1095
            Left            =   6240
            Top             =   240
            Width           =   2295
         End
         Begin VB.Line Line4 
            BorderColor     =   &H000000FF&
            X1              =   255
            X2              =   5055
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Line Line3 
            BorderColor     =   &H000000FF&
            X1              =   255
            X2              =   5055
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Line Line2 
            BorderColor     =   &H000000FF&
            X1              =   255
            X2              =   5055
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Label Label2 
            Caption         =   "Type Reminder      :"
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
            Left            =   1575
            TabIndex        =   41
            Top             =   1035
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.Line Line1 
            BorderColor     =   &H000000FF&
            X1              =   255
            X2              =   5055
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "LIST REMINDER SERVICE"
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
            Left            =   840
            TabIndex        =   44
            Top             =   1500
            Width           =   10335
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00404040&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00404040&
            Height          =   375
            Left            =   135
            Shape           =   4  'Rounded Rectangle
            Top             =   1440
            Width           =   11205
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "DATA KENDARAAN - STNK"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   720
            TabIndex        =   90
            Top             =   8955
            Width           =   10695
         End
         Begin VB.Shape Shape6 
            BackColor       =   &H00404040&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00C0C0FF&
            Height          =   375
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   8880
            Width           =   11220
         End
         Begin VB.Label Label1 
            Caption         =   "Tgl Reminder         :                                 s/d"
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
            Left            =   375
            TabIndex        =   40
            Top             =   300
            Width           =   3600
         End
      End
   End
End
Attribute VB_Name = "FormRemainderServiceMToyota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : FormRemainderServiceMToyota
' Author    : Rudi
' Purpose   :
'---------------------------------------------------------------------------------------

 
Option Explicit

Private OnLoad As Boolean
Private col As TrueOleDBGrid70.Column
Private cols As TrueOleDBGrid70.Columns
Private Urut As String

Private objRem As New ClsRemainder
Private rsRem As ADODB.Recordset
Private rsRem2 As ADODB.Recordset
'.:: RecordSet Temporary ::.
Private rsTem As ADODB.Recordset
Private rsGrid As New ADODB.Recordset
'.:: Item Temporary ::.
Private rsSvc As ADODB.Recordset
Private TglNext As Date
Private valDM As Integer
Private valSM As Integer
Private valCl As Integer

Private rsTempXLS As New ADODB.Recordset
Private rsTempMRS As New ADODB.Recordset
Private rsList As New ADODB.Recordset
Private Remainder As String
Private NewId As Double
Dim ft As New FormTemplateSMS
Dim support As New clsFormDealer

Dim remindSA    As Integer
Dim remindMRA   As Integer

Private tagihanFOP  As String
Private kodeMDL     As String

Public Property Let OnLoaded(value As Boolean)
   OnLoad = value
End Property

Private Sub SetPicture()
   If IsNull(Me.DtRemain.value) = True Then Exit Sub
   cmdDM.Visible = False
   cmdSMS.Visible = False
   cmdCall.Visible = False
   Select Case cboType.Text
      Case "Direct Mail"
         ImgView.Picture = ImageList.ListImages(1).Picture
         lblKet.Caption = "Ket  :  H-14 (Next Service)" ' = " & DateAdd("d", 14, Me.DtRemain.value) & ")"
         'TglNext = DateAdd("d", 14, Me.DtRemain.value)
         cmdDM.Visible = True
      Case "SMS"
         ImgView.Picture = ImageList.ListImages(3).Picture
         lblKet.Caption = "Ket  :  H-8 (Next Service)" ' = " & DateAdd("d", 8, Me.DtRemain.value) & ")"
         'TglNext = DateAdd("d", 8, Me.DtRemain.value)
         cmdSMS.Visible = True
      Case "Call By Phone"
         ImgView.Picture = ImageList.ListImages(2).Picture
         If gTipeUser = "SA" Then
            lblKet.Caption = "Ket  :  H+" & remindSA & " (Jatuh Tempo Service)"
         Else
            lblKet.Caption = "Ket  :  H-" & IIf(remindMRA < 0, remindMRA * -1, remindMRA) & " (Next Service)" ' = " & DateAdd("d", 7, Me.DtRemain.value) & ")"
         End If
         'TglNext = DateAdd("d", 7, Me.DtRemain.value)
         cmdCall.Visible = True
      Case Else
         ImgView.Picture = Nothing
         lblKet.Caption = ""
         'TglNext = Null
   End Select
   lblJml.Caption = "0"
End Sub

Private Sub cboKm_Change()
    Call cmdProcess_Click
End Sub

Private Sub cboKm_Click()
    Call cmdProcess_Click
End Sub

Private Sub cboType_Change()
'    If Left(cboType.text, 1) = "D" Then
'        cboKm.Visible = True
'    Else
'        cboKm.Visible = False
'        Call cmdProcess_Click
'    End If
End Sub

Private Sub cboType_Click()
   Call SetPicture
    Dim nC As Integer
    
        GridRemainder.Width = 11175
        cmdTemplate.Visible = False
        cmdSave.Visible = False
    If Left(cboType.Text, 1) = "D" Then
        cboKm.Visible = True
    Else
        cboKm.Visible = False
        Call cmdProcess_Click
    End If
End Sub

Private Sub chCall_Click()
   Call ListRemainderValue
End Sub

Private Sub chDM_Click()
   Call ListRemainderValue
End Sub

Private Sub chSMS_Click()
   Call ListRemainderValue
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdHistory_Click()
          Dim ft As New FormTransparan
1         If Log.RPLHistori = 1 Then
2             ft.xChassis = txtNoChassis
3             ft.xIsForm = "HIST_HO"
4             ft.show (vbModal)
5         Else
6             If MsgBox("Apakah ingin melihat History Service HO?", vbYesNo, "Ambil Data Dari HO") = vbYes Then
7                 If Log.LogHOConn = 0 Then
                      '' 2015-09-08 | Konek HO ''
8                     If MsgBox("Gagal Koneksi ke Server HO !" & vbCrLf & "Apakah ingin menyambungkan koneksi kembali?", vbYesNo) = vbYes Then
9                         If Get_ConnectionProfil_Bengkel99 Then
10                            Log.LogHOConn = 1
11                        Else
12                            Exit Sub
13                        End If
14                    Else
15                        Exit Sub
16                    End If
              '        MsgBox "Gagal Koneksi ke Server HO !", vbInformation + vbOKOnly, "History On Line HO"
              '        Exit Sub
17                End If
                  'jika mengambil data dari ho
                 
18                ft.xChassis = txtNoChassis
19                ft.xIsForm = "HIST_HO"
20                ft.show (vbModal)
21            Else
                  'jika mengambil data local
                  'Dim ft As New FormTransparan
22                ft.xChassis = txtNoChassis
23                ft.xIsForm = "HIST"
24                ft.show (vbModal)
25            End If
26        End If
End Sub

Private Sub cmdOpen_Click()
   Dim f As New FormTransparan
   With f
      .xChassis = Trim(txtNoChassis.Text)
      .xIsForm = "CUST_EDIT"
'      .Top = 2250
'      .Left = Me.Left
      .isMe = False
      .show (vbModal)
   End With
   Call cmdProcess_Click
End Sub

Private Sub cmdSave_Click()
If GridRemainder.Columns(16).value <> -1 Then
   MsgBox "Silahkan pilih item customer dahulu !", vbCritical + vbOKOnly, "MRS"
   Exit Sub
End If
If txtPesan.Text = "" Then
   MsgBox "Silahkan isi pesan sms anda !" & Chr(13) & "Atau pilih template SMS terlebih dahulu", vbCritical + vbOKOnly, "MRS"
   Exit Sub
End If
If chGSM.value = 0 And chCDMA.value = 0 Then
   MsgBox "Silahkan pilih no telp terlebih dahulu", vbCritical + vbOKOnly, "MRS"
   Exit Sub
End If
    Dim nPesanSMS As String
    nPesanSMS = Trim(Me.txtPesan.Text) & IIf(fot.value = 1, " " & fot.Caption, "")
    
    objRem.RemoveItemSMS (Me.txtNoChassis.Text)
    With objRem
        If Me.chGSM.value = 1 Then
            .NomorRangka = Trim(Me.txtNoChassis.Text)
            .NoGSM = Trim(Replace(Me.chGSM.Caption, "-", ""))
            .NoCDMA = Trim("")
            .TglSendNow = Me.DTSend.value
            .PesanSMS = nPesanSMS
            .AddNewItemSMS_Temp
        End If
        If Me.chCDMA.value = 1 Then
            .NomorRangka = Trim(Me.txtNoChassis.Text)
            .NoGSM = Trim("")
            .NoCDMA = Trim(Replace(Me.chCDMA.Caption, "-", ""))
            .TglSendNow = Me.DTSend.value
            .PesanSMS = nPesanSMS
            .AddNewItemSMS_Temp
        End If
    End With
    Call ShowItemSave_Count
End Sub

Private Sub ShowItemSave_Count()
    Dim rsSave As ADODB.Recordset
    Set rsSave = objRem.GetDataSMS
    lblSave.Caption = "Save Item : " & rsSave.RecordCount
End Sub

Private Sub cmdTemplate_Click()
   ft.OnLoaded = True
   ft.show vbModal
   Remainder = ft.Remainder
   txtPesan.Text = Remainder
   Call PhoneChange
End Sub

Private Sub PhoneChange()
On Error Resume Next
End Sub

Private Sub cmdTop_Click()
   Me.VScrollFUB.value = 1
End Sub

Private Sub CmdXls_Click()
1     On Error GoTo Err_Export
          Dim Er As New ErrInfo
2         If rsList.RecordCount = 0 Then
3             Er.HeadCap = "Export To Excel !"
4             Er.InfoCap = "List Data Reminder Service Kosong !"
5             Beep
6             Er.show (vbModal)
7             Exit Sub
8         End If
          Dim MyPath As String
          Dim onVal As String
9         MyPath = ""
          Dim f As New FormExplorer
10        f.show (vbModal)
11        MyPath = f.PathSaving
12        If MyPath = "" Then Exit Sub
13        If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
14        If f.FileName = "" Then Exit Sub
15        onVal = MyPath & f.FileName & ".XLS"
           
           ''20200513 REQ PAK KOKOK MINTA DITAMBAHKAN LOG UNTUK USER YG MENDOWNLOAD
          ''-----------------------------------------------------------------------------
16        sSQL = "insert into logDownloadExcelMRA(kodeuser,namamenu,tglExport,tglFilter1,tglFilter2,kriteria,totalrecord,destinasi)"
17        sSQL = sSQL & " values ('" & gSA & "','FormRemainderServiceMToyota',getdate(),'" & IIf(IsNull(DtRemain0), "1900-01-01", Format(DtRemain0, "yyyy-mm-dd")) & "','" & Format(DtRemain, "yyyy-mm-dd") & "','" & cboType.Text & "','" & rsList.RecordCount & "','" & onVal & "') "
18        Call objO.InsertData(sSQL)
          
19        GridRemainder.Columns(7).Visible = False
20         GridRemainder.ExportToFile onVal, True
21        GridRemainder.Columns(7).Visible = True
22                Er.HeadCap = "Export to excel !"
23                Er.InfoCap = "Data sudah di Export ke File Excel." & Chr(13) & _
                               "File : " & onVal
                  
24                Beep
25                Er.show (vbModal)
                  'MsgBox "Data sudah di Export ke File Excel.", vbInformation, "Export File >> Excel"
26         Exit Sub
Err_Export:
27          If err.Number = 32755 Or err.Number = 91 Then GoTo Keluar
28          MsgBox err.Number, "Procedure : " & Me.Name & ".CmdXls_Click"
Keluar:

End Sub

Private Sub DtRemain_Change()
   If TglNext > "1/1/1900" And IsNull(TglNext) = False Then
      Call cmdProcess_Click
   Else
      Set Me.GridRemainder.DataSource = Nothing
      With Me
         .txtReminderID.Text = ""
         
         .txtDateLast.Text = ""
         .txtDateNext.Text = ""
         .txtNama.Text = ""
         .txtTipe.Text = ""
         .txtColor.Text = ""
         .txtProgram.Text = ""
         .txtNextSvcStage.Text = ""
         .txtNoPol.Text = ""
         .txtTelp.Text = ""
         .txtReminder.Text = ""
         
         
         .txtKetModel.Text = ""
         .txtNoChassis.Text = ""
         
         valDM = 0
         valSM = 0
         valCl = 0
         .chDM.value = valDM
         .chSMS.value = valSM
         .chCall.value = valCl
         .nDM.Caption = ".  .  .  .  .  .  .  .  ."
         .nSMS.Caption = ".  .  .  .  .  .  .  .  ."
         .nCall.Caption = ".  .  .  .  .  .  .  .  ."
      End With
   End If
End Sub

Private Sub Form_Load()
  ' If OnLoad = False Then Exit Sub
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
   Me.fot.Caption = Log.LogNas
   If Log.hansaRenkei = 1 Then
        remindSA = 30
        remindMRA = -7
        Set rsRem = objO.FetchData("SELECT reminderSA,reminderMRA FROM settingHansaRenkei WITH (NOLOCK)")
        If Not rsRem.EOF Then
            remindSA = CInt(rsRem!remindersa)
            remindMRA = CInt(rsRem!remindermra)
        End If
   End If
   Call ListAllCombo
   Call LoadForm
End Sub

Private Sub ListAllCombo()
   With Me
      .cboType.Clear
      .cboType.AddItem "Call MRA"
      .cboType.AddItem "SMS"
      .cboType.AddItem "Call By Phone"
      .cboType.Text = "Call By Phone"
      .cboType.Visible = False
      
      .cboKm.Clear
      .cboKm.AddItem "10 K"
      .cboKm.AddItem "20 K"
   End With
End Sub

Private Sub LoadForm()
   Me.DtRemain.value = Date
   Me.fot.Caption = gPerush
   Call objRem.RemoveAllItemSMS
   lblSave.Caption = "Save Item : 0"
   Call NullRecord
End Sub


Private Sub cmdProcess_Click()
         
1        If cboType.Text = "" Then
            'MsgBox "Silahkan Isi Type Remainder dan Kelipatan KM Terlebih Dahulu !", vbInformation + vbOKOnly
2           Exit Sub
3        End If
4        Call ListDataRemainder

End Sub
Private Sub errTglNextService()
        Dim Er As New ErrInfo
        Er.HeadCap = "Tanggal Reminder !"
        Er.InfoCap = "Tanggal Remainder tidak valid !"
        Beep
        Er.show (vbModal)
      'MsgBox "Tanggal Remainder tidak valid !", vbExclamation, "Remainder"
End Sub

Private Sub ListDataRemainder()
   Call SetPicture

   If IsNull(DtRemain0) = True And IsNull(DtRemain) = True Then
      'Call errTglNextService
      Exit Sub
   End If
   If DtRemain0 > "01/01/1900" And IsNull(DtRemain) = True Then
      Call errTglNextService
      Exit Sub
   End If
   If DtRemain0 > DtRemain Then
      Call errTglNextService
      Exit Sub
   End If
   
      
   Set GridRemainder.DataSource = Nothing
    Dim nRem As Integer
    If IsNull(DtRemain0) = True And DtRemain.value > "01/01/1900" Then
        nRem = 1
        Select Case cboType.Text
           Case "SMS"
              lblKet.Caption = "Ket  :  H-8 (Next) 01/01/1900 - " & DateAdd("d", 8, Me.DtRemain.value)
           Case "Call By Phone"
              ''lblKet.Caption = "Ket  :  H-7 (Next) 01/01/1900 - " & DateAdd("d", 7, Me.DtRemain.value)
              If gTipeUser = "SA" Then
                   ''lblKet.Caption = "Ket  :  H+" & remindSA & " (Jatuh Tempo Service)"
                   lblKet.Caption = "Ket  :  H+" & remindSA & " (Jatuh Tempo Service) " & DateAdd("d", remindSA * -1, Me.DtRemain0.value) & " - " & DateAdd("d", remindSA * -1, Me.DtRemain.value)
                Else
                   ''lblKet.Caption = "Ket  :  H-" & IIf(remindMRA < 0, remindMRA * -1, remindMRA) & " (Next Service)" ' = " & DateAdd("d", 7, Me.DtRemain.value) & ")"
                   lblKet.Caption = "Ket  :  H-" & IIf(remindMRA < 0, remindMRA * -1, remindMRA) & " (Next Service) " & DateAdd("d", IIf(remindMRA < 0, remindMRA * -1, remindMRA), Me.DtRemain0.value) & " - " & DateAdd("d", IIf(remindMRA < 0, remindMRA * -1, remindMRA), Me.DtRemain.value)
                End If
           Case "Direct Mail"
              lblKet.Caption = "Ket  :  H-14 (Next) 01/01/1900 - " & DateAdd("d", 14, Me.DtRemain.value)
        End Select
    ElseIf DtRemain0.value > "01/01/1900" And DtRemain.value > "01/01/1900" Then
        nRem = 0
        Select Case cboType.Text
           Case "SMS"
              lblKet.Caption = "Ket  :  H-8 (Next) " & DateAdd("d", 8, Me.DtRemain0.value) & " - " & DateAdd("d", 8, Me.DtRemain.value)
           Case "Call By Phone"
              'lblKet.Caption = "Ket  :  H-7 (Next) " & DateAdd("d", 7, Me.DtRemain0.value) & " - " & DateAdd("d", 7, Me.DtRemain.value)
                If gTipeUser = "SA" Then
                   ''lblKet.Caption = "Ket  :  H+" & remindSA & " (Jatuh Tempo Service)"22
                   lblKet.Caption = "Ket  :  H" & IIf(remindSA > 0, "+" & remindSA, remindSA) & " (Jatuh Tempo Service) " & DateAdd("d", remindSA * -1, Me.DtRemain0.value) & " - " & DateAdd("d", remindSA * -1, Me.DtRemain.value)
                Else
                   ''lblKet.Caption = "Ket  :  H-" & IIf(remindMRA < 0, remindMRA * -1, remindMRA) & " (Next Service)" ' = " & DateAdd("d", 7, Me.DtRemain.value) & ")"
                   lblKet.Caption = "Ket  :  H" & IIf(remindMRA > 0, "-" & remindMRA, remindMRA) & " (Next Service) " & DateAdd("d", IIf(remindMRA < 0, remindMRA * -1, remindMRA), Me.DtRemain0.value) & " - " & DateAdd("d", IIf(remindMRA < 0, remindMRA * -1, remindMRA), Me.DtRemain.value)
                End If
           Case "Direct Mail"
              lblKet.Caption = "Ket  :  H-14 (Next) " & DateAdd("d", 14, Me.DtRemain0.value) & " - " & DateAdd("d", 14, Me.DtRemain.value)
        End Select
    End If
    
    Dim tNext0 As Date
    Dim tNext1 As Date
    Dim tRem0 As Date
    Dim tRem1 As Date
    Dim Ket As String
        Select Case cboType.Text
           Case "SMS"
                tRem0 = "01/01/1900"
                tRem1 = "01/01/1900"
                If IsNull(DtRemain0) = True Then
                    tNext0 = "01/01/1900"
                Else
                    tNext0 = DateAdd("d", 8, Me.DtRemain0.value)
                    'tRem0 = Me.DtRemain0.value
                End If
                tNext1 = DateAdd("d", 8, Me.DtRemain.value)
                'tRem1 = Me.DtRemain.value
                Ket = ""
           Case "Call By Phone"
                If IsNull(DtRemain0) = True Then
                    tNext0 = "01/01/1900"
                    tRem0 = "01/01/1900"
                Else
                    If gTipeUser = "SA" Then
                        tNext0 = DateAdd("d", remindSA * -1, Me.DtRemain0.value)
                    Else
                        tNext0 = DateAdd("d", remindMRA * -1, Me.DtRemain0.value)
                    End If
                    
                    
                    tRem0 = Me.DtRemain0.value
                End If
                If gTipeUser = "SA" Then
                    tNext1 = DateAdd("d", remindSA * -1, Me.DtRemain.value)
                Else
                    tNext1 = DateAdd("d", remindMRA * -1, Me.DtRemain.value)
                End If
                ''tNext1 = DateAdd("d", 7, Me.DtRemain.value)
                tRem1 = Me.DtRemain.value
                Ket = ""
           Case "Direct Mail"
                tRem0 = "01/01/1900"
                tRem1 = "01/01/1900"
                If IsNull(DtRemain0) = True Then
                    tNext0 = "01/01/1900"
                Else
                    tNext0 = DateAdd("d", 14, Me.DtRemain0.value)
                    'tRem0 = Me.DtRemain0.value
                End If
                tNext1 = DateAdd("d", 14, Me.DtRemain.value)
                'tRem1 = Me.DtRemain.value
                Ket = " AND SBENEXT = '" & Trim(IIf(cboKm = "10 K", 10000, 20000)) & "' "
        End Select
        
   sSQL = "Select p1.reminder_ID,p1.last_service_date,p1.next_service_date, p1.vin,p1.model,p1.variant,p1.first_name + ' ' + P1.last_name as fullname,"
   sSQL = sSQL & " p1.police_number,p1.phone_number,p1.model,p1.variant,p1.color,p1.program_tab,"
   sSQL = sSQL & " p1.next_service_stage,p1.reminder_message,"
   sSQL = sSQL & " p1.auto_reminder_status,p1.priority_call,p1.extended_warranty_status,"
   sSQL = sSQL & " p1.customer_habit   ,p1.last_service_date,p1.next_service_date,p1.ncs_status from ReplBengkel99..p2_reminder p1"
   
   sSQL = "Select p1.outlet_ID,p1.one_account_ID,p1.dealer_customer_ID,p1.first_name + ' ' + p1.last_name as FullName,p1.phone_number,p1.email,p2.preferred_contact_channel,p1.reminder_ID,p1.activity,p1.activity_plan_scheduled_date,"
   sSQL = sSQL & " p1.auto_reminder_status,p1.reminder_message,p1.priority_call,p1.extended_warranty_status,p1.customer_habit,p1.last_habit,p1.next_service_status,p1.last_service_date,"
   sSQL = sSQL & " P1.next_service_date , P1.ncs_status, P1.program_tab, P1.next_service_stage, P1.vin, P1.police_number, P1.katashiki_suffix, P1.model, P1.variant, P1.color_code, P1.Color"
   sSQL = sSQL & " from ReplBengkel99..p2_reminder p1"
   sSQL = sSQL & " left join ReplBengkel99..p2_reminder_preferred_contact_channel p2 on p1.event_ID=p2.event_ID"
   sSQL = sSQL & " where p1.vin not in (select NoRangka from M_ESTIMASI)"
   sSQL = sSQL & " and p1.activity in ('SERVICE_BOOKING','SERVICE_REMINDER') and p1.reminder_ID not in (select reminder_ID from M_ListReminderCall)"
   sSQL = sSQL & " and (next_service_date between '" & Format(DtRemain0, "yyyy-mm-dd") & "' and '" & Format(DtRemain, "yyyy-mm-dd") & "')"
   Set rsRem = objO.FetchData(sSQL)
   
   If rsRem.RecordCount = 0 Then
        Call NullRecord
   Else
    With Me
        '' 20200512DEV
        If Log.disableExportMra = 1 Then
            If gTipeUser = "KABENG" Or gTipeUser = "ASSKABENG" Or gTipeUser = "ADMIN" Or gTipeUser = "HO" Then
                'cmdCSV.Enabled = True
                .CmdXls.Enabled = True
            Else
                'cmdCSV.Enabled = False
                .CmdXls.Enabled = False
            End If
        Else
            .CmdXls.Enabled = True
        End If
        ''.CmdXls.Enabled = True
        .cmdTemplate.Enabled = True
        .chGSM.Enabled = True
        .chCDMA.Enabled = True
        .cmdSave.Enabled = True
        .cmdSMS.Enabled = True
        .DTSend.Text = "__/__/____"
        '.cmdOpen.Enabled = True
    End With
      Dim NextList As Boolean
      CreateMainList True
      Dim nRangka As String
      nRangka = ""
      
    Dim xNoWoR As String
      If rsRem.RecordCount > 0 Then
        rsRem.MoveFirst
        'ProgBar.Max = rsRem.RecordCount + 1
        'ProgBar.value = 1
        'ProgBar.Visible = True
        Do Until rsRem.EOF
          If nRangka <> Trim(rsRem!vin) Then
              NextList = False
              xNoWoR = "" ''Nul(rsRem.Fields(1).value)
'              If onS.value = 1 Then
'                  If AddToListRemainderSMS(xNoWoR, Trim(rsRem("vin")), 0) = True Then NextList = True
'              End If
'              If onC.value = 1 Then
'                  If AddToListRemainderCall(xNoWoR, Trim(rsRem("vin")), 0) = True Then NextList = True
'              End If
'              If onD.value = 1 Then
'                  If AddToListRemainderDM(xNoWoR, Trim(rsRem("vin")), 0) = True Then NextList = True
'              End If
              If NextList = False Then
                  rsList.AddNew
'                  rsList!reminder_ID = Nul(rsRem!reminder_ID)
'                  rsList!last_service_date = CStr(NulTgl(rsRem!last_service_date))
'                  rsList!next_service_date = CStr(NulTgl(rsRem!next_service_date))
'                  rsList!vin = NulSID(rsRem!vin)
'                  rsList!model = NulSID(rsRem!model)
'                  rsList!variant = Nul(rsRem!variant)
'                  rsList!FullName = NulSID(rsRem!FullName)
'                  rsList!police_number = NulSID(rsRem!police_number)
'                  rsList!phone_number = NulSID(rsRem!phone_number)
'                  rsList!Color = NulSID(rsRem!Color)
'                  rsList!program_tab = NulSID(rsRem!program_tab)
'                  rsList!next_service_stage = NulSID(rsRem!next_service_stage)
'                  rsList!reminder_message = NulSID(rsRem!reminder_message)
                    
                rsList!outlet_ID = Nul(rsRem!outlet_ID)
                rsList!one_account_ID = Nul(rsRem!one_account_ID)
                rsList!dealer_customer_ID = Nul(rsRem!dealer_customer_ID)
                rsList!email = Nul(rsRem!email)
                rsList!preferred_contact_channel = Nul(rsRem!preferred_contact_channel)
                rsList!reminder_ID = Nul(rsRem!reminder_ID)
                rsList!activity = Nul(rsRem!activity)
                rsList!activity_plan_scheduled_date = CStr(NulTgl(rsRem!activity_plan_scheduled_date))
                rsList!auto_reminder_status = Nul(rsRem!auto_reminder_status)
                rsList!priority_call = Nul(rsRem!priority_call)
                rsList!extended_warranty_status = Nul(rsRem!extended_warranty_status)
                rsList!customer_habit = Nul(rsRem!customer_habit)
                rsList!last_habit = Nul(rsRem!last_habit)
                rsList!next_service_status = Nul(rsRem!next_service_status)
                rsList!last_service_date = NulTgl(rsRem!last_service_date)
                rsList!next_service_date = NulTgl(rsRem!next_service_date)
                rsList!ncs_status = Nul(rsRem!ncs_status)
                rsList!katashiki_suffix = Nul(rsRem!katashiki_suffix)
                rsList!vin = Nul(rsRem!vin)
                rsList!model = Nul(rsRem!model)
                rsList!Variant = Nul(rsRem!Variant)
                rsList!FullName = Nul(rsRem!FullName)
                rsList!police_number = Nul(rsRem!police_number)
                rsList!phone_number = Nul(rsRem!phone_number)
                rsList!Color = Nul(rsRem!Color)
                rsList!program_tab = Nul(rsRem!program_tab)
                rsList!next_service_stage = Nul(rsRem!next_service_stage)
                rsList!reminder_message = Nul(rsRem!reminder_message)
                  
              End If
          End If
          If Not rsRem.EOF Then nRangka = Trim(rsRem!vin)
          'If ProgBar.value <= ProgBar.Max Then
            'ProgBar.value = ProgBar.value + 1
            'DoEvents
          'End If
          
          rsRem.MoveNext
        Loop
        'ProgBar.Visible = False
     End If
     
        lblJml.Caption = rsList.RecordCount
        Call BindToControl(rsList)
        GridRemainder.FetchRowStyle = True
   End If
End Sub

Private Sub NullRecord()
      With Me
        .CmdXls.Enabled = False
        .cmdTemplate.Enabled = False
        .chGSM.Enabled = False
        .chCDMA.Enabled = False
        .cmdSave.Enabled = False
        .cmdSMS.Enabled = False
        .DTSend.Text = "__/__/____"
        '.cmdOpen.Enabled = False
       
         .txtReminderID.Text = ""
         
         .txtDateLast.Text = ""
         .txtDateNext.Text = ""
         .txtNama.Text = ""
         .txtTipe.Text = ""
         .txtColor.Text = ""
         .txtProgram.Text = ""
         .txtNextSvcStage.Text = ""
         .txtNoPol.Text = ""
         .txtTelp.Text = ""
         .txtReminder.Text = ""
         
         
         .txtDateNext.Text = ""
         .txtKetModel.Text = ""
         .txtNoChassis.Text = ""
         
         valDM = 0
         valSM = 0
         valCl = 0
         .chDM.value = valDM
         .chSMS.value = valSM
         .chCall.value = valCl
         .nDM.Caption = ".  .  .  .  .  .  .  .  ."
         .nSMS.Caption = ".  .  .  .  .  .  .  .  ."
         .nCall.Caption = ".  .  .  .  .  .  .  .  ."
      End With
End Sub

Private Sub BindToControl(rs As ADODB.Recordset)
   With Me
     .txtReminderID.Text = NulSID(rs!reminder_ID)
      .txtDateLast.Text = NulTgl(NulSID(rs!last_service_date))
      .txtDateNext = NulTgl(NulSID(rs!next_service_date))
      .txtNoChassis.Text = NulSID(rs!vin)
      .txtNama.Text = NulSID(rs!FullName)
      .txtNoPol.Text = NulSID(rs!police_number)
      .txtKetModel.Text = NulSID(rs!model)
      .txtTipe.Text = NulSID(rs!Variant)
      .txtColor.Text = NulSID(rs!Color)
      .txtProgram.Text = NulSID(rs!program_tab)
      .txtNextSvcStage.Text = NulA(NulSID(rs!next_service_stage))
      .txtReminder.Text = NulSID(rs!reminder_message)
   End With
   Call PhoneChange
End Sub

Private Sub ListRemainder()
On Error Resume Next
'   Dim rsListR As ADODB.Recordset
'   Set rsListR = objRem.FetchListRemainder(txtNoWO.Text, txtNoChassis.Text, NulA(txtKmNext.Text), False)
'   With Me
'      valDM = 0
'      valSM = 0
'      valCl = 0
'      .chDM.value = valDM
'      .chSMS.value = valSM
'      .chCall.value = valCl
'      .nDM.Caption = ".  .  .  .  .  .  .  .  ."
'      .nSMS.Caption = ".  .  .  .  .  .  .  .  ."
'      .nCall.Caption = ".  .  .  .  .  .  .  .  ."
'      If rsListR.RecordCount <> 0 Then
'         rsListR.MoveFirst
'         Do Until rsListR.EOF
'            If rsListR.Fields(0).value = "DM" Then
'               valDM = 1
'               .chDM.value = valDM
'               .nDM.Caption = Trim(rsListR.Fields(1).value)
'            ElseIf rsListR.Fields(0).value = "SM" Then
'               valSM = 1
'               .chSMS.value = valSM
'               .nSMS.Caption = Trim(rsListR.Fields(1).value)
'            End If
'            rsListR.MoveNext
'         Loop
'      End If
'   End With
'   Dim rsCallR As ADODB.Recordset
'   Set rsCallR = objRem.FetchListRemainder(txtNoWO.Text, txtNoChassis.Text, NulA(txtKmNext.Text), True)
'   If rsCallR.RecordCount <> 0 Then
'      valCl = 1
'      chCall.value = valCl
'      nCall.Caption = Trim(rsCallR.Fields(0).value) & "  - Call : " & rsCallR.RecordCount
'   End If
End Sub


Private Function AddToListRemainderSMS(tNoWO As String, tNoChassis As String, tKmNext As Double) As Boolean
On Error Resume Next
   AddToListRemainderSMS = False
   Dim rsListR As ADODB.Recordset
   Set rsListR = objRem.FetchListRemainder(tNoWO, tNoChassis, tKmNext, False)
   With Me
      If rsListR.RecordCount <> 0 Then
         rsListR.MoveFirst
         Do Until rsListR.EOF
            If rsListR.Fields(0).value = "SM" Then
               AddToListRemainderSMS = True
            End If
            rsListR.MoveNext
         Loop
      End If
   End With
End Function

Private Function AddToListRemainderDM(tNoWO As String, tNoChassis As String, tKmNext As Double) As Boolean
On Error Resume Next
   AddToListRemainderDM = False
   Dim rsListR As ADODB.Recordset
   Set rsListR = objRem.FetchListRemainder(tNoWO, tNoChassis, tKmNext, False)
   With Me
      If rsListR.RecordCount <> 0 Then
         rsListR.MoveFirst
         Do Until rsListR.EOF
            If rsListR.Fields(0).value = "DM" Then
               AddToListRemainderDM = True
            End If
            rsListR.MoveNext
         Loop
      End If
      
   End With
End Function

Private Function AddToListRemainderCall(tNoWO As String, tNoChassis As String, tKmNext As Double) As Boolean
    AddToListRemainderCall = False
    Dim rsCallR As ADODB.Recordset
    Set rsCallR = objRem.FetchListRemainder(tNoWO, tNoChassis, tKmNext, True)
    If rsCallR.RecordCount <> 0 Then
        AddToListRemainderCall = True
    End If
End Function

Private Sub ListRemainderValue()
    chDM.value = valDM
    chSMS.value = valSM
    chCall.value = valCl
End Sub


Private Sub txtNoChassis_Change()
'Dim rsSPK As ADODB.Recordset
'Dim objSPK As New ClsRemainder
'Set rsSPK = objSPK.FetchDataSPK_Item(Trim(txtNoChassis.Text))
'With Me
'   If rsSPK.RecordCount = 0 Then
'      .DtDO.Text = "__/__/____"
'      .DtDEC.Text = "__/__/____"
'      .DTCR1.Text = "__/__/____"
'      .DTCR2.Text = "__/__/____"
'      .DTLahir.Text = "__/__/____"
'      .txtKMCR1.Text = ""
'      .txtKMCR2.Text = ""
'      .txtNamaS.Text = ""
'      .txtColorS.Text = ""
'      .txtProgramS.Text = ""
'      .txtTempatLahir.Text = ""
'      .txtAgama.Text = ""
'      .txtHoby.Text = ""
'      .txtPekerjaan.Text = ""
'   Else
'      .DtDO.value = Format(rsSPK.Fields(0).value, "DD/MM/YYYY")
'      .DtDEC.Text = Format(rsSPK.Fields(1).value, "DD/MM/YYYY")
'      .DTCR1.Text = Format(rsSPK.Fields(2).value, "DD/MM/YYYY")
'      .DTCR2.Text = Format(rsSPK.Fields(4).value, "DD/MM/YYYY")
'      .DTLahir.Text = Format(rsSPK.Fields(10).value, "DD/MM/YYYY")
'      .txtKMCR1.Text = NullText(rsSPK.Fields(3).value)
'      .txtKMCR2.Text = NullText(rsSPK.Fields(5).value)
'      .txtNamaS.Text = NullText(rsSPK.Fields(6).value)
'      .txtColorS.Text = NullText(rsSPK.Fields(7).value)
'      .txtProgramS.Text = IIf(NullText(rsSPK.Fields(8).value) = "LAINNYA", NullText(rsSPK.Fields(8).value), NullText(rsSPK.Fields(14).value))
'      .txtTempatLahir.Text = NullText(rsSPK.Fields(9).value)
'      .txtAgama.Text = NullText(rsSPK.Fields(11).value)
'      .txtPekerjaan.Text = NullText(rsSPK.Fields(12).value)
'      .txtHoby.Text = NullText(rsSPK.Fields(13).value)
'   End If
'End With
End Sub

Private Sub GridRemainder_FilterChange()
On Error GoTo err
   Set cols = GridRemainder.Columns
   Dim C As Integer
   
   C = GridRemainder.col
   GridRemainder.HoldFields
   rsList.filter = getFilter()
   If rsList.EOF = True Then
      For Each col In GridRemainder.Columns
         col.FilterText = ""
      Next col
      rsList.filter = ""
   End If
   GridRemainder.col = C
   GridRemainder.EditActive = True
   lblJml.Caption = rsList.RecordCount
   Exit Sub
err:
   For Each col In cols
      col.FilterText = ""
   Next col
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

Private Sub GridRemainder_HeadClick(ByVal ColIndex As Integer)
1     On Error GoTo err
2        Urut = IIf(Urut = "DESC", "ASC", "DESC")
3        If ColIndex <= 14 Then rsList.Sort = GridRemainder.Columns(ColIndex).DataField & " " & Urut
4        GridRemainder.SetFocus
5        Sendkeys "{Right " & Val(ColIndex) & "}"
6        lblJml.Caption = rsList.RecordCount
7     Exit Sub
err:
8        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure getFilter of Form Remainder")
End Sub

Private Sub GridRemainder_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo rsKosong
   If rsList.RecordCount = 0 Then Exit Sub
   rsList.MoveFirst
   Call BindToControl(rsList)
rsKosong:
Exit Sub
End Sub

Private Sub VScrollFUB_Change()
    Me.SSFUB.Top = -(VScrollFUB.value + 200)
End Sub
Private Sub VScrollFUB_Scroll()
    Me.SSFUB.Top = -(VScrollFUB.value + 200)
End Sub


Private Sub cmdDM_Click()
On Error GoTo err
    If rsList.RecordCount = 0 Then Exit Sub
    
    If MsgBox("Via Email ?", vbYesNo + vbDefaultButton1, "Email") = vbYes Then
        If MAPISession1.SessionID <> 0 Then MAPISession1.SignOff
        MAPISession1.SignOn
        MAPIMessages1.Compose
        MAPIMessages1.RecipIndex = 1
        MAPIMessages1.RecipDisplayName = txtTipe
        MAPIMessages1.MsgSubject = "Pemberitahuan Reminder Service Berkala " & support.kopOrder(1)
        MAPIMessages1.MsgNoteText = ""
        MAPIMessages1.ResolveName
        MAPIMessages1.Send
    Else
        If cboKm.Text = "" Then
            MsgBox "Silahkan pilih KM service!", vbInformation
            Exit Sub
        Else
            GoTo viaDM
        End If
    End If
    Exit Sub

viaDM:
    Dim f As New FormTransparan
    With f
       .xRsRem = rsList
       .xIsForm = "DM"
       .dmKM = IIf(Trim(cboKm.Text) = "10 K", "10.000", "20.000")
       .show (vbModal)
    End With
    Call cmdProcess_Click
    Exit Sub
    
err:
End Sub
Private Sub cmdSMS_Click()
   If rsList.RecordCount = 0 Then Exit Sub
'   If chSMS.value = 1 Then
'            Dim Er As New ErrInfo
'            Er.HeadCap = "Reminder - SMS !"
'            Er.InfoCap = "Reminder SMS sudah pernah di proses !"
'
'            Beep
'            Er.Show (vbModal)
'      'MsgBox "Remainder - SMS sudah pernah di proses !", vbExclamation, "SMS Remainder"
'      Exit Sub
'   End If

Dim ListSMS As String
Call TambahSMS
    Dim rsInTem As ADODB.Recordset
    Set rsInTem = objRem.GetDataSMS

    If rsInTem.RecordCount = 0 Then
        MsgBox "Data SMS Masih Kosong !", vbInformation + vbOKOnly, "SMS"
        Exit Sub
    End If
    Dim objPul As New ClsRemainder
    rsInTem.MoveFirst
    Do Until rsInTem.EOF
        With objRem
            .NoSMS = NewId
            ListSMS = ListSMS & rsInTem.Fields(0).value
            .pesan = Trim(rsInTem.Fields(4).value)
            If rsInTem.Fields(1).value <> "" Then
                .Tujuan = Trim(rsInTem.Fields(1).value)
                ListSMS = ListSMS & " -  No.HP : " & rsInTem.Fields(1).value
                'Call objRem.AddNew_LogRemainder("SM", rsInTem.Fields(0).value, objRem.MyNoWo(rsInTem.Fields(0).value), TglNext, Val(txtKmNext), rsInTem.Fields(4).value, rsInTem.Fields(1).value, 0, 0, 0, 0, 0)
                Call objRem.AddNew_LogRemainder("SM", rsInTem.Fields(0).value, objRem.MyNoWo(rsInTem.Fields(0).value), objRem.MyTglNext(rsInTem.Fields(0).value), Val(objRem.MyKmNext(rsInTem.Fields(0).value)), rsInTem.Fields(4).value, rsInTem.Fields(1).value, 0, 0, 0, 0, 0)
            End If
            If rsInTem.Fields(2).value <> "" Then
                .Tujuan = Trim(rsInTem.Fields(2).value)
                ListSMS = ListSMS & " -  No.HP : " & rsInTem.Fields(2).value
                'Call objRem.AddNew_LogRemainder("SM", rsInTem.Fields(0).value, objRem.MyNoWo(rsInTem.Fields(0).value), TglNext, Val(txtKmNext), rsInTem.Fields(4).value, rsInTem.Fields(2).value, 0, 0, 0, 0, 0)
                Call objRem.AddNew_LogRemainder("SM", rsInTem.Fields(0).value, objRem.MyNoWo(rsInTem.Fields(0).value), objRem.MyTglNext(rsInTem.Fields(0).value), Val(objRem.MyKmNext(rsInTem.Fields(0).value)), rsInTem.Fields(4).value, rsInTem.Fields(2).value, 0, 0, 0, 0, 0)
            End If
            .waktu = isnow
            .AddNewItemSMS
        End With

        rsInTem.MoveNext
        NewId = Val(NewId) + 1
        ListSMS = ListSMS & Chr(13)
    Loop
    MsgBox "Data List SMS : " & Chr(13) & "----------------------------------------------------------------------" & Chr(13) & ListSMS & Chr(13) & "----------------------------------------------------------------------"
    MsgBox "Data SMS telah dikirim !"

    objRem.RemoveAllItemSMS
    Call cmdProcess_Click
    
    
    
'   If rsRem.RecordCount = 0 Then Exit Sub
'   If chSMS.value = 1 Then
'        Dim Er As New ErrInfo
'        Er.HeadCap = "Reminder SMS !"
'        Er.InfoCap = "Reminder SMS sudah pernah di proses !"""
'        Beep
'        Er.Show (vbModal)
'        Exit Sub
'      'MsgBox "Remainder - SMS sudah pernah di proses !", vbExclamation, "SMS Remainder"
'      'Exit Sub
'   End If
'   Dim f As New FormTransparan
'   With f
'      .xRsRem = rsList
'      .xIsForm = "SMS"
'      .Top = 2250
'      .Left = Me.Left
'      .Show (vbModal)
'   End With
'   Call cmdProcess_Click
End Sub

Private Sub TambahSMS()
    Dim rsSMS As ADODB.Recordset
    Set rsSMS = objRem.GetDataOutbox
    
    If rsSMS.RecordCount = 0 Then
        NewId = "1"
    Else
        rsSMS.MoveLast
        NewId = Val(rsSMS.Fields(0).value) + 1
    End If
End Sub

Private Sub cmdCall_Click()
    On Error GoTo err
   If rsList.RecordCount = 0 Then Exit Sub
   sSQL = "INSERT INTO [M_ListReminderCall] ([LastNoWO],[TglFollow],[RangeFollow],[NoRangka],[DateNextSvc],[KmNextSvc],[K_Telp],[Pesan_Telp],[K_Ulang],[Tanggal],[Jam],[K_Pel],[K_Ken],[K_DM],[K_SMS],[EtaHargaJasa],[EtaHargaPart],[EtaJam],[EtaMenit],[EditPel],[App],[Pesan_App],[Pesan],[TglEntry],[dealer],[NIK],[KDUser],[TipeUser],[reminder_ID])"
   sSQL = sSQL & " VALUES ('',GETDATE(),'0','" & txtNoChassis.Text & "','" & Format(NulTgl(txtDateNext.Text), "yyyy-mm-dd") & "','0','1','',"
   sSQL = sSQL & " '0',CONVERT(VARCHAR(10), GETDATE(), 120),CONVERT(VARCHAR(10), GETDATE(), 108),'0',"
   sSQL = sSQL & " '0','0','',0,0,0,0,0,1,'" & Nul(txtReminder.Text) & "','',GETDATE(),"
   sSQL = sSQL & " '" & Right(gDb, 2) & "','" & gNIK & "','" & gUser & "','','" & txtReminderID.Text & "')"
   Call objO.InsertData(sSQL)
   MsgBox "Data berhasil disimpan", vbInformation
   
   Call cmdProcess_Click
err:
End Sub

Private Sub CreateMainList(Optional Edit As Boolean = False)
    On Error GoTo vb_error
    If rsList.State Then rsList.Close
    Set rsList = Nothing
    With rsList
        .Fields.Append "outlet_ID", adVarChar, 200
        .Fields.Append "one_account_ID", adVarChar, 200
        .Fields.Append "dealer_customer_ID", adVarChar, 200
        .Fields.Append "email", adVarChar, 200
        .Fields.Append "preferred_contact_channel", adVarChar, 200
        .Fields.Append "reminder_ID", adVarChar, 200
        .Fields.Append "activity", adVarChar, 200
        .Fields.Append "activity_plan_scheduled_date", adVarChar, 200
        .Fields.Append "auto_reminder_status", adVarChar, 200
        .Fields.Append "priority_call", adVarChar, 200
        .Fields.Append "extended_warranty_status", adVarChar, 200
        .Fields.Append "customer_habit", adVarChar, 200
        .Fields.Append "last_habit", adVarChar, 200
        .Fields.Append "next_service_status", adVarChar, 200
        .Fields.Append "last_service_date", adVarChar, 200
        .Fields.Append "next_service_date", adVarChar, 200
        .Fields.Append "ncs_status", adVarChar, 200
        .Fields.Append "katashiki_suffix", adVarChar, 200
        .Fields.Append "vin", adVarChar, 50
        .Fields.Append "model", adVarChar, 150
        .Fields.Append "variant", adVarChar, 150
        .Fields.Append "fullname", adVarChar, 200
        .Fields.Append "police_number", adVarChar, 25
        .Fields.Append "phone_number", adVarChar, 50
        .Fields.Append "color", adVarChar, 100
        .Fields.Append "program_tab", adVarChar, 150
        .Fields.Append "next_service_stage", adVarChar, 50
        .Fields.Append "reminder_message", adVarChar, 500
        .Open
    End With

    Me.GridRemainder.DataSource = rsList
    Exit Sub
vb_error:
    err.Raise err.Number, err.Source, err.Description
End Sub
