VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormMainMRSRetention 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   ".:: Maintenance Remainder Service - Nasmoco - Toyota ::."
   ClientHeight    =   10365
   ClientLeft      =   120
   ClientTop       =   0
   ClientWidth     =   15750
   ControlBox      =   0   'False
   Icon            =   "FormMainMRSRetention.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "FormMainMRSRetention.frx":2A8B2
   ScaleHeight     =   10365
   ScaleWidth      =   15750
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Flag Km"
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   10800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   320
      Left            =   13320
      ScaleHeight     =   315
      ScaleWidth      =   615
      TabIndex        =   1
      Top             =   120
      Width           =   615
      Begin VB.Image onClose 
         Height          =   330
         Left            =   0
         Picture         =   "FormMainMRSRetention.frx":4204E
         Top             =   0
         Width           =   630
      End
      Begin VB.Image onMini 
         Height          =   330
         Left            =   0
         Picture         =   "FormMainMRSRetention.frx":42441
         Top             =   120
         Width           =   630
      End
      Begin VB.Image onMax 
         Height          =   330
         Left            =   600
         Picture         =   "FormMainMRSRetention.frx":4275A
         Top             =   0
         Width           =   630
      End
   End
   Begin MSComctlLib.ImageList ImgBarner 
      Left            =   2880
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   1024
      ImageHeight     =   100
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMainMRSRetention.frx":42ADB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMainMRSRetention.frx":51901
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMainMRSRetention.frx":62DB7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMainMRSRetention.frx":68F83
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMainMRSRetention.frx":B3E5D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMainMRSRetention.frx":12B0CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMainMRSRetention.frx":158921
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMainMRSRetention.frx":164420
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image On15 
      Appearance      =   0  'Flat
      Height          =   540
      Left            =   0
      Picture         =   "FormMainMRSRetention.frx":18699F
      Stretch         =   -1  'True
      Top             =   8070
      Width           =   2745
   End
   Begin VB.Image On14 
      Height          =   540
      Left            =   0
      Picture         =   "FormMainMRSRetention.frx":18D64F
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   2745
   End
   Begin VB.Label Off15 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   0
      TabIndex        =   17
      Top             =   8040
      Width           =   2745
   End
   Begin VB.Label Off14 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   0
      TabIndex        =   16
      Top             =   7560
      Width           =   2745
   End
   Begin VB.Image On12 
      Height          =   540
      Left            =   0
      Picture         =   "FormMainMRSRetention.frx":19433D
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   2745
   End
   Begin VB.Image On13 
      Height          =   540
      Left            =   0
      Picture         =   "FormMainMRSRetention.frx":19730C
      Stretch         =   -1  'True
      Top             =   7020
      Width           =   2745
   End
   Begin VB.Label Off13 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   -195
      TabIndex        =   15
      Top             =   7080
      Width           =   2745
   End
   Begin VB.Label Off12 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   -30
      TabIndex        =   14
      Top             =   6430
      Width           =   2745
   End
   Begin VB.Image On11 
      Height          =   540
      Left            =   0
      Picture         =   "FormMainMRSRetention.frx":19A37E
      Stretch         =   -1  'True
      Top             =   5940
      Width           =   2745
   End
   Begin VB.Label Off11 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   -120
      TabIndex        =   13
      Top             =   5910
      Width           =   2745
   End
   Begin VB.Label lblWarning 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      Left            =   3240
      TabIndex        =   12
      Top             =   9600
      Width           =   45
   End
   Begin VB.Image Image1 
      Height          =   6915
      Left            =   4125
      Stretch         =   -1  'True
      Top             =   930
      Width           =   2880
   End
   Begin VB.Image On9 
      Height          =   540
      Left            =   0
      Picture         =   "FormMainMRSRetention.frx":19D42D
      Stretch         =   -1  'True
      Top             =   5455
      Width           =   2745
   End
   Begin VB.Image On10 
      Height          =   540
      Left            =   0
      Picture         =   "FormMainMRSRetention.frx":1A2DC7
      Stretch         =   -1  'True
      Top             =   4935
      Width           =   2730
   End
   Begin VB.Label Off10 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   -75
      TabIndex        =   10
      Top             =   4875
      Width           =   2775
   End
   Begin VB.Image On4 
      Height          =   540
      Left            =   0
      Picture         =   "FormMainMRSRetention.frx":1A84A5
      Stretch         =   -1  'True
      Top             =   3870
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.Image On7 
      Height          =   540
      Left            =   0
      Picture         =   "FormMainMRSRetention.frx":1AE04D
      Stretch         =   -1  'True
      Top             =   4410
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.Image On8 
      Height          =   540
      Left            =   0
      Picture         =   "FormMainMRSRetention.frx":1B366E
      Stretch         =   -1  'True
      Top             =   3360
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.Label Off8 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   15
      TabIndex        =   9
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label Off7 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   -30
      TabIndex        =   8
      Top             =   4365
      Width           =   2775
   End
   Begin VB.Image On6 
      Height          =   540
      Left            =   0
      Picture         =   "FormMainMRSRetention.frx":1B9127
      Stretch         =   -1  'True
      Top             =   1275
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.Image On5 
      Height          =   540
      Left            =   0
      Picture         =   "FormMainMRSRetention.frx":1BEF4E
      Stretch         =   -1  'True
      Top             =   2820
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.Image On3 
      Height          =   540
      Left            =   0
      Picture         =   "FormMainMRSRetention.frx":1C502D
      Stretch         =   -1  'True
      Top             =   2295
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.Label Off3 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   2310
      Width           =   2775
   End
   Begin VB.Image On2 
      Height          =   540
      Left            =   0
      Picture         =   "FormMainMRSRetention.frx":1CAC21
      Stretch         =   -1  'True
      Top             =   1770
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.Image On1 
      Height          =   540
      Left            =   0
      Picture         =   "FormMainMRSRetention.frx":1D09FE
      Stretch         =   -1  'True
      Top             =   780
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.Label Off2 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Line Line1 
      BorderColor     =   &H008080FF&
      X1              =   120
      X2              =   19080
      Y1              =   10320
      Y2              =   10320
   End
   Begin VB.Label Off1 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label Off4 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   -15
      TabIndex        =   5
      Top             =   3870
      Width           =   2775
   End
   Begin VB.Label Off5 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   2850
      Width           =   2775
   End
   Begin VB.Label Off6 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label Off9 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   -30
      TabIndex        =   11
      Top             =   5445
      Width           =   2730
   End
   Begin VB.Image Menu 
      Height          =   8115
      Left            =   -75
      Stretch         =   -1  'True
      Top             =   720
      Width           =   2880
   End
   Begin VB.Image MainImg 
      Height          =   825
      Left            =   15
      Top             =   -150
      Width           =   15360
   End
End
Attribute VB_Name = "FormMainMRSRetention"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================
' IT NEW RATNA MOTOR
'==============================================
Option Explicit
Private onSize As Boolean

Private f1 As New FormDataPelanggan
Private f2 As New FormRemainderService
Private f3 As New FormFollowUpBooking
Private F4 As New FormSMS
Private f5 As New FormGRPAdjustment
Private f6 As New FormReminder1000
Private f7 As New FormTransfer
Private f8 As New FormReminderSSC
Private f9 As New FormKPI
Private f10 As New FormInvalidCR
Private f11 As New FormRecoveryReminderMRA
Private f12 As New FormMRABoard
Private f13 As New FormKPIRetention
Private f14 As New FormRemainderServiceCR
Private f15 As New FormRemainderServiceLub

Private f1n As Boolean
Private f2n As Boolean
Private f3n As Boolean
Private f4n As Boolean
Private f5n As Boolean
Private f6n As Boolean
Private f7n As Boolean
Private f8n As Boolean
Private f9n As Boolean
Private f10n As Boolean
Private f11n As Boolean
Private f12n As Boolean
Private f13n As Boolean
Private f14n As Boolean
Private f15n As Boolean

Private isTgl As Date
Private MyWitdh As Double

Private Sub Command1_Click()
On Error Resume Next
Dim f As New ProsedureFlagKm
f.Show
End Sub

Private Sub Off14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetFalseMenus
    Call SetTrueItemMenu("ReminderCR")
End Sub

Private Sub Off15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetFalseMenus
    Call SetTrueItemMenu("ReminderLub")
End Sub

Private Sub onMax_Click()
On Error Resume Next
   If Me.WindowState = 0 Then
      Me.WindowState = 2
   Else
      Me.WindowState = 0
      Me.Left = 0
      Me.Top = 0
   End If
   If onSize = True Then
      f1.Left = 5100
      f2.Left = 5100
      f3.Left = 5100
      F4.Left = 5100
      f5.Left = 5100
      f6.Left = 5100
      f7.Left = 5100
      f8.Left = 5100
      f9.Left = 5100
      f10.Left = 5100
      f11.Left = 5100
      f12.Left = 5100
      f13.Left = 5100
      f14.Left = 5100
      f15.Left = 5100
   Else
      f1.Left = 3125
      f2.Left = 3125
      f3.Left = 3125
      F4.Left = 3125
      f5.Left = 3125
      f6.Left = 3125
      f7.Left = 3125
      f8.Left = 3125
      f9.Left = 3125
      f10.Left = 3125
      f11.Left = 3125
      f12.Left = 3125
      f13.Left = 3125
      f14.Left = 3125
      f15.Left = 3125
   End If
End Sub

Private Sub onMini_Click()
   Me.WindowState = 1
End Sub

Private Sub onClose_Click()

If f1n = True Then Unload f1
If f2n = True Then Unload f2
If f3n = True Then Unload f3
If f4n = True Then Unload F4
If f5n = True Then Unload f5
If f6n = True Then Unload f6
If f7n = True Then Unload f7
If f8n = True Then Unload f8
If f9n = True Then Unload f9
If f10n = True Then Unload f10
If f11n = True Then Unload f11
If f12n = True Then Unload f12
If f13n = True Then Unload f13
If f14n = True Then Unload f14
If f15n = True Then Unload f15

Unload Me
End Sub

Private Sub Form_Load()
   Call SetFalseMenus
   'Dim obj As New ClsTemplate
   'Me.lblWarning.Caption = "SBI Sms = " & obj.GetIrr_SBI(True) & " , SBI Call = " & obj.GetIrr_SBI(False) & " , SBE Sms = " & obj.GetIrr_SBE(True) & " , SBE Call = " & obj.GetIrr_SBE(False)
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
   Call SetFalseMenus
End Sub
Private Sub Form_Resize()
On Error Resume Next
Me.Width = MyWitdh
   If Me.Width > 15750 Then
      MainImg.Picture = Me.ImgBarner.ListImages(2).Picture
      Me.Picture = Me.ImgBarner.ListImages(5).Picture
      onSize = True
   Else
      MainImg.Picture = Me.ImgBarner.ListImages(1).Picture
      Me.Picture = Me.ImgBarner.ListImages(4).Picture
      onSize = False
   End If
   Menu.Picture = Me.ImgBarner.ListImages(8).Picture
   Me.Picture1.Left = Me.Width - 1000
End Sub


Private Sub Menu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
   Call SetFalseMenus
End Sub

Private Sub Off1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call SetFalseMenus
   Call SetTrueItemMenu("DataPelanggan")
End Sub
Private Sub Off2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call SetFalseMenus
   Call SetTrueItemMenu("RemainderService")
End Sub
Private Sub Off3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call SetFalseMenus
   Call SetTrueItemMenu("FollowUpBooking")
End Sub
Private Sub Off4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call SetFalseMenus
   Call SetTrueItemMenu("AplicationSMS")
End Sub
Private Sub Off5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call SetFalseMenus
   Call SetTrueItemMenu("GRPAdjustment")
End Sub
Private Sub Off6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call SetFalseMenus
   Call SetTrueItemMenu("Reminder1000")
End Sub
Private Sub Off7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call SetFalseMenus
   Call SetTrueItemMenu("CRCTransfer")
End Sub
Private Sub Off8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call SetFalseMenus
   Call SetTrueItemMenu("ReminderSSC")
End Sub
Private Sub Off9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call SetFalseMenus
   Call SetTrueItemMenu("KPI")
End Sub
Private Sub Off10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call SetFalseMenus
   Call SetTrueItemMenu("InvalidCR")
End Sub
Private Sub off11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetFalseMenus
    Call SetTrueItemMenu("RecoverReminder")
End Sub

Private Sub Off12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetFalseMenus
    Call SetTrueItemMenu("MRABoard")
End Sub
Private Sub Off13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetFalseMenus
    Call SetTrueItemMenu("KPIRetention")
End Sub

Private Sub SetFalseMenus()
On Error Resume Next
   With Me
      .On1.Visible = False
      .On2.Visible = False
      .On3.Visible = False
      .On4.Visible = False
      .On5.Visible = False
      .On6.Visible = False
      .On7.Visible = False
      .On8.Visible = False
      .On9.Visible = False
      .On10.Visible = False
      .On11.Visible = False
      .On12.Visible = False
      .On13.Visible = False
      .On14.Visible = False
      .On15.Visible = False
   End With
End Sub
Private Sub SetTrueItemMenu(nMenu As String)
On Error Resume Next
   Select Case nMenu
      Case "DataPelanggan"
         On1.Visible = True
      Case "RemainderService"
         On2.Visible = True
      Case "FollowUpBooking"
         On3.Visible = True
      Case "AplicationSMS"
         On4.Visible = True
      Case "GRPAdjustment"
         On5.Visible = True
      Case "Reminder1000"
         On6.Visible = True
      Case "CRCTransfer"
         On7.Visible = True
      Case "ReminderSSC"
         On8.Visible = True
      Case "KPI"
         On9.Visible = True
      Case "InvalidCR"
         On10.Visible = True
      Case "RecoverReminder"
         On11.Visible = True
      Case "MRABoard"
         On12.Visible = True
      Case "KPIRetention"
         On13.Visible = True
      Case "ReminderCR"
         On14.Visible = True
      Case "ReminderLub"
         On15.Visible = True
   End Select
End Sub


'============= ON CLICK MENU ===========================
Private Sub On1_Click()
On Error Resume Next

If f1n = True Then f1.Hide
If f2n = True Then f2.Hide
If f3n = True Then f3.Hide
If f4n = True Then F4.Hide
If f5n = True Then f5.Hide
If f6n = True Then f6.Hide
If f7n = True Then f7.Hide
If f8n = True Then f8.Hide
If f9n = True Then f9.Hide
If f10n = True Then f10.Hide
If f11n = True Then f11.Hide
If f12n = True Then f12.Hide
If f13n = True Then f13.Hide
If f14n = True Then f14.Hide
If f15n = True Then f15.Hide

   MyWitdh = Me.Width - 1
   f1.OnLoaded = True
   If onSize = True Then
      f1.Left = 5100
   Else
      f1.Left = 2800
   End If
   f1.Top = 700
   Me.Top = 0
   Me.Left = 0
   f1n = True
   Call SetFalseMenus
   f1.Show
End Sub

Private Sub On2_Click()
On Error Resume Next

If f1n = True Then f1.Hide
If f2n = True Then f2.Hide
If f3n = True Then f3.Hide
If f4n = True Then F4.Hide
If f5n = True Then f5.Hide
If f6n = True Then f6.Hide
If f7n = True Then f7.Hide
If f8n = True Then f8.Hide
If f9n = True Then f9.Hide
If f10n = True Then f10.Hide
If f11n = True Then f11.Hide
If f12n = True Then f12.Hide
If f13n = True Then f13.Hide
If f14n = True Then f14.Hide
If f15n = True Then f15.Hide

   MyWitdh = Me.Width - 1
   f2.OnLoaded = True
   If onSize = True Then
      f2.Left = 5100
   Else
      f2.Left = 2800
   End If
   f2.Top = 700
   Me.Top = 0
   Me.Left = 0
   f2n = True
   Call SetFalseMenus
   f2.Show
End Sub

Private Sub On3_Click()
On Error Resume Next

If f1n = True Then f1.Hide
If f2n = True Then f2.Hide
If f3n = True Then f3.Hide
If f4n = True Then F4.Hide
If f5n = True Then f5.Hide
If f6n = True Then f6.Hide
If f7n = True Then f7.Hide
If f8n = True Then f8.Hide
If f9n = True Then f9.Hide
If f10n = True Then f10.Hide
If f11n = True Then f11.Hide
If f12n = True Then f12.Hide
If f13n = True Then f13.Hide
If f14n = True Then f14.Hide
If f15n = True Then f15.Hide

   MyWitdh = Me.Width - 1
   f3.OnLoaded = True
   If onSize = True Then
      f3.Left = 5100
   Else
      f3.Left = 2800
   End If
   f3.Top = 700
   Me.Top = 0
   Me.Left = 0
   f3n = True
   Call SetFalseMenus
   f3.Show
End Sub


Private Sub On4_Click()
On Error Resume Next

If f1n = True Then f1.Hide
If f2n = True Then f2.Hide
If f3n = True Then f3.Hide
If f4n = True Then F4.Hide
If f5n = True Then f5.Hide
If f6n = True Then f6.Hide
If f7n = True Then f7.Hide
If f8n = True Then f8.Hide
If f9n = True Then f9.Hide
If f10n = True Then f10.Hide
If f11n = True Then f11.Hide
If f12n = True Then f12.Hide
If f13n = True Then f13.Hide
If f14n = True Then f14.Hide
If f15n = True Then f15.Hide

   MyWitdh = Me.Width - 1
   F4.OnLoaded = True
   If onSize = True Then
      F4.Left = 5100
   Else
      F4.Left = 2800
   End If
   F4.Top = 700
   Me.Top = 0
   Me.Left = 0
   f4n = True
   Call SetFalseMenus
   F4.Show
End Sub


Private Sub On5_Click()
On Error Resume Next

If f1n = True Then f1.Hide
If f2n = True Then f2.Hide
If f3n = True Then f3.Hide
If f4n = True Then F4.Hide
If f5n = True Then f5.Hide
If f6n = True Then f6.Hide
If f7n = True Then f7.Hide
If f8n = True Then f8.Hide
If f9n = True Then f9.Hide
If f10n = True Then f10.Hide
If f11n = True Then f11.Hide
If f12n = True Then f12.Hide
If f13n = True Then f13.Hide
If f14n = True Then f14.Hide
If f15n = True Then f15.Hide

   MyWitdh = Me.Width - 1
   f5.OnLoaded = True
   If onSize = True Then
      f5.Left = 5100
   Else
      f5.Left = 2800
   End If
   f5.Top = 700
   Me.Top = 0
   Me.Left = 0
   f5n = True
   Call SetFalseMenus
   f5.Show
End Sub


Private Sub On6_Click()
On Error Resume Next

If f1n = True Then f1.Hide
If f2n = True Then f2.Hide
If f3n = True Then f3.Hide
If f4n = True Then F4.Hide
If f5n = True Then f5.Hide
If f6n = True Then f6.Hide
If f7n = True Then f7.Hide
If f8n = True Then f8.Hide
If f9n = True Then f9.Hide
If f10n = True Then f10.Hide
If f11n = True Then f11.Hide
If f12n = True Then f12.Hide
If f13n = True Then f13.Hide
If f14n = True Then f14.Hide
If f15n = True Then f15.Hide

   MyWitdh = Me.Width - 1
   f6.OnLoaded = True
   If onSize = True Then
      f6.Left = 5100
   Else
      f6.Left = 2800
   End If
   f6.Top = 700
   Me.Top = 0
   Me.Left = 0
   f6n = True
   Call SetFalseMenus
   f6.Show
   
End Sub


Private Sub On7_Click()
On Error Resume Next

If f1n = True Then f1.Hide
If f2n = True Then f2.Hide
If f3n = True Then f3.Hide
If f4n = True Then F4.Hide
If f5n = True Then f5.Hide
If f6n = True Then f6.Hide
If f7n = True Then f7.Hide
If f8n = True Then f8.Hide
If f9n = True Then f9.Hide
If f10n = True Then f10.Hide
If f11n = True Then f11.Hide
If f12n = True Then f12.Hide
If f13n = True Then f13.Hide
If f14n = True Then f14.Hide
If f15n = True Then f15.Hide

   MyWitdh = Me.Width - 1
   f7.OnLoaded = True
   If onSize = True Then
      f7.Left = 5100
   Else
      f7.Left = 2800
   End If
   f7.Top = 700
   Me.Top = 0
   Me.Left = 0
   f7n = True
   Call SetFalseMenus
   f7.Show
End Sub



Private Sub On8_Click()
On Error Resume Next
   Dim rsW As ADODB.Recordset
   Dim objc As New ClsSSC
   Set rsW = objc.FetchDataHoWarranty
   If rsW.RecordCount = 0 Then
      MsgBox "Tidak ada data warranty SSC !", vbInformation + vbOKOnly, "Reminder SSC"
      Exit Sub
   End If
   

If f1n = True Then f1.Hide
If f2n = True Then f2.Hide
If f3n = True Then f3.Hide
If f4n = True Then F4.Hide
If f5n = True Then f5.Hide
If f6n = True Then f6.Hide
If f7n = True Then f7.Hide
If f8n = True Then f8.Hide
If f9n = True Then f9.Hide
If f10n = True Then f10.Hide
If f11n = True Then f11.Hide
If f12n = True Then f12.Hide
If f13n = True Then f13.Hide
If f14n = True Then f14.Hide
If f15n = True Then f15.Hide

   MyWitdh = Me.Width - 1
   f8.OnLoaded = True
   If onSize = True Then
      f8.Left = 5100
   Else
      f8.Left = 2800
   End If
   f8.Top = 700
   Me.Top = 0
   Me.Left = 0
   f8n = True
   Call SetFalseMenus
   f8.Show
End Sub

Private Sub On9_Click()
On Error Resume Next

If f1n = True Then f1.Hide
If f2n = True Then f2.Hide
If f3n = True Then f3.Hide
If f4n = True Then F4.Hide
If f5n = True Then f5.Hide
If f6n = True Then f6.Hide
If f7n = True Then f7.Hide
If f8n = True Then f8.Hide
If f9n = True Then f9.Hide
If f10n = True Then f10.Hide
If f11n = True Then f11.Hide
If f12n = True Then f12.Hide
If f13n = True Then f13.Hide
If f14n = True Then f14.Hide
If f15n = True Then f15.Hide

   MyWitdh = Me.Width - 1
   f9.OnLoaded = True
   If onSize = True Then
      f9.Left = 5100
   Else
      f9.Left = 2800
   End If
   f9.Top = 700
   Me.Top = 0
   Me.Left = 0
   f9n = True
   f9.Show
End Sub

Private Sub On10_Click()
On Error Resume Next

If f1n = True Then f1.Hide
If f2n = True Then f2.Hide
If f3n = True Then f3.Hide
If f4n = True Then F4.Hide
If f5n = True Then f5.Hide
If f6n = True Then f6.Hide
If f7n = True Then f7.Hide
If f8n = True Then f8.Hide
If f9n = True Then f9.Hide
If f10n = True Then f10.Hide
If f11n = True Then f11.Hide
If f12n = True Then f12.Hide
If f13n = True Then f13.Hide
If f14n = True Then f14.Hide
If f15n = True Then f15.Hide

   MyWitdh = Me.Width - 1
   f10.OnLoaded = True
   If onSize = True Then
      f10.Left = 5100
   Else
      f10.Left = 2800
   End If
   f10.Top = 700
   Me.Top = 0
   Me.Left = 0
   f10n = True
   f10.Show
End Sub

Private Sub On11_Click()
    On Error Resume Next

If f1n = True Then f1.Hide
If f2n = True Then f2.Hide
If f3n = True Then f3.Hide
If f4n = True Then F4.Hide
If f5n = True Then f5.Hide
If f6n = True Then f6.Hide
If f7n = True Then f7.Hide
If f8n = True Then f8.Hide
If f9n = True Then f9.Hide
If f10n = True Then f10.Hide
If f11n = True Then f11.Hide
If f12n = True Then f12.Hide
If f13n = True Then f13.Hide
If f14n = True Then f14.Hide
If f15n = True Then f15.Hide


   MyWitdh = Me.Width - 1
   f11.OnLoaded = True
   If onSize = True Then
      f11.Left = 5100
   Else
      f11.Left = 2800
   End If
   f11.Top = 700
   Me.Top = 0
   Me.Left = 0
   f11n = True
   f11.Show
End Sub

Private Sub On12_Click()
     On Error Resume Next

If f1n = True Then f1.Hide
If f2n = True Then f2.Hide
If f3n = True Then f3.Hide
If f4n = True Then F4.Hide
If f5n = True Then f5.Hide
If f6n = True Then f6.Hide
If f7n = True Then f7.Hide
If f8n = True Then f8.Hide
If f9n = True Then f9.Hide
If f10n = True Then f10.Hide
If f11n = True Then f11.Hide
If f12n = True Then f12.Hide
If f13n = True Then f13.Hide
If f14n = True Then f14.Hide
If f15n = True Then f15.Hide


   MyWitdh = Me.Width - 1
   '\f12.OnLoaded = True
   If onSize = True Then
      f12.Left = 2800 '5100
   Else
      f12.Left = 2800
   End If
   f12.Top = 300
   Me.Top = 0
   Me.Left = 0
  ' f12.Width = f12.Width + 3500
   f12n = True
   f12.Show
End Sub

Private Sub On13_Click()
   On Error Resume Next

If f1n = True Then f1.Hide
If f2n = True Then f2.Hide
If f3n = True Then f3.Hide
If f4n = True Then F4.Hide
If f5n = True Then f5.Hide
If f6n = True Then f6.Hide
If f7n = True Then f7.Hide
If f8n = True Then f8.Hide
If f9n = True Then f9.Hide
If f10n = True Then f10.Hide
If f11n = True Then f11.Hide
If f12n = True Then f12.Hide
If f13n = True Then f13.Hide
If f14n = True Then f14.Hide
If f15n = True Then f15.Hide


   MyWitdh = Me.Width - 1
   f13.OnLoaded = True
   If onSize = True Then
      f13.Left = 5100
   Else
      f13.Left = 2800
   End If
   f13.Top = 700
   Me.Top = 0
   Me.Left = 0
   f13n = True
   f13.Show
End Sub

Private Sub On14_Click()
   On Error Resume Next

If f1n = True Then f1.Hide
If f2n = True Then f2.Hide
If f3n = True Then f3.Hide
If f4n = True Then F4.Hide
If f5n = True Then f5.Hide
If f6n = True Then f6.Hide
If f7n = True Then f7.Hide
If f8n = True Then f8.Hide
If f9n = True Then f9.Hide
If f10n = True Then f10.Hide
If f11n = True Then f11.Hide
If f12n = True Then f12.Hide
If f13n = True Then f13.Hide
If f14n = True Then f14.Hide
If f15n = True Then f15.Hide

   If Log.CR7 = 0 Then
        MsgBox "Flag CR belum diaktifkan.", vbInformation
        Exit Sub
   End If
   MyWitdh = Me.Width - 1
   f14.OnLoaded = True
   If onSize = True Then
      f14.Left = 5100
   Else
      f14.Left = 2800
   End If
   f14.Top = 700
   Me.Top = 0
   Me.Left = 0
   f14n = True
   f14.Show
End Sub

Private Sub On15_Click()
   On Error Resume Next

If f1n = True Then f1.Hide
If f2n = True Then f2.Hide
If f3n = True Then f3.Hide
If f4n = True Then F4.Hide
If f5n = True Then f5.Hide
If f6n = True Then f6.Hide
If f7n = True Then f7.Hide
If f8n = True Then f8.Hide
If f9n = True Then f9.Hide
If f10n = True Then f10.Hide
If f11n = True Then f11.Hide
If f12n = True Then f12.Hide
If f13n = True Then f13.Hide
If f14n = True Then f14.Hide
If f15n = True Then f15.Hide

   If Log.CR7 = 0 Then
        MsgBox "Flag CR belum diaktifkan.", vbInformation
        Exit Sub
   End If
   MyWitdh = Me.Width - 1
   f15.OnLoaded = True
   If onSize = True Then
      f15.Left = 5100
   Else
      f15.Left = 2800
   End If
   f15.Top = 700
   Me.Top = 0
   Me.Left = 0
   f15n = True
   f15.Show
End Sub
