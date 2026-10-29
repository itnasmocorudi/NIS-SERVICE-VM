VERSION 5.00
Begin VB.Form FormTransparan 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   8355
   ClientLeft      =   0
   ClientTop       =   4005
   ClientWidth     =   12225
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   12225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image Image1 
      Height          =   1665
      Left            =   -120
      Picture         =   "FormTransparan.frx":0000
      Top             =   6600
      Width           =   4500
   End
End
Attribute VB_Name = "FormTransparan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================
'Motion Captured Data Provided By Acep Suryanto
'==============================================
Option Explicit
Private isForm As String
Private isTgl As Date
' === = ====================
Private rsRem As ADODB.Recordset
Private rsAll As ADODB.Recordset
Private isChassis As String
Private isNoPol   As String
Private isKM      As String
' === = ====================

Private Declare Function GetWindowLong Lib "user32.dll" Alias _
"GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias _
"SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal _
dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Const GWL_EXSTYLE As Long = -20
Private Const WS_EX_LAYERED As Long = &H80000
Private Const LWA_ALPHA As Long = &H2

Private xMe As Boolean
Private isSBI As Boolean
Private baris As Integer
Private modelKendaraan As String


Sub AlphaBlend(WindowHandle As Long, Opacity As Long)
    
    If Opacity < 0 Then
        Opacity = 0
    ElseIf Opacity > 100 Then
        Opacity = 100
    End If
    
    If Opacity = 100 Then
        SetWindowLong WindowHandle, GWL_EXSTYLE, GetWindowLong(WindowHandle, GWL_EXSTYLE) And Not WS_EX_LAYERED
    ElseIf Opacity > 0 Then
        SetWindowLong WindowHandle, GWL_EXSTYLE, GetWindowLong(WindowHandle, GWL_EXSTYLE) Or WS_EX_LAYERED
        SetLayeredWindowAttributes WindowHandle, 0, Int(Opacity * 255 / 100), LWA_ALPHA
    End If
End Sub

Private Sub Form_Load()
    AlphaBlend hWnd, 75
End Sub

Public Property Let isMe(value As Boolean)
   xMe = value
End Property

Public Property Let xIsForm(value As String)
   isForm = value
End Property

Public Property Let xIsTgl(value As String)
   isTgl = value
End Property

Public Property Let dmKM(value As String)
   isKM = value
End Property

Public Property Let xBaris(value As Integer)
    baris = value
End Property

' === = ====================
Public Property Let xRsRem(value As ADODB.Recordset)
   Set rsRem = value
End Property

Public Property Let xRsAll(value As ADODB.Recordset)
   Set rsAll = value
End Property

Public Property Let xChassis(value As String)
   isChassis = value
End Property

Public Property Let xNoPol(value As String)
   isNoPol = value
End Property

Public Property Let xSBI_1K(value As Boolean)
   isSBI = value
End Property

Public Property Let xModelKendaraan(value As String)
    modelKendaraan = value
End Property

' === = ====================

Private Sub Form_Activate()
    Call ViewForm
End Sub

Private Sub ViewForm()
   If isForm = "DM" Then
      Dim fDM As New RemainderDM
      With fDM
         .xRsRem = rsRem
         .xSBI = isSBI
         .dataKM = isKM
         .show vbModal, Me
         Unload Me
      End With
   ElseIf isForm = "SMS" Then
      Dim fSM As New RemainderSMS
      With fSM
         .xRsRem = rsRem
         .IsLoadSSC = 0
         .show vbModal, Me
         Unload Me
      End With
   ElseIf isForm = "CALLHANSARENKEI" Then
                             
      Dim fCLH As Object
      If Log.logRetention = 1 Then
        Set fCLH = New RemainderCallRetention
      Else
        Set fCLH = New RemainderCall
      End If
      'Me.Left = 0
      'Me.Top = 750
      Me.Height = 8800
      
      With fCLH
         .xRsRem = rsRem
         .xRsAll = rsAll
         .xTglNext = isTgl
         .xNoPol = isNoPol
         .IsSvc1000 = False
         .IsHansaBySA = True
         '.Left = Me.Left
         '.Top = Me.Top
         .show vbModal, Me
         Unload Me
      End With
   ElseIf isForm = "CALL" Then
                             
      Dim fCL As Object
      If Log.logRetention = 1 Then
        Set fCL = New RemainderCallRetention
      Else
        Set fCL = New RemainderCall
      End If
   ''   Me.Left = 0
      
      With fCL
         .xRsRem = rsRem
         .xRsAll = rsAll
         .xTglNext = isTgl
         .xNoPol = isNoPol
         .IsSvc1000 = False
         .IsHansaBySA = False
         '.Left = Me.Left
         '.Top = Me.Top
         .show vbModal, Me
         Unload Me
      End With
   ElseIf isForm = "CALLCR" Then
                             
      Dim fCr As Object
       
      Set fCr = New RemainderCallCR
      
      With fCr
         .xRsRem = rsRem
         .xRsAll = rsAll
         .xTglNext = isTgl
         .xNoPol = isNoPol
         .IsSvc1000 = False
         '.Left = Me.Left
         '.Top = Me.Top
         .show vbModal, Me
         Unload Me
      End With
   ElseIf isForm = "CALLLUB" Then
                             
      Dim fLub As Object
       
      Set fLub = New RemainderCallLub
      
      With fLub
         .xRsRem = rsRem
         .xRsAll = rsAll
         .xTglNext = isTgl
         .xNoPol = isNoPol
         .IsSvc1000 = False
         '.Left = Me.Left
         '.Top = Me.Top
         .show vbModal, Me
         Unload Me
      End With
   ElseIf isForm = "SMS_SSC" Then
      Dim fSMc As New RemainderSMS
      With fSMc
         .xRsRem = rsRem
         .IsLoadSSC = 1
         .show vbModal, Me
         Unload Me
      End With
   ElseIf isForm = "SMS1000" Then
      Dim fSM1 As New RemainderSMS
      With fSM1
         .xRsRem = rsRem
         .IsLoadSSC = 2
         .show vbModal, Me
         Unload Me
      End With
   ElseIf isForm = "CALL1000" Then
      Dim fCL1 As Object
      If Log.logRetention = 1 Then
        Set fCL1 = New RemainderCallRetention
      Else
        Set fCL1 = New RemainderCall
      End If
      
      With fCL1
         .xRsRem = rsRem
         .xRsAll = rsAll
         .xTglNext = isTgl
         .xNoPol = isNoPol
         .IsSvc1000 = True
         .show vbModal, Me
         Unload Me
      End With
   ElseIf isForm = "CUST_EDIT" Then
      Dim fEC As New FormEditCust
      With fEC
         .NomorRangka = Trim(isChassis)
         .isMe = xMe
         .show vbModal, Me
         Unload Me
      End With
   ElseIf isForm = "HIST" Then
        
      Dim fHI As New FormHistory
      With fHI
         .NomorRangka = Trim(isChassis)
         .isHOService = False
         .show vbModal, Me
         Unload Me
      End With
   ElseIf isForm = "HIST_HO" Then
        Dim fHiHo As New FormHistory
        With fHiHo
            .NomorRangka = Trim(isChassis)
            .isHOService = True
            .show vbModal, Me
            Unload Me
        End With
   ElseIf isForm = "HIST_BOOKING" Then
        Dim fHiBo As New FrmHistory_new
        With fHiBo
            .isOrder = False
            .noBaris = baris
            .modelKendaraan = modelKendaraan
            .show vbModal, Me
            Unload Me
        End With
   End If
End Sub
