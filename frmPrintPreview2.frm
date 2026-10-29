VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPrintPreview 
   Caption         =   "Print Preview"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10155
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   10155
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExit 
      Caption         =   "Kembali ke &Menu Utama"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6135
      TabIndex        =   4
      Top             =   135
      Width           =   2820
   End
   Begin VB.TextBox VSearch 
      Height          =   330
      Left            =   2415
      TabIndex        =   0
      Top             =   165
      Width           =   2880
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   9330
      Top             =   45
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSearch 
      Height          =   390
      Left            =   5355
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   150
      Width           =   405
   End
   Begin VB.CommandButton cmdSetup 
      Caption         =   "&Setup"
      Height          =   570
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Setup Printer Anda ..."
      Top             =   45
      UseMaskColor    =   -1  'True
      Width           =   645
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Cetak"
      Height          =   570
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Cetak Semua Data ..."
      Top             =   45
      UseMaskColor    =   -1  'True
      Width           =   645
   End
   Begin ComctlLib.StatusBar StatBar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   5
      Top             =   5385
      Width           =   10155
      _ExtentX        =   17912
      _ExtentY        =   609
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cari Kata"
      Height          =   195
      Index           =   5
      Left            =   1665
      TabIndex        =   6
      Top             =   270
      Width           =   645
   End
End
Attribute VB_Name = "frmPrintPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FoundPos As Integer

Private Sub SetViewMode(ByVal eViewMode As ERECViewModes)
    Select Case eViewMode
        Case ercWysiwyg
            SendMessageLong txtPreview.hwnd, EM_SETTARGETDEVICE, 0, 360
        Case ercWordWrap
            SendMessageLong txtPreview.hwnd, EM_SETTARGETDEVICE, 0, 0
        Case ercDefault
            SendMessageLong txtPreview.hwnd, EM_SETTARGETDEVICE, 0, 1
    End Select
End Sub


Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Dim TextLine As String
    On Error Resume Next
    Close #1
    Close #2
    Open "c:\my documents\damiprint.txt" For Input As #1
    On Error GoTo bawah
    Open Printer.Port For Output As #2
    Do While Not EOF(1)
        Line Input #1, TextLine
        If Right(TextLine, 1) = Chr(12) Or Right(TextLine, 1) = Chr(13) Then
            Print #2, TextLine;
        Else
            Print #2, TextLine
        End If
    Loop
    Close #2
    Close #1
    Exit Sub
bawah:
    
    yn = MsgBox("Printer Tidak Dapat Dihubungi, Cek Kembali Printer Anda", vbOKOnly + vbExclamation, "Konfirmasi ?")
    On Error GoTo 0
    Close #1
End Sub

Private Sub cmdSearch_Click()
    FoundPos = txtPreview.Find(VSearch, FoundPos + Len(VSearch), , 2)
End Sub

Private Sub cmdSetup_Click()
    On Error Resume Next
    CDialog.CancelError = True
    CDialog.ShowPrinter
    On Error GoTo 0
End Sub

Private Function CekKode(xText As String) As String
    Dim Lompat As Byte
    Dim xxText As String
    Lompat = 0
    For i = 1 To Len(xText)
        If Lompat = 0 Then
            If Mid(xText, i, 1) <> Chr(27) Then
                If Mid(xText, i, 1) <> Chr(12) Then
                    xxText = xxText + Mid(xText, i, 1)
                Else
                    For j = 1 To 10
                        Print #2, ""
                    Next
                End If
            Else
                If Mid(xText, i + 1, 1) = Chr(87) Or Mid(xText, i + 1, 1) = Chr(120) Then
                    Lompat = 2
                Else
                    Lompat = 1
                End If
            End If
        Else
            Lompat = Lompat - 1
        End If
    Next
    CekKode = xxText
End Function


Private Sub Form_Load()
    frmMainMenu.Enabled = False
    Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, (SWP_NOMOVE Or SWP_NOSIZE))
    txtPreview.Left = 1
    txtPreview.Top = 670
    txtPreview.Width = ScaleWidth
    txtPreview.Height = ScaleHeight
    StatBar.Panels(1).Width = Me.Width * 0.3
    StatBar.Panels(2).Width = Me.Width * 0.7
    SetViewMode (2)
    Dim TextLine As String
    Open "c:\my documents\damiprint.txt" For Input As #1
    Open "c:\my documents\damiprint1.txt" For Output As #2
    Do While Not EOF(1)
        Line Input #1, TextLine
        TextLine = CekKode(TextLine)
        Print #2, TextLine
    Loop
    Close #1
    Close #2
    txtPreview.FileName = "c:\my documents\damiprint1.txt"
    txtPreview.Font = "MS Sans Serif"
    txtPreview.Font = "Terminal"
    txtPreview.Font.Size = 10
    txtPreview.Font.Bold = False
    txtPreview.Refresh
End Sub

Private Sub Form_Resize()
    txtPreview.Width = ScaleWidth
    txtPreview.Height = ScaleHeight - 1100
    StatBar.Panels(1).Width = Me.Width * 0.3
    StatBar.Panels(2).Width = Me.Width * 0.7
    txtPreview.Refresh
End Sub


Private Sub Form_Unload(Cancel As Integer)
    frmMainMenu.Enabled = True
End Sub

Private Sub VSearch_Change()
    FoundPos = 0
End Sub
