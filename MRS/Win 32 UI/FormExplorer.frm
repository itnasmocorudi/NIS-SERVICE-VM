VERSION 5.00
Object = "{19B7F2A2-1610-11D3-BF30-1AF820524153}#1.2#0"; "ccrpftv6.OCX"
Begin VB.Form FormExplorer 
   BackColor       =   &H00BFE8EA&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5745
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5070
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   5070
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1320
      TabIndex        =   9
      Top             =   4150
      Width           =   3495
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Ok"
      Height          =   732
      Left            =   1800
      Picture         =   "FormExplorer.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4800
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancel"
      Height          =   735
      Left            =   2520
      Picture         =   "FormExplorer.frx":0655
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4800
      Width           =   735
   End
   Begin VB.FileListBox MyFile 
      Height          =   2820
      Hidden          =   -1  'True
      Left            =   5280
      System          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   2655
   End
   Begin VB.DriveListBox MyDrive 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Width           =   2895
   End
   Begin CCRPFolderTV6.FolderTreeview MyFolder 
      Height          =   2220
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   3916
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama File :"
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
      Left            =   240
      TabIndex        =   8
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      X1              =   120
      X2              =   4920
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ".:: Saving And Export Data ::."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   150
      Width           =   5055
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      Height          =   975
      Left            =   1680
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      X1              =   120
      X2              =   4920
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      X1              =   120
      X2              =   4920
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   120
      X2              =   4920
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Drive :"
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
      Left            =   240
      TabIndex        =   6
      Top             =   630
      Width           =   1815
   End
   Begin VB.Label MyPath 
      BackColor       =   &H00BFE8EA&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   585
      Left            =   240
      TabIndex        =   2
      Top             =   3480
      Width           =   4575
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000040&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "FormExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================
'Motion Captured Data Provided By Acep Suryanto
'==============================================
Option Explicit

Private isDrive As String
Private isPathSave As String
Private isName As String

Public Property Get PathSaving() As String
    PathSaving = isPathSave
End Property
Public Property Get FileName() As String
    FileName = isName
End Property
Private Sub cmdCancel_Click()
    isName = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If txtName.Text = "" Then
        Dim Er As New ErrInfo
        Er.HeadCap = "Export To Excel !"
        Er.InfoCap = "Nama File Masih Kosong !"
        Beep
        Er.show (vbModal)
        Exit Sub
    End If
    If UCase(Right(txtName.Text, 4)) = ".XLS" Then txtName.Text = Mid(txtName, 1, Len(txtName.Text) - 4)
    txtName.Text = Replace(txtName.Text, "'", "")
    txtName.Text = Replace(txtName.Text, ".", "")
    txtName.Text = Replace(txtName.Text, "/", "")
    txtName.Text = Replace(txtName.Text, "\", "")
    isPathSave = MyPath.Caption
    isName = Me.txtName.Text
    Unload Me
End Sub


Private Sub Form_Load()
  Set Icon = Nothing
  Move (Screen.Width - Width) * 0.5, (Screen.Height - Height) * 0.5
  MyFolder.TabIndex = 2
  MyFolder.OverlayIcons = False
  Call MyDrive_Change
End Sub

Private Sub MyDrive_Change()
  Dim sDrive As String
  sDrive = UCase(Left(MyDrive, 2) & "\")
  If IsExplicitFolderPath(sDrive) Then
    MyFolder.RootFolder = sDrive
    isDrive = sDrive
    MyFolder.RootFolder.Selected = True
  Else
    MsgBox sDrive & " is unavailable.", vbExclamation
    MyDrive = isDrive
  End If
End Sub


Private Sub MyFolder_SelectionChange(Folder As CCRPFolderTV6.Folder, PreChange As Boolean, Cancel As Boolean)
  If PreChange Then
    With MyFolder
      If Folder = .GetSpecialFolderName(ftvDesktop) Then
        MyDrive = "c:\"
      End If
    End With
  Else
    MyFile.Path = Folder.FullPath
    If MyFile.ListCount Then
      MyFile.ListIndex = 0
    Else
      MyPath = MyFile.Path
    End If
  End If
End Sub

Private Sub MyFile_Click()
  MyPath = NormalizePath(MyFile.Path)
End Sub

Private Function NormalizePath(ByVal sPath As String) As String
  If Right(sPath, 1) = "\" Then
    NormalizePath = sPath
  Else
    NormalizePath = sPath & "\"
  End If
End Function
  
Private Function IsExplicitFolderPath(ByVal sPath As String) As Boolean
  On Error GoTo Out
  If Left$(sPath, 2) = "\\" Or Mid$(sPath, 2, 2) = ":\" Then
    IsExplicitFolderPath = Len(Dir$(NormalizePath(sPath) & "*.*", vbDirectory Or vbHidden Or vbReadOnly Or vbSystem))   ' 23
  End If
Out:
End Function


