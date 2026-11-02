VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form frmOrderValidasiUpload 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   1380
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   5745
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtStr 
      Height          =   285
      Left            =   2760
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1440
      Top             =   0
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   120
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AccessType      =   1
      Protocol        =   2
      RemotePort      =   21
      URL             =   "ftp://"
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "PROSES UPLOAD..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5535
   End
End
Attribute VB_Name = "frmOrderValidasiUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sUsername As String
Public sPassword As String
Public sHost     As String
Public fileKirim As String
Public fileTujuan As String

Dim FSO          As New Scripting.FileSystemObject
'Private Declare Sub Sleep Lib "kernel32.dll" _
'    (ByVal dwMilliseconds As Long)

Private Sub Timer1_Timer()
   On Error GoTo Timer1_Timer_Error
    Timer1.Enabled = False
    Dim upldFileNm As String
    With Inet1
     .AccessType = icDirect
     .Protocol = icFTP
     .UserName = sUsername
     .Password = sPassword
     .RemotePort = 21
     .RemoteHost = sHost
     '.Execute .URL, NulSID(txtStr.text)
     '.Execute .URL, "PUT ", "C:\Users\Public\Pictures\Sample Pictures\Penguins1.jpg" & " " & "/1810221151.jpg" ' txtStr.text
     upldFileNm = FSO.GetFile(fileKirim).ShortPath
     .Execute .URL, "PUT " & Chr$(34) & upldFileNm & Chr$(34) & " " & Chr$(34) & fileTujuan & ".jpg" & Chr$(34)
    Do While .StillExecuting
       ' Sleep 1000
        DoEvents
        '.ResponseCode
    Loop
    .Execute , "CLOSE"
    End With
    Unload Me
   On Error GoTo 0
   Exit Sub

Timer1_Timer_Error:
    
    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure Timer1_Timer of Form frmOrderValidasiUpload"
    Unload Me
End Sub
