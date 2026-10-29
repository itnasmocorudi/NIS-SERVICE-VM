VERSION 5.00
Begin VB.Form ProsedureFlagKm 
   Caption         =   "Form1"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3120
   ScaleWidth      =   4680
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   720
      Width           =   855
   End
End
Attribute VB_Name = "ProsedureFlagKm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private obj As New ClsEstimasi
Private rs As ADODB.Recordset
Private i As Integer
Dim mConn As New ADODB.Connection

Private Sub Command1_Click()
i = 0
   Set rs = obj.FetchCUSTDAT
   rs.MoveFirst
   MsgBox rs.RecordCount
   
   If mConn.State = adStateClosed Then mConn.Open (SQLCon)
   Do Until rs.EOF
   i = i + 1
      mConn.Execute ("EXEC UpdateFlagKm '" & Trim(rs.Fields(0).value) & "'")
      rs.MoveNext
   Loop
   mConn.Close
   MsgBox "finish"
End Sub
