VERSION 5.00
Begin VB.Form FormXLS 
   Caption         =   "Form1"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "FormXLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo ErrHandler
    Dim onVal As String
    onVal = "D:\Telo\XLS Exp\CustDat.XLS"
    
    
    Dim oRs As ADODB.Recordset
    Dim oRs2 As ADODB.Recordset
    Dim oCnn As ADODB.Connection
    Dim oApp As Excel.Application
    Dim oWB As Excel.Workbook
    Dim i As Integer
    
    Set oCnn = New ADODB.Connection
    oCnn.ConnectionString = SQLCon
    oCnn.Open
    
    '=========================================================================================
    Set oRs = New ADODB.Recordset
    oRs.Open "SELECT * FROM CUSTDAT", oCnn, adOpenKeyset, adLockReadOnly, adCmdText

    Set oApp = New Excel.Application
    oApp.Visible = False
    Set oWB = oApp.Workbooks.Add
    
    For i = 0 To oRs.Fields.Count - 1
        oWB.Sheets(1).Cells(1, i + 1).value = oRs.Fields(i).Name
    Next
    oWB.Sheets(1).Range("1:1").Font.Bold = True
    
    oWB.Sheets(1).Cells(2, 1).CopyFromRecordset oRs
    oWB.Sheets(1).Name = "XX"
    oRs.Close
    Set oRs = Nothing
    '=========================================================================================
    Set oRs2 = New ADODB.Recordset
    oRs2.Open "SELECT * FROM JOBDTA", oCnn, adOpenKeyset, adLockReadOnly, adCmdText
    
    
    For i = 0 To oRs2.Fields.Count - 1
        oWB.Sheets(2).Cells(1, i + 1).value = oRs2.Fields(i).Name
    Next
    oWB.Sheets(2).Range("1:1").Font.Bold = True
    
    oWB.Sheets(2).Cells(2, 1).CopyFromRecordset oRs2
    oRs2.Close
    Set oRs2 = Nothing
    '================
    oCnn.Close
    Set oCnn = Nothing
    oWB.Close SaveChanges:=True, FileName:=onVal
    
    
    
    MsgBox (onVal & ".xls")
    Set oWB = Nothing
    oApp.Quit
    Set oApp = Nothing
ErrHandler:
    Exit Sub
    
End Sub
