VERSION 5.00
Begin VB.Form transferHistory 
   Caption         =   "Form1"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   9750
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "transferHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oconnLama As New ADODB.Connection

Private Function Get_ConnectionLama(DBname As String, _
                                    UsrName As String, _
                                    Pass As String, _
                                    SerName As String) As Boolean
1      On Error GoTo Vb_Error
'2      App.Title = "BENGKEL-Nasmoco"
3      Get_ConnectionLama = True
4      If oconnLama.State Then oconnLama.Close
5      oconnLama.CursorLocation = adUseClient
6      oconnLama.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & UsrName & ";password=" & Pass & ";Initial Catalog=" & DBname & ";server=" & SerName
       LogonStr = oconnLama.ConnectionString
7      adoCon = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & UsrName & ";password=" & Pass & ";Initial Catalog=" & DBname & ";server=" & SerName
       oconnLama.ConnectionTimeout = 60
8      oconnLama.Open
9      Exit Function

Vb_Error:
10    Get_ConnectionLama = False
11    MsgBox ErrMessage(Erl, Err.Number, "Function : Get_Connection"), vbCritical, "Open Connection"
End Function

Private Sub Form_Load()
    Call transfer
End Sub

Private Sub transfer()
Dim rsHlp As New ADODB.Recordset
    
    If Get_ConnectionLama("Bengkelkgw", "SA", "", "THANIA") Then
        If rsHlp.State Then rsHlp.Close
        Set rsHlp = oconnLama.Execute("SELECT * FROM CUSTDAT ORDER BY NO_CHASSIS")
        If Not rsHlp.EOF Then
            Do While Not rsHlp.EOF
                
                rsHlp.MoveNext
            Loop
        End If
    End If
End Sub
