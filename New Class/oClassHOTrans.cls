VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oClassHOTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private hoConn As New ADODB.Connection
Private bengkelConn As New ADODB.Connection
Private profilConn As New ADODB.Connection
Private accConn As New ADODB.Connection
Private spkConn As New ADODB.Connection
Public FakturConn As ADODB.Connection
Private rs As ADODB.Recordset
Private strSQL_AR As String

Private perbaikanConn As New ADODB.Connection

''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''2016-02-23 || perubahan msgbox setelah proses berakhir / rollback
''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Public Sub OpenTransactionHO()
   On Error GoTo OpenTransactionHO_Error
   ' If Log.LogHOConn = 1 Then
        If hoConn.State = adStateClosed Then hoConn.Open (SQLConHO)
        hoConn.Execute ("set xact_abort ON")
        hoConn.BeginTrans
  '  End If
   
   Exit Sub

OpenTransactionHO_Error:
    Log.LogHOConn = 0
End Sub

Public Sub CloseTransactionHO()
    hoConn.CommitTrans
    hoConn.Execute ("set xact_abort OFF")
    If hoConn.State = adStateOpen Then hoConn.Close
End Sub

Public Sub CloseTransactionImportHO()
    hoConn.CommitTrans
    hoConn.Execute ("set xact_abort OFF")
    
    
    bengkelConn.CommitTrans
    bengkelConn.Execute ("set xact_abort OFF")
    
    If hoConn.State = adStateOpen Then hoConn.Close
    If bengkelConn.State = adStateOpen Then bengkelConn.Close
End Sub

Public Sub CloseTransactionHO_err()
    'On Error Resume Next
    hoConn.RollbackTrans
    hoConn.Execute ("set xact_abort OFF")
    If hoConn.State = adStateOpen Then hoConn.Close
End Sub

Public Sub TransDataHO(strSql As String)
      Dim cm As New Command
1     On Error GoTo errHandler
2         cm.CommandText = strSql
3         cm.CommandType = adCmdText
4         Set cm.ActiveConnection = hoConn
5         cm.Execute
6         Set cm = Nothing
7         Exit Sub
errHandler:
8         erTranHO = True
9         hoConn.RollbackTrans
10        If Not cm Is Nothing Then Set cm = Nothing
11        If hoConn.State = adStateOpen Then hoConn.Close
12        err.Raise err.Number, err.Source, err.Description
End Sub

Public Sub OpenTransactionImportHO()
    If hoConn.State = adStateClosed Then hoConn.Open (SQLConHO)
    If bengkelConn.State = adStateClosed Then bengkelConn.Open (SQLCONBengkel)
    
    hoConn.Execute ("set xact_abort ON")
    hoConn.BeginTrans
    
    bengkelConn.Execute ("set xact_abort ON")
    bengkelConn.BeginTrans
End Sub

Public Sub OpenTransactionAll()
    Set FakturConn = New ADODB.Connection
'    If Log.LogHOConn = 1 Then
'        If hoConn.State = adStateClosed Then hoConn.Open (SQLConHO)
'    End If
    If bengkelConn.State = adStateClosed Then bengkelConn.Open (SQLCONBengkel)
    If profilConn.State = adStateClosed Then profilConn.Open (SQLConProfilBK)
    If accConn.State = adStateClosed Then accConn.Open (SQLConACC)
    If spkConn.State = adStateClosed Then spkConn.Open (SQLConSPK)
    If FakturConn.State = adStateClosed Then FakturConn.Open (SQLConFP)
    
'    If Log.LogHOConn = 1 Then
'        hoConn.Execute ("set xact_abort ON")
'        hoConn.BeginTrans
'    End If
    bengkelConn.Execute ("set xact_abort ON")
    bengkelConn.BeginTrans
    
    profilConn.Execute ("set xact_abort ON")
    profilConn.BeginTrans
    
    accConn.Execute ("set xact_abort ON")
    accConn.BeginTrans
    
    spkConn.Execute ("set xact_abort ON")
    spkConn.BeginTrans
    
    FakturConn.Execute ("set xact_abort ON")
    FakturConn.BeginTrans
End Sub

Public Sub CloseTransactionAll()
'    If Log.LogHOConn = 1 Then
'        hoConn.CommitTrans
'        hoConn.Execute ("set xact_abort OFF")
'    End If
    bengkelConn.CommitTrans
    bengkelConn.Execute ("set xact_abort OFF")
    
    profilConn.Execute ("set xact_abort OFF")
    profilConn.CommitTrans
    
    accConn.Execute ("set xact_abort OFF")
    accConn.CommitTrans
    
    spkConn.Execute ("set xact_abort OFF")
    spkConn.CommitTrans
    
    FakturConn.Execute ("set xact_abort OFF")
    FakturConn.CommitTrans
'    If Log.LogHOConn = 1 Then
'        If hoConn.State = adStateOpen Then hoConn.Close
'    End If
    If bengkelConn.State = adStateOpen Then bengkelConn.Close
    If profilConn.State = adStateOpen Then profilConn.Close
    If accConn.State = adStateOpen Then accConn.Close
    If spkConn.State = adStateOpen Then spkConn.Close
    If FakturConn.State = adStateOpen Then FakturConn.Close
End Sub

Public Sub CloseTransactionAll_Err()
    On Error Resume Next
'    If Log.LogHOConn = 1 Then
'        hoConn.RollbackTrans
'        hoConn.Execute ("set xact_abort OFF")
'    End If
    bengkelConn.RollbackTrans
    bengkelConn.Execute ("set xact_abort OFF")
    
    profilConn.RollbackTrans
    profilConn.Execute ("set xact_abort OFF")
        
    accConn.RollbackTrans
    accConn.Execute ("set xact_abort OFF")
    
    spkConn.RollbackTrans
    spkConn.Execute ("set xact_abort OFF")
    
    FakturConn.RollbackTrans
    FakturConn.Execute ("set xact_abort OFF")
'
'    If Log.LogHOConn = 1 Then
'        If hoConn.State = adStateOpen Then hoConn.Close
'    End If
    If bengkelConn.State = adStateOpen Then bengkelConn.Close
    If profilConn.State = adStateOpen Then profilConn.Close
    If accConn.State = adStateOpen Then accConn.Close
    If spkConn.State = adStateOpen Then spkConn.Close
    If FakturConn.State = adStateOpen Then FakturConn.Close
End Sub


'HO

Public Sub UpdateData_Ho(strSql As String)
      Dim cm As New Command
1     On Error GoTo errHandler
2         cm.CommandText = strSql
3         cm.CommandType = adCmdText
4         Set cm.ActiveConnection = hoConn
5         cm.Execute
6         Set cm = Nothing
7         Exit Sub
errHandler:
8         Call CloseTransactionAll_Err
9         erTranAll = True
10        If Not cm Is Nothing Then Set cm = Nothing
11            err.Raise err.Number, err.Source, err.Description
              
End Sub

Public Sub InsertData_Ho(strSql As String)
      Dim cm As New Command
1     On Error GoTo errHandler
2         cm.CommandText = strSql
3         cm.CommandType = adCmdText
4         Set cm.ActiveConnection = hoConn
5         cm.Execute
6         Set cm = Nothing
7         Exit Sub
errHandler:
8         Call CloseTransactionHO_err
9         erTranAll = True
10        If Not cm Is Nothing Then Set cm = Nothing
11            err.Raise err.Number, err.Source, err.Description
              
End Sub

Public Function FetchDataHO(strSql As String) As ADODB.Recordset
1         On Error GoTo errHandler
2         Set rs = New ADODB.Recordset
3         rs.CursorLocation = adUseClient
4         rs.Open strSql, hoConn, adOpenStatic, adLockBatchOptimistic, adCmdText
5         Set rs.ActiveConnection = Nothing
6         Set FetchDataHO = rs
7         Exit Function
errHandler:
8         Call CloseTransactionAll_Err
9         erTranAll = True
10            err.Raise err.Number, err.Source, err.Description
              
End Function

Public Function FetchDataHO_String(strSql As String) As String
1         On Error GoTo errHandler
2         Set rs = New ADODB.Recordset
3         rs.CursorLocation = adUseClient
4         rs.Open strSql, hoConn, adOpenStatic, adLockBatchOptimistic, adCmdText
5         Set rs.ActiveConnection = Nothing
6         If rs.RecordCount > 0 Then
7              FetchDataHO_String = Nul(rs.Fields(0).value)
8         Else
9              FetchDataHO_String = ""
10        End If
          
11        Exit Function
errHandler:
12        Call CloseTransactionAll_Err
13                erTranAll = True
14            err.Raise err.Number, err.Source, err.Description

End Function

Public Sub DeleteData_Ho(strSql As String)
      Dim cm As New Command
1     On Error GoTo errHandler
2         cm.CommandText = strSql
3         cm.CommandType = adCmdText
4         Set cm.ActiveConnection = hoConn
5         cm.Execute
6         Set cm = Nothing
7         Exit Sub
errHandler:
8         Call CloseTransactionAll_Err
9         erTranAll = True
10        If Not cm Is Nothing Then Set cm = Nothing
11            err.Raise err.Number, err.Source, err.Description
              
End Sub

'BENGKEL

Public Sub UpdateData_Bengkel(strSql As String)
      Dim cm As New Command
1     On Error GoTo errHandler
     
2         cm.CommandText = strSql
3         cm.CommandType = adCmdText
4         Set cm.ActiveConnection = bengkelConn
5         cm.Execute
6         Set cm = Nothing
7         Exit Sub

errHandler:
8         Call CloseTransactionAll_Err
9         erTranAll = True
10        If Not cm Is Nothing Then Set cm = Nothing
11            err.Raise Erl, err.Source, err.Description
              
End Sub

Public Sub UpdateData_Faktur(strSql As String)
      Dim cm As New Command
1     On Error GoTo errHandler
2         cm.CommandText = strSql
3         cm.CommandType = adCmdText
4         Set cm.ActiveConnection = FakturConn
5         cm.Execute
6         Set cm = Nothing
7         Exit Sub
errHandler:
8         Call CloseTransactionAll_Err
9         erTranAll = True
10        If Not cm Is Nothing Then Set cm = Nothing
11            err.Raise err.Number, err.Source, err.Description
              
End Sub

Public Sub AddNew_Valid_AR_80(kode As String, Cur As Double, Pla As Double, Ket As String, noWO As String)
1        On Error GoTo errHandler
            Dim cm As New Command
2           cm.CommandText = "Insert Into LogAR  Values ('" & kode & "', " & Cur & ", " & Pla & ", 1 , '" & gSA & "', Getdate(), '" & noWO & "', '" & Ket & " %')"
3           cm.CommandType = adCmdText
              
4           Set cm.ActiveConnection = bengkelConn
5           cm.Execute
6           Set cm = Nothing
7           Exit Sub
errHandler:
8         Call CloseTransactionAll_Err
9         erTranAll = True
10        If Not cm Is Nothing Then Set cm = Nothing
11            err.Raise err.Number, err.Source, err.Description
              
End Sub

Public Sub Update_LogAR(kode As String, noWO As String, nKet As String)
1        On Error GoTo errHandler
            Dim cm As New Command
            
2           strSQL_AR = "Update LogAR Set Valid = 1 , NOWO = '" & noWO & "', Ket  = Ket + ' - " & nKet & "' Where KodeLgn = '" & kode & "' and valid =0"
            
3           cm.CommandText = strSQL_AR
4           cm.CommandType = adCmdText
              
5           Set cm.ActiveConnection = bengkelConn
6           cm.Execute
7           Set cm = Nothing
8           Exit Sub
errHandler:
9         Call CloseTransactionAll_Err
10        erTranAll = True
11        If Not cm Is Nothing Then Set cm = Nothing
12            err.Raise err.Number, err.Source, err.Description
              
End Sub

Public Sub Update_LogAR_BatalWO(noWO As String)
1        On Error GoTo errHandler
            Dim cm As New Command
            
2           strSQL_AR = "Update LogAR Set Valid = 0, NOWO = '',  Ket  = Left(Ket + ' - BATAL WO',100) Where NOWO = '" & noWO & "'"
            
3           cm.CommandText = strSQL_AR
4           cm.CommandType = adCmdText
              
5           Set cm.ActiveConnection = bengkelConn
6           cm.Execute
7           Set cm = Nothing
8           Exit Sub
errHandler:
9         Call CloseTransactionAll_Err
10        erTranAll = True
11        If Not cm Is Nothing Then Set cm = Nothing
12            err.Raise err.Number, err.Source, err.Description
              
End Sub

Public Sub InsertData_Bengkel(strSql As String)
      Dim cm As New Command
1     On Error GoTo errHandler
2         cm.CommandText = strSql
3         cm.CommandType = adCmdText
4         Set cm.ActiveConnection = bengkelConn
5         cm.Execute
6         Set cm = Nothing
7         Exit Sub
errHandler:
8         Call CloseTransactionAll_Err
9         erTranAll = True
10        If Not cm Is Nothing Then Set cm = Nothing
11            err.Raise Erl, err.Source, err.Description
              
End Sub


Public Sub InsertData_Faktur(strSql As String)
      Dim cm As New Command
1     On Error GoTo errHandler
2         cm.CommandText = strSql
3         cm.CommandType = adCmdText
4         Set cm.ActiveConnection = FakturConn
5         cm.Execute
6         Set cm = Nothing
7         Exit Sub
errHandler:
8         Call CloseTransactionAll_Err
9         erTranAll = True
10        If Not cm Is Nothing Then Set cm = Nothing
11            err.Raise err.Number, err.Source, err.Description
              
End Sub

Public Sub DeleteData_Bengkel(strSql As String)
      Dim cm As New Command
1     On Error GoTo errHandler
2         cm.CommandText = strSql
3         cm.CommandType = adCmdText
4         Set cm.ActiveConnection = bengkelConn
5         cm.Execute
6         Set cm = Nothing
7         Exit Sub
errHandler:
8         Call CloseTransactionAll_Err
9         erTranAll = True
10        If Not cm Is Nothing Then Set cm = Nothing
11            err.Raise Erl, err.Source, err.Description
              
End Sub

'ACC

Public Sub UpdateData_ACC(strSql As String)
      Dim cm As New Command
1     On Error GoTo errHandler
2         cm.CommandText = strSql
3         cm.CommandType = adCmdText
4         Set cm.ActiveConnection = accConn
5         cm.Execute
6         Set cm = Nothing
7         Exit Sub
errHandler:
8         Call CloseTransactionAll_Err
9         erTranAll = True
10        If Not cm Is Nothing Then Set cm = Nothing
11            err.Raise err.Number, err.Source, err.Description
              
End Sub

Public Sub InsertData_ACC(strSql As String)
      Dim cm As New Command
1     On Error GoTo errHandler
2         cm.CommandText = strSql
3         cm.CommandType = adCmdText
4         Set cm.ActiveConnection = accConn
5         cm.Execute
6         Set cm = Nothing
7         Exit Sub
errHandler:
8         Call CloseTransactionAll_Err
9         erTranAll = True
10        If Not cm Is Nothing Then Set cm = Nothing
11            err.Raise err.Number, err.Source, err.Description
              
End Sub

Public Function FetchDataBengkel_String(strSql As String) As String
1         On Error GoTo errHandler
2         Set rs = New ADODB.Recordset
3         rs.CursorLocation = adUseClient
4         rs.Open strSql, bengkelConn, adOpenStatic, adLockBatchOptimistic, adCmdText
5         Set rs.ActiveConnection = Nothing
6         If rs.RecordCount > 0 Then
7              FetchDataBengkel_String = Nul(rs.Fields(0).value)
8         Else
9              FetchDataBengkel_String = ""
10        End If
          
11        Exit Function
errHandler:
12        Call CloseTransactionAll_Err
13                erTranAll = True
14            err.Raise err.Number, err.Source, err.Description

End Function

Public Function FetchDataACC_String(strSql As String) As String
1         On Error GoTo errHandler
2         Set rs = New ADODB.Recordset
3         rs.CursorLocation = adUseClient
4         rs.Open strSql, accConn, adOpenStatic, adLockBatchOptimistic, adCmdText
5         Set rs.ActiveConnection = Nothing
6         If rs.RecordCount > 0 Then
7              FetchDataACC_String = Nul(rs.Fields(0).value)
8         Else
9              FetchDataACC_String = ""
10        End If
          
11        Exit Function
errHandler:
12        Call CloseTransactionAll_Err
13                erTranAll = True
14            err.Raise err.Number, err.Source, err.Description

End Function


Public Function FetchDataBengkel(strSql As String) As ADODB.Recordset
1         On Error GoTo errHandler
2         Set rs = New ADODB.Recordset
3         rs.CursorLocation = adUseClient
4         rs.Open strSql, bengkelConn, adOpenStatic, adLockBatchOptimistic, adCmdText
5         Set rs.ActiveConnection = Nothing
6         Set FetchDataBengkel = rs
7         Exit Function
errHandler:
8         Call CloseTransactionAll_Err
9         erTranAll = True
10            err.Raise err.Number, err.Source, err.Description
              
End Function

Public Function FetchDataACC(strSql As String) As ADODB.Recordset
1         On Error GoTo errHandler
2         Set rs = New ADODB.Recordset
3         rs.CursorLocation = adUseClient
4         rs.Open strSql, accConn, adOpenStatic, adLockBatchOptimistic, adCmdText
5         Set rs.ActiveConnection = Nothing
6         Set FetchDataACC = rs
7         Exit Function
errHandler:
8         Call CloseTransactionAll_Err
9          erTranAll = True
10            err.Raise err.Number, err.Source, err.Description
             
End Function


Public Sub DeleteData_ACC(strSql As String)
      Dim cm As New Command
1     On Error GoTo errHandler
2         cm.CommandText = strSql
3         cm.CommandType = adCmdText
4         Set cm.ActiveConnection = accConn
5         cm.Execute
6         Set cm = Nothing
7         Exit Sub
errHandler:
8         Call CloseTransactionAll_Err
9          erTranAll = True
10        If Not cm Is Nothing Then Set cm = Nothing
11            err.Raise err.Number, err.Source, err.Description
             
End Sub

'SPK

Public Sub UpdateData_SPK(strSql As String)
      Dim cm As New Command
1     On Error GoTo errHandler
2         cm.CommandText = strSql
3         cm.CommandType = adCmdText
4         Set cm.ActiveConnection = spkConn
5         cm.Execute
6         Set cm = Nothing
7         Exit Sub
errHandler:
8         Call CloseTransactionAll_Err
9          erTranAll = True
10        If Not cm Is Nothing Then Set cm = Nothing
11            err.Raise err.Number, err.Source, err.Description
             
End Sub

Public Sub InsertData_SPK(strSql As String)
      Dim cm As New Command
1     On Error GoTo errHandler
2         cm.CommandText = strSql
3         cm.CommandType = adCmdText
4         Set cm.ActiveConnection = spkConn
5         cm.Execute
6         Set cm = Nothing
7         Exit Sub
errHandler:
8         Call CloseTransactionAll_Err
9         erTranAll = True
10        If Not cm Is Nothing Then Set cm = Nothing
11            err.Raise err.Number, err.Source, err.Description
              
End Sub

Public Sub DeleteData_SPK(strSql As String)
      Dim cm As New Command
1     On Error GoTo errHandler
2         cm.CommandText = strSql
3         cm.CommandType = adCmdText
4         Set cm.ActiveConnection = spkConn
5         cm.Execute
6         Set cm = Nothing
7         Exit Sub
errHandler:
8         Call CloseTransactionAll_Err
9         erTranAll = True
10        If Not cm Is Nothing Then Set cm = Nothing
11            err.Raise err.Number, err.Source, err.Description
              
End Sub


Public Function FetchDataFP(strSql As String) As ADODB.Recordset
1         On Error GoTo errHandler
2         Set rs = New ADODB.Recordset
3         rs.CursorLocation = adUseClient
4         rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
5         Set rs.ActiveConnection = Nothing
6         Set FetchDataFP = rs
7         Exit Function
errHandler:
8         Call CloseTransactionAll_Err
9         erTranAll = True
10            err.Raise err.Number, err.Source, err.Description
              
End Function

'BENGKEL SAJA

Public Sub OpenTransactionBengkel()
    If bengkelConn.State = adStateClosed Then bengkelConn.Open (SQLCONBengkel)
    bengkelConn.Execute ("set xact_abort ON")
    bengkelConn.BeginTrans
End Sub

Public Sub CloseTransactionBengkel()
    bengkelConn.CommitTrans
    bengkelConn.Execute ("set xact_abort OFF")
    If bengkelConn.State = adStateOpen Then bengkelConn.Close
End Sub

Public Sub CloseTransactionBengkel_err()
    On Error Resume Next
    bengkelConn.RollbackTrans
    bengkelConn.Execute ("set xact_abort OFF")
    If bengkelConn.State = adStateOpen Then bengkelConn.Close
End Sub


Public Sub UpdateData_BengkelOnly(strSql As String)
      Dim cm As New Command
1     On Error GoTo errHandler
2         cm.CommandText = strSql
3         cm.CommandType = adCmdText
4         Set cm.ActiveConnection = bengkelConn
5         cm.Execute
6         Set cm = Nothing
7         Exit Sub
errHandler:
8         Call CloseTransactionBengkel_err
9         erTranAll = True
10        If Not cm Is Nothing Then Set cm = Nothing
11            err.Raise err.Number, err.Source, err.Description
              
End Sub

Public Sub InsertData_BengkelOnly(strSql As String)
      Dim cm As New Command
1     On Error GoTo errHandler
2         cm.CommandText = strSql
3         cm.CommandType = adCmdText
4         Set cm.ActiveConnection = bengkelConn
5         cm.Execute
6         Set cm = Nothing
7         Exit Sub
errHandler:
8         Call CloseTransactionBengkel_err
9           erTranAll = True
10        If Not cm Is Nothing Then Set cm = Nothing
11            err.Raise err.Number, err.Source, err.Description
            
End Sub

Public Sub DeleteData_BengkelOnly(strSql As String)
      Dim cm As New Command
1     On Error GoTo errHandler
2         cm.CommandText = strSql
3         cm.CommandType = adCmdText
4         Set cm.ActiveConnection = bengkelConn
5         cm.Execute
6         Set cm = Nothing
7         Exit Sub
errHandler:
8         Call CloseTransactionBengkel_err
9          erTranAll = True
10        If Not cm Is Nothing Then Set cm = Nothing
11            err.Raise err.Number, err.Source, err.Description
             
End Sub

'HO

'---------------------------------------------------------------------------------------
' Procedure : UpdateData_Ho_Only
' DateTime  : 23/02/2016 10:03
' Author    : RUDI KRISTANTO
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub UpdateData_Ho_Only(strSql As String)
      Dim cm As New Command
1     On Error GoTo errHandler
2         cm.CommandText = strSql
3         cm.CommandType = adCmdText
4         Set cm.ActiveConnection = hoConn
5         cm.Execute
6         Set cm = Nothing
7         Exit Sub
errHandler:
8         Call CloseTransactionHO_err
9          erTranAll = True
10        If Not cm Is Nothing Then Set cm = Nothing
11            err.Raise err.Number, err.Source, err.Description

         
12       Exit Sub

       
End Sub

Public Sub InsertData_Ho_Only(strSql As String)
      Dim cm As New Command
1     On Error GoTo errHandler
2         cm.CommandText = strSql
3         cm.CommandType = adCmdText
4         Set cm.ActiveConnection = hoConn
5         cm.Execute
6         Set cm = Nothing
7         Exit Sub
errHandler:
8         Call CloseTransactionHO_err
9          erTranAll = True
10        If Not cm Is Nothing Then Set cm = Nothing
11            err.Raise err.Number, err.Source, err.Description
             
End Sub


Public Sub DeleteData_Ho_Only(strSql As String)
      Dim cm As New Command
1     On Error GoTo errHandler
2         cm.CommandText = strSql
3         cm.CommandType = adCmdText
4         Set cm.ActiveConnection = hoConn
5         cm.Execute
6         Set cm = Nothing
7         Exit Sub
errHandler:
8         Call CloseTransactionHO_err
9           erTranAll = True
10        If Not cm Is Nothing Then Set cm = Nothing
11            err.Raise err.Number, err.Source, err.Description
            
End Sub

'PERBAIKAN DATA OLEH HO

Public Sub OpenTransactionPerbaikanDataHO(koneksi As String)
    If perbaikanConn.State = adStateClosed Then perbaikanConn.Open (koneksi)
    If hoConn.State = adStateClosed Then hoConn.Open (SQLConHO)
    hoConn.Execute ("set xact_abort ON")
    hoConn.BeginTrans
    
    perbaikanConn.Execute ("set xact_abort ON")
    perbaikanConn.BeginTrans
End Sub

Public Sub CloseTransactionPerbaikanDataHO()
    perbaikanConn.CommitTrans
    perbaikanConn.Execute ("set xact_abort OFF")
    
    hoConn.CommitTrans
    hoConn.Execute ("set xact_abort OFF")
    
    If hoConn.State = adStateOpen Then hoConn.Close
    If perbaikanConn.State = adStateOpen Then perbaikanConn.Close
End Sub

Public Sub CloseTransactionImportHO_err()
    On Error Resume Next
    hoConn.RollbackTrans
    hoConn.Execute ("set xact_abort OFF")
    
    
    bengkelConn.RollbackTrans
    bengkelConn.Execute ("set xact_abort OFF")
    
    If hoConn.State = adStateOpen Then hoConn.Close
    If bengkelConn.State = adStateOpen Then bengkelConn.Close
End Sub


Public Sub CloseTransactionPerbaikanDataHO_err()
    On Error Resume Next
    perbaikanConn.RollbackTrans
    perbaikanConn.Execute ("set xact_abort OFF")
    
    hoConn.RollbackTrans
    hoConn.Execute ("set xact_abort OFF")
    
    If hoConn.State = adStateOpen Then hoConn.Close
    If perbaikanConn.State = adStateOpen Then perbaikanConn.Close
End Sub

Public Sub UpdatePerbaikanDataHoConnectionHO(strSql As String)
      Dim cm As New Command
1     On Error GoTo errHandler
2         cm.CommandText = strSql
3         cm.CommandType = adCmdText
4         Set cm.ActiveConnection = hoConn
5         cm.Execute
6         Set cm = Nothing
7         Exit Sub
errHandler:
8         Call CloseTransactionPerbaikanDataHO_err
9          erTranAll = True
10        If Not cm Is Nothing Then Set cm = Nothing
11            err.Raise err.Number, err.Source, err.Description
             
End Sub

Public Function FetchDataPerbaikanDataHO(strSql As String) As ADODB.Recordset
1         On Error GoTo errHandler
2         Set rs = New ADODB.Recordset
3         rs.CursorLocation = adUseClient
4         rs.Open strSql, perbaikanConn, adOpenStatic, adLockBatchOptimistic, adCmdText
5         Set rs.ActiveConnection = Nothing
6         Set FetchDataPerbaikanDataHO = rs
7         Exit Function
errHandler:
8         Call CloseTransactionAll_Err
9            erTranAll = True
10            err.Raise err.Number, err.Source, err.Description
           
End Function

Public Sub UpdatePerbaikanDataHO(strSql As String)
      Dim cm As New Command
1     On Error GoTo errHandler
2         cm.CommandText = strSql
3         cm.CommandType = adCmdText
4         Set cm.ActiveConnection = perbaikanConn
5         cm.Execute
6         Set cm = Nothing
7         Exit Sub
errHandler:
8         Call CloseTransactionPerbaikanDataHO_err
9          erTranAll = True
10        If Not cm Is Nothing Then Set cm = Nothing
11            err.Raise err.Number, err.Source, err.Description
             
End Sub

Public Sub InsertDataImport_Ho(strSql As String)
      Dim cm As New Command
1     On Error GoTo errHandler
2         cm.CommandText = strSql
3         cm.CommandType = adCmdText
4         Set cm.ActiveConnection = hoConn
5         cm.Execute
6         Set cm = Nothing
7         Exit Sub
errHandler:
8         Call CloseTransactionImportHO_err
9          erTranAll = True
10        If Not cm Is Nothing Then Set cm = Nothing
11            err.Raise err.Number, err.Source, err.Description
             
End Sub

Public Sub InsertDataImport_Bengkel(strSql As String)
      Dim cm As New Command
1     On Error GoTo errHandler
2         cm.CommandText = strSql
3         cm.CommandType = adCmdText
4         Set cm.ActiveConnection = bengkelConn
5         cm.Execute
6         Set cm = Nothing
7         Exit Sub
errHandler:
8         Call CloseTransactionImportHO_err
9          erTranAll = True
10        If Not cm Is Nothing Then Set cm = Nothing
11            err.Raise err.Number, err.Source, err.Description
             
End Sub

