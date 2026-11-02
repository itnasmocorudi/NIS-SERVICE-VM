VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oClassNonTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private aConn As New ADODB.Connection
Private mConn As New ADODB.Connection
Private rs As ADODB.Recordset
Private strSql As String

Private txtAcc(90) As String
Private isID As Integer

Public Property Let nID(value As Integer)
    isID = value
End Property

Public Property Let nTxtAcc(value As String)
    txtAcc(isID) = value
End Property


Public Function FetchData(strSql As String) As ADODB.Recordset
1         On Error GoTo errHandler
          
2         If aConn.State = adStateOpen Then aConn.Close
3         aConn.ConnectionTimeout = 120
4         aConn.CommandTimeout = 120
5         If aConn.State = adStateClosed Then aConn.Open (SQLCONBengkel)
          
6         Set rs = New ADODB.Recordset
7         rs.CursorLocation = adUseClient
          
        '  rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
8         rs.Open strSql, aConn, adOpenStatic, adLockBatchOptimistic, adCmdText
9         Set rs.ActiveConnection = Nothing
10        Set FetchData = rs
11        If aConn.State = adStateOpen Then aConn.Close
12        Exit Function
errHandler:
13    If aConn.State = adStateOpen Then aConn.Close
14    MsgBox ErrMessage(Erl, err.Number, "Procedure : Ambil Data Bengkel"), vbInformation
      'err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchDataFP(strSql As String) As ADODB.Recordset
    On Error GoTo errHandler
    If aConn.State = adStateOpen Then aConn.Close
    If aConn.State = adStateClosed Then aConn.Open (SQLConFP)
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchDataFP = rs
    If aConn.State = adStateOpen Then aConn.Close
    Exit Function
errHandler:

If aConn.State = adStateOpen Then aConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchData_MEKANIK(strSql As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchData_MEKANIK = rs
    If mConn.State = adStateOpen Then mConn.Close
    Exit Function
errHandler:
If aConn.State = adStateOpen Then aConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchData_String(strSql As String) As String
    On Error GoTo errHandler
    If aConn.State = adStateClosed Then aConn.Open (SQLCONBengkel)
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    If rs.RecordCount > 0 Then
        FetchData_String = Nul(Trim(rs.Fields(0).value))
    Else
        FetchData_String = ""
    End If
    If aConn.State = adStateOpen Then aConn.Close
    Exit Function
errHandler:
If aConn.State = adStateOpen Then aConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchData_String_Acc(strSql As String) As String
    On Error GoTo errHandler
    If aConn.State = adStateClosed Then aConn.Open (SQLConACC)
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConACC, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    If rs.RecordCount > 0 Then
        FetchData_String_Acc = Nul(Trim(rs.Fields(0).value))
    Else
        FetchData_String_Acc = ""
    End If
    If aConn.State = adStateOpen Then aConn.Close
    Exit Function
errHandler:
If aConn.State = adStateOpen Then aConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchData_Double(strSql As String) As Double
1         On Error GoTo errHandler
2         If aConn.State = adStateClosed Then aConn.Open (SQLCONBengkel)
3         Set rs = New ADODB.Recordset
4         rs.CursorLocation = adUseClient
5         rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
6         Set rs.ActiveConnection = Nothing
7         If rs.RecordCount > 0 Then
8             FetchData_Double = CDbl(rs.Fields(0).value)
9         Else
10            FetchData_Double = 0
11        End If
12        If aConn.State = adStateOpen Then aConn.Close
13        Exit Function
errHandler:
14    If aConn.State = adStateOpen Then aConn.Close
15    err.Raise Erl, err.Source, err.Description
End Function


Public Function FetchData_DoubleACC(strSql As String) As Double
    On Error GoTo errHandler
    If aConn.State = adStateClosed Then aConn.Open (SQLConACC)
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConACC, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    If rs.RecordCount > 0 Then
        FetchData_DoubleACC = CDbl(rs.Fields(0).value)
    Else
        FetchData_DoubleACC = 0
    End If
    If aConn.State = adStateOpen Then aConn.Close
    Exit Function
errHandler:
If aConn.State = adStateOpen Then aConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchData_Integer(strSql As String) As Integer
    On Error GoTo errHandler
    If aConn.State = adStateClosed Then aConn.Open (SQLCONBengkel)
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    If rs.RecordCount > 0 Then
        FetchData_Integer = CInt(rs.Fields(0).value)
    Else
        FetchData_Integer = 0
    End If
    If aConn.State = adStateOpen Then aConn.Close
    Exit Function
errHandler:
If aConn.State = adStateOpen Then aConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchData_Count(strSql As String) As Integer
    On Error GoTo errHandler
    If aConn.State = adStateClosed Then aConn.Open (SQLCONBengkel)
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    FetchData_Count = CInt(rs.RecordCount)
    If aConn.State = adStateOpen Then aConn.Close
    Exit Function
errHandler:
If aConn.State = adStateOpen Then aConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchData_Boolean(strSql As String) As Boolean
    On Error GoTo errHandler
    If aConn.State = adStateClosed Then aConn.Open (SQLCONBengkel)
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    If Not rs.EOF Then
        FetchData_Boolean = True
    Else
        FetchData_Boolean = False
    End If
    If aConn.State = adStateOpen Then aConn.Close
    Exit Function
errHandler:
If aConn.State = adStateOpen Then aConn.Close
err.Raise err.Number, err.Source, err.Description
End Function



Public Sub UpdateData(strSql As String)
Dim cm As New Command
On Error GoTo errHandler
    If aConn.State = adStateClosed Then aConn.Open (SQLCONBengkel)
    cm.CommandText = strSql
    cm.CommandType = adCmdText
    Set cm.ActiveConnection = aConn
    cm.Execute
    Set cm = Nothing
    aConn.Close
    Exit Sub
errHandler:
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub

Public Sub InsertData(strSql As String)
Dim cm As New Command
On Error GoTo errHandler
    If aConn.State = adStateClosed Then aConn.Open (SQLCONBengkel)
    cm.CommandText = strSql
    cm.CommandType = adCmdText
    Set cm.ActiveConnection = aConn
    cm.Execute
    
    Set cm = Nothing
    If aConn.State = adStateOpen Then aConn.Close
    Exit Sub
errHandler:
If aConn.State = adStateOpen Then aConn.Close
    If Not cm Is Nothing Then Set cm = Nothing
              err.Raise err.Number, err.Source, err.Description
End Sub

Public Sub DeleteData(strSql As String)
Dim cm As New Command
On Error GoTo errHandler
    If aConn.State = adStateClosed Then aConn.Open (SQLCONBengkel)
    cm.CommandText = strSql
    cm.CommandType = adCmdText
    Set cm.ActiveConnection = aConn
    cm.Execute
    Set cm = Nothing
    If aConn.State = adStateOpen Then aConn.Close
    Exit Sub
errHandler:
If aConn.State = adStateOpen Then aConn.Close
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub


'===============================================================================================
'HO HO HO HO HO HO HO HO HO HO HO HO HO HO HO HO HO HO HO HO HO HO HO HO HO HO HO HO HO HO HO HO
'===============================================================================================
Public Function FetchDataHO(strSql As String) As ADODB.Recordset
1         On Error GoTo errHandler
2         If aConn.State = adStateOpen Then aConn.Close
3         aConn.ConnectionTimeout = 120
4         aConn.CommandTimeout = 120
5         If aConn.State = adStateClosed Then aConn.Open (SQLConHO)
6         Set rs = New ADODB.Recordset
7         rs.CursorLocation = adUseClient
8         rs.Open strSql, aConn, adOpenStatic, adLockBatchOptimistic, adCmdText
9         Set rs.ActiveConnection = Nothing
10        Set FetchDataHO = rs
11        If aConn.State = adStateOpen Then aConn.Close
12        Exit Function
errHandler:
13    If aConn.State = adStateOpen Then aConn.Close
14    Log.LogHOConn = 0
      'err.Raise err.Number, err.Source, err.Description
15    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure FetchDataHO of Class Module oClassNonTrans")
End Function

Public Function FetchDataHOUnlimitedTime(strSql As String) As ADODB.Recordset
1         On Error GoTo errHandler
2         If aConn.State = adStateOpen Then aConn.Close
3         aConn.ConnectionTimeout = 180
4         aConn.CommandTimeout = 180
5         If aConn.State = adStateClosed Then aConn.Open (SQLConHO)
          
6         Set rs = New ADODB.Recordset
7         rs.CursorLocation = adUseClient
          
        '  rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
8         rs.Open strSql, aConn, adOpenStatic, adLockBatchOptimistic, adCmdText
9         Set rs.ActiveConnection = Nothing
10        Set FetchDataHOUnlimitedTime = rs
11        If aConn.State = adStateOpen Then aConn.Close
12        Exit Function
errHandler:
13    If aConn.State = adStateOpen Then aConn.Close
14    err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchDataHO_NoError(strSql As String) As ADODB.Recordset
    On Error GoTo errHandler
    nHis = True
    If aConn.State = adStateClosed Then aConn.Open (SQLConHO)
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConHO, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchDataHO_NoError = rs
    If aConn.State = adStateOpen Then aConn.Close
    Exit Function
errHandler:
nHis = False
If aConn.State = adStateOpen Then aConn.Close
End Function


Public Function FetchDataHO_Double(strSql As String) As Double
1         On Error GoTo errHandler
2         If aConn.State = adStateClosed Then aConn.Open (SQLConHO)
3         Set rs = New ADODB.Recordset
4         rs.CursorLocation = adUseClient
5         rs.Open strSql, aConn, adOpenStatic, adLockBatchOptimistic, adCmdText
6         Set rs.ActiveConnection = Nothing
7         If rs.RecordCount > 0 Then
8             FetchDataHO_Double = CDbl(rs.Fields(0).value)
9         Else
10            FetchDataHO_Double = 0
11        End If
12        If aConn.State = adStateOpen Then aConn.Close
13        Exit Function
errHandler:
14    If aConn.State = adStateOpen Then aConn.Close
15    err.Raise Erl, err.Source, err.Description
End Function

Public Function FetchDataHO_String(strSql As String) As String
    On Error GoTo errHandler
    If aConn.State = adStateClosed Then aConn.Open (SQLConHO)
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, aConn, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    If rs.RecordCount > 0 Then
        FetchDataHO_String = Nul(Trim(rs.Fields(0).value))
    Else
        FetchDataHO_String = ""
    End If
    If aConn.State = adStateOpen Then aConn.Close
    Exit Function
errHandler:
If aConn.State = adStateOpen Then aConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchDataHO_Count(strSql As String) As Integer
    On Error GoTo errHandler
    If aConn.State = adStateClosed Then aConn.Open (SQLConHO)
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConHO, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    FetchDataHO_Count = CInt(rs.RecordCount)
    If aConn.State = adStateOpen Then aConn.Close
    Exit Function
errHandler:
If aConn.State = adStateOpen Then aConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchDataHO_Count_NoError(strSql As String) As Integer
    On Error GoTo errHandler
    If aConn.State = adStateClosed Then aConn.Open (SQLConHO)
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConHO, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    FetchDataHO_Count_NoError = CInt(rs.RecordCount)
    If aConn.State = adStateOpen Then aConn.Close
    Exit Function
errHandler:
If aConn.State = adStateOpen Then aConn.Close
    If aConn.State = adStateOpen Then aConn.Close
End Function

Public Sub InsertDataHO(strSql As String)
Dim cm As New Command
On Error GoTo errHandler
    If aConn.State = adStateClosed Then aConn.Open (SQLConHO)
    cm.CommandText = strSql
    cm.CommandType = adCmdText
    Set cm.ActiveConnection = aConn
    cm.Execute
    Set cm = Nothing
    If aConn.State = adStateOpen Then aConn.Close
    Exit Sub
errHandler:
If aConn.State = adStateOpen Then aConn.Close
    If Not cm Is Nothing Then Set cm = Nothing
    err.Raise err.Number, err.Source, err.Description
End Sub

Public Sub InsertDataTempKirim(strSql As String, SQLConect)
Dim cm As New Command
On Error GoTo errHandler
    If aConn.State = adStateOpen Then aConn.Close
    If aConn.State = adStateClosed Then aConn.Open (SQLConect)
    cm.CommandText = strSql
    cm.CommandType = adCmdText
    Set cm.ActiveConnection = aConn
    cm.Execute
    Set cm = Nothing
    If aConn.State = adStateOpen Then aConn.Close
    Exit Sub
errHandler:
If aConn.State = adStateOpen Then aConn.Close
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub

Public Sub UpdateDataHO(strSql As String)
Dim cm As New Command
On Error GoTo errHandler
    If aConn.State = adStateClosed Then aConn.Open (SQLConHO)
    cm.CommandText = strSql
    cm.CommandType = adCmdText
    Set cm.ActiveConnection = aConn
    cm.Execute
    Set cm = Nothing
    If aConn.State = adStateOpen Then aConn.Close
    Exit Sub
errHandler:
If aConn.State = adStateOpen Then aConn.Close
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub

Public Sub DeleteDataHO(strSql As String)
Dim cm As New Command
On Error GoTo errHandler
    If aConn.State = adStateClosed Then aConn.Open (SQLConHO)
    cm.CommandText = strSql
    cm.CommandType = adCmdText
    Set cm.ActiveConnection = aConn
    cm.Execute
    Set cm = Nothing
    If aConn.State = adStateOpen Then aConn.Close
    Exit Sub
errHandler:
If aConn.State = adStateOpen Then aConn.Close
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub

'===============================================================================================
'ACC ACC ACC ACC ACC ACC ACC ACC ACC ACC ACC ACC ACC ACC ACC ACC ACC ACC ACC ACC ACC ACC ACC ACC
'===============================================================================================

Public Function FetchData_StringACC(strSql As String) As String
    On Error GoTo errHandler
    If aConn.State = adStateClosed Then aConn.Open (SQLConACC)
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConACC, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    If rs.RecordCount > 0 Then
        FetchData_StringACC = Nul(Trim(rs.Fields(0).value))
    Else
        FetchData_StringACC = ""
    End If
    If aConn.State = adStateOpen Then aConn.Close
    Exit Function
errHandler:
If aConn.State = adStateOpen Then aConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchDataACC(strSql As String) As ADODB.Recordset
    On Error GoTo errHandler
    If aConn.State = adStateClosed Then aConn.Open (SQLConACC)
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConACC, adOpenStatic, adLockBatchOptimistic, adCmdText
    'Set rs.ActiveConnection = Nothing
    Set FetchDataACC = rs
    
    If aConn.State = adStateOpen Then aConn.Close
    Exit Function
errHandler:
If aConn.State = adStateOpen Then aConn.Close
MsgBox ErrMessage(Erl, err.Number, "Procedure : Ambil Data ACC"), vbInformation
'err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchDataACCX(strSql As String, nCon As String) As ADODB.Recordset
    On Error GoTo errHandler
    If aConn.State = adStateClosed Then aConn.Open (nCon)
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, nCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchDataACCX = rs
    If aConn.State = adStateOpen Then aConn.Close
    Exit Function
errHandler:
If aConn.State = adStateOpen Then aConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

'===============================================================================================
'SPK SPK SPK SPK SPK SPK SPK SPK SPK SPK SPK SPK SPK SPK SPK SPK SPK SPK SPK SPK SPK SPK SPK SPK
'===============================================================================================

Public Function FetchDataSPK(strSql As String) As ADODB.Recordset
    On Error GoTo errHandler
    If aConn.State = adStateClosed Then aConn.Open (SQLConSPK)
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConSPK, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchDataSPK = rs
    If aConn.State = adStateOpen Then aConn.Close
    Exit Function
errHandler:
If aConn.State = adStateOpen Then aConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchDataSPK_Double(strSql As String) As Double
    On Error GoTo errHandler
    If aConn.State = adStateClosed Then aConn.Open (SQLConSPK)
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, aConn, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    If rs.RecordCount > 0 Then
        FetchDataSPK_Double = CDbl(rs.Fields(0).value)
    Else
        FetchDataSPK_Double = 0
    End If
    If aConn.State = adStateOpen Then aConn.Close
    Exit Function
errHandler:
If aConn.State = adStateOpen Then aConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchDataSPKHO(strSql As String) As ADODB.Recordset
    On Error GoTo errHandler
    If aConn.State = adStateClosed Then aConn.Open (SQLConSPKHO)
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConSPKHO, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchDataSPKHO = rs
    If aConn.State = adStateOpen Then aConn.Close
    Exit Function
errHandler:
If aConn.State = adStateOpen Then aConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

'===============================================================================================
'LOG LOG LOG LOG LOG LOG LOG LOG LOG LOG LOG LOG LOG LOG LOG LOG LOG LOG LOG LOG LOG LOG LOG LOG
'===============================================================================================

Public Function FetchData_Profil(strSql As String) As ADODB.Recordset
    On Error GoTo errHandler
    If aConn.State = adStateClosed Then aConn.Open (SQLConProfilBK)
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConProfilBK, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchData_Profil = rs
    If aConn.State = adStateOpen Then aConn.Close
    Exit Function
errHandler:
If aConn.State = adStateOpen Then aConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

Public Sub UpdateData_Profil(strSql As String)
Dim cm As New Command
On Error GoTo errHandler
    If aConn.State = adStateClosed Then aConn.Open (SQLConProfilBK)
    cm.CommandText = strSql
    cm.CommandType = adCmdText
    Set cm.ActiveConnection = aConn
    cm.Execute
    Set cm = Nothing
    If aConn.State = adStateOpen Then aConn.Close
    Exit Sub
errHandler:
If aConn.State = adStateOpen Then aConn.Close
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub

Public Sub InsertData_Profil(strSql As String)
Dim cm As New Command
On Error GoTo errHandler
    If aConn.State = adStateClosed Then aConn.Open (SQLConProfilBK)
    cm.CommandText = strSql
    cm.CommandType = adCmdText
    Set cm.ActiveConnection = aConn
    cm.Execute
    Set cm = Nothing
If aConn.State = adStateOpen Then aConn.Close
    Exit Sub
errHandler:
If aConn.State = adStateOpen Then aConn.Close
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub

Public Sub InsertDataACC(strSql As String)
Dim cm As New Command
On Error GoTo errHandler
    If aConn.State = adStateClosed Then aConn.Open (SQLConACC)
    cm.CommandText = strSql
    cm.CommandType = adCmdText
    Set cm.ActiveConnection = aConn
    cm.Execute
    Set cm = Nothing
    If aConn.State = adStateOpen Then aConn.Close
    Exit Sub
errHandler:
    If aConn.State = adStateOpen Then aConn.Close
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub

Public Sub UpdateDataACC(strSql As String)
Dim cm As New Command
On Error GoTo errHandler
    If aConn.State = adStateClosed Then aConn.Open (SQLConACC)
    cm.CommandText = strSql
    cm.CommandType = adCmdText
    Set cm.ActiveConnection = aConn
    cm.Execute
    Set cm = Nothing
    If aConn.State = adStateOpen Then aConn.Close
    Exit Sub
errHandler:
    If aConn.State = adStateOpen Then aConn.Close
    If Not cm Is Nothing Then Set cm = Nothing
    err.Raise err.Number, err.Source, err.Description
End Sub

Public Sub DeleteData_Profil(strSql As String)
Dim cm As New Command
On Error GoTo errHandler
    If aConn.State = adStateClosed Then aConn.Open (SQLConProfilBK)
    cm.CommandText = strSql
    cm.CommandType = adCmdText
    Set cm.ActiveConnection = aConn
    cm.Execute
    Set cm = Nothing
    If aConn.State = adStateOpen Then aConn.Close
    Exit Sub
errHandler:
    If aConn.State = adStateOpen Then aConn.Close
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub

Public Function FetchDataAAS(strSql As String) As ADODB.Recordset
    On Error GoTo errHandler
    If aConn.State = adStateClosed Then aConn.Open (SQLConAAS)
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConAAS, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchDataAAS = rs
    If aConn.State = adStateOpen Then aConn.Close
    Exit Function
errHandler:
    If aConn.State = adStateOpen Then aConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

Public Sub UpdateDataAAS(strSql As String)
Dim cm As New Command
On Error GoTo errHandler
    If aConn.State = adStateClosed Then aConn.Open (SQLConAAS)
    cm.CommandText = strSql
    cm.CommandType = adCmdText
    Set cm.ActiveConnection = aConn
    cm.Execute
    Set cm = Nothing
    If aConn.State = adStateOpen Then aConn.Close
    Exit Sub
errHandler:
    If aConn.State = adStateOpen Then aConn.Close
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub

Public Sub Update_LOG(nkode As String, nStatus As String, nTipe As String)
   On Error GoTo errHandler
      If errHO = False Then
        If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
        Dim cm As New Command
        cm.CommandText = "Update LogHO Set LogHo = 1, transfer = Getdate() Where Kode = '" & nkode & "' And Status = '" & nStatus & "' And Tipe = '" & nTipe & "'"
        cm.CommandType = adCmdText
          
        Set cm.ActiveConnection = mConn
        cm.Execute
        Set cm = Nothing
        If mConn.State = adStateOpen Then mConn.Close
      End If
    Exit Sub
errHandler:
    If mConn.State = adStateOpen Then mConn.Close
    If Not cm Is Nothing Then Set cm = Nothing
    err.Raise err.Number, err.Source, err.Description
End Sub

Public Function FetchOverdueAR(KodeLg As String) As ADODB.Recordset
1         On Error GoTo errHandler
2         If mConn.State = adStateClosed Then mConn.Open (SQLConACC)
          
      '    strSql = "Select top 1 X.NilaiAR FROM " & _
      '             "(Select Sum(JumlahTRn) as NilaiAR from ARTRN WITH(NOLOCK) " & _
      '             "Where KodeLgn = '" & KodeLg & "' AND left(NoFaktur,2) = 'E-' " & _
      '             "And NoFaktur Not in (Select NoFaktur From ARTRN WITH(NOLOCK) Where KodeLgn = '" & KodeLg & "' and TglJthTmp > '" & Format(isnow, "mm/dd/YYYY") & "' ) " & _
      '             "Group By NoFaktur) X " & _
      '             "Where x.NilaiAR > 0"
      '    strSql = "Select top 1 X.NilaiAR FROM " & _
      '             " (" & _
      '             " Select Sum(JumlahTRn) as NilaiAR " & _
      '             " from ARTRN WITH(NOLOCK) " & _
      '             " Where KodeLgn = '" & KodeLg & "'  And TglJthTmp > '" & Format(isnow, "mm/dd/YYYY") & "' AND left(NoFaktur,2) = 'E-' " & _
      '             " Group By NoFaktur) X Where x.NilaiAR > 0"

         ''2016-03-16 || ganti tanda keawal
3        strSql = "Select top 1 X.NilaiAR FROM " & _
                   " (" & _
                   " Select Sum(JumlahTRn) as NilaiAR " & _
                   " from ARTRN WITH(NOLOCK) " & _
                   " Where KodeLgn = '" & KodeLg & "'  And TglJthTmp <= '" & Format(isnow, "mm/dd/YYYY") & "' AND left(NoFaktur,2) = 'E-' " & _
                   " Group By NoFaktur) X Where x.NilaiAR > 0"
          
          ''20220121 revisi mb thania terkait log ar
4         strSql = "Select top 1 x.NilaiAR , a.NoFaktur, TglJthTmp" & _
                   "   from ARTRN a WITH(NOLOCK) " & _
                   "   left join ( " & _
                   "       Select NoFaktur, Sum(JumlahTRn) as NilaiAR " & _
                   "       from ARTRN WITH(NOLOCK) " & _
                   "       Where KodeLgn = '" & KodeLg & "' AND left(NoFaktur,2) = 'E-' " & _
                   "       Group By NoFaktur " & _
                   "   ) X on x.NoFaktur = a.NoFaktur " & _
                   "   Where x.NilaiAR > 0 And a.TglJthTmp <= '" & Format(isnow, "mm/dd/YYYY") & "'"
5         Set rs = New ADODB.Recordset
6         rs.CursorLocation = adUseClient
7         rs.Open strSql, SQLConACC, adOpenStatic, adLockReadOnly, adCmdText
8         Set rs.ActiveConnection = Nothing
9         Set FetchOverdueAR = rs
10        If mConn.State = adStateOpen Then mConn.Close
            
11        Exit Function
errHandler:
12        If mConn.State = adStateOpen Then mConn.Close
13    err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchOverdueAR_Data(KodeLg As String) As ADODB.Recordset
1         On Error GoTo errHandler
2         If mConn.State = adStateClosed Then mConn.Open (SQLConACC)
      '
      '    strSql = "Select top 1 X.NilaiAR FROM " & _
      '             "(Select Sum(JumlahTRn) as NilaiAR from ARTRN WITH(NOLOCK) " & _
      '             "Where KodeLgn = '" & KodeLg & "' And TglJthTmp < '" & Format(isnow, "mm/dd/YYYY") & "' AND left(NoFaktur,2) = 'E-' " & _
      '             "And NoFaktur Not in (Select NoFaktur From ARTRN WITH(NOLOCK) Where KodeLgn = '" & KodeLg & "' and TglJthTmp > '" & Format(isnow, "mm/dd/YYYY") & "' ) " & _
      '             "Group By NoFaktur) X " & _
      '             "Where x.NilaiAR > 0"
                   
      '    strSql = "Select top 1 X.NilaiAR FROM " & _
      '             " (" & _
      '             " Select Sum(JumlahTRn) as NilaiAR " & _
      '             " from ARTRN WITH(NOLOCK) " & _
      '             " Where KodeLgn = '" & KodeLg & "'  And TglJthTmp > '" & Format(isnow, "mm/dd/YYYY") & "' AND left(NoFaktur,2) = 'E-' " & _
      '             " Group By NoFaktur) X Where x.NilaiAR > 0"

          ''2016-03-16 || ganti kembali keawal
3         strSql = "Select top 1 X.NilaiAR FROM " & _
                   " (" & _
                   " Select Sum(JumlahTRn) as NilaiAR " & _
                   " from ARTRN WITH(NOLOCK) " & _
                   " Where KodeLgn = '" & KodeLg & "'  And TglJthTmp <= '" & Format(isnow, "mm/dd/YYYY") & "' AND left(NoFaktur,2) = 'E-' " & _
                   " Group By NoFaktur) X Where x.NilaiAR > 0"
          ''20220121 revisi mb thania terkait log ar
4         strSql = "Select top 1 x.NilaiAR , a.NoFaktur, TglJthTmp" & _
                   "   from ARTRN a WITH(NOLOCK) " & _
                   "   left join ( " & _
                   "       Select NoFaktur, Sum(JumlahTRn) as NilaiAR " & _
                   "       from ARTRN WITH(NOLOCK) " & _
                   "       Where KodeLgn = '" & KodeLg & "' AND left(NoFaktur,2) = 'E-' " & _
                   "       Group By NoFaktur " & _
                   "   ) X on x.NoFaktur = a.NoFaktur " & _
                   "   Where x.NilaiAR > 0 And a.TglJthTmp <= '" & Format(isnow, "mm/dd/YYYY") & "'"
5         Set rs = New ADODB.Recordset
6         rs.CursorLocation = adUseClient
7         rs.Open strSql, SQLConACC, adOpenStatic, adLockReadOnly, adCmdText
8         Set rs.ActiveConnection = Nothing
9         Set FetchOverdueAR_Data = rs
10        If mConn.State = adStateOpen Then mConn.Close
              
11        Exit Function
errHandler:
12        If mConn.State = adStateOpen Then mConn.Close
13    err.Raise err.Number, err.Source, err.Description
End Function

Public Function Fetch_BatasKredit_Ar(nkode As String) As Double
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConACC)
    If OnTopSV = 0 Then
        strSql = "Select BatasKredit From ARMST WITH(NOLOCK) Where KodeLgn = '" & Trim(nkode) & "'"
    Else
    'UNTUK FEBUARI
        strSql = "Select BatasGRP as BatasKredit From ARMST WITH(NOLOCK) Where KodeLgn = '" & Trim(nkode) & "'"
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConACC, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    Fetch_BatasKredit_Ar = Val(rs.Fields(0).value)
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
    If mConn.State = adStateOpen Then mConn.Close
    Fetch_BatasKredit_Ar = 0
    If mConn.State = adStateOpen Then mConn.Close
End Function

Public Function Fetch_Current_Ar(nkode As String) As Double
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConACC)
    
    strSql = "Select Sum(JumlahTrn) From ARTRN WITH(NOLOCK) Where KodeLgn = '" & nkode & "' AND left(NoFaktur,2) = 'E-'"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConACC, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    Fetch_Current_Ar = Val(rs.Fields(0).value)
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
    Fetch_Current_Ar = 0
    If mConn.State = adStateOpen Then mConn.Close
End Function

'========== CEX VALIDASI AR - BY KABENG ==============================
Public Function FetchIsValid_AR_Kabeng(KodeLg As String, noWO As String) As Boolean
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    FetchIsValid_AR_Kabeng = False
    strSql = "Select * FROM LogAR WITH(NOLOCK) Where KodeLgn = '" & KodeLg & "' And (Valid =0 or nowo='" & noWO & "')"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    If rs.RecordCount > 0 Then
        FetchIsValid_AR_Kabeng = True
    Else
        FetchIsValid_AR_Kabeng = False
    End If
    If mConn.State = adStateOpen Then mConn.Close
    Exit Function
errHandler:
    If mConn.State = adStateOpen Then mConn.Close
err.Raise err.Number, err.Source, err.Description
End Function


Public Function FetchIsOpenLockAR(KodeLg As String) As Boolean
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    FetchIsOpenLockAR = False
    strSql = "Select * FROM OpenLockAR WITH(NOLOCK) Where KodeLgn = '" & KodeLg & "' And USERDEL = ''"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    If rs.RecordCount > 0 Then
        FetchIsOpenLockAR = True
    Else
        FetchIsOpenLockAR = False
    End If
    If mConn.State = adStateOpen Then mConn.Close
    Exit Function
errHandler:
    If mConn.State = adStateOpen Then mConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchData_Valid_AR() As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "Select Top 50 * from LogAR WITH(NOLOCK) order by TglLog DESC"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchData_Valid_AR = rs
    Set rs = Nothing
    If mConn.State = adStateOpen Then mConn.Close
    Exit Function
errHandler:
    If mConn.State = adStateOpen Then mConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchPelangganACC(param As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConACC)
    
    strSql = "Select a.KodeLgn, NamaLgn ,(Alamat1 + ' ' + Alamat2) as Alamat, Alamat3 as Kota, Telepon as Telp, NoFax as Fax, eMail,b.idcust FROM ARMST a WITH(NOLOCK) LEFT JOIN [singleidori]..bridging b with (nolock) on a.KodeLgn = b.KodeLgn and b.kodedealer = '" & Right(gDb, 2) & "' " & param
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConACC, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchPelangganACC = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
    If mConn.State = adStateOpen Then mConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

Public Function Fetch_Valid_AR(KodLgn As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "Select * from LogAR WITH(NOLOCK) Where KodeLgn = '" & KodLgn & "' And Valid = 0"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    Set Fetch_Valid_AR = rs
    Set rs = Nothing
    If mConn.State = adStateOpen Then mConn.Close
    Exit Function
errHandler:
    If mConn.State = adStateOpen Then mConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

Public Sub AddNew_Valid_AR(kode As String, Cur As Double, Pla As Double, Ket As String)
   On Error GoTo errHandler
      Dim cm As New Command
      If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
      cm.CommandText = "Insert Into LogAR  Values ('" & kode & "', " & Cur & ", " & Pla & ", 0 , '" & gSA & "', Getdate(), '', '" & Ket & "')"
      cm.CommandType = adCmdText
        
      Set cm.ActiveConnection = mConn
      cm.Execute
      Set cm = Nothing
      If mConn.State = adStateOpen Then mConn.Close
      Exit Sub
errHandler:
    If mConn.State = adStateOpen Then mConn.Close
    If Not cm Is Nothing Then Set cm = Nothing
err.Raise err.Number, err.Source, err.Description
End Sub



Public Function FetchIsValid_AR_Kabeng_WO(noWO As String) As Boolean
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    FetchIsValid_AR_Kabeng_WO = False
    strSql = "Select * FROM LogAR WITH(NOLOCK) Where NOWO = '" & noWO & "'"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    If rs.RecordCount > 0 Then FetchIsValid_AR_Kabeng_WO = True
    If mConn.State = adStateOpen Then mConn.Close
    Exit Function
errHandler:
    If mConn.State = adStateOpen Then mConn.Close
End Function


Public Function FetchValidasiAR_Report(tgl0 As Date, Tgl1 As Date) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)

    strSql = "Select * from LogAR WITH(NOLOCK) Where TGLLOG Between '" & TglNul2(tgl0) & " 00:00:00.001' And '" & TglNul2(Tgl1) & " 23:59:59.999'"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchValidasiAR_Report = rs
    If mConn.State = adStateOpen Then mConn.Close

    Exit Function
errHandler:
    If mConn.State = adStateOpen Then mConn.Close
err.Raise err.Number, err.Source, err.Description
End Function


Public Function FetchData_ValidasiNEAPS(nTgl As Date, nTgl1 As Date, nKri As String) As ADODB.Recordset
    On Error GoTo errHandler
    If aConn.State = adStateOpen Then aConn.Close
    If aConn.State = adStateClosed Then aConn.Open (SQLCONBengkel)
    strSql = "select A.NoWO, A.NEA, A.PS, A.Info, A.LogSA, A.TglLog, A.UserValid, A.TglValid, Replace(B.NoInv,'-A','') as NoInv, B.InvDate, B.No_polisi, B.FMDL as Model " & _
                               "FROm NEA_PS A WITH(NOLOCK) Inner Join T_invoice B WITH(NOLOCK) On a.NoWo = B.No_WO " & _
                               "Where B.FS_Log_DEL = '' and Right(B.NoInv,1) <> 'B' and (B.InvDate between '" & TglNul2(nTgl) & "' and '" & TglNul2(nTgl1) & "')"
    Select Case nKri
        Case "Data Belum Validasi"
            strSql = strSql & " And A.UserValid = ''"
        Case "Data Sudah Validasi"
            strSql = strSql & " And A.UserValid <> ''"
    End Select
    strSql = strSql & " Order By B.InvDate ASC"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchData_ValidasiNEAPS = rs
    If aConn.State = adStateOpen Then aConn.Close
    Exit Function
errHandler:
If aConn.State = adStateOpen Then aConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchData_LapNEA(nTgl As Date, nTgl1 As Date) As ADODB.Recordset
    On Error GoTo errHandler
    If aConn.State = adStateOpen Then aConn.Close
    If aConn.State = adStateClosed Then aConn.Open (SQLCONBengkel)
    
    strSql = "Select B.NOINV, (B.InvDate + B.InvTime) as Tgl_INV,A.NO_WO, A.[DATE] + A.[TIME] as Tgl_WO, A.No_Polisi, A.FMDl as TYPE, " & _
           "(A.Req1 + ', ' + A.Req2 + ', ' + A.Req3 + ', ' + A.Req4 + ', ' + A.Req5) as Problem_Keluhan, " & _
           "(A.Rep1 + ', ' + A.Rep2 + ', ' + A.Rep3 + ', ' + A.Rep4 + ', ' + A.Rep5 + ', ' + A.Rep6 + ', ' + A.Rep7) as Solusi_Masalah, " & _
           "A.FRONT as PIC, B.Labor as Jasa, (B.Part_G + B.Part_NG + B.Oil_G + B.Oil_NG) as Part, B.SBLT as Sublet, B.Total  as Revenue, " & _
           "C.Info as Remark, A.Mech,A.Customer from JobDTA A WITH(NOLOCK) Inner Join T_invoice B WITH(NOLOCK) On A.NO_WO = B.No_WO Inner Join NEA_PS C WITH(NOLOCK) On A.NO_WO = C.NOWO " & _
           "Where C.UserValid <> '' And C.NEA = 1 and (B.InvDate between '" & TglNul2(nTgl) & "' and '" & TglNul2(nTgl1) & "') " & _
           " "
    strSql = strSql & " UNION "
    strSql = strSql & "Select B.FS_NO_RETUR as NoINV,(B.FD_Retur) as Tgl_INV,A.NO_WO, A.[DATE] + A.[TIME] as Tgl_WO, A.No_Polisi, A.FMDl as TYPE, " & _
           "(A.Req1 + ', ' + A.Req2 + ', ' + A.Req3 + ', ' + A.Req4 + ', ' + A.Req5) as Problem_Keluhan, " & _
           "(A.Rep1 + ', ' + A.Rep2 + ', ' + A.Rep3 + ', ' + A.Rep4 + ', ' + A.Rep5 + ', ' + A.Rep6 + ', ' + A.Rep7) as Solusi_Masalah, " & _
           "A.FRONT as PIC, B.Labor as Jasa, (B.Part_G + B.Part_NG + B.Oil_G + B.Oil_NG) as Part, B.SBLT as Sublet, B.Total  as Revenue, " & _
           "C.Info as Remark, A.Mech,A.Customer from JobDTA A WITH(NOLOCK) Inner Join T_invoice B WITH(NOLOCK) On A.NO_WO = B.No_WO Inner Join NEA_PS C WITH(NOLOCK) On A.NO_WO = C.NOWO " & _
           "Where B.FS_LOG_DEL <> '' and C.UserValid <> '' And C.NEA = 1 and (B.FD_RETUR between '" & TglNul2(nTgl) & "' and '" & TglNul2(nTgl1) & "') " & _
           "Order By Tgl_INV, B.NOINV ASC"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchData_LapNEA = rs
    If aConn.State = adStateOpen Then aConn.Close
    Exit Function
errHandler:
If aConn.State = adStateOpen Then aConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchData_LapPS(nTgl As Date, nTgl1 As Date) As ADODB.Recordset
    On Error GoTo errHandler
    If aConn.State = adStateOpen Then aConn.Close
    If aConn.State = adStateClosed Then aConn.Open (SQLCONBengkel)
    
    strSql = "Select A.No_polisi, A.Customer, A.FMDl as Type, A.NO_WO, B.NoInv, (B.InvDate + B.InvTime) as TglInv, " & _
             "(A.Rep1 + ', ' + A.Rep2 + ', ' + A.Rep3 + ', ' + A.Rep4 + ', ' + A.Rep5 + ', ' + A.Rep6 + ', ' + A.Rep7) as Jenis_Pekerjaan, " & _
             "A.[DATE] + A.[TIME] as Tgl_Masuk, B.Labor as Jasa, (B.Part_G + B.Part_NG + B.Oil_G + B.Oil_NG) as Part, B.SBLT as Sublet, B.Total  as Revenue " & _
             "from JobDTA A WITH(NOLOCK) Inner Join T_invoice B WITH(NOLOCK) On A.NO_WO = B.No_WO Inner Join NEA_PS C WITH(NOLOCK) On A.NO_WO = c.NOWO " & _
             "Where C.UserValid <> '' And C.PS = 1 and (B.InvDate between '" & TglNul2(nTgl) & "' and '" & TglNul2(nTgl1) & "') " & _
             " "
    strSql = strSql & " UNION "
    strSql = strSql & "Select A.No_polisi, A.Customer, A.FMDl as Type, A.NO_WO, B.FS_NO_RETUR as NoInv, (B.FD_Retur) as TglInv, " & _
             "(A.Rep1 + ', ' + A.Rep2 + ', ' + A.Rep3 + ', ' + A.Rep4 + ', ' + A.Rep5 + ', ' + A.Rep6 + ', ' + A.Rep7) as Jenis_Pekerjaan, " & _
             "A.[DATE] + A.[TIME] as Tgl_Masuk, B.Labor as Jasa, (B.Part_G + B.Part_NG + B.Oil_G + B.Oil_NG) as Part, B.SBLT as Sublet, B.Total  as Revenue " & _
             "from JobDTA A WITH(NOLOCK) Inner Join T_invoice B WITH(NOLOCK) On A.NO_WO = B.No_WO Inner Join NEA_PS C WITH(NOLOCK) On A.NO_WO = c.NOWO " & _
             "Where B.FS_LOG_DEL <> '' and C.UserValid <> '' And C.PS = 1 and (B.FD_Retur between '" & TglNul2(nTgl) & "' and '" & TglNul2(nTgl1) & "') " & _
             "Order By TglInv, B.NOINV ASC"
    
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchData_LapPS = rs
    If aConn.State = adStateOpen Then aConn.Close
    Exit Function
errHandler:
If aConn.State = adStateOpen Then aConn.Close
err.Raise err.Number, err.Source, err.Description
End Function


Public Sub InsertData_Penjualan_Vehicle()
Dim cm As New Command
On Error GoTo errHandler
    If aConn.State = adStateClosed Then aConn.Open (SQLCONBengkel)
'    cm.CommandText = "insert into SalesUnit " & _
'                     "select NORANGKA, TANGGAL from OPENROWSET('SQLOLEDB','DRIVER={SQL Server};SERVER=" & Log.LogHOKompSPK & ";Uid=" & Log.LogHOUserSPK & ";PWD=" & Log.LogHOPassSPK & "',[" & Log.LogHODbSPK & "].dbo.SUMMARY) " & _
'                     "WHERE TIPETR = 'ND' AND STATUS = 1 AND ISNULL(HAPUS,0) = 0 AND TANGGAL between '" & TglNul2(DateAdd("d", -7, isnow)) & "' and '" & TglNul2(DateAdd("d", -1, isnow)) & "' " & _
'                     "AND NORANGKA NOT IN (SELECT NORANGKA FROM SALESUNIT Where TglEntry between '" & TglNul2(DateAdd("d", -7, isnow)) & "' and '" & TglNul2(DateAdd("d", -1, isnow)) & "')"
'     cm.CommandText = "insert into SalesUnit " & _
'                     "select NORANGKA, TGL_BPK AS TANGGAL, fs_kd_dealer, fs_nm_dealer from OPENROWSET('SQLOLEDB','DRIVER={SQL Server};SERVER=" & Log.LogHOKompSPK & ";Uid=" & Log.LogHOUserSPK & ";PWD=" & Log.LogHOPassSPK & "',[" & Log.LogHODbSPK & "].dbo.DATACUST) A LEFT JOIN T_DEALER b on a.dealer = b.KODECABANG  " & _
'                     "WHERE ISNULL(BATAL,0) = 0 AND TGL_BPK between '" & TglNul2(DateAdd("d", -7, isnow)) & "' and '" & TglNul2(DateAdd("d", -1, isnow)) & "' " & _
'                     "AND NORANGKA NOT IN (SELECT NORANGKA FROM SALESUNIT Where TglEntry between '" & TglNul2(DateAdd("d", -7, isnow)) & "' and '" & TglNul2(DateAdd("d", -1, isnow)) & "')"

   ''2016-07-29 "" DATA DATACUST DIGANTI SUMMARY -> KARENA DIDATACUST TIDAK ADA PENJUALAN OTHERS
'   cm.CommandText = "insert into SalesUnit " & _
'                     "select NORANGKA, TANGGAL AS TANGGAL, fs_kd_dealer, fs_nm_dealer from OPENROWSET('SQLOLEDB','DRIVER={SQL Server};SERVER=" & Log.LogHOKompSPK & ";Uid=" & Log.LogHOUserSPK & ";PWD=" & Log.LogHOPassSPK & "',[" & Log.LogHODbSPK & "].dbo.SUMMARY) A LEFT JOIN T_DEALER b on a.KODEDEALER = b.KODECABANG  " & _
'                     "WHERE ISNULL(HAPUS,0) = 0 and TIPETR = 'ND' AND [STATUS] = 1 AND TANGGAL between '" & TglNul2(DateAdd("d", -7, isnow)) & "' and '" & TglNul2(DateAdd("d", -1, isnow)) & "' " & _
'                     "AND NORANGKA NOT IN (SELECT NORANGKA FROM SALESUNIT Where TglEntry between '" & TglNul2(DateAdd("d", -7, isnow)) & "' and '" & TglNul2(DateAdd("d", -1, isnow)) & "')"

   ''2016-09-07
   ''status dihilangkan agar mendapatkan data penjualan terakhir jika terjadi retur penjualan disales
   ''
   cm.CommandText = "insert into SalesUnit " & _
                     "select NORANGKA, TANGGAL AS TANGGAL, fs_kd_dealer, fs_nm_dealer from OPENROWSET('SQLOLEDB','DRIVER={SQL Server};SERVER=" & Log.LogHOKompSPK & ";Uid=" & Log.LogHOUserSPK & ";PWD=" & Log.LogHOPassSPK & "',[" & Log.LogHODbSPK & "].dbo.SUMMARY) A LEFT JOIN T_DEALER b on a.KODEDEALER = b.KODECABANG  " & _
                     "WHERE ISNULL(HAPUS,0) = 0 and TIPETR = 'ND' AND TANGGAL between '" & TglNul2(DateAdd("d", -7, isnow)) & "' and '" & TglNul2(DateAdd("d", -1, isnow)) & "' " & _
                     "AND NORANGKA NOT IN (SELECT NORANGKA FROM SALESUNIT Where TglEntry between '" & TglNul2(DateAdd("d", -7, isnow)) & "' and '" & TglNul2(DateAdd("d", -1, isnow)) & "')"

     cm.CommandType = adCmdText
    Set cm.ActiveConnection = aConn
    cm.Execute
    Set cm = Nothing
    If aConn.State = adStateOpen Then aConn.Close
    Exit Sub
errHandler:
If aConn.State = adStateOpen Then aConn.Close
    If Not cm Is Nothing Then Set cm = Nothing
End Sub


'-------------- MASTER AKUN ACC NEW ----------------------------------------------------
Public Function Get_AkunNew() As ADODB.Recordset
   On Error GoTo errHandler
    If aConn.State = adStateClosed Then aConn.Open (SQLCONBengkel)
    
    strSql = "Select * FROM T_AkunNew WITH(NOLOCK) "
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    Set Get_AkunNew = rs
    If aConn.State = adStateOpen Then aConn.Close
        
    Exit Function
errHandler:
    If aConn.State = adStateOpen Then aConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

'Public Function Get_AkunGL_NEW(nTipe As String) As ADODB.Recordset
'    On Error GoTo errHandler
'    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
'
'    strSQL = "Select PiutangUmum,PiutangAsuransi,PenjJasa,PenjSublet,PenjPart,PenjOlie,PenjAccs,PPN,Materai,HPPJasa,HutangBorong,HPPSublet,BpInProses FROM T_AkunNew Where TIPE = '" & nTipe & "'"
'
'    Set rs = New ADODB.Recordset
'    rs.CursorLocation = adUseClient
'    rs.Open strSQL, SQLConBP, adOpenStatic, adLockReadOnly, adCmdText
'    Set rs.ActiveConnection = Nothing
'    Set Get_AkunGL_NEW = rs
'    If mConn.State = adStateOpen Then mConn.Close
'    Exit Function
'errHandler:
'err.Raise err.Number, err.Source, err.Description
'End Function


Public Sub CreateNew_AkunNew(nID As Integer, isAdd As Boolean)
   On Error GoTo errHandler
    If aConn.State = adStateClosed Then aConn.Open (SQLCONBengkel)
      Dim cm As New Command
      
      If isAdd = True Then
        Select Case nID
            Case 0  'Dyna Ext
                strSql = "Insert into T_AkunNew values ('01DynaExt','" & txtAcc((nID + 0)) & "','" & txtAcc((nID + 0) + 1) & "','" & txtAcc((nID + 0) + 2) & "','" & txtAcc((nID + 0) + 3) & "','" & txtAcc((nID + 0) + 4) & "','" & txtAcc((nID + 0) + 5) & "','" & txtAcc((nID + 0) + 6) & "','" & txtAcc((nID + 0) + 7) & "','" & txtAcc((nID + 0) + 8) & "','" & txtAcc((nID + 0) + 9) & "')"
            Case 1  'Dyna Int
                strSql = "Insert into T_AkunNew values ('02DynaInt','" & txtAcc((nID + 9)) & "','" & txtAcc((nID + 9) + 1) & "','" & txtAcc((nID + 9) + 2) & "','" & txtAcc((nID + 9) + 3) & "','" & txtAcc((nID + 9) + 4) & "','" & txtAcc((nID + 9) + 5) & "','" & txtAcc((nID + 9) + 6) & "','" & txtAcc((nID + 9) + 7) & "','" & txtAcc((nID + 9) + 8) & "','" & txtAcc((nID + 9) + 9) & "')"
            Case 2  'Dyna Afl
                strSql = "Insert into T_AkunNew values ('03DynaAfl','" & txtAcc((nID + 18)) & "','" & txtAcc((nID + 18) + 1) & "','" & txtAcc((nID + 18) + 2) & "','" & txtAcc((nID + 18) + 3) & "','" & txtAcc((nID + 18) + 4) & "','" & txtAcc((nID + 18) + 5) & "','" & txtAcc((nID + 18) + 6) & "','" & txtAcc((nID + 18) + 7) & "','" & txtAcc((nID + 18) + 8) & "','" & txtAcc((nID + 18) + 9) & "')"
            Case 3  'Non Dyna Ext
                strSql = "Insert into T_AkunNew values ('04NonDynaExt','" & txtAcc((nID + 27)) & "','" & txtAcc((nID + 27) + 1) & "','" & txtAcc((nID + 27) + 2) & "','" & txtAcc((nID + 27) + 3) & "','" & txtAcc((nID + 27) + 4) & "','" & txtAcc((nID + 27) + 5) & "','" & txtAcc((nID + 27) + 6) & "','" & txtAcc((nID + 27) + 7) & "','" & txtAcc((nID + 27) + 8) & "','" & txtAcc((nID + 27) + 9) & "')"
            Case 4  'Non Dyna Int
                strSql = "Insert into T_AkunNew values ('05NonDynaInt','" & txtAcc((nID + 36)) & "','" & txtAcc((nID + 36) + 1) & "','" & txtAcc((nID + 36) + 2) & "','" & txtAcc((nID + 36) + 3) & "','" & txtAcc((nID + 36) + 4) & "','" & txtAcc((nID + 36) + 5) & "','" & txtAcc((nID + 36) + 6) & "','" & txtAcc((nID + 36) + 7) & "','" & txtAcc((nID + 36) + 8) & "','" & txtAcc((nID + 36) + 9) & "')"
            Case 5  'Non Dyna Afl
                strSql = "Insert into T_AkunNew values ('06NonDynaAfl','" & txtAcc((nID + 45)) & "','" & txtAcc((nID + 45) + 1) & "','" & txtAcc((nID + 45) + 2) & "','" & txtAcc((nID + 45) + 3) & "','" & txtAcc((nID + 45) + 4) & "','" & txtAcc((nID + 45) + 5) & "','" & txtAcc((nID + 45) + 6) & "','" & txtAcc((nID + 45) + 7) & "','" & txtAcc((nID + 45) + 8) & "','" & txtAcc((nID + 45) + 9) & "')"
            Case 6  'Sublet
                strSql = "Insert into T_AkunNew values ('07Sublet','" & txtAcc((nID + 54)) & "','" & txtAcc((nID + 54) + 1) & "','" & txtAcc((nID + 54) + 2) & "','" & txtAcc((nID + 54) + 3) & "','" & txtAcc((nID + 54) + 4) & "','" & txtAcc((nID + 54) + 5) & "','" & txtAcc((nID + 54) + 6) & "','" & txtAcc((nID + 54) + 7) & "','" & txtAcc((nID + 54) + 8) & "','" & txtAcc((nID + 54) + 9) & "')"
        End Select
      Else
        Select Case nID
            Case 0  'Dyna Ext
                strSql = "Update T_AkunNew Set Piutang='" & txtAcc((nID + 0)) & "',PenjJasa='" & txtAcc((nID + 0) + 1) & "',PenjSublet='" & txtAcc((nID + 0) + 2) & "',PenjPart='" & txtAcc((nID + 0) + 3) & "',PenjOlie='" & txtAcc((nID + 0) + 4) & "',PenjAccs='" & txtAcc((nID + 0) + 5) & "',PPN='" & txtAcc((nID + 0) + 6) & "',Materai='" & txtAcc((nID + 0) + 7) & "',HPPSublet='" & txtAcc((nID + 0) + 8) & "',BpInProses='" & txtAcc((nID + 0) + 9) & "' Where Tipe = '01DynaExt'"
            Case 1  'Dyna Int
                strSql = "Update T_AkunNew Set Piutang='" & txtAcc((nID + 9)) & "',PenjJasa='" & txtAcc((nID + 9) + 1) & "',PenjSublet='" & txtAcc((nID + 9) + 2) & "',PenjPart='" & txtAcc((nID + 9) + 3) & "',PenjOlie='" & txtAcc((nID + 9) + 4) & "',PenjAccs='" & txtAcc((nID + 9) + 5) & "',PPN='" & txtAcc((nID + 9) + 6) & "',Materai='" & txtAcc((nID + 9) + 7) & "',HPPSublet='" & txtAcc((nID + 9) + 8) & "',BpInProses='" & txtAcc((nID + 9) + 9) & "' Where Tipe = '02DynaInt'"
            Case 2  'Dyna Afl
                strSql = "Update T_AkunNew Set Piutang='" & txtAcc((nID + 18)) & "',PenjJasa='" & txtAcc((nID + 18) + 1) & "',PenjSublet='" & txtAcc((nID + 18) + 2) & "',PenjPart='" & txtAcc((nID + 18) + 3) & "',PenjOlie='" & txtAcc((nID + 18) + 4) & "',PenjAccs='" & txtAcc((nID + 18) + 5) & "',PPN='" & txtAcc((nID + 18) + 6) & "',Materai='" & txtAcc((nID + 18) + 7) & "',HPPSublet='" & txtAcc((nID + 18) + 8) & "',BpInProses='" & txtAcc((nID + 18) + 9) & "' Where Tipe = '03DynaAfl'"
            Case 3  'Non Dyna Ext
                strSql = "Update T_AkunNew Set Piutang='" & txtAcc((nID + 27)) & "',PenjJasa='" & txtAcc((nID + 27) + 1) & "',PenjSublet='" & txtAcc((nID + 27) + 2) & "',PenjPart='" & txtAcc((nID + 27) + 3) & "',PenjOlie='" & txtAcc((nID + 27) + 4) & "',PenjAccs='" & txtAcc((nID + 27) + 5) & "',PPN='" & txtAcc((nID + 27) + 6) & "',Materai='" & txtAcc((nID + 27) + 7) & "',HPPSublet='" & txtAcc((nID + 27) + 8) & "',BpInProses='" & txtAcc((nID + 27) + 9) & "' Where Tipe = '04NonDynaExt'"
            Case 4  'Non Dyna Int
                strSql = "Update T_AkunNew Set Piutang='" & txtAcc((nID + 36)) & "',PenjJasa='" & txtAcc((nID + 36) + 1) & "',PenjSublet='" & txtAcc((nID + 36) + 2) & "',PenjPart='" & txtAcc((nID + 36) + 3) & "',PenjOlie='" & txtAcc((nID + 36) + 4) & "',PenjAccs='" & txtAcc((nID + 36) + 5) & "',PPN='" & txtAcc((nID + 36) + 6) & "',Materai='" & txtAcc((nID + 36) + 7) & "',HPPSublet='" & txtAcc((nID + 36) + 8) & "',BpInProses='" & txtAcc((nID + 36) + 9) & "' Where Tipe = '05NonDynaInt'"
            Case 5  'Non Dyna Afl
                strSql = "Update T_AkunNew Set Piutang='" & txtAcc((nID + 45)) & "',PenjJasa='" & txtAcc((nID + 45) + 1) & "',PenjSublet='" & txtAcc((nID + 45) + 2) & "',PenjPart='" & txtAcc((nID + 45) + 3) & "',PenjOlie='" & txtAcc((nID + 45) + 4) & "',PenjAccs='" & txtAcc((nID + 45) + 5) & "',PPN='" & txtAcc((nID + 45) + 6) & "',Materai='" & txtAcc((nID + 45) + 7) & "',HPPSublet='" & txtAcc((nID + 45) + 8) & "',BpInProses='" & txtAcc((nID + 45) + 9) & "' Where Tipe = '06NonDynaAfl'"
            Case 6  'Sublet
                strSql = "Update T_AkunNew Set Piutang='" & txtAcc((nID + 54)) & "',PenjJasa='" & txtAcc((nID + 54) + 1) & "',PenjSublet='" & txtAcc((nID + 54) + 2) & "',PenjPart='" & txtAcc((nID + 54) + 3) & "',PenjOlie='" & txtAcc((nID + 54) + 4) & "',PenjAccs='" & txtAcc((nID + 54) + 5) & "',PPN='" & txtAcc((nID + 54) + 6) & "',Materai='" & txtAcc((nID + 54) + 7) & "',HPPSublet='" & txtAcc((nID + 54) + 8) & "',BpInProses='" & txtAcc((nID + 54) + 9) & "' Where Tipe = '07Sublet'"
        End Select
      End If
      
      cm.CommandText = strSql
      cm.CommandType = adCmdText
        
      Set cm.ActiveConnection = aConn
      cm.Execute
      Set cm = Nothing
      If aConn.State = adStateOpen Then aConn.Close
      Exit Sub
errHandler:
    If Not cm Is Nothing Then Set cm = Nothing
      If aConn.State = adStateOpen Then aConn.Close
         err.Raise err.Number, err.Source, err.Description
    
End Sub

Public Function Get_KodeAkun(kode As String) As ADODB.Recordset
    On Error GoTo errHandler
    If aConn.State = adStateClosed Then aConn.Open (SQLConACC)
    
    strSql = "Select kodegl, namagl from glmst WITH(NOLOCK) " & kode
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConACC, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    Set Get_KodeAkun = rs
    If aConn.State = adStateOpen Then aConn.Close
        
    Exit Function
errHandler:
    If aConn.State = adStateOpen Then aConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

Public Function Fetch_TipeCust_Afiliasi_ARMST(nkode As String) As Boolean
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConACC)
    
    strSql = "Select Afiliasi From ARMST WITH(NOLOCK) Where KodeLgn = '" & nkode & "'"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConACC, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
'    If rs.Fields(0).value = "1" Then
'        Fetch_TipeCust_Afiliasi_ARMST = True
'    Else
'        Fetch_TipeCust_Afiliasi_ARMST = False
'    End If
    If Not rs.EOF Then
        Fetch_TipeCust_Afiliasi_ARMST = rs.Fields(0).value
    End If
    
    If mConn.State = adStateOpen Then mConn.Close
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function Fetch_TipeCust_Afiliasi_SIJURNALNEW(nkode As String) As Boolean
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConACC)
    
    strSql = "Select KodeLgn From SiJurnalNew WITH(NOLOCK) Where KodeLgn = '" & nkode & "'"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConACC, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    If rs.RecordCount > 0 Then
        Fetch_TipeCust_Afiliasi_SIJURNALNEW = True
    Else
        Fetch_TipeCust_Afiliasi_SIJURNALNEW = False
    End If
    If mConn.State = adStateOpen Then mConn.Close
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function Fetch_TipeCust_Internal_SIJURNALNEW(nkode As String) As Boolean
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConACC)
    
    strSql = "Select KodeLgn From SiJurnalNew WITH(NOLOCK) Where KodeLgn = '" & nkode & "'"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConACC, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    If rs.RecordCount > 0 Then
        Fetch_TipeCust_Internal_SIJURNALNEW = True
    Else
        Fetch_TipeCust_Internal_SIJURNALNEW = False
    End If
    If mConn.State = adStateOpen Then mConn.Close
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function


Public Function Get_AkunGL_NEW(nKodeLgn As String, isDyna As Boolean, isRetur As Boolean) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConACC)
    
    If isRetur = False Then
        If isDyna = True Then
            strSql = "Select NoJPiutangGrpUmum, " & _
                            "NoJPenjGrpDyna, NoJPenjSublGrpDyna, NoJPenjPartGrpDyna, NoJPenjOliGrpDyna, NoJPenjAccGrpDyna, " & _
                            "NoJPPnBengkel, NojPPnPart, NojPPnAcc, NoJBiayaMtr, " & _
                            "NoJHppSublGrpDyna , NoJPersSublGrpDyna " & _
                            "From sijurnalnew WITH(NOLOCK)  Where KodeLgn = '" & nKodeLgn & "'"
        Else
            strSql = "Select NoJPiutangGrpUmum, " & _
                            "NoJPenjGrpNonDyna, NoJPenjSublGrpNonDyna, NoJPenjPartGrpNonDyna, NoJPenjOliGrpNonDyna, NoJPenjAccGrpNonDyna, " & _
                            "NoJPPnBengkel, NojPPnPart, NojPPnAcc, NoJBiayaMtr, " & _
                            "NoJHppSublGrpNonDyna , NoJPersSublGrpNonDyna " & _
                            "From sijurnalnew WITH(NOLOCK) Where KodeLgn = '" & nKodeLgn & "'"
        End If
    Else
        If isDyna = True Then
            strSql = "Select NoJPiutangGrpUmum, " & _
                            "NoRPenjGrpDyna, NoRPenjSublGrpDyna, NoRPenjPartGrpDyna, NoRPenjOliGrpDyna, NoRPenjAccGrpDyna, " & _
                            "NoJPPnBengkel, NojPPnPart, NojPPnAcc, NoJBiayaMtr, " & _
                            "NoJHppSublGrpDyna , NoJPersSublGrpDyna " & _
                            "From sijurnalnew WITH(NOLOCK) Where KodeLgn = '" & nKodeLgn & "'"
        Else
            strSql = "Select NoJPiutangGrpUmum, " & _
                            "NoRPenjGrpNonDyna, NoRPenjSublGrpNonDyna, NoRPenjPartGrpNonDyna, NoRPenjOliGrpNonDyna, NoRPenjAccGrpNonDyna, " & _
                            "NoJPPnBengkel, NojPPnPart, NojPPnAcc, NoJBiayaMtr, " & _
                            "NoJHppSublGrpNonDyna , NoJPersSublGrpNonDyna " & _
                            "From sijurnalnew WITH(NOLOCK) Where KodeLgn = '" & nKodeLgn & "'"
        End If
    End If
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConACC, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    Set Get_AkunGL_NEW = rs
    If mConn.State = adStateOpen Then mConn.Close
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function Get_AkunGL_NEW_Sublet(nKodeLgn As String, isDyna As Boolean, isRetur As Boolean) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConACC)
    
    If isRetur = False Then
        If isDyna = True Then
            strSql = "Select NoJPersSublGRPDyna, NoJPPnSublet, NoJHutangSubletGRP " & _
                            "From sijurnalnew WITH(NOLOCK) Where KodeLgn = '" & nKodeLgn & "'"
        Else
            strSql = "Select NoJPersSublGRPNonDyna, NoJPPnSublet, NoJHutangSubletGRP " & _
                            "From sijurnalnew WITH(NOLOCK) Where KodeLgn = '" & nKodeLgn & "'"
        End If
    Else
        If isDyna = True Then
            strSql = "Select NoJPersSublGRPNonDyna, NoJPPnSublet, NoJHutangSubletGRP " & _
                            "From sijurnalnew WITH(NOLOCK) Where KodeLgn = '" & nKodeLgn & "'"
        Else
            strSql = "Select NoJPersSublGRPNonDyna, NoJPPnSublet, NoJHutangSubletGRP " & _
                            "From sijurnalnew WITH(NOLOCK) Where KodeLgn = '" & nKodeLgn & "'"
        End If
    End If
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConACC, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    Set Get_AkunGL_NEW_Sublet = rs
    If mConn.State = adStateOpen Then mConn.Close
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function


