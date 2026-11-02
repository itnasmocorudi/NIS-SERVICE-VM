VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsSSC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mConn As New ADODB.Connection
Private mAConn As New ADODB.Connection
Private rs As ADODB.Recordset
Private strSql As String

Public Function FetchDataHoWarranty() As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    If Log.RPLHistori = 0 Then
        strSql = "Select a.Id_warranty, b.MDL, b.Model, b.no_Rangka1, b.no_Rangka2, a.warranty, a.Keterangan, " & _
                 "a.TglMulai, a.tglselesai FROM HO_Warranty a WITH(NOLOCK) INNER JOIN HO_WARDTL b WITH(NOLOCK) " & _
                 "ON a.Id_Warranty = b.Id_Warranty and A.FS_LOG_DEL = '' Order By a.Id_warranty ASC"
    Else
        strSql = "Select a.Id_warranty, b.MDL, b.Model, b.no_Rangka1, b.no_Rangka2, a.warranty, a.Keterangan, " & _
                 "a.TglMulai, a.tglselesai FROM [" & dbReplikasiHO & "]..HO_Warranty a WITH(NOLOCK) INNER JOIN [" & dbReplikasiHO & "]..HO_WARDTL b WITH(NOLOCK) " & _
                 "ON a.Id_Warranty = b.Id_Warranty and A.FS_LOG_DEL = '' Order By a.Id_warranty ASC"

    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchDataHoWarranty = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
      If mConn.State = adStateOpen Then mConn.Close
err.Raise err.Number, err.Source, err.Description
End Function


Public Function FetchDataCustdatWarranty(idRangka As String, LenId As Integer, Rangka1 As Double, Rangka2 As Double) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    strSql = "SELECT Customer, Cust_adres, Kota, KodePos, No_Polisi, No_Engine, No_Chassis, " & _
             "Telpon, GSM, CDMA FROM CUSTDAT WITH(NOLOCK) " & _
             "WHERE SubString(No_Chassis,1," & LenId & ") = '" & idRangka & "'"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchDataCustdatWarranty = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
      If mConn.State = adStateOpen Then mConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchRecordItem(noRangka As String, Id_Warr As String, nTipe As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    strSql = "SELECT * FROM M_ListReminderSSC WITH(NOLOCK) " & _
             "WHERE NoRangka = '" & noRangka & "' AND Id_Warranty = '" & Id_Warr & "' AND Tipe = '" & nTipe & "'"
     
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchRecordItem = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
      If mConn.State = adStateOpen Then mConn.Close
err.Raise err.Number, err.Source, err.Description
End Function

Public Sub AddNew_LogRemainderSSC_Call(K_Telp As Integer, Pesan_Telp As String, nTipe As String, noRangka As String, Pesan1 As String, Pesan2 As String, _
                                       Id_Warr As String, App As String, Pesan_App As String)
   On Error GoTo ErrAddCall
      Dim cm As New Command
      If mConn.State = adStateClosed Then mConn.Open (SQLCon)
      strSql = "Insert Into M_ListReminderSSC(TglProses, Diterima, KetDiterima, Tipe, NoRangka, Pesan1, Pesan2, Id_Warranty, App, KetApp) " & _
               "Values(GetDate()," & K_Telp & " , '" & Pesan_Telp & "', '" & nTipe & "', '" & noRangka & "', '" & Pesan1 & "', '" & Pesan2 & "', '" & Id_Warr & "', '" & App & "', '" & Pesan_App & "' )"
               
      cm.CommandText = strSql
      cm.CommandType = adCmdText
        
      Set cm.ActiveConnection = mConn
      cm.Execute
      Set cm = Nothing
      If mConn.State = adStateOpen Then mConn.Close
      Exit Sub
ErrAddCall:
      If mConn.State = adStateOpen Then mConn.Close
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub


Public Function FetchDataReminderSSC(idRangka As String, Id_Warr As String, isCall As Boolean) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    If isCall = False Then
      strSql = "SELECT TglProses, Tipe FROM M_ListReminder WITH(NOLOCK) WHERE NoRangka = '" & idRangka & "' And LastNoWo = '" & Id_Warr & "' Order By TglProses ASC"
    Else
      strSql = "SELECT TglProses, Tipe FROM M_ListReminderSSC WITH(NOLOCK) WHERE NoRangka = '" & idRangka & "' And Id_Warranty = '" & Id_Warr & "' Order By TglProses ASC"
    End If
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchDataReminderSSC = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function
