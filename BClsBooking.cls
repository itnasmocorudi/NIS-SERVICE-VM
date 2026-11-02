VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BClsBooking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mConn As New ADODB.Connection
Private rs As ADODB.Recordset
Private strSQL As String

Public Function isBookingExist(numberBooking As String) As Boolean
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSQL = "SELECT * FROM BOOK WITH(NOLOCK) WHERE NOFAK = '" & numberBooking & "' AND DELIV <> 'OKE'"

    isBookingExist = False
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    If Not rs.EOF Then
        isBookingExist = True
    End If
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function searchBooking(numberBooking As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSQL = "SELECT * FROM BOOK WITH(NOLOCK) WHERE NOFAK = '" & numberBooking & "' AND DELIV <> 'OKE'"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set searchBooking = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function searchBookingPart(numberBooking As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSQL = "SELECT * FROM BOOK_PART WITH(NOLOCK) WHERE NOFAK = '" & numberBooking & "'"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set searchBookingPart = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function
