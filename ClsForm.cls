VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mConn As New ADODB.Connection
Private rs As ADODB.Recordset
Private strSql As String

Private mTop As Double
Private mLeft As Double

Public Property Let SetTop(value As Double)
    mTop = value
End Property
Public Property Get MyTop() As Double
    MyTop = mTop
End Property

Public Property Let SetLeft(value As Double)
    mLeft = value
End Property
Public Property Get MyLeft() As Double
    MyLeft = mLeft
End Property

Public Function searchVehicleModel() As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConHO)
    
    strSql = "SELECT a.fs_grp_mdl, a.mdl, a.[desc], ISNULL(id_program, '') as PRO FROM model a WITH(NOLOCK) LEFT JOIN ho_PROmdl b WITH(NOLOCK) ON a.MDL = b.MDL ORDER BY a.fs_grp_mdl, a.mdl"
            
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConHO, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set searchVehicleModel = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
    err.Raise err.Number, err.Source, err.Description
End Function

Public Function searchEvent(eventCode As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConHO)
    
    'If Log.RPLHistori = 0 Then
        strSql = "SELECT * FROM HO_EVENT WITH(NOLOCK) WHERE NOEVENT = '" & Trim(eventCode) & "' and FS_LOG_DEL = ''"
    'Else
    '    strSql = "SELECT * FROM [" & dbReplikasiHO & "]..HO_EVENT WITH(NOLOCK) WHERE NOEVENT = '" & Trim(eventCode) & "' and FS_LOG_DEL = ''"
    'End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConHO, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set searchEvent = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
    err.Raise err.Number, err.Source, err.Description
End Function

Public Function searchEventState(eventCode As String) As String
    On Error GoTo errHandler
    
    searchEventState = ""
    If mConn.State = adStateClosed Then mConn.Open (SQLConHO)
    
    strSql = "SELECT aktif FROM HO_EVENT WITH(NOLOCK) WHERE NOEVENT = '" & Trim(eventCode) & "' and FS_LOG_DEL = ''"
            
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConHO, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    searchEventState = Trim(rs("aktif"))
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
    searchEventState = ""
    err.Raise err.Number, err.Source, err.Description
    Exit Function
End Function

Public Function displayEvent(eventState As Integer) As ADODB.Recordset
Dim statEvent As String

    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConHO)
    
    If eventState = 1 Then
        statEvent = " And aktif = " & CDbl(eventState) & ""
    Else
        statEvent = ""
    End If
    strSql = "SELECT noEvent as No,namaEvent as Event, textEvent as Keterangan,tipeOrder as 'Tipe_Booking',thVehicle1 as 'Tahun_Mulai',thVehicle2 as 'Tahun_Akhir',KM1 as 'KM_Mulai',KM2 as 'KM_Akhir',lastService1, lastService2,birthDay1 as 'Ultah_Mulai',birthDay2 as 'Ultah_Selesai',hobby as 'Hobby', " _
           & "DelivDay1 as 'Delivery_Mulai',DelivDay2 as 'Delivery_Akhir',modelVehichle as 'Model',dealerBerlaku as 'Dealer',berlaku1 as 'Tanggal_Mulai',berlaku2 as 'Tanggal_Selesai', " _
           & "case when aktif = 1 then 'Aktif' else 'Non Aktif' end as 'Status_Event',fs_Log_Add as 'User_Add',fs_Log_DeAktifasi as 'Deactivation', kodeProgram as 'Kode', direminder FROM HO_EVENT WITH(NOLOCK) WHERE FS_LOG_DEL = '' " & statEvent & ""
            
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConHO, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set displayEvent = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
    err.Raise err.Number, err.Source, err.Description
End Function

