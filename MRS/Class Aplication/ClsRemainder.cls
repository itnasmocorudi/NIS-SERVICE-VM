VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsRemainder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mConn As New ADODB.Connection
Private mAConn As New ADODB.Connection
Private mSConn As New ADODB.Connection
Private rs As ADODB.Recordset
Private strSql As String

Private mNomorRangka As String
Private mNoGSM As String
Private mNoCDMA As String
Private mTglSendNow As Date
Private mPesanSMS As String

Private mNoSMS As String
Private mPesan  As String
Private mTujuan As String
Private mWaktu  As Date
Private mMilik As String
Private mPengirim As String
Private mTanggal1 As Date
Private mTanggal2 As Date

Private mSearchBy As String
Private mSearchText As String


Public Property Let NoSMS(value As String)
    mNoSMS = value
End Property
Public Property Let pesan(value As String)
    mPesan = value
End Property
Public Property Let Tujuan(value As String)
    mTujuan = value
End Property
Public Property Let waktu(value As String)
    mWaktu = value
End Property
Public Property Let Milik(value As String)
    mMilik = value
End Property
Public Property Let Pengirim(value As String)
    mPengirim = value
End Property
Public Property Let tanggal1(value As Date)
    mTanggal1 = value
End Property
Public Property Let tanggal2(value As Date)
    mTanggal2 = value
End Property


Public Property Let NomorRangka(value As String)
    mNomorRangka = value
End Property
Public Property Let NoGSM(value As String)
    mNoGSM = value
End Property
Public Property Let NoCDMA(value As String)
    mNoCDMA = value
End Property
Public Property Let TglSendNow(value As Variant)
    mTglSendNow = value
End Property
Public Property Let PesanSMS(value As String)
    mPesanSMS = value
End Property


Public Function FetchDataRemainder(TglNext0 As Date, TglNext1 As Date, nRem As Integer, Optional TglRem0 As Date, Optional TglRem1 As Date, Optional ketr As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
'    strSQL = "select a.NoInv, a.NoWo, a.TglSvcNow as DateLastSvc,a.KmNow, " & _
'             "(CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END) as DateNextSvc, " & _
'             "a.KmNext as KmNextSvc, c.UseName, c.UseAddr, c.UseCity, c.UsePos, c.No_Polisi as NoPolisi, b.No_Engine as NoEngine, a.NoRangka as NoChassis, " & _
'             "c.UsePSTN as Telpon , c.UseGSM, c.UseCDMA, a.NoRangka as NoChassisSMS, c.FS_KD_PLG as KodeLGN, c.*, b.FS_RC  " & _
'             "FROM M_Estimasi a INNER JOIN History b ON a.NoWo = b.No_Wo AND B.FS_LOG_DEL = '' INNER JOIN Custdat c On a.NoRangka = c.No_Chassis " & _
'             "WHERE a.JT_Tgl is Not Null AND "
             
    strSql = "SELECT a.NoInv, a.NoWo, a.TglSvcNow as DateLastSvc,a.KmNow, " & _
             "(CASE WHEN A.JT_Tgl > A.JT_Bln then A.JT_Bln WHEN A.JT_Tgl < A.JT_Bln then A.JT_Tgl WHEN A.JT_Tgl = A.JT_Bln then A.Jt_Bln END) as DateNextSvc, " & _
             "a.KmNext as KmNextSvc, c.UseName, c.UseAddr, c.UseCity, c.UsePos, c.No_Polisi as NoPolisi, C.FMDL as KetModel, " & _
             "a.NoRangka as NoChassis, c.UsePSTN as Telpon , c.UseGSM, c.UseCDMA, a.NoRangka as NoChassisSMS, c.FS_KD_PLG as KodeLGN, " & _
             "b.FS_RC, A.SbeNext, C.COLOR as WrnModel, C.USEEMAIL AS EMAIL, C.MDL, C.MODEL, C.FS_TH_MOBIL, C.COLOR, C.FS_KD_DEALER, C.FS_NM_DEALER FROM M_ESTIMASI A WITH(NOLOCK) Inner Join " & _
             "(SELECT G.NORANGKA, (SELECT MAX(SBENEXT) FROM m_estimasi WITH(NOLOCK) WHERE NORANGKA=G.NORANGKA AND ISNULL(TGLSVCNOW,'') <> '') S, MAX((CASE WHEN G.JT_Tgl > G.JT_Bln then G.JT_Bln WHEN G.JT_Tgl < G.JT_Bln then G.JT_Tgl WHEN G.JT_Tgl = G.JT_Bln then G.JT_Tgl END)) AS JT_TGL " & _
             "FROM M_ESTIMASI G WITH(NOLOCK) Where (CASE WHEN G.JT_Tgl > G.JT_Bln then G.JT_Bln WHEN G.JT_Tgl < G.JT_Bln then G.JT_Tgl WHEN G.JT_Tgl = G.JT_Bln then G.JT_Tgl END) Is Not Null GROUP BY G.NORANGKA) AS Z ON A.NORANGKA = Z.NORANGKA AND A.SBENEXT = Z.S " & _
             "INNER JOIN History b WITH(NOLOCK) ON a.NoWo = b.No_Wo AND A.NORANGKA = B.NO_CHASSIS AND B.FS_LOG_DEL = '' " & _
             "INNER JOIN Custdat c WITH(NOLOCK) On a.NoRangka = c.No_Chassis Where "
             
'             If nRem = 1 Then
'                strSQL = strSQL & "(CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END) <= '" & FormatTgl(TglNext1) & "' " & _
'                         "Order By a.noRangka, a.SBENext DESC"
'             Else
'                strSQL = strSQL & "(CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END) between '" & FormatTgl(TglNext0) & "' And '" & FormatTgl(TglNext1) & "' " & _
'                         "Order By a.noRangka, a.SBENext DESC"
'             End If
             
             If nRem = 1 Then
                strSql = strSql & "(CASE WHEN A.JT_Tgl > A.JT_Bln then  A.JT_Bln WHEN A.JT_Tgl < A.JT_Bln then  A.JT_Tgl WHEN A.JT_Tgl = A.JT_Bln then  A.Jt_Bln END) <= '" & FormatTgl(TglNext1) & "' "
                         '"Order By a.noRangka, a.SBENext DESC"
             Else
                strSql = strSql & "(CASE WHEN A.JT_Tgl > A.JT_Bln then  A.JT_Bln WHEN A.JT_Tgl < A.JT_Bln then  A.JT_Tgl WHEN A.JT_Tgl = A.JT_Bln then  A.Jt_Bln END) between '" & FormatTgl(TglNext0) & "' And '" & FormatTgl(TglNext1) & "' "
                         '"Order By a.noRangka, a.SBENext DESC"
             End If
             strSql = strSql & ketr
             
    strSql = strSql & "UNION SELECT NOINV, NO_WO AS NOWO, DATE AS DATELASTSVC, KM AS KMNOW, DATENEXTSVC, KMNEXTSVC, L.UseName, L.UseAddr, L.UseCity, L.UsePos, L.No_Polisi as NoPolisi, L.FMDL as KetModel, " & _
                      "K.NoRangka as NoChassis, L.UsePSTN as Telpon , L.UseGSM, L.UseCDMA, K.NoRangka as NoChassisSMS, L.FS_KD_PLG as KodeLGN, " & _
                      "M.FS_RC, CAST ((CAST(RIGHT(M.FS_RC,3) AS INTEGER) + (KMNEXTSVC - KM )/1000) * 1000 AS INTEGER) AS SBENEXT, L.COLOR as WrnModel, L.USEEMAIL AS EMAIL, L.MDL, L.MODEL, L.FS_TH_MOBIL, L.COLOR, L.FS_KD_DEALER, L.FS_NM_DEALER " & _
                      "FROM M_LISTREMINDERCALL K WITH(NOLOCK) LEFT JOIN HISTORY M WITH(NOLOCK) ON K.NORANGKA = M.NO_CHASSIS AND K.LASTNOWO = M.NO_WO " & _
                      "INNER JOIN Custdat L WITH(NOLOCK) On K.NoRangka = L.No_Chassis WHERE TANGGAL BETWEEN '" & FormatTgl(TglRem0) & "' AND '" & FormatTgl(TglRem1) & "' AND K.NoRangka <> '' " & _
                      "ORDER BY a.noRangka, a.SBENext DESC"
             
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchDataRemainder = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function


Public Function FetchDataRemainder_DM(Data As Boolean, TglNext0 As Date, TglNext1 As Date, nRem As Integer, Optional TglRem0 As Date, Optional TglRem1 As Date, Optional ketr As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
'    strSQL = "select a.NoInv, a.NoWo, a.TglSvcNow as DateLastSvc,a.KmNow, " & _
'             "(CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END) as DateNextSvc, " & _
'             "a.KmNext as KmNextSvc, c.UseName, c.UseAddr, c.UseCity, c.UsePos, c.No_Polisi as NoPolisi, b.No_Engine as NoEngine, a.NoRangka as NoChassis, " & _
'             "c.UsePSTN as Telpon , c.UseGSM, c.UseCDMA, a.NoRangka as NoChassisSMS, c.FS_KD_PLG as KodeLGN, c.*, b.FS_RC  " & _
'             "FROM M_Estimasi a INNER JOIN History b ON a.NoWo = b.No_Wo AND B.FS_LOG_DEL = '' INNER JOIN Custdat c On a.NoRangka = c.No_Chassis " & _
'             "WHERE a.JT_Tgl is Not Null AND "
    If Data = True Then
        strSql = "SELECT a.NoInv, a.NoWo, a.TglSvcNow as DateLastSvc,a.KmNow, " & _
             "(CASE WHEN A.JT_Tgl > A.JT_Bln then A.JT_Bln WHEN A.JT_Tgl < A.JT_Bln then A.JT_Tgl WHEN A.JT_Tgl = A.JT_Bln then A.Jt_Bln END) as DateNextSvc, " & _
             "a.KmNext as KmNextSvc, c.UseName, c.UseAddr, c.UseCity, c.UsePos, c.No_Polisi as NoPolisi, C.FMDL as KetModel, " & _
             "a.NoRangka as NoChassis, c.UsePSTN as Telpon , c.UseGSM, c.UseCDMA, a.NoRangka as NoChassisSMS, c.FS_KD_PLG as KodeLGN, " & _
             "b.FS_RC, A.SbeNext, C.COLOR as WrnModel, C.USEEMAIL AS EMAIL, C.MDL, C.MODEL, C.FS_TH_MOBIL, C.COLOR, C.FS_KD_DEALER, C.FS_NM_DEALER FROM M_ESTIMASI A WITH(NOLOCK) Inner Join " & _
             "(SELECT G.NORANGKA, (SELECT MAX(SBENEXT) FROM m_estimasi WITH(NOLOCK) WHERE NORANGKA=G.NORANGKA AND ISNULL(TGLSVCNOW,'') <> '') S, MAX((CASE WHEN G.JT_Tgl > G.JT_Bln then G.JT_Bln WHEN G.JT_Tgl < G.JT_Bln then G.JT_Tgl WHEN G.JT_Tgl = G.JT_Bln then G.JT_Tgl END)) AS JT_TGL " & _
             "FROM M_ESTIMASI G WITH(NOLOCK) Where (CASE WHEN G.JT_Tgl > G.JT_Bln then G.JT_Bln WHEN G.JT_Tgl < G.JT_Bln then G.JT_Tgl WHEN G.JT_Tgl = G.JT_Bln then G.JT_Tgl END) Is Not Null GROUP BY G.NORANGKA) AS Z ON A.NORANGKA = Z.NORANGKA AND A.SBENEXT = Z.S " & _
             "INNER JOIN History b WITH(NOLOCK) ON a.NoWo = b.No_Wo AND A.NORANGKA = B.NO_CHASSIS AND B.FS_LOG_DEL = '' " & _
             "INNER JOIN Custdat c WITH(NOLOCK) On a.NoRangka = c.No_Chassis Where "
             
'             If nRem = 1 Then
'                strSQL = strSQL & "(CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END) <= '" & FormatTgl(TglNext1) & "' " & _
'                         "Order By a.noRangka, a.SBENext DESC"
'             Else
'                strSQL = strSQL & "(CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END) between '" & FormatTgl(TglNext0) & "' And '" & FormatTgl(TglNext1) & "' " & _
'                         "Order By a.noRangka, a.SBENext DESC"
'             End If
             
             If nRem = 1 Then
                strSql = strSql & "(CASE WHEN A.JT_Tgl > A.JT_Bln then  A.JT_Bln WHEN A.JT_Tgl < A.JT_Bln then  A.JT_Tgl WHEN A.JT_Tgl = A.JT_Bln then  A.Jt_Bln END) <= '" & FormatTgl(TglNext1) & "' "
                         '"Order By a.noRangka, a.SBENext DESC"
             Else
                strSql = strSql & "(CASE WHEN A.JT_Tgl > A.JT_Bln then  A.JT_Bln WHEN A.JT_Tgl < A.JT_Bln then  A.JT_Tgl WHEN A.JT_Tgl = A.JT_Bln then  A.Jt_Bln END) between '" & FormatTgl(TglNext0) & "' And '" & FormatTgl(TglNext1) & "' "
                         '"Order By a.noRangka, a.SBENext DESC"
             End If
             strSql = strSql & ketr
    Else
        strSql = "SELECT NOINV, NO_WO AS NOWO, DATE AS DATELASTSVC, KM AS KMNOW, DATENEXTSVC, KMNEXTSVC, L.UseName, L.UseAddr, L.UseCity, L.UsePos, L.No_Polisi as NoPolisi, L.FMDL as KetModel, " & _
                      "K.NoRangka as NoChassis, L.UsePSTN as Telpon , L.UseGSM, L.UseCDMA, K.NoRangka as NoChassisSMS, L.FS_KD_PLG as KodeLGN, " & _
                      "M.FS_RC, CAST ((CAST(RIGHT(M.FS_RC,3) AS INTEGER) + (KMNEXTSVC - KM )/1000) * 1000 AS INTEGER) AS SBENEXT, L.COLOR as WrnModel, L.USEEMAIL AS EMAIL, L.MDL, L.MODEL, L.FS_TH_MOBIL, L.COLOR, L.FS_KD_DEALER, L.FS_NM_DEALER " & _
                      "FROM M_LISTREMINDERCALL K WITH(NOLOCK) LEFT JOIN HISTORY M WITH(NOLOCK) ON K.NORANGKA = M.NO_CHASSIS AND K.LASTNOWO = M.NO_WO " & _
                      "INNER JOIN Custdat L WITH(NOLOCK) On K.NoRangka = L.No_Chassis WHERE TANGGAL BETWEEN '" & FormatTgl(TglRem0) & "' AND '" & FormatTgl(TglRem1) & "' AND K.NoRangka <> '' And K.KMNextSvc <> 1000 " & _
                      "ORDER BY k.noRangka, KMNEXTSVC DESC"
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchDataRemainder_DM = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

                   
Public Function FetchDataRemainder_Item(TglNext As Date, xNoWo As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    strSql = "select a.NoInv, a.NoWo, a.TglSvcNow as DateLastSvc, a.KmNow, " & _
             "(CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END) as DateNextSvc, " & _
             "a.KmNext as KmNextSvc, c.UseName, c.UseAddr, c.UseCity, c.UsePos, c.No_Polisi as NoPolisi, b.No_Engine as NoEngine, a.NoRangka as NoChassis, " & _
             "c.UsePSTN as Telpon , c.UseGSM, c.UseCDMA " & _
             "FROM M_Estimasi a WITH(NOLOCK) INNER JOIN History b WITH(NOLOCK) ON a.NoWo = b.No_Wo AND B.FS_LOG_DEL = '' INNER JOIN Custdat c WITH(NOLOCK) On a.NoRangka = c.No_Chassis " & _
             "WHERE a.JT_Tgl is Not Null AND " & _
             "(CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END)= '" & FormatTgl(TglNext) & "' " & _
             "AND a.NoWo = '" & xNoWo & "' Order By a.noRangka, a.SBENext"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchDataRemainder_Item = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchDataSPK_Item(noRangka As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConSPK)

    strSql = "Select b.Tgl_BPK as TglDO, b.Tgl_BPK as TglDEC, c.TglCR1, c.KmCR1, c.TglCR2, c.KmCR2, a.Nama_Faktur, " & _
             "(a.Alamat1_Faktur + a.Alamat2_Faktur + a.Alamat3_Faktur + a.Alamat4_Faktur + a.Alamat5_Faktur ) as Alamat, a.Kotamadya, " & _
             "c.TmptLahir , c.TglLahir, c.Agama, c.Pekerjaan, c.Hoby, a.Wil_Lain " & _
             "FROM SPK a WITH(NOLOCK) INNER JOIN BPK b WITH(NOLOCK) On a.SPK_No = b.SPK_No INNER JOIN AFI c WITH(NOLOCK) On a.SPK_No = c.SPK_No " & _
             "WHERE a.NoRangka  = '" & noRangka & "'"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConSPK, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchDataSPK_Item = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchLastReminderCall(noWO As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    strSql = "Select Top 1 Left(Pesan_Telp,1) as isID, Pesan_Telp from M_ListReminderCall WITH(NOLOCK) Where LastNoWO = '" & noWO & "' Order By TglEntry DESC"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchLastReminderCall = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchEta(noWO As String, IsTipe As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    strSql = "SELECT EtaHargaJasa, EtaHargaPart, EtaJam, EtaMenit FROM M_ListReminder WITH(NOLOCK) WHERE LastNoWo = '" & noWO & "' And Tipe = '" & IsTipe & "'"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchEta = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchListRemainder(noWO As String, noChassis As String, KmNext As Double, isCall As Boolean) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    If isCall = False Then
      strSql = "SELECT Tipe, TglProses FROM M_ListReminder WITH(NOLOCK) WHERE LastNoWo = '" & noWO & "' AND NORANGKA = '" & noChassis & "' AND KMNEXTSVC = '" & KmNext & "'"
    Else
      strSql = "SELECT TglFollow FROM M_ListReminderCall WITH(NOLOCK) WHERE LastNoWo = '" & noWO & "' AND NORANGKA = '" & noChassis & "' AND KMNEXTSVC = '" & KmNext & "' Order By TglFollow DESC"
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchListRemainder = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchListRemainderCR(noWO As String, noChassis As String, KmNext As Double, isCall As Boolean) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    If isCall = False Then
      strSql = "SELECT Tipe, TglProses FROM [M_ListReminderCallCR] WITH(NOLOCK) WHERE LastNoWo = '" & noWO & "' AND NORANGKA = '" & noChassis & "' AND KMNEXTSVC = '" & KmNext & "'"
    Else
      strSql = "SELECT TglFollow FROM [M_ListReminderCallCR] WITH(NOLOCK) WHERE LastNoWo = '" & noWO & "' AND NORANGKA = '" & noChassis & "' AND KMNEXTSVC = '" & KmNext & "' Order By TglFollow DESC"
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchListRemainderCR = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Sub AddNew_LogRemainder(tipe As String, noRangka As String, LastNoWo As String, DateNextSvc As Date, KmNextSvc As Double, _
                               Pesan1 As String, Pesan2 As String, Remainder As Integer, EtaHargaJasa As Double, EtaHargaPart As Double, EtaJam As Integer, EtaMenit As Integer)
   On Error GoTo ErrAddLog
      
      Call DeleteItemDM(LastNoWo, tipe, KmNextSvc)
      Dim cm As New Command
      If mConn.State = adStateClosed Then mConn.Open (SQLCon)
      strSql = "Insert Into M_ListReminder(LastNoWo, Tipe, TglProses, NoRangka, DateNextSvc, KmNextSvc, Pesan1, Pesan2, Remainder, EtaHargaJasa, EtaHargaPart, EtaJam, EtaMenit,NIK,KDUser,TipeUser) " & _
               "Values( '" & LastNoWo & "' , '" & tipe & "' , GetDate() , '" & noRangka & "' , '" & FormatTgl(DateNextSvc) & "' , " & KmNextSvc & " , '" & SingleToDouble(Pesan1) & "' , '" & SingleToDouble(Pesan2) & "' , " & Remainder & " , " & k2t(EtaHargaJasa) & " , " & k2t(EtaHargaPart) & " , " & EtaJam & " , " & EtaMenit & ",'" & gNIK & "','" & gSA & "','" & gTipeUser & "')"
                                  
      cm.CommandText = strSql
      cm.CommandType = adCmdText
        
      Set cm.ActiveConnection = mConn
      cm.Execute
      Set cm = Nothing
      If mConn.State = adStateOpen Then mConn.Close
      Exit Sub
ErrAddLog:
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub



Public Sub AddNew_LogRemainderCR(tipe As String, noRangka As String, LastNoWo As String, DateNextSvc As Date, KmNextSvc As Double, _
                               Pesan1 As String, Pesan2 As String, Remainder As Integer, EtaHargaJasa As Double, EtaHargaPart As Double, EtaJam As Integer, EtaMenit As Integer)
   On Error GoTo ErrAddLog
      
      Call DeleteItemDM(LastNoWo, tipe, KmNextSvc)
      Dim cm As New Command
      If mConn.State = adStateClosed Then mConn.Open (SQLCon)
      strSql = "Insert Into M_ListReminderCR(LastNoWo, Tipe, TglProses, NoRangka, DateNextSvc, KmNextSvc, Pesan1, Pesan2, Remainder, EtaHargaJasa, EtaHargaPart, EtaJam, EtaMenit) " & _
               "Values( '" & LastNoWo & "' , '" & tipe & "' , GetDate() , '" & noRangka & "' , '" & FormatTgl(DateNextSvc) & "' , " & KmNextSvc & " , '" & SingleToDouble(Pesan1) & "' , '" & SingleToDouble(Pesan2) & "' , " & Remainder & " , " & k2t(EtaHargaJasa) & " , " & k2t(EtaHargaPart) & " , " & EtaJam & " , " & EtaMenit & ")"
                                  
      cm.CommandText = strSql
      cm.CommandType = adCmdText
        
      Set cm.ActiveConnection = mConn
      cm.Execute
      Set cm = Nothing
      If mConn.State = adStateOpen Then mConn.Close
      Exit Sub
ErrAddLog:
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub

Public Sub AddNew_LogRemainderLub(tipe As String, noRangka As String, LastNoWo As String, DateNextSvc As Date, KmNextSvc As Double, _
                               Pesan1 As String, Pesan2 As String, Remainder As Integer, EtaHargaJasa As Double, EtaHargaPart As Double, EtaJam As Integer, EtaMenit As Integer)
   On Error GoTo ErrAddLog
      
      Call DeleteItemDM(LastNoWo, tipe, KmNextSvc)
      Dim cm As New Command
      If mConn.State = adStateClosed Then mConn.Open (SQLCon)
      strSql = "Insert Into M_ListReminderLub(LastNoWo, Tipe, TglProses, NoRangka, DateNextSvc, KmNextSvc, Pesan1, Pesan2, Remainder, EtaHargaJasa, EtaHargaPart, EtaJam, EtaMenit) " & _
               "Values( '" & LastNoWo & "' , '" & tipe & "' , GetDate() , '" & noRangka & "' , '" & FormatTgl(DateNextSvc) & "' , " & KmNextSvc & " , '" & SingleToDouble(Pesan1) & "' , '" & SingleToDouble(Pesan2) & "' , " & Remainder & " , " & k2t(EtaHargaJasa) & " , " & k2t(EtaHargaPart) & " , " & EtaJam & " , " & EtaMenit & ")"
                                  
      cm.CommandText = strSql
      cm.CommandType = adCmdText
        
      Set cm.ActiveConnection = mConn
      cm.Execute
      Set cm = Nothing
      If mConn.State = adStateOpen Then mConn.Close
      Exit Sub
ErrAddLog:
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub

Public Sub DeleteItemDM(noWO As String, tipe As String, nKMNext As Double)
On Error GoTo errHandler
    Dim cm As New Command
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    strSql = "DELETE FROM M_ListReminder Where LastNoWo = '" & noWO & "' And Tipe = '" & tipe & "' and KMNextSVC = " & nKMNext & " "
             
    cm.CommandText = strSql
    cm.CommandType = adCmdText
    
    Set cm.ActiveConnection = mConn
    cm.Execute
    Set cm = Nothing
    If mConn.State = adStateOpen Then mConn.Close
    Exit Sub
errHandler:
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub

Public Sub AddNew_LogRemainder_Call(LastNoWo As String, TglFollow As Date, RangeFollow As Integer, noRangka As String, DateNextSvc As Date, _
                                    KmNextSvc As Double, K_Telp As Integer, Pesan_Telp As String, K_Ulang As Integer, Tanggal As Variant, jam As Variant, _
                                    K_Pel As Integer, K_Ken As Integer, K_DM As Integer, K_SMS As Integer, _
                                    EtaHargaJasa As Double, EtaHargaPart As Double, EtaJam As Integer, EtaMenit As Integer, EditPel As Integer, App As Integer, Pesan_App As String, pesan As String, reminderID As String)
   On Error GoTo ErrAddCall
      Dim cm As New Command
      If mConn.State = adStateClosed Then mConn.Open (SQLCon)
      strSql = "Insert Into M_ListReminderCall (LastNoWO,TglFollow,RangeFollow,NoRangka,DateNextSvc,KmNextSvc,K_Telp" & _
               " ,Pesan_Telp,K_Ulang,Tanggal,Jam,K_Pel,K_Ken,K_DM,K_SMS,EtaHargaJasa " & _
               " ,EtaHargaPart,EtaJam,EtaMenit,EditPel,App,Pesan_App,Pesan,TglEntry,NIK,KDUser,TipeUser,reminder_ID) " & _
               "Values( '" & LastNoWo & "' , '" & FormatTgl(TglFollow) & "' , " & RangeFollow & " , '" & noRangka & "', '" & FormatTgl(DateNextSvc) & "' , " & _
               " " & KmNextSvc & " , " & K_Telp & " , '" & Pesan_Telp & "' , " & K_Ulang & " , " & NullTgl(Tanggal) & " , '" & Nul(jam) & "' , " & _
               " " & K_Pel & " , " & K_Ken & " , " & K_DM & " , " & K_SMS & " , " & k2t(EtaHargaJasa) & ", " & k2t(EtaHargaPart) & " , " & EtaJam & " , " & EtaMenit & " , " & EditPel & " , " & App & " , '" & Pesan_App & "', '" & pesan & "', GetDate(),'" & gNIK & "','" & gSA & "','" & gTipeUser & "','" & reminderID & "')"
               
      cm.CommandText = strSql
      cm.CommandType = adCmdText
        
      Set cm.ActiveConnection = mConn
      cm.Execute
      Set cm = Nothing
      If mConn.State = adStateOpen Then mConn.Close
      Exit Sub
ErrAddCall:
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub

Public Sub AddNew_LogRemainder_CallCR(LastNoWo As String, TglFollow As Date, RangeFollow As Integer, noRangka As String, DateNextSvc As Date, _
                                    KmNextSvc As Double, K_Telp As Integer, Pesan_Telp As String, K_Ulang As Integer, Tanggal As Variant, jam As Variant, _
                                    K_Pel As Integer, K_Ken As Integer, K_DM As Integer, K_SMS As Integer, _
                                    EtaHargaJasa As Double, EtaHargaPart As Double, EtaJam As Integer, EtaMenit As Integer, EditPel As Integer, App As Integer, Pesan_App As String, pesan As String)
   On Error GoTo ErrAddCall
      Dim cm As New Command
      If mConn.State = adStateClosed Then mConn.Open (SQLCon)
      strSql = "Insert Into M_ListReminderCallCR (LastNoWO,TglFollow,RangeFollow,NoRangka,DateNextSvc,KmNextSvc,K_Telp" & _
               " ,Pesan_Telp,K_Ulang,Tanggal,Jam,K_Pel,K_Ken,K_DM,K_SMS,EtaHargaJasa " & _
               " ,EtaHargaPart,EtaJam,EtaMenit,EditPel,App,Pesan_App,Pesan,TglEntry) " & _
               "Values( '" & LastNoWo & "' , '" & FormatTgl(TglFollow) & "' , " & RangeFollow & " , '" & noRangka & "', '" & FormatTgl(DateNextSvc) & "' , " & _
               " " & KmNextSvc & " , " & K_Telp & " , '" & Pesan_Telp & "' , " & K_Ulang & " , " & NullTgl(Tanggal) & " , '" & Nul(jam) & "' , " & _
               " " & K_Pel & " , " & K_Ken & " , " & K_DM & " , " & K_SMS & " , " & k2t(EtaHargaJasa) & ", " & k2t(EtaHargaPart) & " , " & EtaJam & " , " & EtaMenit & " , " & EditPel & " , " & App & " , '" & Pesan_App & "', '" & pesan & "', GetDate())"
               
      cm.CommandText = strSql
      cm.CommandType = adCmdText
        
      Set cm.ActiveConnection = mConn
      cm.Execute
      Set cm = Nothing
      If mConn.State = adStateOpen Then mConn.Close
      Exit Sub
ErrAddCall:
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub

Public Sub AddNew_LogRemainder_CallLub(LastNoWo As String, TglFollow As Date, RangeFollow As Integer, noRangka As String, DateNextSvc As Date, _
                                    KmNextSvc As Double, K_Telp As Integer, Pesan_Telp As String, K_Ulang As Integer, Tanggal As Variant, jam As Variant, _
                                    K_Pel As Integer, K_Ken As Integer, K_DM As Integer, K_SMS As Integer, _
                                    EtaHargaJasa As Double, EtaHargaPart As Double, EtaJam As Integer, EtaMenit As Integer, EditPel As Integer, App As Integer, Pesan_App As String, pesan As String)
   On Error GoTo ErrAddCall
      Dim cm As New Command
      If mConn.State = adStateClosed Then mConn.Open (SQLCon)
      strSql = "Insert Into M_ListReminderCallLub (LastNoWO,TglFollow,RangeFollow,NoRangka,DateNextSvc,KmNextSvc,K_Telp" & _
               " ,Pesan_Telp,K_Ulang,Tanggal,Jam,K_Pel,K_Ken,K_DM,K_SMS,EtaHargaJasa " & _
               " ,EtaHargaPart,EtaJam,EtaMenit,EditPel,App,Pesan_App,Pesan,TglEntry) " & _
               "Values( '" & LastNoWo & "' , '" & FormatTgl(TglFollow) & "' , " & RangeFollow & " , '" & noRangka & "', '" & FormatTgl(DateNextSvc) & "' , " & _
               " " & KmNextSvc & " , " & K_Telp & " , '" & Pesan_Telp & "' , " & K_Ulang & " , " & NullTgl(Tanggal) & " , '" & Nul(jam) & "' , " & _
               " " & K_Pel & " , " & K_Ken & " , " & K_DM & " , " & K_SMS & " , " & k2t(EtaHargaJasa) & ", " & k2t(EtaHargaPart) & " , " & EtaJam & " , " & EtaMenit & " , " & EditPel & " , " & App & " , '" & Pesan_App & "', '" & pesan & "', GetDate())"
               
      cm.CommandText = strSql
      cm.CommandType = adCmdText
        
      Set cm.ActiveConnection = mConn
      cm.Execute
      Set cm = Nothing
      If mConn.State = adStateOpen Then mConn.Close
      Exit Sub
ErrAddCall:
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub

Public Sub AddNew_LogRemainder_CallMapping(TglFollow As Date, RangeFollow As Integer, noRangka As String, _
                                    K_Telp As Integer, Pesan_Telp As String, K_Ulang As Integer, Tanggal As Variant, jam As Variant, _
                                    K_Pel As Integer, K_Ken As Integer, K_DM As Integer, K_SMS As Integer, _
                                    App As Integer, Pesan_App As String, pesan As String)
   On Error GoTo ErrAddCall
      Dim cm As New Command
      If mConn.State = adStateClosed Then mConn.Open (SQLCon)
      strSql = "Insert Into M_ListReminderCallMapping ([TglFollow],[RangeFollow],[NoRangka],[K_Telp],[Pesan_Telp],[K_Ulang],[Tanggal],[Jam],[K_Pel],[K_Ken],[K_DM],[K_SMS],[App],[Pesan_App],[Pesan],[TglEntry]) " & _
               "Values('" & FormatTgl(TglFollow) & "' , " & RangeFollow & " , '" & noRangka & "'," & _
               " " & K_Telp & " , '" & Pesan_Telp & "' , " & K_Ulang & " , " & NullTgl(Tanggal) & " , '" & Nul(jam) & "' , " & _
               " " & K_Pel & " , " & K_Ken & " , " & K_DM & " , " & K_SMS & " , " & App & " , '" & Pesan_App & "', '" & pesan & "', GetDate())"
               
      cm.CommandText = strSql
      cm.CommandType = adCmdText
        
      Set cm.ActiveConnection = mConn
      cm.Execute
      Set cm = Nothing
      If mConn.State = adStateOpen Then mConn.Close
      Exit Sub
ErrAddCall:
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub

Public Sub AddNewItemSMS()
    Dim cm As New Command
    On Error GoTo errHandler
    'SMS... If mAConn.State = adStateClosed Then mAConn.Open (AccessCon)
    'SMS... strSQL = "Insert Into tboutbox Values(" & mNoSMS & ",'" & mPesan & "','" & mTujuan & "','" & mWaktu & "','admin')"
    
    If mAConn.State = adStateClosed Then mAConn.Open (SQLConSMS)
    strSql = "Insert Into OutBox (Tujuan,Pesan,modul,Tanggal,UserLog) Values('" & mTujuan & "', '" & mPesan & "','MRS',getdate(),'" & gSA & "')"

    cm.CommandText = strSql
    cm.CommandType = adCmdText

    Set cm.ActiveConnection = mAConn
    cm.Execute
    Set cm = Nothing
    If mAConn.State = adStateOpen Then mAConn.Close
    Exit Sub
errHandler:
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub

Public Function GetDataOutbox() As ADODB.Recordset
    On Error GoTo errHandler
    
    'SMS... If mAConn.State = adStateClosed Then mAConn.Open (AccessCon)
    'SMS... strSQL = "SELECT * FROm tboutbox" ' order by obx_id"
    
    If mAConn.State = adStateClosed Then mAConn.Open (SQLConSMS)
    strSql = "SELECT [id] FROm outbox"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    'SMS... rs.Open strSQL, AccessCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    rs.Open strSql, SQLConSMS, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set GetDataOutbox = rs
    mAConn.Close
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function GetDataSMS() As ADODB.Recordset
    On Error GoTo errHandler
    If mSConn.State = adStateClosed Then mSConn.Open (SQLConSPK)
    
    strSql = "SELECT * FROM T_TempSMS WITH(NOLOCK) Order By NomorRangka ASC"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConSPK, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set GetDataSMS = rs
    mSConn.Close
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function


Public Function FetchItemRemainderCall(noWO As String, noRangka As String, KmNext As Double) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    strSql = "SELECT * FROM M_ListReminderCall WITH(NOLOCK) WHERE LastNoWo = '" & noWO & "' and NoRangka = '" & noRangka & "' and KmNextSvc = '" & CDbl(KmNext) & "' " & _
             "Order By TglEntry ASC"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchItemRemainderCall = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Sub UpdateItemCustdat_Use(uNama As String, uAlamat As String, uKota As String, uPos As String, _
                                 uTelp As String, uGSM As String, uCDMA As String, uNoPol As String, Rangka As String)
    Dim cm As New Command
On Error GoTo errHandler

    If mConn.State = adStateClosed Then mConn.Open (SQLCon)

    strSql = "Update CUSTDAT set UseName = '" & uNama & "', UseAddr = '" & uAlamat & "', UseCity = '" & uKota & "', " & _
             "UsePos = '" & uPos & "', UsePSTN = '" & uTelp & "', UseGSM = '" & uGSM & "', UseCDMA = '" & uCDMA & "', " & _
             "No_Polisi = '" & uNoPol & "' Where No_Chassis = '" & Rangka & "'"
   
    cm.CommandText = strSql
    cm.CommandType = adCmdText
    Set cm.ActiveConnection = mConn
    cm.Execute
    
    Set cm = Nothing
    If mConn.State = adStateOpen Then mConn.Close
    Exit Sub
errHandler:
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub

Public Function Fetch_SMSnDM_Remainder(noWO As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    strSql = "SELECT Tipe, TglProses FROM M_ListReminder WITH(NOLOCK) WHERE LastNoWo = '" & noWO & "' "
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set Fetch_SMSnDM_Remainder = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Sub UpdateReminder(noWO As String, tipe As String, Remainder As Integer, Range As Integer, HargaJasa As Double, HargaPart As Double, jam As Integer, menit As Integer)
    Dim cm As New Command
On Error GoTo errHandler

    If mConn.State = adStateClosed Then mConn.Open (SQLCon)

    strSql = "Update M_ListReminder set Remainder = " & Remainder & ", TglFollow = GetDate(), RangeFollow = " & Range & ", " & _
             "EtaHargaJasa = " & Replace(HargaJasa, ",", "") & ", EtaHargaPart = " & Replace(HargaPart, ",", "") & " , EtaJam = " & jam & " , EtaMenit = " & menit & " Where LastNoWo = '" & noWO & "' And Tipe = '" & tipe & "'"
   
    cm.CommandText = strSql
    cm.CommandType = adCmdText
    Set cm.ActiveConnection = mConn
    cm.Execute
    
    Set cm = Nothing
    If mConn.State = adStateOpen Then mConn.Close
    Exit Sub
errHandler:
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub

Public Function FetchHargaJasa(noRangka As String, xHrg As Boolean, isSBI As Boolean) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    If isSBI Then
        If xHrg = True Then
             strSql = "Select c.Price, b.MDL FROM " & Log.LogSPK & "..SPK A WITH(NOLOCK) INNER JOIN " & Log.LogBengkel & "..MODEL B WITH(NOLOCK) ON A.KODE_TYPE = B.MDL " & _
                      "INNER JOIN " & Log.LogBengkel & "..CarClass c WITH(NOLOCK) On c.CLASSCODE = b.MODEL Where A.NoRangka = '" & noRangka & "'"
        Else
             strSql = "Select '1.000' as Fs_Keterangan"
        End If
    Else
        If xHrg = True Then
            
             If Log.RPLHistori = 0 Then
                strSql = "Select r.Price, c.MDL FROM Custdat c WITH(NOLOCK) INNER JOIN CarClass r On c.Model = r.ClassCode " & _
                      "Where c.No_Chassis = '" & noRangka & "'"
             Else
                strSql = "Select r.Price, c.MDL FROM Custdat c WITH(NOLOCK) INNER JOIN [" & dbReplikasiHO & "]..CarClass r On c.Model = r.ClassCode " & _
                      "Where c.No_Chassis = '" & noRangka & "'"
             End If
        Else
            
            If Log.RPLHistori = 0 Then
                strSql = "Select g.Fs_Keterangan FROM T_Regular g WITH(NOLOCK) Inner Join M_Estimasi m WITH(NOLOCK) On g.FS_NIlai = m.SBENext " & _
                      "Where m.noRangka = '" & noRangka & "'"
            Else
                strSql = "Select g.Fs_Keterangan FROM [" & dbReplikasiHO & "]..T_Regular g WITH(NOLOCK) Inner Join M_Estimasi m WITH(NOLOCK) On g.FS_NIlai = m.SBENext " & _
                      "Where m.noRangka = '" & noRangka & "'"
            End If
        End If
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchHargaJasa = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchRate(nMDL As String, nFsKet As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
 
         strSql = "Select Rate From HO_Flate WITH(NOLOCK) Where MDL = '" & nMDL & "' And FS_Item = '" & "SERVIS BERKALA " & nFsKet & " KM" & "' and FS_LOG_DEL = ''"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchRate = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchItemPart(nMDL As String, nFsKet As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    
    If Log.RPLHistori = 0 Then
        strSql = "Select Item_Part, Qty From Ho_PartPack WITH(NOLOCK) Where MDL = '" & nMDL & "' And FS_Item = '" & "SERVIS BERKALA " & nFsKet & " KM" & "'"
    Else
        strSql = "Select Item_Part, Qty From [" & dbReplikasiHO & "]..Ho_PartPack WITH(NOLOCK) Where MDL = '" & nMDL & "' And FS_Item = '" & "SERVIS BERKALA " & nFsKet & " KM" & "'"
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchItemPart = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchHeavyUse(noRangka As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
 
         strSql = "Select HeavyUse From Custdat WITH(NOLOCK) Where No_Chassis = '" & Trim(noRangka) & "'"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchHeavyUse = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchDataRemainder1000(TglNext0 As Date, TglNext1 As Date, nRem As Integer, TglRem0 As Date, TglRem1 As Date) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConSPK)
    
    strSql = "Select DISTINCT " & _
                "b.SPK_No, f.Tgl_BPK as TglDO, isnull(i.TglCR1,'01/01/1900') as TglCR1, isnull(i.KmCR1,0) as KmCR1, i.TglCR2, isNull(i.KMCR2,0) as KmCR2 , b.SNama as UseName, " & _
                "b.SAlamat as UseAddr, b.SKota as UseCity, b.SKodePos as UsePos, a.NoPol as NoPolisi, " & _
                "h.NoMesin as NoEngine, b.NoRangka as NoChassis, b.STelp as Telpon, b.SGsm as GSM, b.SCdma as CDMA, " & _
                "d.KodeKategory as KodeType, d.NamaKategory as NamaType, e.KodeWarna, e.NamaEnglish as NamaWarna, " & _
                "g.NamaSales as Wiraniaga, b.Nama_Faktur as NamaSTNK, " & _
                "(a.Alamat1 + a.Alamat2 + a.Alamat3 + a.Alamat4 + a.Alamat5) As AlamatSTNK, " & _
                "b.Kota as KotaSTNK, a.KodePos as KodePosSTNK, a.TmptLahir as TempatLahir, " & _
                "a.TglLahir as TanggalLahir, a.Agama, a.Pekerjaan, a.Hoby, i.Tgl1, b.NoRangka as No_Chassis, " & _
                "B.Kode_Lgn , h.KodeTahun, e.NamaInd "
        ''2017-08-14
        strSql = strSql & " ,B.nama_pemesan 'Namaspk', b.alamat1+' '+b.alamat2 'Alamatspk', b.telp_rumah 'Telpspk', b.hp1 'GSM1spk', b.hp2 'GSM2spk'"
   
        strSql = strSql & " From " & Log.LogSPK & "..AFI a WITH(NOLOCK) INNER JOIN  " & Log.LogSPK & "..SPK b WITH(NOLOCK) ON a.SPK_No = b.SPK_No " & _
                "INNER JOIN " & Log.LogSPK & "..INMerek c WITH(NOLOCK) ON b.Kode_Merk = c.KodeMerek " & _
                "INNER JOIN " & Log.LogSPK & "..InKategory d WITH(NOLOCK) ON b.Kode_Type = d.KodeKategory " & _
                "INNER JOIN " & Log.LogSPK & "..InWarnaExt e WITH(NOLOCK) ON b.Kode_Warna = e.KodeWarna " & _
                "INNER JOIN " & Log.LogSPK & "..BPK f WITH(NOLOCK) ON a.SPK_No = f.SPK_No " & _
                "INNER JOIN " & Log.LogSPK & "..ArSales g WITH(NOLOCK) ON b.Kode_Sales = g.KodeSales " & _
                "INNER JOIN " & Log.LogSPK & "..INMst h WITH(NOLOCK) ON b.NoRangka = h.kodeItem " & _
                "INNER JOIN " & Log.LogSPK & "..T_MRS i WITH(NOLOCK) ON b.SPK_No = i.NoSPK " & _
                "LEFT JOIN " & Log.LogBengkel & "..History y WITH(NOLOCK) ON b.NoRangka = y.No_Chassis And b.norangka not in (SELECT NO_CHASSIS FROM " & Log.LogBengkel & "..HISTORY WHERE FS_RC = 'R001') " & _
                "LEFT JOIN " & Log.LogBengkel & "..M_ListReminderCall l WITH(NOLOCK) ON b.NoRangka = l.NoRangka AND K_Ulang = 1 AND KMNEXTSVC = '1000' " & _
                "Where i.Km1 = '1000' And a.batal = 0 AND "
            
            
            
            '(y.FS_RC='')
             If nRem = 1 Then
                strSql = strSql & "(i.Tgl1 <= '" & FormatTgl(TglNext1) & "') OR (l.Tanggal <= '" & FormatTgl(TglRem1) & "') " & _
                         "ORDER BY i.Tgl1, b.SPK_No ASC"
             Else
                strSql = strSql & "(i.Tgl1 between '" & FormatTgl(TglNext0) & "' And '" & FormatTgl(TglNext1) & "') OR (l.Tanggal between '" & FormatTgl(TglRem0) & "' And '" & FormatTgl(TglRem1) & "') " & _
                         "ORDER BY i.Tgl1, b.SPK_No ASC"
             End If
             
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConSPK, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchDataRemainder1000 = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchListRemainder1000(noWO As String, isCall As Integer) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    Select Case isCall
        Case 0
            ' ALL
            strSql = "SELECT Tipe, TglProses FROM M_ListReminder WITH(NOLOCK) WHERE LastNoWo = '" & noWO & "'"
        Case 1
            ' SMS
            strSql = "SELECT Tipe, TglProses FROM M_ListReminder WITH(NOLOCK) WHERE LastNoWo = '" & noWO & "' AND TIPE = '1K'"
        Case 2
            ' CALL
            strSql = "SELECT TglFollow FROM M_ListReminderCall WITH(NOLOCK) WHERE LastNoWo = '" & noWO & "' Order By TglFollow DESC"
        Case 3
            ' DM
            strSql = "SELECT Tipe, TglProses FROM M_ListReminder WITH(NOLOCK) WHERE LastNoWo = '" & noWO & "' AND TIPE = 'DM'"
    End Select
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchListRemainder1000 = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Sub RemoveItemSMS(noRangka)
    Dim cm As New ADODB.Command
    Dim strSql As String

    On Error GoTo errHandler

    If mSConn.State = adStateClosed Then mSConn.Open (SQLConSPK)

    strSql = "DELETE From T_TempSMS WHERE NomorRangka=?"
    
    cm.CommandText = strSql
    cm.CommandType = adCmdText
    cm.Parameters.Append cm.CreateParameter(, adVarChar, adParamInput, 30, noRangka)
    
    Set cm.ActiveConnection = mSConn
    cm.Execute
    Set cm = Nothing
    If mSConn.State <> adStateClosed Then mSConn.Close
    Exit Sub
errHandler:
    If Not cm Is Nothing Then Set cm = Nothing
    err.Raise err.Number, err.Source, err.Description
End Sub

Public Sub RemoveAllItemSMS()
    Dim cm As New ADODB.Command
    Dim strSql As String

    On Error GoTo errHandler
    If mSConn.State = adStateClosed Then mSConn.Open (SQLConSPK)

    strSql = "DELETE From T_TempSMS"
    
    cm.CommandText = strSql
    cm.CommandType = adCmdText
    
    Set cm.ActiveConnection = mSConn
    cm.Execute
    Set cm = Nothing
    If mSConn.State <> adStateClosed Then mSConn.Close
    Exit Sub
errHandler:
    If Not cm Is Nothing Then Set cm = Nothing
    err.Raise err.Number, err.Source, err.Description
End Sub


Public Sub AddNewItemSMS_Temp()
    Dim cm As New Command
    On Error GoTo errHandler
    If mSConn.State = adStateClosed Then mSConn.Open (SQLConSPK)
    strSql = "Insert Into T_TempSMS Values('" & mNomorRangka & "','" & mNoGSM & "','" & mNoCDMA & "','" & FormatTgl(mTglSendNow) & "','" & mPesanSMS & "')"
    
    cm.CommandText = strSql
    cm.CommandType = adCmdText

    Set cm.ActiveConnection = mSConn
    cm.Execute
    Set cm = Nothing
    If mSConn.State <> adStateClosed Then mSConn.Close
    Exit Sub
errHandler:
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub

Public Function MyNoSPK(xNoRangka) As String
    On Error GoTo errHandler
    If mSConn.State = adStateClosed Then mSConn.Open (SQLConSPK)
    
    strSql = "Select SPK_No FROM SPK WITH(NOLOCK) Where noRangka= '" & xNoRangka & "'"

             
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConSPK, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    MyNoSPK = rs.Fields(0).value
    If mSConn.State = adStateOpen Then mSConn.Close
        
    Exit Function
errHandler:
End Function

Public Function MyNoWo(xNoRangka) As String
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    strSql = "Select NoWo FROM M_Estimasi WITH(NOLOCK) Where noRangka= '" & xNoRangka & "' AND NOINV <> '' And KmNow <> 0 Order By SBENEXT DESC"

             
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    If rs.RecordCount <> 0 Then rs.MoveFirst
    MyNoWo = rs.Fields(0).value
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
End Function

Public Function MyKmNext(xNoRangka) As String
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    strSql = "Select KMNEXT FROM M_Estimasi WITH(NOLOCK) Where noRangka= '" & xNoRangka & "' AND NOINV <> '' And KmNow <> 0 Order By SBENEXT DESC"

             
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    If rs.RecordCount <> 0 Then rs.MoveFirst
    MyKmNext = rs.Fields(0).value
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
End Function

Public Function MyTglNext(xNoRangka) As Date
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    strSql = "Select ((CASE WHEN JT_Tgl > JT_Bln then JT_Bln WHEN JT_Tgl < JT_Bln then JT_Tgl WHEN JT_Tgl = JT_Bln then Jt_Bln END) AS TGLNEXT FROM M_Estimasi WITH(NOLOCK) Where noRangka= '" & xNoRangka & "' AND NOINV <> '' And KmNow <> 0 Order By SBENEXT DESC"

             
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    If rs.RecordCount <> 0 Then rs.MoveFirst
    MyTglNext = Format(rs.Fields(0).value, "DD/MM/YYYY")
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
End Function

Public Function myModel(xNoRangka) As String
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    strSql = "SELECT MDL FROM CUSTDAT WITH(NOLOCK) WHERE NO_CHASSIS = '" & Trim(xNoRangka) & "'"
             
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    If rs.RecordCount <> 0 Then rs.MoveFirst
    myModel = rs.Fields(0).value
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
End Function

Public Sub UpdatePulsa()
    Dim cm As New Command
On Error GoTo errHandler

    If mConn.State = adStateClosed Then mConn.Open (SQLCon)

    strSql = "Update M_Pulsa Set Pulsa = (select Pulsa From M_Pulsa) - (select Tarif From M_Pulsa)"
   
    cm.CommandText = strSql
    cm.CommandType = adCmdText
    Set cm.ActiveConnection = mConn
    cm.Execute
    
    Set cm = Nothing
    If mConn.State = adStateOpen Then mConn.Close
    Exit Sub
errHandler:
    If Not cm Is Nothing Then Set cm = Nothing
        err.Raise err.Number, err.Source, err.Description
End Sub

Public Function FetchItem_1K(nRangka As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    strSql = "Select FS_RC FROM JobDTA WITH(NOLOCK) Where No_Chassis = '" & nRangka & "'"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchItem_1K = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function GetData1K(nRangka As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    strSql = "Select d.NoInv, d.No_Wo, CAST(d.[Date] as Char(40)) as Tanggal, " & _
             "d.KM, d.Customer, d.Cust_adres, d.No_Polisi, d.No_Engine, d.No_Chassis, " & _
             "b.Telpon, b.GSM, b.CDMA, d.REP1, d.REP2, d.REP3, d.REP4, d.REP5, d.REP6, d.REP7, d.REP8, " & _
             "d.RTE1 , d.RTE2, d.RTE3, d.RTE4, d.RTE5, d.RTE6, d.RTE7, d.RTE8, useEmail " & _
             "FROM JOBDTA d WITH(NOLOCK) Inner Join Custdat b WITH(NOLOCK) On d.No_Chassis = b.No_Chassis " & _
             "AND d.FS_LOG_DEL = '' LEFT JOIN HISTORY a WITH(NOLOCK) on a.No_Wo = d.No_Wo AND a.FS_LOG_DEL = '' " & _
             "where d.FS_RC = 'R001' And  d.No_Chassis = '" & nRangka & "'"
 
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set GetData1K = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function pricePartMonth(xNoPart As String, xQtyPart As Integer) As Double
    On Error GoTo errHandler
    adoConNis = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & NisUSER & ";password=" & NisPass & ";Initial Catalog=" & NisDb & ";server=" & NisKomp
    If mConn.State = adStateClosed Then mConn.Open (adoConNis)
    
    pricePartMonth = 0
    strSql = "SELECT ISNULL(HARGAJUALA,0) AS Price FROM INMST WITH(NOLOCK) WHERE KODEITEM LIKE '%" & Trim(xNoPart) & "' AND KODEGDG = '03'"
             
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, adoConNis, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    If rs.RecordCount <> 0 Then rs.MoveFirst
    pricePartMonth = CDbl(rs.Fields(0).value) * CDbl(xQtyPart)
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
End Function

Public Function pasifCustomer(xNoChassis As String) As Boolean
Dim nCount As Integer

    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    nCount = 0
    pasifCustomer = False
    strSql = "SELECT TOP 3 * FROM M_ESTIMASI WITH(NOLOCK) WHERE NoRangka = '" & Trim(xNoChassis) & "' AND ISNULL(TGLSVCNOW,'') <> '' AND (CASE WHEN JT_TGL < JT_BLN THEN JT_TGL ELSE JT_BLN END) <= '" & Format(isnow, "MM/DD/YYYY") & "'  ORDER BY SBENEXT DESC"
             
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    If rs.RecordCount <> 0 Then rs.MoveLast
    Do While Not rs.BOF
        If Trim(rs("NOINV")) = "" Then
            nCount = nCount + 1
        Else
            nCount = 0
        End If
        rs.MovePrevious
    Loop
    If nCount >= 2 Then pasifCustomer = True
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
End Function

Public Function searchLastSBE(xNoChassis As String) As String

    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    searchLastSBE = ""
    
    
    If Log.RPLHistori = 0 Then
        strSql = "SELECT TOP 2 SBENEXT, FS_RC FROM M_ESTIMASI A WITH(NOLOCK) INNER JOIN T_REGULAR B WITH(NOLOCK) ON A.SBENEXT = B.FS_NILAI WHERE NoRangka = '" & Trim(xNoChassis) & "' AND ISNULL(TGLSVCNOW,'') <> '' ORDER BY SBENEXT DESC"
    Else
        strSql = "SELECT TOP 2 SBENEXT, FS_RC FROM M_ESTIMASI A WITH(NOLOCK) INNER JOIN [" & dbReplikasiHO & "]..T_REGULAR B WITH(NOLOCK) ON A.SBENEXT = B.FS_NILAI WHERE NoRangka = '" & Trim(xNoChassis) & "' AND ISNULL(TGLSVCNOW,'') <> '' ORDER BY SBENEXT DESC"
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    If rs.RecordCount <> 0 Then rs.MoveLast
    If Not rs.BOF Then
        searchLastSBE = Trim(rs("FS_RC"))
    End If
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
End Function

Public Function searchLastSBE2(xNoChassis As String) As Date

    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    searchLastSBE2 = "01/01/1900"
    strSql = "SELECT TOP 1 TGLSVCNOW FROM M_ESTIMASI WITH(NOLOCK) WHERE NoRangka = '" & Trim(xNoChassis) & "' AND ISNULL(TGLSVCNOW,'') <> '' ORDER BY SBENEXT DESC"
             
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    If rs.RecordCount <> 0 Then rs.MoveLast
    If Not rs.BOF Then
        searchLastSBE2 = Format(rs("TGLSVCNOW"), "MM/DD/YYYY")
    End If
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
End Function

Public Function searchLastSBE3(xNoChassis As String) As Double

    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
    
    searchLastSBE3 = 0
    strSql = "SELECT TOP 1 KMNOW FROM M_ESTIMASI WITH(NOLOCK) WHERE NoRangka = '" & Trim(xNoChassis) & "' AND ISNULL(TGLSVCNOW,'') <> '' ORDER BY SBENEXT DESC"
             
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    If rs.RecordCount <> 0 Then rs.MoveLast
    If Not rs.BOF Then
        searchLastSBE3 = CDbl(rs("KMNOW"))
    End If
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
End Function

'' -------------------------------------------------
'' HANSA RENKEI
'' -------------------------------------------------

Public Function FetchDataRemainder_HansaRenkeiMRA(Data As Boolean, TglNext0 As Date, TglNext1 As Date, nRem As Integer, Optional TglRem0 As Date, Optional TglRem1 As Date, Optional ketr As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCon)
     
    If Data = True Then
        strSql = "SELECT a.NoInv, a.NoWo, a.TglSvcNow as DateLastSvc,a.KmNow, " & _
             "(CASE WHEN A.JT_Tgl > A.JT_Bln then A.JT_Bln WHEN A.JT_Tgl < A.JT_Bln then A.JT_Tgl WHEN A.JT_Tgl = A.JT_Bln then A.Jt_Bln END) as DateNextSvc, " & _
             "a.KmNext as KmNextSvc, c.UseName, c.UseAddr, c.UseCity, c.UsePos, c.No_Polisi as NoPolisi, C.FMDL as KetModel, " & _
             "a.NoRangka as NoChassis, c.UsePSTN as Telpon , c.UseGSM, c.UseCDMA, a.NoRangka as NoChassisSMS, c.FS_KD_PLG as KodeLGN, " & _
             "b.FS_RC, A.SbeNext, C.COLOR as WrnModel, C.USEEMAIL AS EMAIL, C.MDL, C.MODEL, C.FS_TH_MOBIL, C.COLOR, C.FS_KD_DEALER, C.FS_NM_DEALER FROM M_ESTIMASI A WITH(NOLOCK) Inner Join " & _
             "(SELECT G.NORANGKA, (SELECT MAX(SBENEXT) FROM m_estimasi WITH(NOLOCK) WHERE NORANGKA=G.NORANGKA AND ISNULL(TGLSVCNOW,'') <> '') S, MAX((CASE WHEN G.JT_Tgl > G.JT_Bln then G.JT_Bln WHEN G.JT_Tgl < G.JT_Bln then G.JT_Tgl WHEN G.JT_Tgl = G.JT_Bln then G.JT_Tgl END)) AS JT_TGL " & _
             "FROM M_ESTIMASI G WITH(NOLOCK) Where (CASE WHEN G.JT_Tgl > G.JT_Bln then G.JT_Bln WHEN G.JT_Tgl < G.JT_Bln then G.JT_Tgl WHEN G.JT_Tgl = G.JT_Bln then G.JT_Tgl END) Is Not Null GROUP BY G.NORANGKA) AS Z ON A.NORANGKA = Z.NORANGKA AND A.SBENEXT = Z.S " & _
             "INNER JOIN History b WITH(NOLOCK) ON a.NoWo = b.No_Wo AND A.NORANGKA = B.NO_CHASSIS AND B.FS_LOG_DEL = '' " & _
             "INNER JOIN Custdat c WITH(NOLOCK) On a.NoRangka = c.No_Chassis Where "
             
        strSql = "SELECT a.NoInv, a.NoWo, a.TglSvcNow as DateLastSvc,a.KmNow, " & _
             "(CASE WHEN A.JT_Tgl > A.JT_Bln then A.JT_Bln WHEN A.JT_Tgl < A.JT_Bln then A.JT_Tgl WHEN A.JT_Tgl = A.JT_Bln then A.Jt_Bln END) as DateNextSvc, " & _
             "a.KmNext as KmNextSvc, c.UseName, c.UseAddr, c.UseCity, c.UsePos, c.No_Polisi as NoPolisi, C.FMDL as KetModel, " & _
             "a.NoRangka as NoChassis, c.UsePSTN as Telpon , c.UseGSM, c.UseCDMA, a.NoRangka as NoChassisSMS, c.FS_KD_PLG as KodeLGN, " & _
             "b.FS_RC, A.SbeNext, C.COLOR as WrnModel, C.USEEMAIL AS EMAIL, C.MDL, C.MODEL, C.FS_TH_MOBIL, C.COLOR, C.FS_KD_DEALER, C.FS_NM_DEALER " & _
             " ,r.program_tab,r2.preferred_contact_channel," & _
             " r.reminder_message,r.priority_call,r.extended_warranty_status,r.customer_habit,r.last_habit," & _
             " r.next_service_status,r.ncs_status,r.next_service_stage,r.reminder_ID FROM M_ESTIMASI A WITH(NOLOCK) Inner Join " & _
             "(SELECT G.NORANGKA, (SELECT MAX(SBENEXT) FROM m_estimasi WITH(NOLOCK) WHERE NORANGKA=G.NORANGKA AND ISNULL(TGLSVCNOW,'') <> '') S, MAX((CASE WHEN G.JT_Tgl > G.JT_Bln then G.JT_Bln WHEN G.JT_Tgl < G.JT_Bln then G.JT_Tgl WHEN G.JT_Tgl = G.JT_Bln then G.JT_Tgl END)) AS JT_TGL " & _
             "FROM M_ESTIMASI G WITH(NOLOCK) Where (CASE WHEN G.JT_Tgl > G.JT_Bln then G.JT_Bln WHEN G.JT_Tgl < G.JT_Bln then G.JT_Tgl WHEN G.JT_Tgl = G.JT_Bln then G.JT_Tgl END) Is Not Null GROUP BY G.NORANGKA) AS Z ON A.NORANGKA = Z.NORANGKA AND A.SBENEXT = Z.S " & _
             "INNER JOIN History b WITH(NOLOCK) ON a.NoWo = b.No_Wo AND A.NORANGKA = B.NO_CHASSIS AND B.FS_LOG_DEL = '' " & _
             "INNER JOIN Custdat c WITH(NOLOCK) On a.NoRangka = c.No_Chassis left join ReplBengkel99..p2_reminder r on a.NoRangka=r.vin" & _
             " left join ReplBengkel99..p2_reminder_preferred_contact_channel r2 on r.event_ID=r2.event_ID Where "
        '' menambahkan tidak menambahkan norangka jika sudah / lbh dari sbe yang direminder
        strSql = strSql & " A.SBENEXT > isnull((select MAX(BB.FS_NILAI) from [" & dbReplikasiHO & "]..JOBDTA AA left join [" & dbReplikasiHO & "]..T_REGULAR BB on BB.FS_RC = AA.FS_RC where AA.NO_CHASSIS = A.NORANGKA and FS_LOG_DEL = ''),0) and "
             '' menambahkan norangka dan sbenext yang tidak ada di tbSudahServiceHansaRenkei
             
             ''strSql = strSql & " A.SBENEXT+A.NORANGKA NOT IN (SELECT SbeNext+norangka FROM tbSudahServiceHansaRenkei BB WITH (NOLOCK) WHERE bb.norangka = A.NORANGKA ) and "
             ''server majapahit sbenext error converting
             ''strSql = strSql & " A.SBENEXT+A.NORANGKA NOT IN (SELECT SbeNext+norangka FROM tbSudahServiceHansaRenkei BB WITH (NOLOCK) WHERE bb.norangka = A.NORANGKA ) and "
             strSql = strSql & " cast(A.SBENEXT as varchar(200))+A.NORANGKA NOT IN (SELECT SbeNext+norangka FROM tbSudahServiceHansaRenkei BB WITH (NOLOCK) WHERE bb.norangka = A.NORANGKA ) and "
             
             If nRem = 1 Then
                strSql = strSql & "(CASE WHEN A.JT_Tgl > A.JT_Bln then  A.JT_Bln WHEN A.JT_Tgl < A.JT_Bln then  A.JT_Tgl WHEN A.JT_Tgl = A.JT_Bln then  A.Jt_Bln END) <= '" & FormatTgl(TglNext1) & "' "
                         '"Order By a.noRangka, a.SBENext DESC"
             Else
                strSql = strSql & "(CASE WHEN A.JT_Tgl > A.JT_Bln then  A.JT_Bln WHEN A.JT_Tgl < A.JT_Bln then  A.JT_Tgl WHEN A.JT_Tgl = A.JT_Bln then  A.Jt_Bln END) between '" & FormatTgl(TglNext0) & "' And '" & FormatTgl(TglNext1) & "' "
                         '"Order By a.noRangka, a.SBENext DESC"
             End If
             strSql = strSql & ketr
    Else
        strSql = "SELECT NOINV, NO_WO AS NOWO, DATE AS DATELASTSVC, KM AS KMNOW, DATENEXTSVC, KMNEXTSVC, L.UseName, L.UseAddr, L.UseCity, L.UsePos, L.No_Polisi as NoPolisi, L.FMDL as KetModel, " & _
                      "K.NoRangka as NoChassis, L.UsePSTN as Telpon , L.UseGSM, L.UseCDMA, K.NoRangka as NoChassisSMS, L.FS_KD_PLG as KodeLGN, " & _
                      "M.FS_RC, CAST ((CAST(RIGHT(M.FS_RC,3) AS INTEGER) + (KMNEXTSVC - KM )/1000) * 1000 AS INTEGER) AS SBENEXT, L.COLOR as WrnModel, L.USEEMAIL AS EMAIL, L.MDL, L.MODEL, L.FS_TH_MOBIL, L.COLOR, L.FS_KD_DEALER, L.FS_NM_DEALER " & _
                      "FROM M_LISTREMINDERCALL K WITH(NOLOCK) LEFT JOIN HISTORY M WITH(NOLOCK) ON K.NORANGKA = M.NO_CHASSIS AND K.LASTNOWO = M.NO_WO " & _
                      "INNER JOIN Custdat L WITH(NOLOCK) On K.NoRangka = L.No_Chassis WHERE TANGGAL BETWEEN '" & FormatTgl(TglRem0) & "' AND '" & FormatTgl(TglRem1) & "' AND K.NoRangka <> '' And K.KMNextSvc <> 1000 " & _
                      "ORDER BY k.noRangka, KMNEXTSVC DESC"
        strSql = "SELECT NOINV, NO_WO AS NOWO, DATE AS DATELASTSVC, KM AS KMNOW, DATENEXTSVC, KMNEXTSVC, L.UseName, L.UseAddr, L.UseCity, L.UsePos, L.No_Polisi as NoPolisi, L.FMDL as KetModel, " & _
                      "K.NoRangka as NoChassis, L.UsePSTN as Telpon , L.UseGSM, L.UseCDMA, K.NoRangka as NoChassisSMS, L.FS_KD_PLG as KodeLGN, " & _
                      "M.FS_RC, CAST ((CAST(RIGHT(M.FS_RC,3) AS INTEGER) + (KMNEXTSVC - KM )/1000) * 1000 AS INTEGER) AS SBENEXT, L.COLOR as WrnModel, L.USEEMAIL AS EMAIL, L.MDL, L.MODEL, L.FS_TH_MOBIL, L.COLOR, L.FS_KD_DEALER, L.FS_NM_DEALER " & _
                      "FROM M_LISTREMINDERCALL K WITH(NOLOCK) LEFT JOIN HISTORY M WITH(NOLOCK) ON K.NORANGKA = M.NO_CHASSIS AND K.LASTNOWO = M.NO_WO " & _
                      "INNER JOIN Custdat L WITH(NOLOCK) On K.NoRangka = L.No_Chassis WHERE 1=0"
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchDataRemainder_HansaRenkeiMRA = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function


Public Function FetchDataRemainder_HansaRenkeiSA(Data As Boolean, TglNext0 As Date, TglNext1 As Date, nRem As Integer, Optional TglRem0 As Date, Optional TglRem1 As Date, Optional ketr As String) As ADODB.Recordset
1         On Error GoTo errHandler
2         If mConn.State = adStateClosed Then mConn.Open (SQLCon)
           
3         If Data = True Then
              
4             strSql = "SELECT * FROM ( SELECT a.NoInv, a.NoWo, a.TglSvcNow as DateLastSvc,a.KmNow, " & _
                   "(CASE WHEN A.JT_Tgl > A.JT_Bln then A.JT_Bln WHEN A.JT_Tgl < A.JT_Bln then A.JT_Tgl WHEN A.JT_Tgl = A.JT_Bln then A.Jt_Bln END) as DateNextSvc, " & _
                   "a.KmNext as KmNextSvc, c.UseName, c.UseAddr, c.UseCity, c.UsePos, c.No_Polisi as NoPolisi, C.FMDL as KetModel, " & _
                   "a.NoRangka as NoChassis, c.UsePSTN as Telpon , c.UseGSM, c.UseCDMA, a.NoRangka as NoChassisSMS, c.FS_KD_PLG as KodeLGN, " & _
                   "b.FS_RC, A.SbeNext, C.COLOR as WrnModel, C.USEEMAIL AS EMAIL, C.MDL, C.MODEL, C.FS_TH_MOBIL, C.COLOR, C.FS_KD_DEALER, C.FS_NM_DEALER,(SELECT TOP 1 FRONT FROM T_INVOICE WHERE NO_CHASSIS = A.NoRangka AND FS_LOG_DEL = '' ORDER BY INVDATE desc) lastSa FROM M_ESTIMASI A WITH(NOLOCK) Inner Join " & _
                   "(SELECT G.NORANGKA, (SELECT MAX(SBENEXT) FROM m_estimasi WITH(NOLOCK) WHERE NORANGKA=G.NORANGKA AND ISNULL(TGLSVCNOW,'') <> '') S, MAX((CASE WHEN G.JT_Tgl > G.JT_Bln then G.JT_Bln WHEN G.JT_Tgl < G.JT_Bln then G.JT_Tgl WHEN G.JT_Tgl = G.JT_Bln then G.JT_Tgl END)) AS JT_TGL " & _
                   "FROM M_ESTIMASI G WITH(NOLOCK) Where (CASE WHEN G.JT_Tgl > G.JT_Bln then G.JT_Bln WHEN G.JT_Tgl < G.JT_Bln then G.JT_Tgl WHEN G.JT_Tgl = G.JT_Bln then G.JT_Tgl END) Is Not Null GROUP BY G.NORANGKA) AS Z ON A.NORANGKA = Z.NORANGKA AND A.SBENEXT = Z.S " & _
                   "INNER JOIN History b WITH(NOLOCK) ON a.NoWo = b.No_Wo AND A.NORANGKA = B.NO_CHASSIS AND B.FS_LOG_DEL = '' " & _
                   "INNER JOIN Custdat c WITH(NOLOCK) On a.NoRangka = c.No_Chassis "
                   ''20220121 revisi mb thania terkait reminder sa
               
5             strSql = "SELECT * FROM ( SELECT a.NoInv, a.NoWo, a.TglSvcNow as DateLastSvc,a.KmNow, " & _
                   "(CASE WHEN A.JT_Tgl > A.JT_Bln then A.JT_Bln WHEN A.JT_Tgl < A.JT_Bln then A.JT_Tgl WHEN A.JT_Tgl = A.JT_Bln then A.Jt_Bln END) as DateNextSvc, " & _
                   "a.KmNext as KmNextSvc, c.UseName, c.UseAddr, c.UseCity, c.UsePos, c.No_Polisi as NoPolisi, C.FMDL as KetModel, " & _
                   "a.NoRangka as NoChassis, c.UsePSTN as Telpon , c.UseGSM, c.UseCDMA, a.NoRangka as NoChassisSMS, c.FS_KD_PLG as KodeLGN, " & _
                   "b.FS_RC, A.SbeNext, C.COLOR as WrnModel, C.USEEMAIL AS EMAIL, C.MDL, C.MODEL, C.FS_TH_MOBIL, C.COLOR, C.FS_KD_DEALER, C.FS_NM_DEALER,(SELECT TOP 1 FRONT FROM T_INVOICE WHERE NO_CHASSIS = A.NoRangka AND FS_LOG_DEL = '' ORDER BY INVDATE desc) lastSa " & _
                   " FROM M_ESTIMASI A WITH(NOLOCK) Inner Join " & _
                   "(SELECT G.NORANGKA, b.KDUser MRA, c.KDUser SA, MAX(SBENEXT) S, " & _
                   "   MAX((CASE WHEN G.JT_Tgl > G.JT_Bln then G.JT_Bln WHEN G.JT_Tgl < G.JT_Bln then G.JT_Tgl WHEN G.JT_Tgl = G.JT_Bln then G.JT_Tgl END)) AS JT_TGL " & _
                   "   FROM M_ESTIMASI G WITH(NOLOCK) " & _
                   "  left join ( " & _
                   "       select norangka, LastNoWO, TipeUser, KDUser, KmNextSvc " & _
                   "       From M_ListRemindercall " & _
                   "       where TipeUser not in ('SA') " & _
                   "   )b on g.NoRangka = b.NoRangka and g.KmNext = b.KmNextSvc " & _
                   "   left join ( " & _
                   "       select norangka, LastNoWO, TipeUser, KDUser, KmNextSvc " & _
                   "       From M_ListRemindercall " & _
                   "       where TipeUser in ('SA') " & _
                   "   )c on g.NoRangka = c.NoRangka and g.KmNext = c.KmNextSvc " & _
                   "   Where (CASE WHEN G.JT_Tgl > G.JT_Bln then G.JT_Bln WHEN G.JT_Tgl < G.JT_Bln then G.JT_Tgl WHEN G.JT_Tgl = G.JT_Bln then G.JT_Tgl END) Is Not Null " & _
                   "   and ISNULL(TGLSVCNOW,'') <> ''  and c.KDUser is null " & _
                   "   GROUP BY G.NORANGKA, b.KDUser, c.KDUser) AS Z ON A.NORANGKA = Z.NORANGKA AND A.SBENEXT = Z.S  " & _
                   "INNER JOIN History b WITH(NOLOCK) ON a.NoWo = b.No_Wo AND A.NORANGKA = B.NO_CHASSIS AND B.FS_LOG_DEL = '' " & _
                   "INNER JOIN Custdat c WITH(NOLOCK) On a.NoRangka = c.No_Chassis left join ReplBengkel99..p2_reminder r on a.NoRangka=r.vin" & _
                   " left join ReplBengkel99..p2_reminder_preferred_contact_channel r2 on r.event_ID=r2.event_ID "
              '' menambahkan tidak menambahkan norangka jika sudah / lbh dari sbe yang direminder
6                 strSql = strSql & " where A.SBENEXT > isnull((select MAX(BB.FS_NILAI) from [" & dbReplikasiHO & "]..T_INVOICEnew AA left join [" & dbReplikasiHO & "]..T_REGULAR BB on BB.FS_RC = AA.FS_RC where AA.NO_CHASSIS = A.NORANGKA and FS_LOG_DEL = ''),0) and "
                  
                   '' menambahkan norangka dan sbenext yang tidak ada di tbSudahServiceHansaRenkei
                   ''server majapahit sbenext error converting
                   ''strSql = strSql & " A.SBENEXT+A.NORANGKA NOT IN (SELECT SbeNext+norangka FROM tbSudahServiceHansaRenkei BB WITH (NOLOCK) WHERE bb.norangka = A.NORANGKA ) and "
                   
7                  strSql = strSql & " cast(A.SBENEXT as varchar(200))+A.NORANGKA NOT IN (SELECT SbeNext+norangka FROM tbSudahServiceHansaRenkei BB WITH (NOLOCK) WHERE bb.norangka = A.NORANGKA ) and "
                   
8                  If nRem = 1 Then
9                     strSql = strSql & "(CASE WHEN A.JT_Tgl > A.JT_Bln then  A.JT_Bln WHEN A.JT_Tgl < A.JT_Bln then  A.JT_Tgl WHEN A.JT_Tgl = A.JT_Bln then  A.Jt_Bln END) <= '" & FormatTgl(TglNext1) & "' "
                               '"Order By a.noRangka, a.SBENext DESC"
10                 Else
11                    strSql = strSql & "(CASE WHEN A.JT_Tgl > A.JT_Bln then  A.JT_Bln WHEN A.JT_Tgl < A.JT_Bln then  A.JT_Tgl WHEN A.JT_Tgl = A.JT_Bln then  A.Jt_Bln END) between '" & FormatTgl(TglNext0) & "' And '" & FormatTgl(TglNext1) & "' "
                               '"Order By a.noRangka, a.SBENext DESC"
12                 End If
13                 strSql = strSql & ketr
14                 strSql = strSql & " ) Z WHERE LASTSA = '" & gSA & "'"
15                 strSql = strSql & " Union"
16                 strSql = strSql & " select  a.NoInv,a.NoWo,a.TglSvcNow as datelastsvc,a.KmNow,a.JT_Tgl datenextsvc, a.KmNext as kmnextsvc,"
17                 strSql = strSql & " c.UseName, c.UseAddr, c.UseCity, c.UsePos, c.No_Polisi as NoPolisi, C.FMDL as KetModel,"
18                 strSql = strSql & " a.NoRangka as NoChassis, c.UsePSTN as Telpon , c.UseGSM, c.UseCDMA, a.NoRangka as NoChassisSMS,"
19                 strSql = strSql & " c.FS_KD_PLG as KodeLGN, b.FS_RC, A.SbeNext, C.COLOR as WrnModel, C.USEEMAIL AS EMAIL, C.MDL,"
20                 strSql = strSql & " C.model , C.FS_TH_MOBIL, C.Color, C.FS_KD_DEALER, C.FS_NM_DEALER,a.SA AS LASTSA"
21                 strSql = strSql & " FROM M_ESTIMASIInject a"
22                 strSql = strSql & " INNER JOIN History b WITH(NOLOCK) ON a.NoWo = b.No_Wo AND A.NORANGKA = B.NO_CHASSIS AND B.FS_LOG_DEL = ''"
23                 strSql = strSql & " INNER JOIN Custdat c WITH(NOLOCK) On a.NoRangka = c.No_Chassis"
24                 strSql = strSql & " where"
25                 strSql = strSql & " A.SBENEXT > isnull((select MAX(BB.FS_NILAI) from [ReplBengkel99]..T_INVOICEnew AA left join [ReplBengkel99]..T_REGULAR BB on BB.FS_RC = AA.FS_RC where AA.NO_CHASSIS = A.NORANGKA and FS_LOG_DEL = ''),0)"
                   
26                 strSql = strSql & " and  cast(A.SBENEXT as varchar(200))+A.NORANGKA NOT IN (SELECT SbeNext+norangka FROM tbSudahServiceHansaRenkei BB WITH (NOLOCK) WHERE bb.norangka = A.NORANGKA )"
27                 If nRem = 1 Then
28                    strSql = strSql & " and A.JT_Tgl <= '" & FormatTgl(TglNext1) & "' and a.SA = '" & gSA & "'"
29                 Else
30                    strSql = strSql & " and A.JT_Tgl Between '" & FormatTgl(TglNext0) & "' And '" & FormatTgl(TglNext1) & "' and a.SA = '" & gSA & "'"
31                 End If
                   
32        Else
33            strSql = "SELECT NOINV, NO_WO AS NOWO, DATE AS DATELASTSVC, KM AS KMNOW, DATENEXTSVC, KMNEXTSVC, L.UseName, L.UseAddr, L.UseCity, L.UsePos, L.No_Polisi as NoPolisi, L.FMDL as KetModel, " & _
                            "K.NoRangka as NoChassis, L.UsePSTN as Telpon , L.UseGSM, L.UseCDMA, K.NoRangka as NoChassisSMS, L.FS_KD_PLG as KodeLGN, " & _
                            "M.FS_RC, CAST ((CAST(RIGHT(M.FS_RC,3) AS INTEGER) + (KMNEXTSVC - KM )/1000) * 1000 AS INTEGER) AS SBENEXT, L.COLOR as WrnModel, L.USEEMAIL AS EMAIL, L.MDL, L.MODEL, L.FS_TH_MOBIL, L.COLOR, L.FS_KD_DEALER, L.FS_NM_DEALER " & _
                            "FROM M_LISTREMINDERCALL K WITH(NOLOCK) LEFT JOIN HISTORY M WITH(NOLOCK) ON K.NORANGKA = M.NO_CHASSIS AND K.LASTNOWO = M.NO_WO " & _
                            "INNER JOIN Custdat L WITH(NOLOCK) On K.NoRangka = L.No_Chassis WHERE TANGGAL BETWEEN '" & FormatTgl(TglRem0) & "' AND '" & FormatTgl(TglRem1) & "' AND K.NoRangka <> '' And K.KMNextSvc <> 1000 " & _
                            "ORDER BY k.noRangka, KMNEXTSVC DESC"
34            strSql = "SELECT NOINV, NO_WO AS NOWO, DATE AS DATELASTSVC, KM AS KMNOW, DATENEXTSVC, KMNEXTSVC, L.UseName, L.UseAddr, L.UseCity, L.UsePos, L.No_Polisi as NoPolisi, L.FMDL as KetModel, " & _
                            "K.NoRangka as NoChassis, L.UsePSTN as Telpon , L.UseGSM, L.UseCDMA, K.NoRangka as NoChassisSMS, L.FS_KD_PLG as KodeLGN, " & _
                            "M.FS_RC, CAST ((CAST(RIGHT(M.FS_RC,3) AS INTEGER) + (KMNEXTSVC - KM )/1000) * 1000 AS INTEGER) AS SBENEXT, L.COLOR as WrnModel, L.USEEMAIL AS EMAIL, L.MDL, L.MODEL, L.FS_TH_MOBIL, L.COLOR, L.FS_KD_DEALER, L.FS_NM_DEALER " & _
                            "FROM M_LISTREMINDERCALL K WITH(NOLOCK) LEFT JOIN HISTORY M WITH(NOLOCK) ON K.NORANGKA = M.NO_CHASSIS AND K.LASTNOWO = M.NO_WO " & _
                            "INNER JOIN Custdat L WITH(NOLOCK) On K.NoRangka = L.No_Chassis WHERE 1=0"
35        End If
36        Set rs = New ADODB.Recordset
37        rs.CursorLocation = adUseClient
38        rs.Open strSql, SQLCon, adOpenStatic, adLockBatchOptimistic, adCmdText
39        Set rs.ActiveConnection = Nothing
40        Set FetchDataRemainder_HansaRenkeiSA = rs
41        If mConn.State = adStateOpen Then mConn.Close
              
42        Exit Function
errHandler:
43    err.Raise err.Number, err.Source, err.Description
End Function



