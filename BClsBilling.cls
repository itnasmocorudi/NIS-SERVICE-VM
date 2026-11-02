VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BClsBilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mConn As New ADODB.Connection
Private rs As ADODB.Recordset
Private strSql As String

Public Function takeInvoiceName(nChassis As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "SELECT A.FS_KD_DEALER, B.FS_NM_TAGIH AS FS_NM_DEALER, B.FS_NM_DEALER AS INIS_DEALER, FS_ALAMAT, FS_KOTA, FS_NPWP, FS_PKP, ISNULL(KDPLG,0) AS KDPLG, INTERNAL FROM CUSTDAT A WITH(NOLOCK) LEFT JOIN T_DEALER B WITH(NOLOCK) ON A.FS_KD_DEALER = B.FS_KD_DEALER WHERE NO_CHASSIS = '" & nChassis & "'" ' AND A.FS_KD_DEALER <> ''"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set takeInvoiceName = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function takeKodeDealer(nkode As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "SELECT FS_KD_DEALER, FS_NM_TAGIH AS FS_NM_DEALER, FS_NM_DEALER AS INIS_DEALER, FS_ALAMAT, FS_KOTA, FS_NPWP, FS_PKP, ISNULL(KDPLG,0) AS KDPLG, INTERNAL FROM T_DEALER WITH(NOLOCK) WHERE FS_KD_DEALER = '" & Trim(nkode) & "'"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set takeKodeDealer = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description

End Function

Public Function takeClassPrice(nameModel As String, idx As String) As Double
1         On Error GoTo errHandler
2         If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
          
3         If Log.RPLHistori = 1 Then
4             Select Case idx
                  Case 1
5                     strSql = "SELECT price FROM [" & dbReplikasiHO & "]..carclass WITH(NOLOCK) WHERE CLASSCODE = '" & nameModel & "'"
6                 Case 2
7                     strSql = "SELECT priceTAM AS PRICE FROM [" & dbReplikasiHO & "]..carclass WITH(NOLOCK) WHERE CLASSCODE = '" & nameModel & "'"
8                 Case 3
9                     strSql = "SELECT price FROM [" & dbReplikasiHO & "]..carclass WITH(NOLOCK) WHERE WAR = 1"
10                Case 4
11                    strSql = "SELECT price FROM [" & dbReplikasiHO & "]..carclass WITH(NOLOCK) WHERE CLASSCODE = 'PDS'"
12            End Select

13        Else
14            Select Case idx
                  Case 1
15                    strSql = "SELECT price FROM carclass WITH(NOLOCK) WHERE CLASSCODE = '" & nameModel & "'"
16                Case 2
17                    strSql = "SELECT priceTAM AS PRICE FROM carclass WITH(NOLOCK) WHERE CLASSCODE = '" & nameModel & "'"
18                Case 3
19                    strSql = "SELECT price FROM carclass WITH(NOLOCK) WHERE WAR = 1"
20                Case 4
21                    strSql = "SELECT price FROM carclass WITH(NOLOCK) WHERE CLASSCODE = 'PDS'"
22            End Select
23        End If
          

24        takeClassPrice = 0
25        Set rs = New ADODB.Recordset
26        rs.CursorLocation = adUseClient
27        rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
28        Set rs.ActiveConnection = Nothing
29        If Not rs.EOF Then
30            Call LoadtipeDisc(nameModel)
31            If idx = 1 Then
32                If CexChassis_Discon(nameModel) = True Then
33                    takeClassPrice = CDbl(rs("PRICE")) * ((100 - nDiscNRM) / 100)
34                Else
35                    takeClassPrice = CDbl(rs("PRICE")) * ((100 - nDiscNON) / 100)
36                End If
37            Else
38                takeClassPrice = CDbl(rs("PRICE")) ' * ((100 - nDiscNON) / 100)
39            End If
40        End If
41        If mConn.State <> adStateClosed Then mConn.Close
42        Exit Function
          
errHandler:
43    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
End Function

Public Function takeCheapsClassName() As String
    On Error GoTo errHandler
'    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
'
'    strSQL = "SELECT CLASSCODE FROM carclass WHERE priCE = (SELECT MIN(PRICE) FROM CARCLASS " & _
'             "WHERE FS_LOG_DEL = '' AND WAR = 0 AND PRICETAM > 0)"
'
'
'    Set rs = New ADodb.Recordset
'    rs.CursorLocation = adUseClient
'    rs.Open strSQL, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
'    Set rs.ActiveConnection = Nothing
'    takeCheapsClassName = Trim(rs("CLASSCODE"))
'    If mConn.State <> adStateClosed Then mConn.Close
        If gGenerasi = "" Then gGenerasi = "ECO"
        takeCheapsClassName = gGenerasi
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function checkDealerStatus(codeDealer As String) As Integer
1         On Error GoTo errHandler
2         If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
          
3         strSql = "SELECT * FROM T_DEALER WITH(NOLOCK) WHERE FS_KD_DEALER = '" & codeDealer & "'"

4         Set rs = New ADODB.Recordset
5         rs.CursorLocation = adUseClient
6         rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
7         Set rs.ActiveConnection = Nothing
8         If Not rs.EOF Then
9             If Trim(rs("INTERNAL")) = 1 And gNpwp = Trim(rs("FS_NPWP")) Then
                  ' NASMOCO SAMA NPWP
10                checkDealerStatus = 1
11            ElseIf Trim(rs("INTERNAL")) = 1 And gNpwp <> Trim(rs("FS_NPWP")) Then
                  ' NASMOCO BEDA NPWP
12                checkDealerStatus = 2
13            Else
14                If Trim(rs("HARGANRM")) = 1 Then
                      ' NASMOCO BEDA NPWP
15                    checkDealerStatus = 2
16                Else
                      ' MAIN DEALER LAIN
17                    checkDealerStatus = 3
18                End If
19            End If
20        Else
              ' DIANGGAP MAIN DEALER LAIN
21            checkDealerStatus = 3
22        End If
23        If mConn.State <> adStateClosed Then mConn.Close
24        Exit Function
          
errHandler:
25    err.Raise err.Number, err.Source, err.Description
End Function

Public Function checkKdDealer(codeDealer As String) As String
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "SELECT KDPLG FROM T_DEALER WITH(NOLOCK) WHERE FS_KD_DEALER = '" & codeDealer & "'"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    If Not rs.EOF Then
        checkKdDealer = Trim(rs("KDPLG"))
    End If
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function takeInvoiceCounter() As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "SELECT FN_NO_INV FROM T_Profil WITH(NOLOCK)"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set takeInvoiceCounter = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function searchOrderDetail(orderNumber As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "SELECT TOP 1 a.*,b.gsm,b.cdma,b.fs_TH_mobil,b.dotam, B.FS_NM_DEALER, B.FS_KD_DEALER  FROM jobDta a WITH(NOLOCK) INNER JOIN custdat b WITH(NOLOCK) ON a.no_chassis = b.no_chassis WHERE no_WO = '" & orderNumber & "'" ' AND NOINV = '' "

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set searchOrderDetail = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function searchWarrantyDetail(warrantyCode As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    If Log.RPLHistori = 0 Then
        strSql = "SELECT * FROM ho_Warranty WITH(NOLOCK) WHERE id_warranty = '" & warrantyCode & "' and FS_LOG_DEL = ''"
    Else
        
        strSql = "SELECT * FROM [" & dbReplikasiHO & "]..ho_Warranty WITH(NOLOCK) WHERE id_warranty = '" & warrantyCode & "' and FS_LOG_DEL = ''"
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set searchWarrantyDetail = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function searchProgramDetail(programCode As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    
    If Log.RPLHistori = 0 Then
        strSql = "SELECT ID_PROGRAM, PROGRAM, DISCJASA FROM ho_program WITH(NOLOCK) WHERE id_program = '" & programCode & "' and FS_LOG_DEL = ''"
    Else
        strSql = "SELECT ID_PROGRAM, PROGRAM, DISCJASA FROM [" & dbReplikasiHO & "]..ho_program WITH(NOLOCK) WHERE id_program = '" & programCode & "' and FS_LOG_DEL = ''"
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set searchProgramDetail = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function searchMOUDetail(dealerCode As String, modelType As String, nChassis As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    ' ================================================================================================================================
    ' 01 Mar 2017 | Thania | buka kuncian pencarian mou berdasarkan model hanya berdasarkan noka (req. 13 Feb 2017 - Setyo)
    ' ================================================================================================================================
'    strSql = "SELECT A.*,B.*,C.KODEDEALER, C.NAMADEALER, ISNULL(C.DISCJASA,0) AS DISCJASA, ISNULL(C.DISCPART,0) AS DISCPART FROM mouhdr A WITH(NOLOCK) INNER JOIN moudtl B WITH(NOLOCK) ON A.NO_mou = B.NO_mou LEFT JOIN moucbg C WITH(NOLOCK) ON A.NO_mou = C.NO_mou WHERE " _
'           & "kodedealer = '" & dealerCode & "' AND mdl = '" & modelType & "' AND No_rangka = '" & nChassis & "'"
    
    If Log.RPLHistori = 0 Then
        strSql = "SELECT A.*,B.*,C.KODEDEALER, C.NAMADEALER, ISNULL(C.DISCJASA,0) AS DISCJASA, ISNULL(C.DISCPART,0) AS DISCPART FROM mouhdr A WITH(NOLOCK) INNER JOIN moudtl B WITH(NOLOCK) ON A.NO_mou = B.NO_mou LEFT JOIN moucbg C WITH(NOLOCK) ON A.NO_mou = C.NO_mou WHERE " _
               & "kodedealer = '" & dealerCode & "' AND No_rangka = '" & nChassis & "'"
    Else
        strSql = "SELECT A.*,B.*,C.KODEDEALER, C.NAMADEALER, ISNULL(C.DISCJASA,0) AS DISCJASA, ISNULL(C.DISCPART,0) AS DISCPART FROM [" & dbReplikasiHO & "]..mouhdr A WITH(NOLOCK) INNER JOIN [" & dbReplikasiHO & "]..moudtl B WITH(NOLOCK) ON A.NO_mou = B.NO_mou LEFT JOIN [" & dbReplikasiHO & "]..moucbg C WITH(NOLOCK) ON A.NO_mou = C.NO_mou WHERE " _
               & "kodedealer = '" & dealerCode & "' AND No_rangka = '" & nChassis & "'"

    End If
    ' ================================================================================================================================
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set searchMOUDetail = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function isMOUCustomer(dealerCode As String, modelType As String, chassisNumber As String) As Boolean
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    isMOUCustomer = False
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    Set rs = searchMOUDetail(dealerCode, modelType, chassisNumber)
    Set rs.ActiveConnection = Nothing
    If Not rs.EOF Then
        isMOUCustomer = True
    End If
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function takeDiscountMOU(dealerCode As String, modelType As String, chassisNumber As String) As Integer
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    takeDiscountMOU = 0
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    Set rs = searchMOUDetail(dealerCode, modelType, chassisNumber)
    Set rs.ActiveConnection = Nothing
    If Not rs.EOF Then
        takeDiscountMOU = rs("DISCJASA")
    End If
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function searchPaymentType(orderNumber As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "SELECT PEMBAYARAN, VERIFIKATOR, KARAKTER, DTLKARAKTER FROM JOBDTL WITH(NOLOCK) WHERE NO_WO = '" & orderNumber & "'"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set searchPaymentType = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function searchDtlOrder(orderNumber As String, chassisNumber As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    strSql = "SELECT * FROM JOBDTL WITH(NOLOCK) WHERE NO_WO = '" & orderNumber & "' AND NO_CHASSIS = '" & chassisNumber & "'"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    Set searchDtlOrder = rs
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function isDuoInvoice(idProgram As String) As Boolean
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
    
    If Log.RPLHistori = 0 Then
        strSql = "SELECT * FROM HO_PROGRAM WITH(NOLOCK) WHERE ID_PROGRAM = '" & idProgram & "'"
    Else
        strSql = "SELECT * FROM [" & dbReplikasiHO & "]..HO_PROGRAM WITH(NOLOCK) WHERE ID_PROGRAM = '" & idProgram & "'"
    End If
    isDuoInvoice = False
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockBatchOptimistic, adCmdText
    Set rs.ActiveConnection = Nothing
    If Not rs.EOF Then
        If rs("DOUBLEINV") = "1" Then isDuoInvoice = True
    End If
    If mConn.State <> adStateClosed Then mConn.Close
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function takeDiscountProgram(idProgram As String) As Double
    On Error GoTo errHandler
    
    Set rs = searchProgramDetail(idProgram)
    If Not rs.EOF Then
        takeDiscountProgram = CDbl(rs("DISCJASA"))
    Else
        takeDiscountProgram = 0
    End If
    Exit Function
    
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchJatuhTempo(nkode As String) As String
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConACC)
    If OnTopSV = 0 Then
        strSql = "Select SyaratBayar FROM ARMST WITH(NOLOCK) Where KodeLgn = '" & nkode & "'"
    Else
    'UNTUK FEBUARI - TOP PELANGGAN
        strSql = "Select IsNull(TOPGRP,'') as SyaratBayar FROM ARMST WITH(NOLOCK) Where KodeLgn = '" & nkode & "'"
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConACC, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    If rs.RecordCount = 0 Then
        FetchJatuhTempo = "D0"
    Else:
        If Trim(rs.Fields(0).value) = "" Then
            FetchJatuhTempo = "D0"
        Else
            FetchJatuhTempo = Trim(rs.Fields(0).value)
        End If
    End If
    If mConn.State = adStateOpen Then mConn.Close
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchAvailable_FPNew() As Boolean
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConFP)
    
    strSql = "Select top 1 Right(NoSekarang,8), Nomor From FPKonter WITH(NOLOCK) Where Tahun = '" & Format(isnow, "YYYY") & "' and NoSekarang <> NoAkhir Order By Nomor ASC"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConFP, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    
    If rs.RecordCount > 0 Then
        FetchAvailable_FPNew = True
    Else
        FetchAvailable_FPNew = False
    End If
    
    Set rs = Nothing
    If mConn.State = adStateOpen Then mConn.Close
    Exit Function
errHandler:
    err.Raise err.Number, err.Source, err.Description
    If mConn.State = adStateOpen Then mConn.Close
End Function

Public Function FetchMinRange_FPNew() As Double
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConFP)
    
    strSql = "Select RemindSisaFP from Dealer WITH(NOLOCK)"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConFP, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    
    FetchMinRange_FPNew = Val(rs.Fields(0).value)
    
    Set rs = Nothing
    If mConn.State = adStateOpen Then mConn.Close
    Exit Function
errHandler:
    err.Raise err.Number, err.Source, err.Description
    If mConn.State = adStateOpen Then mConn.Close
End Function



'FetchRange_FPNew
'Tania IT: strSQL = "Select sum(cast(Right(NoAkhir,8) as numeric) - cast(Right(NoSekarang,8) as numeric)) as sisa from FPKonter Where Tahun = '" & Format(isnow, "YYYY") & "' and NoSekarang <> NoAkhir "
'Tania IT sedang mengetik...
'Tania IT: FetchRange_FPNew = Val(rs.Fields(0).value)
'

Public Function FetchRange_FPNew() As Double
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConFP)
    
    'strSQL = "Select Right(NoSekarang,8) , Right(NoAkhir,8)  from FPKonter Where Tahun = '" & Format(isnow, "YYYY") & "' and NoSekarang <> NoAkhir Order By Nomor ASC"
    strSql = "Select isNull(sum(cast(Right(NoAkhir,8) as numeric) - cast(Right(NoSekarang,8) as numeric)),0) as sisa from FPKonter WITH(NOLOCK) Where Tahun = '" & Format(isnow, "YYYY") & "' and NoSekarang <> NoAkhir "
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConFP, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    
    If rs.RecordCount = 0 Then
        FetchRange_FPNew = 0
    Else
        'FetchRange_FPNew = Val(rs.Fields(1).value) - Val(rs.Fields(0).value)
        FetchRange_FPNew = Val(rs.Fields(0).value)
    End If
    
    Set rs = Nothing
    If mConn.State = adStateOpen Then mConn.Close
    Exit Function
errHandler:
    err.Raise err.Number, err.Source, err.Description
    If mConn.State = adStateOpen Then mConn.Close
End Function



'ALTER TABLE [dbo].[CarClass]
'ADD [DiscNRM] numeric(18, 0) NULL
'GO
'ALTER TABLE [dbo].[CarClass]
'ADD [DiscNON] numeric(18, 0) NULL
'GO
