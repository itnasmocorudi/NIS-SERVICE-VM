VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mConn   As New ADODB.Connection
Private rs      As ADODB.Recordset
Private strSql  As String

Public Function FetchDataSID(SQLf As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConSID)
        
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open SQLf, SQLConSID, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchDataSID = rs
    If mConn.State = adStateOpen Then mConn.Close
    Exit Function
errHandler:
err.Raise err.Number, err.Source, err.Description
End Function

Public Function FetchListAllPelanggan(param As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConBP)
    
    'If Profil.nIpAcc <> Log.logServer Then
    If Log.logServBP <> Log.LogServ Then
        strSql = "Select TOP 100 c.NoRangka, Isnull(e.IDCust,'') IDCust, '' as NoKTP, '' as NoNPWP, '' as Title, Substring(c.Nama,1,30) as Nama1, " & _
                 "Substring(c.Nama,31,61) as Nama2, Substring(c.Nama,62,212) as Nama3, '01/01/3000' as TglLahir, '' as Agama, '' as Pekerjaan, " & _
                 "'' as Propinsi, c.Kota as Kota, '' as Kecamatan, '' as Kelurahan, c.Pos as KodePos, c.Alamat, '' RT, '' RW, '' KodeTelp, c.PSTN Telp, " & _
                 "c.GSM as HP1, c.CDMA as HP2, c.EMail, '' as Hobi, '' as Olahraga, '' Suku, '' Preferensi, c.PoNomor as NoPolis, c.PoPolish0 as tglAwalPolis, " & _
                 "c.PoPolish1 as TglAkhirPolis, c.NoPolisi, c.NoMesin as NoEngine, '' NoKey, c.Model as MDL, m.Model as Class, m.[DESC] as Fmdl, c.Warna as Color, " & _
                 "c.Dealer as KdDlrPenjual, d.FS_NM_DEALER as NmDlrPenjual, c.DDate, c.ThnMobil as Thmobil,c.KodeLg as KodeLgn " & _
                 "FROM " & Log.logDbBP & "..Custdat c WITH (NOLOCK)  " & _
                 "LEFT JOIN (select * from OPENROWSET('SQLOLEDB','DRIVER={SQL Server};SERVER=" & Log.LogServ & ";Uid=" & Log.LogAccUser & ";PWD=" & Log.LogAccPass & "', " & Log.LogSPKDb & ".dbo.InWarnaExt)) w On c.Warna = w.KodeWarna " & _
                 "LEFT JOIN " & Log.logDbBP & "..M_Asuransi a WITH (NOLOCK)  On c.ASURANSI = a.Kode " & _
                 "LEFT JOIN " & Log.logDbBP & ".dbo.Model m WITH (NOLOCK)  On c.Model = m.MDL " & _
                 "LEFT JOIN OPENROWSET('SQLOLEDB','DRIVER={SQL Server};SERVER=" & Log.LogServ & ";Uid=" & Log.LogAccUser & ";PWD=" & Log.LogAccPass & "', " & Log.LogBengkel & ".dbo.T_Dealer) d " & _
                 "On c.Dealer = d.FS_KD_DEALER LEFT JOIN SingleIDOri..Bridging e On c.KodeLg = e.KodeLgn " & param & " Order by c.TglEntry DESC"
    Else
        strSql = "SELECT * FROM (Select TOP 100 c.NoRangka, Isnull(e.IDCust,'') IDCust, '' as NoKTP, '' as NoNPWP, Substring(c.Nama,1,30) as Nama1, " & _
                 "Substring(c.Nama,31,61) as Nama2, Substring(c.Nama,62,212) as Nama3, '01/01/3000' as TglLahir, " & _
                 " c.Kota as Kota, c.Alamat, " & _
                 " c.GSM as HP1, " & _
                 " c.NoPolisi,c.KodeLg as KodeLgn " & _
                 "FROM " & Log.logDbBP & "..Custdat c  WITH (NOLOCK) " & _
                 "LEFT JOIN (select * from " & Log.LogSPKDb & "..InWarnaExt) w On c.Warna = w.KodeWarna " & _
                 "LEFT JOIN " & Log.logDbBP & "..M_Asuransi a  WITH (NOLOCK) On c.ASURANSI = a.Kode " & _
                 "LEFT JOIN " & Log.logDbBP & "..Model m  WITH (NOLOCK) On c.Model = m.MDL " & _
                 "LEFT JOIN " & Log.LogBengkel & "..T_Dealer d  WITH (NOLOCK) On c.Dealer = d.FS_KD_DEALER " & _
                 "LEFT JOIN SingleIDOri..Bridging e  WITH (NOLOCK) On c.KodeLg = e.KodeLgn and e.KodeDealer='" & Right(gDb, 2) & "') a " & param & " Order by A.IDCust asc "
    End If
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConBP, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchListAllPelanggan = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
    If err.Number = -2147217900 Then
        Set FetchListAllPelanggan = Nothing
    Else
        err.Raise err.Number, err.Source, err.Description
    End If
End Function

 

Public Function FetchPelangganGRP(param As String, paramAwal As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
'    If Profil.nIpAcc <> Log.logServer Then
'        strSql = "Select Distinct TOP 10 c.NO_CHASSIS as NoRangka, Isnull(e.IDCust,'') IDCust, '' as NoKTP, '' as NoNPWP, '' as Title, Substring(c.Customer,1,30) as Nama1, Substring(c.Customer,31,61) as Nama2, " & _
'                 "Substring(c.Customer,62,212) as Nama3, '01/01/3000' as TglLahir, '' as Agama, '' as Pekerjaan, '' as Propinsi, c.Kota as Kota, '' as Kecamatan, '' as Kelurahan, c.KodePos, c.Cust_Adres as Alamat, " & _
'                 "'' RT, '' RW, '' KodeTelp, c.UsePstn Telp, c.UseGSM as HP1, c.UseCDMA as HP2, c.USEEMAIL as Email, '' as Hobi, '' as Olahraga, '' Suku, '' Preferensi, '' as NoPolis, '01/01/3000' as tglAwalpolis, " & _
'                 "'01/01/3000' as tglAkhirpolis, c.No_Polisi as NoPolisi, c.No_Engine as NoEngine, '' NoKey, c.Mdl, c.Model as Class, c.fmdl, (Select top 1 KodeWarna FROM SPK18..InWarnaExt Where NamaInd = c.color) Color, " & _
'                 "c.FS_KD_DEALER as KdDlrPenjual, c.FS_NM_DEALER as NmDlrPenjual, c.DDate, c.Fs_Th_Mobil Thmobil " & _
'                 "FROM OPENROWSET('SQLOLEDB','DRIVER={SQL Server};SERVER=" & Profil.nIpAcc & ";Uid=" & Profil.nUserAcc & ";PWD=" & Profil.nPassAcc & "', " & DB_Bengkel & ".dbo.Custdat) c INNER JOIN " & _
'                 "OPENROWSET('SQLOLEDB','DRIVER={SQL Server};SERVER=" & Profil.nIpAcc & ";Uid=" & Profil.nUserAcc & ";PWD=" & Profil.nPassAcc & "', " & DB_Bengkel & ".dbo.JobDTA) j On c.No_chassis = j.No_Chassis " & _
'                 "LEFT JOIN SingleIDOri..Bridging e On c.FS_KD_PLG = e.KodeLgn " & param
'    Else
        strSql = "SELECT * FROM ( Select Distinct TOP 100 c.NO_CHASSIS as NoRangka, Isnull(e.IDCust,'') IDCust, '' as NoKTP, '' as NoNPWP, Substring(c.Customer,1,30) as Nama1, Substring(c.Customer,31,61) as Nama2, " & _
                 "Substring(c.Customer,62,212) as Nama3, '01/01/3000' as TglLahir, c.Kota as Kota,c.Cust_Adres as Alamat, " & _
                 " c.UseGSM as HP1," & _
                 " c.No_Polisi as NoPolisi,c.fs_kd_plg as KodeLgn FROM " & Log.LogBengkel & ".dbo.Custdat c WITH (NOLOCK)  left JOIN " & Log.LogBengkel & ".dbo.JobDTA j  WITH (NOLOCK) On c.No_chassis = j.No_Chassis " & _
                 "LEFT JOIN SingleIDOri..Bridging e  WITH (NOLOCK) On c.FS_KD_PLG = e.KodeLgn and e.KodeDealer='" & Right(gDb, 2) & "' " & paramAwal & ") c " & param
'    End If
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchPelangganGRP = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
If err.Number = -2147217900 Then
    Set FetchPelangganGRP = Nothing
Else
    err.Raise err.Number, err.Source, err.Description
End If
End Function

Public Function fetchPelangganSPK(param As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLCONBengkel)
    
'    If Profil.nIpAcc <> Log.logServer Then
'        strSql = "Select c.Kode_Lgn as KodeLgn, IsNull(c.Nama_Pemesan,'') as NamaCust, c.TglLahir, IsNull(c.NPWP,'') NPWP, (c.Alamat1 + ' ' + c.Alamat2) as Alamat, IsNull(c.Kota,'') as Kotamadya, IsNull(c.Kode_Pos,'') as KodePos, " & _
'                 "IsNull(c.Telp_Rumah,'') as PSTN, IsNull(c.HP1,'') as GSM, IsNull(c.HP2,'') as CDMA, isNull(c.Email,'') as EMAIL, IsNull(a.NoPol,'')as NoPol, isNull(a.Tahun,'') as Tahun, IsNull(c.Kode_Type,'') as KodeTipe, m.Model, " & _
'                 "m.[DESC], isNull(a.Warna,'') as Warna, IsNull(w.NamaEnglish,'') as NamaEnglish, isNull(c.NoRangka,'') as NoRangka, IsNull(i.NoMesin,'') as NoMesin, IsNull(c.SNama,'') as SNama, IsNull(c.SAlamat,'') as sAlamat, IsNull(c.SKota,'') as SKota, " & _
'                 "IsNull(c.SKodePos,'') as SKodePos, IsNull(c.STelp,'') as sTelp, IsNull(c.SGSM,'') as sGsm, IsNull(c.SCdma,'') as sCDMA, c.Tgl_Entry as TglEntry " & _
'                 "FROM (select Kode_lgn,Nama_Pemesan,TglLahir,Alamat1,Alamat2,Kota,Kode_Pos,NPWP,Telp_Rumah,HP1,HP2,Email,Kode_Type,NoRangka,SNama,SAlamat,SKota,SKodePos,STelp,SGSM,SCdma,SPK_No,Tgl_Entry from " & _
'                 "OPENROWSET('SQLOLEDB','DRIVER={SQL Server};SERVER=" & Profil.nIpAcc & ";Uid=" & Profil.nUserAcc & ";PWD=" & Profil.nPassAcc & "', " & DB_SPK & ".dbo.SPK)) c " & _
'                 "Left JOIN OPENROWSET('SQLOLEDB','DRIVER={SQL Server};SERVER=" & Profil.nIpAcc & ";Uid=" & Profil.nUserAcc & ";PWD=" & Profil.nPassAcc & "', " & DB_SPK & ".dbo.AFI) a On c.SPK_No = a.SPK_No " & _
'                 "Left JOIN OPENROWSET('SQLOLEDB','DRIVER={SQL Server};SERVER=" & Profil.nIpAcc & ";Uid=" & Profil.nUserAcc & ";PWD=" & Profil.nPassAcc & "', " & DB_SPK & ".dbo.BPK) b On b.SPK_No = a.SPK_No " & _
'                 "Left JOIN OPENROWSET('SQLOLEDB','DRIVER={SQL Server};SERVER=" & Profil.nIpAcc & ";Uid=" & Profil.nUserAcc & ";PWD=" & Profil.nPassAcc & "', " & DB_SPK & ".dbo.INMST) i On c.NoRangka = i.KodeItem " & _
'                 "Left JOIN OPENROWSET('SQLOLEDB','DRIVER={SQL Server};SERVER=" & Profil.nIpAcc & ";Uid=" & Profil.nUserAcc & ";PWD=" & Profil.nPassAcc & "', " & DB_SPK & ".dbo.InWarnaExt) w On a.Warna = w.KodeWarna " & _
'                 "Left JOIN " & DB_BP & ".dbo.Model m On c.Kode_Type = m.MDL " & param
'    Else
        strSql = "Select c.Kode_Lgn as KodeLgn, IsNull(c.Nama_Pemesan,'') as NamaCust, c.TglLahir, IsNull(c.NPWP,'') NPWP, (c.Alamat1 + ' ' + c.Alamat2) as Alamat, IsNull(c.Kota,'') as Kotamadya, IsNull(c.Kode_Pos,'') as KodePos, " & _
                 "IsNull(c.Telp_Rumah,'') as PSTN, IsNull(c.HP1,'') as GSM, IsNull(c.HP2,'') as CDMA, isNull(c.Email,'') as EMAIL, IsNull(a.NoPol,'')as NoPol, isNull(a.Tahun,'') as Tahun, IsNull(c.Kode_Type,'') as KodeTipe, m.Model, " & _
                 "m.[DESC], isNull(a.Warna,'') as Warna, IsNull(w.NamaEnglish,'') as NamaEnglish, isNull(c.NoRangka,'') as NoRangka, IsNull(i.NoMesin,'') as NoMesin, IsNull(c.SNama,'') as SNama, IsNull(c.SAlamat,'') as sAlamat, IsNull(c.SKota,'') as SKota, " & _
                 "IsNull(c.SKodePos,'') as SKodePos, IsNull(c.STelp,'') as sTelp, IsNull(c.SGSM,'') as sGsm, IsNull(c.SCdma,'') as sCDMA, c.Tgl_Entry as TglEntry " & _
                 "FROM (select Kode_lgn,Nama_Pemesan,TglLahir,Alamat1,Alamat2,Kota,Kode_Pos,NPWP,Telp_Rumah,HP1,HP2,Email,Kode_Type,NoRangka,SNama,SAlamat,SKota,SKodePos,STelp,SGSM,SCdma,SPK_No,Tgl_Entry from [spk" & Right(gDb, 2) & "].dbo.SPK) c " & _
                 "Left JOIN [spk" & Right(gDb, 2) & "].dbo.AFI a On c.SPK_No = a.SPK_No " & _
                 "Left JOIN [spk" & Right(gDb, 2) & "].dbo.BPK b On b.SPK_No = a.SPK_No " & _
                 "Left JOIN [spk" & Right(gDb, 2) & "].dbo.INMST i On c.NoRangka = i.KodeItem " & _
                 "Left JOIN [spk" & Right(gDb, 2) & "].dbo.InWarnaExt w On a.Warna = w.KodeWarna " & _
                 "Left JOIN [bp" & Right(gDb, 2) & "].dbo.Model m On c.Kode_Type = m.MDL " & param
'    End If

    strSql = "SELECT DISTINCT * FROM (Select IsNull(NoRangka,'') NoRangka, IsNull(c.Kode_Lgn,'') as KodeLgn, Isnull(e.IDCust,'') as IDCust, '' as TipeCust, IsNull(c.NoKTP,'') as NoKTP, IsNull(c.NPWP,'') NoNPWP, '' Title, IsNull(c.Nama_Pemesan,'') as NamaFull, Case When IsNull(c.Nama_Pemesan,'')='' Then '' Else Case When CHARINDEX(' ',IsNull(c.Nama_Pemesan,'')) = 0 Then " & _
                 "Substring(IsNull(c.Nama_Pemesan,''),1,30) Else Substring(IsNull(c.Nama_Pemesan,''),1,CHARINDEX(' ',IsNull(c.Nama_Pemesan,''))) End End as Nama1, Case When CHARINDEX(' ',IsNull(c.Nama_Pemesan,'')) = 0 Then '' Else Case When CHARINDEX(' ',Ltrim(Substring(IsNull(c.Nama_Pemesan,'')," & _
                 "CHARINDEX(' ',IsNull(c.Nama_Pemesan,'')),61)))=0 Then Ltrim(Substring(IsNull(c.Nama_Pemesan,''),CHARINDEX(' ',IsNull(c.Nama_Pemesan,'')),61)) Else substring(Ltrim(Substring(IsNull(c.Nama_Pemesan,''),CHARINDEX(' ',IsNull(c.Nama_Pemesan,'')),61)),1,CHARINDEX(' ',Ltrim(Substring(IsNull(c.Nama_Pemesan,'')," & _
                 "CHARINDEX(' ',IsNull(c.Nama_Pemesan,'')),61)))) end end as Nama2, Case When CHARINDEX(' ',Ltrim(Substring(IsNull(c.Nama_Pemesan,''),CHARINDEX(' ',IsNull(c.Nama_Pemesan,'')),61)))=0 Then '' Else Ltrim(Substring(Ltrim(Substring(IsNull(c.Nama_Pemesan,''),CHARINDEX(' ',IsNull(c.Nama_Pemesan,'')),61))," & _
                 "CHARINDEX(' ',Ltrim(Substring(IsNull(c.Nama_Pemesan,''),CHARINDEX(' ',IsNull(c.Nama_Pemesan,'')),61))),212)) end as Nama3, IsNull(c.TglLahir,'01/01/3000') as TglLahir, IsNull(c.AgamaSpk,'') as Agama, IsNull(c.Pekerjaan,'') Pekerjaan, IsNull(c.Propinsi,'') Propinsi, IsNull(c.Kota,'') as Kota, " & _
                 "'' Kecamatan, '' Kelurahan, IsNull(c.Kode_Pos,'') as KodePos, (IsNull(c.Alamat1,'') + ' ' + IsNull(c.Alamat2,'')) as Alamat, '' RT, '' RW, '' KodeTelp, IsNull(c.Telp_Rumah,'') as Telp, IsNull(c.HP1,'') as HP1, IsNull(c.HP2,'') as HP2, isNull(c.Email,'') as Email, '' Hobi, '' Olahraga, '' Suku, " & _
                 "'' Preferensi, '' NoPolis, '01/01/3000' tglAwalPolis, '01/01/3000' tglAkhirPolis, IsNull(a.NoPol,'')as NoPolisi, IsNull(i.NoMesin,'') as NoEngine, '' NoKey, IsNull(c.Kode_Type,'') as Mdl, m.Model as Class, m.[desc] as Fmdl, isNull(c.Kode_Warna,'') as Color, IsNull(w.NamaEnglish,'') as ColorNama, " & _
                 "'' KdDlrPenjual, '' NmDlrPenjual, '' DDate, isNull(a.Tahun,'') as ThMobil, 0 as CexSurat, '' NonToyota FROM ( select Kode_lgn,Nama_Pemesan,TglLahir,Alamat1,Alamat2,NoKTP,Propinsi,Kota,Kode_Pos,NPWP,Telp_Rumah,HP1,HP2,Email,Kode_Type,NoRangka,SPK_No,Tgl_Entry,AgamaSpk,Pekerjaan, Kode_Warna from [spk" & Right(gDb, 2) & "].dbo.SPK) c " & _
                 "Left JOIN [spk" & Right(gDb, 2) & "].dbo.AFI a On c.SPK_No = a.SPK_No Left JOIN [spk" & Right(gDb, 2) & "].dbo.BPK b On b.SPK_No = a.SPK_No Left JOIN [spk" & Right(gDb, 2) & "].dbo.INMST i On c.NoRangka = i.KodeItem " & _
                 "Left JOIN [spk" & Right(gDb, 2) & "].dbo.InWarnaExt w On a.Warna = w.KodeWarna Left JOIN [bengkel" & Right(gDb, 2) & "].dbo.model m On c.Kode_Type = m.MDL LEFT JOIN [SINGLEIDORI]..Bridging e On c.Kode_Lgn = e.KodeLgn AND e.KodeDealer='" & Right(gDb, 2) & "' ) c " & param

    strSql = "SELECT DISTINCT * FROM (Select IsNull(NoRangka,'') NoRangka, IsNull(c.Kode_Lgn,'') as KodeLgn, Isnull(e.IDCust,'') as IDCust, '' as TipeCust, '' as NoKTP, IsNull(c.NPWP,'') NoNPWP, '' Title, IsNull(c.Nama_Pemesan,'') as NamaFull, Case When IsNull(c.Nama_Pemesan,'')='' Then '' Else Case When CHARINDEX(' ',IsNull(c.Nama_Pemesan,'')) = 0 Then " & _
                 "Substring(IsNull(c.Nama_Pemesan,''),1,30) Else Substring(IsNull(c.Nama_Pemesan,''),1,CHARINDEX(' ',IsNull(c.Nama_Pemesan,''))) End End as Nama1, Case When CHARINDEX(' ',IsNull(c.Nama_Pemesan,'')) = 0 Then '' Else Case When CHARINDEX(' ',Ltrim(Substring(IsNull(c.Nama_Pemesan,'')," & _
                 "CHARINDEX(' ',IsNull(c.Nama_Pemesan,'')),61)))=0 Then Ltrim(Substring(IsNull(c.Nama_Pemesan,''),CHARINDEX(' ',IsNull(c.Nama_Pemesan,'')),61)) Else substring(Ltrim(Substring(IsNull(c.Nama_Pemesan,''),CHARINDEX(' ',IsNull(c.Nama_Pemesan,'')),61)),1,CHARINDEX(' ',Ltrim(Substring(IsNull(c.Nama_Pemesan,'')," & _
                 "CHARINDEX(' ',IsNull(c.Nama_Pemesan,'')),61)))) end end as Nama2, Case When CHARINDEX(' ',Ltrim(Substring(IsNull(c.Nama_Pemesan,''),CHARINDEX(' ',IsNull(c.Nama_Pemesan,'')),61)))=0 Then '' Else Ltrim(Substring(Ltrim(Substring(IsNull(c.Nama_Pemesan,''),CHARINDEX(' ',IsNull(c.Nama_Pemesan,'')),61))," & _
                 "CHARINDEX(' ',Ltrim(Substring(IsNull(c.Nama_Pemesan,''),CHARINDEX(' ',IsNull(c.Nama_Pemesan,'')),61))),212)) end as Nama3, IsNull(c.TglLahir,'01/01/3000') as TglLahir, IsNull(c.AgamaSpk,'') as Agama, IsNull(c.Pekerjaan,'') Pekerjaan, IsNull(c.Propinsi,'') Propinsi, IsNull(c.Kota,'') as Kota, " & _
                 "'' Kecamatan, '' Kelurahan, IsNull(c.Kode_Pos,'') as KodePos, (IsNull(c.Alamat1,'') + ' ' + IsNull(c.Alamat2,'')) as Alamat, '' RT, '' RW, '' KodeTelp, IsNull(c.Telp_Rumah,'') as Telp, IsNull(c.HP1,'') as HP1, IsNull(c.HP2,'') as HP2, isNull(c.Email,'') as Email, '' Hobi, '' Olahraga, '' Suku, " & _
                 "'' Preferensi, '' NoPolis, '01/01/3000' tglAwalPolis, '01/01/3000' tglAkhirPolis, IsNull(a.NoPol,'')as NoPolisi, IsNull(i.NoMesin,'') as NoEngine, '' NoKey, IsNull(c.Kode_Type,'') as Mdl, m.Model as Class, m.[desc] as Fmdl, isNull(c.Kode_Warna,'') as Color, IsNull(w.NamaEnglish,'') as ColorNama, " & _
                 "'' KdDlrPenjual, '' NmDlrPenjual, '' DDate, isNull(a.Tahun,'') as ThMobil, 0 as CexSurat, '' NonToyota FROM ( select Kode_lgn,Nama_Pemesan,TglLahir,Alamat1,Alamat2,NoKTP,Propinsi,Kota,Kode_Pos,NPWP,Telp_Rumah,HP1,HP2,Email,Kode_Type,NoRangka,SPK_No,Tgl_Entry,AgamaSpk,Pekerjaan, Kode_Warna from [spk" & Right(gDb, 2) & "].dbo.SPK) c " & _
                 "Left JOIN [spk" & Right(gDb, 2) & "].dbo.AFI a On c.SPK_No = a.SPK_No Left JOIN [spk" & Right(gDb, 2) & "].dbo.BPK b On b.SPK_No = a.SPK_No Left JOIN [spk" & Right(gDb, 2) & "].dbo.INMST i On c.NoRangka = i.KodeItem " & _
                 "Left JOIN [spk" & Right(gDb, 2) & "].dbo.InWarnaExt w On a.Warna = w.KodeWarna Left JOIN [bengkel" & Right(gDb, 2) & "].dbo.model m On c.Kode_Type = m.MDL LEFT JOIN [SINGLEIDORI]..Bridging e On c.Kode_Lgn = e.KodeLgn AND e.KodeDealer='" & Right(gDb, 2) & "' ) c " & param
     
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLCONBengkel, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    Set fetchPelangganSPK = rs
    If mConn.State = adStateOpen Then mConn.Close
        
    Exit Function
errHandler:
If err.Number = -2147217900 Then
    Set fetchPelangganSPK = Nothing
Else
    err.Raise err.Number, err.Source, err.Description
End If
End Function

Public Function FetchPelangganSPK00(param As String) As ADODB.Recordset
    On Error GoTo errHandler

    mConn.CommandTimeout = 60
    If mConn.State = adStateClosed Then mConn.Open (SQLConHO)
    
    strSql = "Select Top 100 NamaFaktur, IsNull(TglLahir,'01/01/3000') TglLahir, AlamatFaktur, Kota, IsNull(KodePos,'') KodePos, IsNull(Telp,'') Telp, " & _
             "IsNull(GSMFaktur,'') GSMFaktur, IsNull(CDMAFaktur,'') CDMAFaktur, IsNull(NoPol,'') NoPol, a.KodeKategori, a.KodeWarna, c.NamaEnglish, a.NoRangka, a.NoMesin, " & _
             "a.dealer, o.NamaDealer, o.DbName, IsNull(sNama,'') sNama, IsNull(sAlamat,'') sAlamat, IsNull(sKota,'') sKota, IsNull(sKodePos,'') sKodePos, " & _
             "IsNull(sTelp,'') sTelp, IsNull(sGSM,'') sGSM, IsNull(sCDMA,'') sCDMA " & _
             "FROM SPK99..DATACUST a INNER JOIN SPK00..INKATEGORY b ON a.KODEKATEGORI = b.KODEKATEGORY " & _
             "INNER JOIN SPK00..INWARNAEXT c ON a.KODEWARNA = c.KODEWARNA " & _
             "INNER JOIN SPK00..INMST d ON a.NORANGKA = d.KODEITEM " & _
             "INNER JOIN SPK00..DODEALER o On a.Dealer = o.KodeDealer " & param
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConHO, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchPelangganSPK00 = rs
    If mConn.State = adStateOpen Then mConn.Close
    ConHO = True
    Exit Function
errHandler:
    Set FetchPelangganSPK00 = Nothing
    ConHO = False
End Function

Public Function FetchLastID(tipe As String) As ADODB.Recordset
    On Error GoTo errHandler
    If mConn.State = adStateClosed Then mConn.Open (SQLConSID)
    
'    strSql = "Select Top 1 Case When LEN(a.Kode)=1 Then '0'+a.Kode Else Kode End + 'S' + c.KodeCabang As Kode, Convert(Numeric(19,0),Right(IsNull(d.IDCust,0),7)) as LastUrut " & _
'             "From TipeCustomer a Left Join ProfilBP..Profil c On c.KodeCabang<>'' Left Join NewCustomer d On LEFT(d.IDCust,5) = Case " & _
'             "When LEN(a.Kode)=1 Then '0' + a.Kode Else Kode End + 'S' + c.KodeCabang Where Upper(a.Tipe)='" & UCase(tipe) & "' Order By Convert(Numeric(19,0),Right(IsNull(d.IDCust,0),7)) Desc "
    strSql = "Select Top 1 Case When LEN(a.Kode)=1 Then '0'+a.Kode Else Kode End + 'S' + '" & Right(gDb, 2) & "' As Kode, Convert(Numeric(19,0),Right(IsNull(d.IDCust,0),7)) as LastUrut " & _
             "From TipeCustomer a Left Join NewCustomer d On LEFT(d.IDCust,5) = Case " & _
             "When LEN(a.Kode)=1 Then '0' + a.Kode Else Kode End + 'S' + '" & Right(gDb, 2) & "' Where Upper(a.Tipe)='" & UCase(tipe) & "' Order By Convert(Numeric(19,0),Right(IsNull(d.IDCust,0),7)) Desc "
             
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSql, SQLConSID, adOpenStatic, adLockReadOnly, adCmdText
    Set rs.ActiveConnection = Nothing
    Set FetchLastID = rs
    If mConn.State = adStateOpen Then mConn.Close
    Exit Function
    
errHandler:
    If err.Number = -2147217900 Then
        Set FetchLastID = Nothing
    Else
        err.Raise err.Number, err.Source, err.Description
    End If
End Function



