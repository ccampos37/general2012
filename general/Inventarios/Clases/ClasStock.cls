VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClasStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim adodc1 As ADODB.Recordset
Dim Adodc2 As ADODB.Recordset
Dim c_cantid As Double
Dim nTra As Integer
Dim cAnoMes As String
Dim dfecha As Date
Dim Init As Boolean
Public Sub RestauraStock(ByVal arAlma As String, ByVal arCodArti As String, ByVal Fechtran As Date)
Dim cSql1 As String, CSQL2 As String
'On Error GoTo ErrCarga
'*******************************************************
'**Carga todos los Movimientos al respectivo Recordset a partir de 1er dia del Mes en que se realizo la transacciojn0
'*******************************************************
cAnoMes = Format(Year(Fechtran), "0000") + Format(Month(Fechtran), "00")
dfecha = Format(Fechtran, "dd/mm/yyyy")
Set adodc1 = New ADODB.Recordset
cSql1 = "Select DETIPCAM,CATIPCAM,DECANTID,DECODIGO,CATIPMOV,DEALMA,DETD,DENUMDOC,DEITEM,CATD,DEPRECIO,CACODMON,CAFECDOC,DETIPCAM From MovAlmDet A Inner Join MovAlmCab B On B.CAALMA = A.DEALMA And B.CATD = A.DETD  And B.CANUMDOC = A.DENUMDOC "
cSql1 = cSql1 & " Where CAALMA = '" & arAlma & "' and not (CACODMOV='GF' And CASITGUI='F') "
'cSql1 = cSql1 & " And CASITGUI<>'A'  and DECODIGO='" & arCodArti & "' and FORMAT( YEAR(CAFECDOC),'0000' )+ FORMAT( MONTH(CAFECDOC),'00' ) >='" & cAnoMes & "' Order By DECODIGO,CAFECDOC,CAHORA"
cSql1 = cSql1 & " And CASITGUI<>'A'  and DECODIGO='" & arCodArti & "' and cast(year(CAFECDOC) as varchar(4)) + replicate('0',2-len(month(CAFECDOC))) + cast(month(CAFECDOC) as varchar(2)) >='" & cAnoMes & "' Order By DECODIGO,CAFECDOC,catipmov,CAHORA"



adodc1.Open cSql1, VGCNx, adOpenStatic, adLockReadOnly
'*******************************************************

Set Adodc2 = New ADODB.Recordset
Adodc2.Open "Select * from MORESMES where SMMESPRO >='" & AnioMesAnterior(cAnoMes) & "' AND SMALMA='" & arAlma & "' AND SMCODIGO='" & arCodArti & "'", VGCNx, adOpenStatic, adLockReadOnly


CSQL2 = "delete From MORESMES Where SMMESPRO>='" & cAnoMes & "' and SMALMA='" & arAlma & "' and SMCODIGO='" & arCodArti & "'"
VGCNx.Execute CSQL2
Call ValorizaXArticulo(arCodArti, cAnoMes, arAlma) 'Procedimiento que Valoriza por Articulo

Exit Sub
ErrCarga:
        MsgBox Err.Description
        
End Sub
Public Sub MovimientosdeEstablecimientos(ByVal Empresa As String, ByVal Establec As String, ByVal arCodArti As String, ByVal Fechtran As Date)
Dim cSql1 As String, CSQL2 As String

cAnoMes = Format(Year(Fechtran), "0000") + Format(Month(Fechtran), "00")
dfecha = Format(Fechtran, "dd/mm/yyyy")
Set adodc1 = New ADODB.Recordset

SQL = "Select tipodecosto=isnull(tipodecosto,'P'),DETIPCAM,CATIPCAM,DECANTID,DECODIGO,CATIPMOV,DEALMA,DETD,DENUMDOC,DEITEM,CATD,DEPRECIO,CACODMON,"
SQL = SQL & "CAFECDOC From MovAlmDet A Inner Join MovAlmCab B On B.CAALMA +B.CATD+B.CANUMDOC =A.DEALMA +A.DETD+A.DENUMDOC "
SQL = SQL & " inner join tabalm al on b.caalma=al.taalma "
SQL = SQL & " inner join tabtransa c on b.cacodmov=c.tt_codmov Where al.empresacodigo = '" & Empresa & "' and decodigo='" & arCodArti & "'"
SQL = SQL & " and  isnull(estadocosto,0)=1 And CASITGUI<>'A' and al.puntovtacodigo='" & Establec & "'"
SQL = SQL & " and cast(year(CAFECDOC) as varchar(4)) + replicate('0',2-len(month(CAFECDOC))) + cast(month(CAFECDOC) as varchar(2)) ='" & cAnoMes & "' Order By DECODIGO,CAFECDOC,catipmov,CAHORA"



Set adodc1 = VGCNx.Execute(SQL)
'*******************************************************
Set Adodc2 = New ADODB.Recordset
   SQL = "Select top 1 * from al_MOvRESMES where SMMESPRO <='" & AnioMesAnterior(cAnoMes) & "' AND empresacodigo='" & VGparametros.empresacodigo & "'"
   SQL = SQL & " AND SMCODIGO='" & arCodArti & "' and puntovtacodigo='" & Establec & "'"
   SQL = SQL & " order by smmespro desc "
   Set Adodc2 = VGCNx.Execute(SQL)
   Call ValorizaEmpresaestablecXArticulo(arCodArti, Establec, cAnoMes, Empresa)
Exit Sub
ErrCarga:
        MsgBox Err.Description
        
End Sub

Public Sub MovimientosdeEmpresa(ByVal Empresa As String, ByVal arCodArti As String, ByVal Fechtran As Date)
Dim cSql1 As String, CSQL2 As String

cAnoMes = Format(Year(Fechtran), "0000") + Format(Month(Fechtran), "00")
dfecha = Format(Fechtran, "dd/mm/yyyy")
Set adodc1 = New ADODB.Recordset

SQL = "Select tipodecosto,DETIPCAM,CATIPCAM,DECANTID,DECODIGO,CATIPMOV,DEALMA,DETD,DENUMDOC,DEITEM,CATD,DEPRECIO,CACODMON,"
SQL = SQL & "CAFECDOC From v_kardexvalorizado Where empresacodigo = '" & Empresa & "' and decodigo='" & arCodArti & "'"
SQL = SQL & " and  isnull(estadocosto,0)=1 And almacenvalorizado=1  "
SQL = SQL & " and cast(year(CAFECDOC) as varchar(4)) + replicate('0',2-len(month(CAFECDOC))) + cast(month(CAFECDOC) as varchar(2)) ='" & cAnoMes & "'"
SQL = SQL & " Order By DECODIGO,CAFECDOC,catipmov,tipodecosto,cafecact"

Set adodc1 = VGCNx.Execute(SQL)
'*******************************************************

Set Adodc2 = New ADODB.Recordset
SQL = "Select top 1 * from al_MOvRESMES where SMMESPRO <='" & AnioMesAnterior(cAnoMes) & "' AND empresacodigo='" & VGparametros.empresacodigo & "' AND SMCODIGO='" & arCodArti & "'"
SQL = SQL & " order by smmespro desc "
Set Adodc2 = VGCNx.Execute(SQL)
Call ValorizaEmpresaXArticulo(arCodArti, cAnoMes, Empresa)

Exit Sub
ErrCarga:
        MsgBox Err.Description
        
End Sub
Private Sub ValorizaXArticulo(ByVal vCodArt As String, ByVal arMes As String, ByVal almacen As String)
Dim TCamb As Double
Dim rrsql As New ADODB.Recordset
Dim Li As Integer
Dim nCambio, nSaldo As Double, nCosPro, nCosProUS As Double
Dim nPrecio, nPrecioUS, xPrecio As Double, nCantid As Double
Dim cMesPro, cMesActu, cMesAnte As String
Dim Rsql1 As String
Dim nTipCam, cSql1 As String
Dim cAnoMes As String
'**********Roberto
Dim VALMOV, VALANTE, VALMOVUS, VALANTEUS, valor As Double
Dim nMes, nYear As Long
Dim nSal, nIng, nSaldoInicial As Double
Dim dfecha As Date
Dim csql As String
'On Local Error GoTo ERRAR
xPrecio = 0
nPrecio = 0: nCantid = 0
nPrecioUS = 0
nSal = 0: nIng = 0
nSaldoInicial = 0
cAnoMes = arMes
nCosProUS = 0: nCosPro = 0
If Not adodc1.EOF Then
   nMes = Month(adodc1("CAFECDOC"))
   nYear = Year(adodc1("CAFECDOC"))
   dfecha = adodc1("CAFECDOC")
Else
   dfecha = CDate("01/01/1500") 'fecha no valida o Vacia
End If

Adodc2.Filter = "SMCODIGO = '" & vCodArt & "' AND SMMESPRO ='" & AnioMesAnterior(cAnoMes) & "' and smalma='" & almacen & "'"
If Not Adodc2.EOF Then
   nSaldo = IIf(IsNull(Adodc2!SMantcan), 0 + (Adodc2!SMCANENT - Adodc2!SMCANSAL), Adodc2!SMantcan) + (Adodc2!SMCANENT - Adodc2!SMCANSAL)
   nSaldoInicial = nSaldo
   nCosPro = Adodc2("SMMNPREUNI")
   nCosProUS = numero(Adodc2("SMUSPREUNI"))
   VALANTE = nCosPro * nSaldo
   VALANTEUS = nCosProUS * nSaldo
   cMesPro = AnioMesSiguiente(Adodc2!SMMESPRO)
Else
   nSaldo = 0: nSaldoInicial = 0: nCosPro = 0: nCosProUS = 0
   VALANTE = 0: VALANTEUS = 0
   cMesPro = cAnoMes
End If

Do While cMesPro < Format(Year(dfecha), "0000") + Format(Month(dfecha), "00")
   Rsql1 = "INSERT INTO  MoResMes (SMALMA,SMCODIGO,SMMESPRO,smantcan,SMmnantval,SMCANENT,SMCANSAL,SMMNPREUNI,SMUSPREUNI)VALUES('" & almacen & "','" & vCodArt & "','" & cMesPro & "'," & nSaldoInicial & " ," & VALANTE & "," & nIng & "," & nSal & "," & nCosPro & "," & nCosProUS & ")"
   VGCNx.BeginTrans
  VGCNx.Execute Rsql1
  VGCNx.CommitTrans
  cMesPro = AnioMesSiguiente(cMesPro)
Loop

Do While Not adodc1.EOF
    
   If Year(adodc1("CAFECDOC")) <> nYear Or Month(adodc1("CAFECDOC")) <> nMes Then
  
      cMesPro = Format(nYear, "0000") & Format(nMes, "00")
      Rsql1 = "INSERT INTO  MoResMes (SMALMA,SMCODIGO,SMMESPRO,smantcan,SMmnantval,SMCANENT,SMCANSAL,SMMNPREUNI,SMUSPREUNI)VALUES('" & almacen & "','" & vCodArt & "','" & cMesPro & "'," & nSaldoInicial & "," & VALANTE & "," & nIng & "," & nSal & "," & nCosPro & "," & nCosProUS & ")"
      VGCNx.BeginTrans
      VGCNx.Execute Rsql1
      VGCNx.CommitTrans
      
      cMesActu = (Format(Year(adodc1("CAFECDOC")), "0000") & Format(Month(adodc1("CAFECDOC")), "00"))
      nSaldoInicial = nSaldoInicial + (nIng - nSal)
      nIng = 0
      nSal = 0
      cMesAnte = AnioMesSiguiente(cMesPro)
      While cMesAnte <> cMesActu
            Rsql1 = "INSERT INTO  MoResMes (SMALMA,SMCODIGO,SMMESPRO,smantcan,SMmnantval,SMCANENT,SMCANSAL,SMMNPREUNI,SMUSPREUNI)VALUES('" & almacen & "','" & vCodArt & "','" & cMesAnte & "'," & nSaldoInicial & "," & VALANTE & "," & nIng & "," & nSal & "," & nCosPro & "," & nCosProUS & ")"
            VGCNx.BeginTrans
            VGCNx.Execute Rsql1
            VGCNx.CommitTrans
            cMesAnte = AnioMesSiguiente(cMesAnte)
      Wend
      '*************************************************
      dfecha = adodc1("CAFECDOC")
      nMes = Month(adodc1("CAFECDOC"))
      nYear = Year(adodc1("CAFECDOC"))
            
  Else
     '*************************************************
      If adodc1!DETIPCAM = 0 Or adodc1!DETIPCAM = 1 Then
          TCamb = DevolverTCambio(adodc1!CAFECDOC)
      Else
          TCamb = IIf(IsNull(adodc1!DETIPCAM), 0, adodc1!DETIPCAM)
      End If
      '*************************************************
      
      nCantid = adodc1("DECANTID")
      '*************************************************
      '***DOCUMENTOS EN  DOLARES
      If cNull(adodc1!CACODMON) = "02" Then
         nPrecio = numero(adodc1("DEPRECIO")) * TCamb
      Else
         nPrecio = numero(adodc1("DEPRECIO"))
      End If
      '*************************************************
      '*************************************************
      '***DOCUMENTOS EN SOLES
      If cNull(adodc1!CACODMON) = "02" Then
         nPrecioUS = numero(adodc1("DEPRECIO"))
      Else
         If Round(TCamb, 3) > 0 Then
            nPrecioUS = numero(adodc1("DEPRECIO")) / TCamb
         Else
            nPrecioUS = 0
         End If
      End If
      '*************************************************
      
      If adodc1("CATIPMOV") = "I" Then
         nSaldo = nSaldo + nCantid
         VALMOV = nCantid * nPrecio
         VALMOVUS = nCantid * nPrecioUS 'valorizacion en dolares
         nIng = nIng + nCantid
      Else
         nSaldo = nSaldo - nCantid
         VALMOV = nCantid * nCosPro
         VALMOVUS = nCantid * nCosProUS 'valorizacion en dolares
         nSal = nSal + nCantid
      End If
      
      If adodc1("CATIPMOV") = "I" Then
         If nSaldo <> 0 Then
            nCosPro = numero((VALMOV + VALANTE)) / nSaldo
            nCosProUS = (VALMOVUS + VALANTEUS) / nSaldo
         End If
       Else
        SQL = " update movalmdet set deprecio=" & nCosPro & " where dealma='" & adodc1("dealma") & "'"
        SQL = SQL & " and detd='" & adodc1("detd") & "' and denumdoc='" & adodc1("denumdoc") & "'"
        SQL = SQL & " and deitem='" & adodc1("deitem") & "'"
        Set rrsql = VGCNx.Execute(SQL)
      End If
     
      VALANTE = nCosPro * nSaldo
      VALANTEUS = nCosProUS * nSaldo
      dfecha = adodc1("CAFECDOC")
      adodc1.MoveNext
   End If
   
   
Loop
    
   If IsDate(dfecha) Then
     If Not Year(dfecha) = 1500 Then
        cMesPro = Format(Year(dfecha), "0000") & Format(Month(dfecha), "00")
     Else
        cMesPro = cMesPro
     End If
   '*************************************************
        Rsql1 = "INSERT INTO  MoResMes (SMALMA,SMCODIGO,SMMESPRO,smantcan,SMmnantval,SMCANENT,SMCANSAL,SMMNPREUNI,SMUSPREUNI)VALUES('" & almacen & "','" & vCodArt & "','" & cMesPro & "'," & nSaldoInicial & "," & VALANTE & "," & nIng & "," & nSal & "," & nCosPro & "," & nCosProUS & ")"
   
        VGCNx.BeginTrans
        VGCNx.Execute Rsql1
        VGCNx.CommitTrans
    '*************************************************
         nSaldoInicial = nSaldoInicial + (nIng - nSal)
         cMesActu = AnioMesSiguiente(Format(Year(Now), "0000") & Format(Month(Now), "00"))
         nIng = 0
         nSal = 0
         cMesAnte = AnioMesSiguiente(cMesPro)
         
         While cMesAnte <= cMesActu
               Rsql1 = "INSERT INTO  MoResMes (SMALMA,SMCODIGO,SMMESPRO,smantcan,SMmnantval,SMCANENT,SMCANSAL,SMMNPREUNI,SMUSPREUNI)VALUES('" & almacen & "','" & vCodArt & "','" & cMesAnte & "'," & nSaldoInicial & "," & VALANTE & "," & nIng & "," & nSal & "," & nCosPro & "," & nCosProUS & ")"
               VGCNx.BeginTrans
               VGCNx.Execute Rsql1
               VGCNx.CommitTrans
               cMesAnte = AnioMesSiguiente(cMesAnte)
         Wend
      End If
      Dim cCampo As String
      Dim rs As New ADODB.Recordset

Exit Sub

ERRAR:
MsgBox Err.Description
Resume
End Sub

Private Sub ValorizaEmpresaXArticulo(ByVal vCodArt As String, ByVal arMes As String, ByVal empresa1 As String)
Dim TCamb As Double
Dim rrsql As New ADODB.Recordset
Dim Li As Integer
Dim nCambio, nSaldo As Double, nCosPro, nCosProUS As Double
Dim nPrecio, nPrecioUS, xPrecio As Double, nCantid As Double
Dim cMesPro, cMesActu, cMesAnte As String
Dim Rsql1 As String
Dim nTipCam, cSql1 As String
Dim graba As Integer
Dim cAnoMes As String
Dim VALMOV, VALANTE, VALMOVUS, VALANTEUS, valor As Double
Dim nMes, nYear As Long
Dim nSal, nIng, nSaldoInicial As Double
Dim valing, valsal As Double
Dim dfecha As Date
Dim csql As String
Dim nN As Integer
Dim RSQL As New ADODB.Recordset
'On Local Error GoTo ERRAR
graba = 0
xPrecio = 0
nPrecio = 0: nCantid = 0
nPrecioUS = 0
nSal = 0: nIng = 0
valing = 0: valsal = 0
nSaldoInicial = 0
cAnoMes = arMes
nCosProUS = 0: nCosPro = 0
If Not adodc1.EOF Then
   nMes = Month(adodc1("CAFECDOC"))
   nYear = Year(adodc1("CAFECDOC"))
   dfecha = adodc1("CAFECDOC")
   graba = 1
Else
   dfecha = CDate("01/01/1500") 'fecha no valida o Vacia
End If
cMesPro = cAnoMes

If Not Adodc2.EOF Then
   nSaldo = ESNULO(Adodc2!SMSALDOINI, 0) + ESNULO(Adodc2!SMCANENT, 0) - ESNULO(Adodc2!SMCANSAL, 0)
   nSaldoInicial = nSaldo
  If nSaldo > 0 Then
      nCosPro = Round((ESNULO(Adodc2!smmnvali, 0) + ESNULO(Adodc2!smmnent, 0) - ESNULO(Adodc2!smmnsal, 0)) / nSaldo, 4)
     Else
      nCosPro = Round(numero(Adodc2("SMMNPREFIN")), 4)
   End If
   nCosProUS = Round(numero(Adodc2("SMUSPREFIN")), 4)
   VALANTE = nCosPro * nSaldo
   VALANTEUS = nCosProUS * nSaldo
   If nSaldo > 0 Then
      graba = 1
   End If
Else
   nSaldo = 0: nSaldoInicial = 0: nCosPro = 0: nCosProUS = 0
   VALANTE = 0: VALANTEUS = 0: valing = 0: valsal = 0
End If
nN = 0
If graba = 1 Then
   SQL = "INSERT INTO  al_MovResMes (empresacodigo,SMCODIGO,SMMESPRO,smsaldoini,SMmnvali,SMMNPREINI,SMUSPREINI)VALUES"
   SQL = SQL & " ('" & VGparametros.empresacodigo & "','" & vCodArt & "','" & cMesPro & "'," & nSaldoInicial & "," & VALANTE & "," & nCosPro & "," & nCosProUS & ")"
   Set RSQL = New ADODB.Recordset
   VGCNx.BeginTrans
   Set RSQL = VGCNx.Execute(SQL)
   VGCNx.CommitTrans
End If
      nIng = 0
      nSal = 0

   Do While Not adodc1.EOF
   If nN = 0 And nCosPro = 0 Then
      nCosPro = adodc1("DEPRECIO")
    Else
    '  If nN = 0 And adodc1("DEPRECIO") > 0 Then nCosPro = adodc1("DEPRECIO")
   End If
   nN = 1
   dfecha = adodc1("CAFECDOC")
   nMes = Month(adodc1("CAFECDOC"))
   nYear = Year(adodc1("CAFECDOC"))
            
  If adodc1!DETIPCAM = 0 Or adodc1!DETIPCAM = 1 Then
          TCamb = DevolverTCambio(adodc1!CAFECDOC)
      Else
          TCamb = IIf(IsNull(adodc1!DETIPCAM), 0, adodc1!DETIPCAM)
  End If
      
  nCantid = adodc1("DECANTID")
      '***DOCUMENTOS EN  DOLARES
  If cNull(adodc1!CACODMON) = "02" Then
         nPrecio = numero(adodc1("DEPRECIO")) * TCamb
      Else
         nPrecio = numero(adodc1("DEPRECIO"))
  End If
      '***DOCUMENTOS EN SOLES
  If cNull(adodc1!CACODMON) = "02" Then
         nPrecioUS = numero(adodc1("DEPRECIO"))
      Else
         If Round(TCamb, 3) > 0 Then
            nPrecioUS = numero(adodc1("DEPRECIO")) / TCamb
         Else
            nPrecioUS = 0
         End If
   End If
 
 
 If adodc1("CATIPMOV") = "I" Then
      If TCamb = 0 Then TCamb = 1
         nSaldo = nSaldo + nCantid
         nIng = nIng + nCantid
         If adodc1("tipodecosto") = "C" Or adodc1("tipodecosto") = "A" Or adodc1("tipodecosto") = "D" Then
             If nPrecio <> 0 Then
                VALMOV = nCantid * nPrecio
                VALMOVUS = nCantid * nPrecioUS 'valorizacion en dolares
              Else
                VALMOV = nCantid * nCosPro
                VALMOVUS = nCantid * nPrecioUS 'valorizacion en dolares
             End If
          ElseIf adodc1("tipodecosto") = "P" Then
                 VALMOV = nCantid * nCosPro
                 VALMOVUS = nCantid * nCosPro / TCamb  'valorizacion en dolares
               Else
                 VALMOV = nCantid * nCosPro
                 VALMOVUS = nCantid * nCosPro / TCamb  'valorizacion en dolares
         End If
         valing = valing + VALMOV
      Else
         nSaldo = nSaldo - nCantid
         VALMOV = nCantid * nCosPro
         VALMOVUS = nCantid * nCosProUS 'valorizacion en dolares
         nSal = nSal + nCantid
         valsal = valsal + VALMOV
      End If
      
      If adodc1("CATIPMOV") = "I" Then
         If adodc1("tipodecosto") = "P" Or (adodc1("tipodecosto") <> "C" And nPrecio = 0) Then
            SQL = " update movalmdet set deprecio=" & nCosPro & " where dealma='" & adodc1("dealma") & "'"
            SQL = SQL & " and detd='" & adodc1("detd") & "' and denumdoc='" & adodc1("denumdoc") & "'"
            SQL = SQL & " and deitem='" & adodc1("deitem") & "'"
            Set rrsql = VGCNx.Execute(SQL)
          End If
         If nSaldo <> 0 Then
            nCosPro = numero((VALMOV + VALANTE)) / nSaldo
            nCosProUS = (VALMOVUS + VALANTEUS) / nSaldo
         End If
       Else
        SQL = " update movalmdet set deprecio=" & nCosPro & " where dealma='" & adodc1("dealma") & "'"
        SQL = SQL & " and detd='" & adodc1("detd") & "' and denumdoc='" & adodc1("denumdoc") & "'"
        SQL = SQL & " and deitem='" & adodc1("deitem") & "'"
        Set rrsql = VGCNx.Execute(SQL)
      End If
     
      VALANTE = nCosPro * nSaldo
      VALANTEUS = nCosProUS * nSaldo
      dfecha = adodc1("CAFECDOC")
      adodc1.MoveNext
   
   
  Loop
If graba = 1 Then
   SQL = "update al_MovResMes set SMCANENT=" & nIng & " ,SMCANSAL=" & nSal & ","
   SQL = SQL & " smmnent=" & valing & ",smmnsal=" & valsal & ","
   SQL = SQL & " SMMNPREFIN=" & nCosPro & ",SMUSPREFIN=" & nCosProUS & ""
   SQL = SQL & " where empresacodigo='" & VGparametros.empresacodigo & "' and smcodigo='" & vCodArt & "' and "
   SQL = SQL & " smmespro='" & cMesPro & "' "
    VGCNx.BeginTrans
   Set RSQL = VGCNx.Execute(SQL)
   VGCNx.CommitTrans
End If
 Exit Sub

ERRAR:
MsgBox Err.Description
Resume
End Sub
Private Sub ValorizaEmpresaestablecXArticulo(ByVal vCodArt As String, ByRef estab As String, ByVal arMes As String, ByVal empresa1 As String)
Dim TCamb As Double
Dim rrsql As New ADODB.Recordset
Dim Li As Integer
Dim nCambio, nSaldo As Double, nCosPro, nCosProUS As Double
Dim nPrecio, nPrecioUS, xPrecio As Double, nCantid As Double
Dim cMesPro, cMesActu, cMesAnte As String
Dim Rsql1 As String
Dim nTipCam, cSql1 As String
Dim graba As Integer
Dim cAnoMes As String
Dim VALMOV, VALANTE, VALMOVUS, VALANTEUS, valor As Double
Dim nMes, nYear As Long
Dim nSal, nIng, nSaldoInicial As Double
Dim valing, valsal As Double
Dim dfecha As Date
Dim csql As String
Dim nN As Integer
Dim RSQL As New ADODB.Recordset
'On Local Error GoTo ERRAR
graba = 0
xPrecio = 0
nPrecio = 0: nCantid = 0
nPrecioUS = 0
nSal = 0: nIng = 0
valing = 0: valsal = 0
nSaldoInicial = 0
cAnoMes = arMes
nCosProUS = 0: nCosPro = 0
If Not adodc1.EOF Then
   nMes = Month(adodc1("CAFECDOC"))
   nYear = Year(adodc1("CAFECDOC"))
   dfecha = adodc1("CAFECDOC")
   graba = 1
Else
   dfecha = CDate("01/01/1500") 'fecha no valida o Vacia
End If
cMesPro = cAnoMes
  

If Not Adodc2.EOF Then
   nSaldo = Round(ESNULO(Adodc2!SMSALDOINI, 0) + ESNULO(Adodc2!SMCANENT, 0) - ESNULO(Adodc2!SMCANSAL, 0), 4)
   nSaldoInicial = nSaldo
  If nSaldo > 0 Then
      nCosPro = Round((ESNULO(Adodc2!smmnvali, 0) + ESNULO(Adodc2!smmnent, 0) - ESNULO(Adodc2!smmnsal, 0)) / nSaldo, 4)
     Else
      nCosPro = 0
   End If
   nCosProUS = numero(Adodc2("SMUSPREUNI"))
   VALANTE = nCosPro * nSaldo
   VALANTEUS = nCosProUS * nSaldo
   If nSaldo > 0 Then
      graba = 1
   End If
Else
   nSaldo = 0: nSaldoInicial = 0: nCosPro = 0: nCosProUS = 0
   VALANTE = 0: VALANTEUS = 0: valing = 0: valsal = 0
End If
nN = adodc1.RecordCount
nN = 0
If graba = 1 Then
   SQL = "INSERT INTO  al_MovResMes (empresacodigo,SMCODIGO,SMMESPRO,puntovtacodigo,smsaldoini,SMmnvali,SMMNPREUNI,SMUSPREUNI)VALUES"
   SQL = SQL & " ('" & VGparametros.empresacodigo & "','" & vCodArt & "','" & cMesPro & "','" & estab & "'," & nSaldoInicial & "," & VALANTE & "," & nCosPro & "," & nCosProUS & ")"
   Set RSQL = New ADODB.Recordset
   VGCNx.BeginTrans
   Set RSQL = VGCNx.Execute(SQL)
   VGCNx.CommitTrans
End If
      nIng = 0
      nSal = 0

   Do While Not adodc1.EOF
   If nN = 0 And nCosPro = 0 Then
      nCosPro = adodc1("DEPRECIO")
    Else
    '  If nN = 0 And adodc1("DEPRECIO") > 0 Then nCosPro = adodc1("DEPRECIO")
   End If
   nN = 1
   dfecha = adodc1("CAFECDOC")
   nMes = Month(adodc1("CAFECDOC"))
   nYear = Year(adodc1("CAFECDOC"))
            
  If adodc1!DETIPCAM = 0 Or adodc1!DETIPCAM = 1 Then
          TCamb = DevolverTCambio(adodc1!CAFECDOC)
      Else
          TCamb = IIf(IsNull(adodc1!DETIPCAM), 0, adodc1!DETIPCAM)
  End If
      
  nCantid = adodc1("DECANTID")
      '***DOCUMENTOS EN  DOLARES
  If cNull(adodc1!CACODMON) = "02" Then
         nPrecio = numero(adodc1("DEPRECIO")) * TCamb
      Else
         nPrecio = numero(adodc1("DEPRECIO"))
  End If
      '***DOCUMENTOS EN SOLES
  If cNull(adodc1!CACODMON) = "02" Then
         nPrecioUS = numero(adodc1("DEPRECIO"))
      Else
         If Round(TCamb, 3) > 0 Then
            nPrecioUS = numero(adodc1("DEPRECIO")) / TCamb
         Else
            nPrecioUS = 0
         End If
 End If
      
 If adodc1("CATIPMOV") = "I" Then
      If TCamb = 0 Then TCamb = 1
         nSaldo = nSaldo + nCantid
         nIng = nIng + nCantid
         If adodc1("tipodecosto") = "C" Or adodc1("tipodecosto") = "A" Or adodc1("tipodecosto") = "D" Then
             If nPrecio <> 0 Then
                VALMOV = nCantid * nPrecio
                VALMOVUS = nCantid * nPrecioUS 'valorizacion en dolares
              Else
                VALMOV = nCantid * nCosPro
                VALMOVUS = nCantid * nPrecioUS 'valorizacion en dolares
             End If
          ElseIf adodc1("tipodecosto") = "P" Then
                 VALMOV = nCantid * nCosPro
                 VALMOVUS = nCantid * nCosPro / TCamb  'valorizacion en dolares
               Else
                 VALMOV = nCantid * nCosPro
                 VALMOVUS = nCantid * nCosPro / TCamb  'valorizacion en dolares
         End If
         valing = valing + VALMOV
      Else
         nSaldo = nSaldo - nCantid
         VALMOV = nCantid * nCosPro
         VALMOVUS = nCantid * nCosProUS 'valorizacion en dolares
         nSal = nSal + nCantid
         valsal = valsal + VALMOV
      End If
      
      If adodc1("CATIPMOV") = "I" Then
         If adodc1("tipodecosto") = "P" Or (adodc1("tipodecosto") <> "C" And nPrecio = 0) Then
            SQL = " update movalmdet set deprecio=" & nCosPro & " where dealma='" & adodc1("dealma") & "'"
            SQL = SQL & " and detd='" & adodc1("detd") & "' and denumdoc='" & adodc1("denumdoc") & "'"
            SQL = SQL & " and deitem='" & adodc1("deitem") & "'"
            Set rrsql = VGCNx.Execute(SQL)
          End If
         If nSaldo <> 0 Then
            nCosPro = numero((VALMOV + VALANTE)) / nSaldo
            nCosProUS = (VALMOVUS + VALANTEUS) / nSaldo
         End If
       Else
        SQL = " update movalmdet set deprecio=" & nCosPro & " where dealma='" & adodc1("dealma") & "'"
        SQL = SQL & " and detd='" & adodc1("detd") & "' and denumdoc='" & adodc1("denumdoc") & "'"
        SQL = SQL & " and deitem='" & adodc1("deitem") & "'"
        Set rrsql = VGCNx.Execute(SQL)
      End If
     
      VALANTE = nCosPro * nSaldo
      VALANTEUS = nCosProUS * nSaldo
      dfecha = adodc1("CAFECDOC")
      adodc1.MoveNext
   
  Loop
If graba = 1 Then
   SQL = "update al_MovResMes set SMCANENT=" & nIng & " ,SMCANSAL=" & nSal & ","
   SQL = SQL & " smmnent=" & Round(valing, 4) & ",smmnsal=" & Round(valsal, 4) & ""
'    SQL = SQL & " SMMNPREUNI=" & nCosPro & ",SMUSPREUNI=" & nCosProUS & ""
   SQL = SQL & " where empresacodigo='" & VGparametros.empresacodigo & "' and smcodigo='" & vCodArt & "' and "
   SQL = SQL & " smmespro='" & cMesPro & "' and puntovtacodigo='" & estab & "'"
    VGCNx.BeginTrans
   Set RSQL = VGCNx.Execute(SQL)
   VGCNx.CommitTrans
End If
 Exit Sub

ERRAR:
MsgBox Err.Description
Resume
End Sub
Public Function SaldoArti(ByVal cAlma As String, ByVal cArti As String, ByVal cn As ADODB.Connection) As Double
Dim rs As New ADODB.Recordset
rs.Open "SELECT STKART.STSKDIS FROM STKART where STALMA='" & cAlma & "' AND STCODIGO='" & cArti & "'", VGCNx, adOpenForwardOnly, adLockReadOnly
If Not rs.EOF Then
   SaldoArti = rs!STSKDIS
Else
   SaldoArti = 0
End If
End Function

Public Sub CalcularStock(ByVal arAlma As String, ByVal arCodArti As String, ByVal Fechtran As Date)
Dim cSql1 As String, CSQL2 As String
On Error GoTo ErrCarga
'*******************************************************
'**Carga todos los Movimientos al respectivo Recordset a partir de 1er dia del Mes en que se realizo la transacciojn0
'*******************************************************
cAnoMes = Format(Year(Fechtran), "0000") + Format(Month(Fechtran), "00")

Set adodc1 = New ADODB.Recordset
cSql1 = "Select CATIPCAM,DETIPCAM,DECANTID,DECODIGO,CATIPMOV,DEALMA,DETD,DENUMDOC,DEITEM,CATD,DEPRECIO,CACODMON,CAFECDOC From MovAlmDet A Inner Join MovAlmCab B On B.CAALMA = A.DEALMA And B.CATD = A.DETD  And B.CANUMDOC = A.DENUMDOC "
cSql1 = cSql1 & " Where CAALMA = '" & arAlma & "' and not (CATD='GS' And CACODMOV='GF' And CASITGUI='F') "
'cSql1 = cSql1 & " And CASITGUI<>'A'  and DECODIGO='" & arCodArti & "' and FORMAT( YEAR(CAFECDOC),'0000' )+ FORMAT( MONTH(CAFECDOC),'00' ) >='" & cAnoMes & "' Order By DECODIGO,CAFECDOC,CAHORA"
cSql1 = cSql1 & " And CASITGUI<>'A'  and DECODIGO='" & arCodArti & "' and right('0000'+ltrim(str(YEAR(CAFECDOC))),4)+ "
cSql1 = cSql1 & " right('00'+ltrim(str(MONTH(CAFECDOC))),2) >='" & cAnoMes & "' Order By DECODIGO,CAFECDOC,CAHORA"
adodc1.Open cSql1, VGCNx, adOpenForwardOnly, adLockReadOnly
'*******************************************************

Set Adodc2 = New ADODB.Recordset
Adodc2.Open "Select * from MORESMES where SMMESPRO >='" & AnioMesAnterior(cAnoMes) & "' AND SMALMA='" & arAlma & "' AND SMCODIGO='" & arCodArti & "'", VGCNx, adOpenStatic
CSQL2 = "delete From MORESMES Where SMMESPRO>='" & cAnoMes & "' and SMALMA='" & arAlma & "' and SMCODIGO='" & arCodArti & "'"
VGCNx.Execute CSQL2
Call ValorizaXArticulo(arCodArti, cAnoMes, arAlma) 'Procedimiento que Valoriza por Articulo

Exit Sub
ErrCarga:
        MsgBox Err.Description
        Resume
        
End Sub
'***************************************************
Public Function ExisteEnStock(ByVal cAlma As String, ByVal cArti As String, ByVal cn As ADODB.Connection) As Boolean
Dim rs As New ADODB.Recordset
rs.Open "SELECT STSKDIS FROM STKART where STALMA='" & cAlma & "' AND STCODIGO='" & cArti & "'", VGCNx, adOpenForwardOnly, adLockReadOnly
ExisteEnStock = IIf(Not rs.EOF, True, False)
rs.Close
End Function
Public Function ExisteEnStockSerie(ByVal cAlma As String, ByVal cArti As String, ByVal arSerie As String, ByVal cn As ADODB.Connection) As Boolean
Dim rs As New ADODB.Recordset
rs.Open "SELECT STSSKDIS FROM STKSERI where STSALMA='" & cAlma & "' AND STSCODIGO='" & cArti & "' AND STSSERIE='" & arSerie & "'", VGCNx, adOpenForwardOnly, adLockReadOnly
ExisteEnStockSerie = IIf(Not rs.EOF, True, False)
rs.Close
End Function

Public Function ExisteEnStockAlmacenes(ByVal cArti As String, ByVal cn As ADODB.Connection) As Boolean
Dim rs As New ADODB.Recordset
rs.Open "SELECT STSKDIS FROM STKART where  STCODIGO='" & cArti & "' And STSKDIS<>0 ", VGCNx, adOpenForwardOnly, adLockReadOnly
ExisteEnStockAlmacenes = IIf(Not rs.EOF, True, False)
rs.Close
End Function

'Public Function GrabaSerie(ByVal arAlma As String, ByVal cArti As String, ByVal cSerie As String, ByVal TipoDoc As String, ByVal cn As ADODB.Connection, CANTIDAD) As Boolean
'End Function

Public Sub CalculaStockSerie(ByVal arAlma As String, ByVal arCodArti As String, Optional ByVal Fechtran As Variant)
Dim cSql1, xserie As String, CSQL2 As String
On Error GoTo ErrCarga
Dim valor As Double
Dim n As Long
Dim cCampo As String
'*******************************************************
'**Carga todos los Movimientos al respectivo Recordset a partir de 1er dia del Mes en que se realizo la transacciojn0
'*******************************************************
Set adodc1 = New ADODB.Recordset
If Not IsDate(Fechtran) Then
   cSql1 = "Select CATIPCAM,DECANTID,DECODIGO,CATIPMOV,DETD,CATD,DEPRECIO,CACODMON,CAFECDOC,DESERIE From MovAlmDet A Inner Join MovAlmCab B On B.CAALMA = A.DEALMA And B.CATD = A.DETD  And B.CANUMDOC = A.DENUMDOC "
   cSql1 = cSql1 & " Where CAALMA = '" & arAlma & "' and not (CATD='GS' And CACODMOV='GF' And CASITGUI='F') "
   cSql1 = cSql1 & " And CASITGUI<>'A'  and DECODIGO='" & arCodArti & "'  Order By DECODIGO,CAFECDOC,CAHORA"
Else
   cAnoMes = Format(Year(Fechtran), "0000") + Format(Month(Fechtran), "00")
   dfecha = Format(Fechtran, "dd/mm/yyyy")
   cSql1 = "Select CATIPCAM,DECANTID,DECODIGO,CATIPMOV,DETD,CATD,DEPRECIO,CACODMON,CAFECDOC,DESERIE From MovAlmDet A Inner Join MovAlmCab B On B.CAALMA = A.DEALMA And B.CATD = A.DETD  And B.CANUMDOC = A.DENUMDOC "
   cSql1 = cSql1 & " Where CAALMA = '" & arAlma & "' and not (CATD='GS' And CACODMOV='GF' And CASITGUI='F') "
   cSql1 = cSql1 & " And CASITGUI<>'A'  and DECODIGO='" & arCodArti & "' AND  FORMAT( YEAR(CAFECDOC),'0000' )+ FORMAT( MONTH(CAFECDOC),'00' ) >='" & cAnoMes & "' Order By DECODIGO,CAFECDOC,CAHORA"
End If
adodc1.Open cSql1, VGCNx, adOpenStatic, adLockReadOnly
'*******************************************************
cSql1 = "Delete from STKSERI WHERE  STSALMA='" & arAlma & "' and STSCODIGO='" & arCodArti & "'"
VGCNx.Execute cSql1

cSql1 = "UPDATE STKSERI SET STSSKDIS=0 WHERE  STSALMA='" & arAlma & "' and STSCODIGO='" & arCodArti & "'"
VGCNx.Execute cSql1

Do While Not adodc1.EOF
    If ExisteEnStockSerie(arAlma, arCodArti, cNull(adodc1!DESERIE), VGCNx) Then
       If adodc1!catipmov = "I" Then
          valor = adodc1!DECANTID
          cCampo = " STSSKDIS = STSSKDIS+" & valor
       Else
          valor = adodc1!DECANTID
          cCampo = " STSSKDIS = STSSKDIS-" & valor
       End If
       cSql1 = "update STKSERI set " & cCampo & " WHERE  STSALMA='" & arAlma & "' and STSCODIGO='" & arCodArti & "'AND STSSERIE='" & adodc1!DESERIE & "'"
    Else
       If IsNull(adodc1!DESERIE) Then
          If adodc1!catipmov = "I" Then
             xserie = "INGRESO SIN SERIE" & Format(n, "00")
          Else
             xserie = "SALIDA SIN SERIE" & Format(n, "00")
          End If
       Else
          xserie = adodc1!DESERIE
       End If
       
      cSql1 = "insert into STKSERI (STSALMA,STSCODIGO,STSSERIE,STSSKDIS)  VALUES ('" & arAlma & "','" & arCodArti & "','" & xserie & "' ," & adodc1!DECANTID & ")"

       
    End If
    
    VGCNx.Execute cSql1
    adodc1.MoveNext
Loop

adodc1.Close

Exit Sub
ErrCarga:
        MsgBox Err.Description
        
End Sub

'*************************************************
'NO UTLIZAR EL ARGUMENTO FECHTRAN ESTA POR REVISAR
'*************************************************
Public Sub CalculaStockLOTE(ByVal arAlma As String, ByVal arCodArti As String, Optional ByVal Fechtran As Variant)
Dim cSql1 As String, CSQL2 As String
On Error GoTo ErrCarga
Dim nSaldo As Double
Dim cCampo As String
'*******************************************************
'**Carga todos los Movimientos al respectivo Recordset a partir de 1er dia del Mes en que se realizo la transacciojn0
'*******************************************************
Set adodc1 = New ADODB.Recordset
If Not IsDate(Fechtran) Then
   cSql1 = "Select CATIPCAM,DECANTID,DECODIGO,CATIPMOV,DETD,CATD,DEPRECIO,CACODMON,CAFECDOC,DELOTE From MovAlmDet A Inner Join MovAlmCab B On B.CAALMA = A.DEALMA And B.CATD = A.DETD  And B.CANUMDOC = A.DENUMDOC "
   cSql1 = cSql1 & " Where CAALMA = '" & arAlma & "' and not (CATD='GS' And CACODMOV='GF' And CASITGUI='F') "
   cSql1 = cSql1 & " And CASITGUI<>'A'  and DECODIGO='" & arCodArti & "'  Order By DECODIGO,CAFECDOC,CAHORA"
Else
   cAnoMes = Format(Year(Fechtran), "0000") + Format(Month(Fechtran), "00")
   dfecha = Format(Fechtran, "dd/mm/yyyy")
   cSql1 = "Select CATIPCAM,DECANTID,DECODIGO,CATIPMOV,DETD,CATD,DEPRECIO,CACODMON,CAFECDOC,DELOTE From MovAlmDet A Inner Join MovAlmCab B On B.CAALMA = A.DEALMA And B.CATD = A.DETD  And B.CANUMDOC = A.DENUMDOC "
   cSql1 = cSql1 & " Where CAALMA = '" & arAlma & "' and not (CATD='GS' And CACODMOV='GF' And CASITGUI='F') "
   cSql1 = cSql1 & " And CASITGUI<>'A'  and DECODIGO='" & arCodArti & "' AND FORMAT( YEAR(CAFECDOC),'0000' )+ FORMAT( MONTH(CAFECDOC),'00' ) >='" & cAnoMes & "' Order By DECODIGO,CAFECDOC,CAHORA"
End If
adodc1.Open cSql1, VGCNx, adOpenStatic, adLockReadOnly
'*******************************************************
''''''''''''''''SELECT STSALMA, STSCODIGO,STSLOTE, STSLKDIS, STSFECFAB, STSFECVEN, STSLKCOM FROM STKLOTE;
VGCNx.Execute "update  STKLOTE set STSLKDIS=0 where STSALMA='" & arAlma & "' AND STSCODIGO='" & arCodArti & "'"
nSaldo = 0
Do While Not adodc1.EOF
   If adodc1!catipmov = "I" Then
      nSaldo = adodc1!DECANTID
      cCampo = "STSLKDIS =STSLKDIS + " & nSaldo
   Else
      nSaldo = adodc1!DECANTID
      cCampo = "STSLKDIS =STSLKDIS - " & nSaldo
   End If
    
   If ExisteEnStockLote(arAlma, arCodArti, cNull(adodc1!DELOTE), VGCNx) Then
      cSql1 = "update STKLOTE set " & cCampo & " WHERE  STSALMA='" & arAlma & "' and STSCODIGO='" & arCodArti & "'AND STSLOTE='" & adodc1!DELOTE & "'"
   Else
      cSql1 = "insert into STKLOTE (STSALMA, STSCODIGO,STSLOTE, STSLKDIS)  VALUES ('" & arAlma & "','" & arCodArti & "','" & adodc1!DELOTE & "' ," & nSaldo & ")"
   End If

   VGCNx.Execute cSql1
   adodc1.MoveNext
Loop

adodc1.Close

Exit Sub
ErrCarga:
        MsgBox Err.Description
        
End Sub

Public Function ExisteEnStockLote(ByVal cAlma As String, ByVal cArti As String, ByVal arLOTE As String, ByVal cn As ADODB.Connection) As Boolean
Dim rs As New ADODB.Recordset
rs.Open "SELECT STSLKDIS FROM STKLOTE where STSALMA='" & cAlma & "' AND STSCODIGO='" & cArti & "' AND STSLOTE='" & arLOTE & "'", VGCNx, adOpenForwardOnly, adLockReadOnly
ExisteEnStockLote = IIf(Not rs.EOF, True, False)
rs.Close
End Function
''


Public Sub CalculaSaldoNoValorizado(ByVal arAlma As String, ByVal arCodArti As String, ByVal Fechtran As Date)
Dim cSql1 As String
On Error GoTo ErrCarga
Dim nIng As Double
Dim nSal, nSaldo As Double
'*******************************************************
'**Carga todos los Movimientos al respectivo Recordset a partir de 1er dia del Mes en que se realizo la transacciojn0
'*******************************************************
'**********INGRESOS*************
Set adodc1 = New ADODB.Recordset
cSql1 = "Select SUM(DECANTID) as Ing From MovAlmDet A Inner Join MovAlmCab B On B.CAALMA = A.DEALMA And B.CATD = A.DETD  And B.CANUMDOC = A.DENUMDOC "
cSql1 = cSql1 & " Where CAALMA = '" & arAlma & "' and not (CATD='GS' And CACODMOV='GF' And CASITGUI='F') "
cSql1 = cSql1 & " And CASITGUI<>'A'  and DECODIGO='" & arCodArti & "' and CATIPMOV='I'"
adodc1.Open cSql1, VGCNx, adOpenForwardOnly, adLockReadOnly
If Not adodc1.EOF Then
   nIng = IIf(IsNull(adodc1!ING), 0, adodc1!ING)
Else
   nIng = 0
End If
'**********SALIDAS*************
Set adodc1 = New ADODB.Recordset
cSql1 = "Select SUM(DECANTID) as Sal From MovAlmDet A Inner Join MovAlmCab B On B.CAALMA = A.DEALMA And B.CATD = A.DETD  And B.CANUMDOC = A.DENUMDOC "
cSql1 = cSql1 & " Where CAALMA = '" & arAlma & "' and not (CATD='GS' And CACODMOV='GF' And CASITGUI='F') "
cSql1 = cSql1 & " And CASITGUI<>'A'  and DECODIGO='" & arCodArti & "' and CATIPMOV='S'"
adodc1.Open cSql1, VGCNx, adOpenForwardOnly, adLockReadOnly
If Not adodc1.EOF Then
   nSal = IIf(IsNull(adodc1!SAL), 0, adodc1!SAL)
Else
   nSal = 0
End If

nSaldo = nIng - nSal

If ExisteEnStock(arAlma, arCodArti, VGCNx) Then
   VGCNx.Execute "Update STKART SET STSKDIS=" & nSaldo & " WHERE STALMA='" & arAlma & "' AND STCODIGO='" & arCodArti & "'"
Else
   VGCNx.Execute "Insert Into STKART (STALMA,STCODIGO,STSKDIS)Values('" & arAlma & "','" & arCodArti & "'," & nSaldo & ")"
End If

adodc1.Close

Exit Sub
ErrCarga:
        MsgBox Err.Description
        
End Sub

Public Function ArticuloConMovimiento(ByVal arCodArti As String, ByVal cn As ADODB.Connection) As Boolean
Dim rs As New ADODB.Recordset
Dim cSql1 As String
'************Veririca en Todos los Almacenes
'cSql1 = "Select top 1 Decodigo  From MovAlmDet A Inner Join MovAlmCab B On B.CAALMA = A.DEALMA And B.CATD = A.DETD  And B.CANUMDOC = A.DENUMDOC "
cSql1 = "Select Decodigo  From MovAlmDet A Inner Join MovAlmCab B On B.CAALMA = A.DEALMA And B.CATD = A.DETD  And B.CANUMDOC = A.DENUMDOC "
cSql1 = cSql1 & " Where not (CATD='GS' And CACODMOV='GF' And CASITGUI='F') "
cSql1 = cSql1 & " And CASITGUI<>'A'  and DECODIGO='" & arCodArti & "'"
rs.Open cSql1, VGCNx, adOpenForwardOnly, adLockReadOnly
ArticuloConMovimiento = IIf(Not rs.EOF, True, False)
rs.Close
End Function

Public Function ArticuloConLotes(ByVal arCodArti As String, ByVal Lote As String, ByVal cn As ADODB.Connection) As Boolean
Dim rs As New ADODB.Recordset
Dim cSql1 As String
'************Veririca en Todos los Almacenes
cSql1 = "Select top 1 Decodigo  From MovAlmDet A Inner Join MovAlmCab B On B.CAALMA = A.DEALMA And B.CATD = A.DETD  And B.CANUMDOC = A.DENUMDOC "
cSql1 = cSql1 & " Where not (CATD='GS' And CACODMOV='GF' And CASITGUI='F') "
cSql1 = cSql1 & " And CASITGUI<>'A'  and DECODIGO='" & arCodArti & "' and DELOTE='" & Lote & "'"
rs.Open cSql1, VGCNx, adOpenForwardOnly, adLockReadOnly
ArticuloConLotes = IIf(Not rs.EOF, True, False)
rs.Close
End Function

Public Function SaldoLote(ByVal cAlma As String, ByVal cArti As String, ByVal Lote As String, ByVal cn As ADODB.Connection) As Double
Dim rs As New ADODB.Recordset
rs.Open "SELECT STSLKDIS FROM STKlote where STSALMA='" & cAlma & "' AND STSCODIGO='" & cArti & "' and STSLOTE='" & Lote & "'", VGCNx, adOpenForwardOnly, adLockReadOnly
If Not rs.EOF Then
   SaldoLote = rs!STSLKDIS
Else
   SaldoLote = 0
End If
rs.Close
End Function
'*********SI NO EXISTE EN EL ALMACEN LO REGISTRA EN STKART
Public Sub VerificaKIT(ByVal arAlma As String, ByVal arKit As String, ByVal cn As ADODB.Connection)
If Not ExisteEnStock(arAlma, arKit, VGCNx) Then
   VGCNx.Execute "Insert Into STKART (STALMA,STCODIGO,STSKDIS,STKPREULT,STKPREPRO)Values('" & arAlma & "','" & arKit & "',0,0,0)"
End If
End Sub

'S= serie     L=Lote
Public Function EsSerie_Lote(ByVal arCodigo As String, ByVal cn As ADODB.Connection) As String
Dim rs As New ADODB.Recordset
rs.Open "SELECT AFSERIE,AFLOTE FROM MAEART where ACODIGO='" & arCodigo & "'", VGCNx, adOpenForwardOnly, adLockReadOnly
If Not rs.EOF Then
   If cNull(rs!afserie) = "S" Then EsSerie_Lote = "S"
   If cNull(rs!afLOTE) = "S" Then EsSerie_Lote = "L"
Else
   MsgBox "El Articulo " & arCodigo & " Registrado en el Documento no existe en el maestro", vbInformation, "Error al Eliminar"
   EsSerie_Lote = ""
End If
rs.Close
End Function

'.ArticuloConOrdFabri
Public Function ArticuloConOrdFabri(ByVal arAlma As String, ByVal arCodArti As String, ByVal OrdFab As String, ByVal cn As ADODB.Connection) As Boolean
Dim rs As New ADODB.Recordset
Dim cSql1 As String
'************Veririca en Todos los Almacenes
cSql1 = "Select top 1 Decodigo  From MovAlmDet A Inner Join MovAlmCab B On B.CAALMA = A.DEALMA And B.CATD = A.DETD  And B.CANUMDOC = A.DENUMDOC "
cSql1 = cSql1 & " Where not (CATD='GS' And CACODMOV='GF' And CASITGUI='F') "
cSql1 = cSql1 & " And CASITGUI<>'A'  and DEALMA='" & arAlma & "' AND DECODIGO='" & arCodArti & "' and DEORDFAB='" & OrdFab & "'"
rs.Open cSql1, VGCNx, adOpenStatic, adLockReadOnly
ArticuloConOrdFabri = IIf(Not rs.EOF, True, False)
rs.Close
End Function

Public Function ExisteOrdenFab(ByVal cAlma As String, ByVal cArti As String, ByVal arOrdenFab As String, ByVal cn As ADODB.Connection) As Boolean
Dim rs As New ADODB.Recordset
rs.Open "SELECT ORD_CODART FROM ORDE_FAB where ORD_ALMA='" & cAlma & "' AND ORD_CODART='" & cArti & "' AND ORD_FABNUM='" & arOrdenFab & "'", VGCNx, adOpenStatic, adLockReadOnly
ExisteOrdenFab = IIf(Not rs.EOF, True, False)
rs.Close
End Function

Public Sub BorrarServicios(ByVal cAlma As String, ByVal cn As ADODB.Connection)
Dim rs As New ADODB.Recordset
Dim SQL As String
VGCNx.Execute "DELETE STKART  FROM STKART INNER JOIN MAEART ON STKART.STCODIGO = MAEART.ACODIGO WHERE isnull(MAEART.afstock,1)=0  "
VGCNx.Execute "DELETE al_MOvRESMES  FROM al_MOvRESMES a INNER JOIN MAEART b ON a.SMCODIGO = b.ACODIGO WHERE isnull(b.afstock,1)=0 AND a.empresacodigo='" & VGparametros.empresacodigo & "'"
End Sub

Public Function CCostoconSalidas(ByVal cAlma As String, ByVal cCosto As String, ByVal cn As ADODB.Connection, Optional ByVal artCod As String) As Boolean
Dim rs As New ADODB.Recordset
Dim cSql1 As String
'************Veririca si hubo Salidas con ese centro de Costos
   cSql1 = "Select top 1 Decodigo  From MovAlmDet A Inner Join MovAlmCab B On B.CAALMA = A.DEALMA And B.CATD = A.DETD  And B.CANUMDOC = A.DENUMDOC "
   cSql1 = cSql1 & " Where not (CATD='GS' And CACODMOV='GF' And CASITGUI='F') "

If artCod = "" Then
   cSql1 = cSql1 & " And CASITGUI<>'A'  and DEALMA='" & cAlma & "' AND CATIPMOV='S' AND DECENCOS='" & cCosto & "'"
Else
   cSql1 = cSql1 & " And CASITGUI<>'A'  and DEALMA='" & cAlma & "' AND CATIPMOV='S' AND DECENCOS='" & cCosto & "' and decodigo='" & artCod & "'"
End If
rs.Open cSql1, VGCNx, adOpenStatic, adLockReadOnly
CCostoconSalidas = IIf(Not rs.EOF, True, False)
rs.Close
End Function

Public Sub GeneraSaldosxAlmacen(Empresa As String, mesant As String, mesact As String)
   Set VGCommandoSP = New ADODB.Command
   VGCommandoSP.ActiveConnection = VGgeneral
   VGCommandoSP.CommandType = adCmdStoredProc
   VGCommandoSP.CommandText = "al_SaldoMensualxAlmacen_pro"
   VGCommandoSP.Parameters.Refresh
   With VGCommandoSP
       .Parameters("@base") = VGParamSistem.BDEmpresa
       .Parameters("@empresa") = VGparametros.empresacodigo
       .Parameters("@mesant") = mesant
       .Parameters("@mesact") = mesact
       .Parameters("@computer") = VGComputer
       .Execute
   End With
End Sub
