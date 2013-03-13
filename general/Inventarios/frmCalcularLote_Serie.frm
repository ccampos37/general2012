VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCalcularLote_Serie 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calcular Lotes"
   ClientHeight    =   2430
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   6675
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1875
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5655
      Begin VB.ComboBox Combo1 
         Height          =   288
         ItemData        =   "frmCalcularLote_Serie.frx":0000
         Left            =   1080
         List            =   "frmCalcularLote_Serie.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   270
         Width           =   4335
      End
      Begin VB.Label Label2 
         Caption         =   "Total ==>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   780
         Width           =   915
      End
      Begin VB.Label Label2 
         Caption         =   "Transcurridos==>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2580
         TabIndex        =   3
         Top             =   810
         Width           =   1635
      End
   End
End
Attribute VB_Name = "frmCalcularLote_Serie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''Option Explicit
'''Dim PCount As Long
'''Dim cConexAux As ADODB.Connection
'''Dim Adodc1 As ADODB.Recordset
'''Dim Adodc2 As ADODB.Recordset
'''Dim rS As ADODB.Recordset
'''Dim cRt As String
'''Dim almacen As String
'''Dim nTra As Integer
'''
'''Private Sub cmd_exit_Click()
'''Unload Me
'''End Sub
'''
'''Private Sub Form_Load()
''' central Me
''' Call Carga_Almacen
''' dFecVal.Value = Format(Now, "dd/mm/yyyy")
'''End Sub
'''
'''Private Sub opt1_Click()
'''If opt2.Value = True Then
'''    Frame2.Enabled = True
'''Else
'''    Frame2.Enabled = False
'''End If
'''End Sub
'''
'''Private Sub opt2_Click()
'''If opt2.Value = True Then
'''    Frame2.Enabled = True
'''Else
'''    Frame2.Enabled = False
'''End If
'''End Sub
'''
'''Private Sub Text1_DblClick()
'''Dim Adodc2 As New ADODB.Recordset
'''
'''    'Adodc2.Open "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD from MaeArt p, StkArt n Where n.STSKDIS<> 0 and  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & VGAlma & "'  ", cConexCom, adOpenStatic, adLockOptimistic
'''    Adodc2.Open "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD from MaeArt p, StkArt n Where  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & VGAlma & "'  ", cConexCom, adOpenStatic, adLockOptimistic
'''    'frmReferencia.conectar Adodc2, "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD from MaeArt p, StkArt n Where n.STSKDIS<> 0 and  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & VGAlma & "'"
'''    frmReferencia.Conectar Adodc2, "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD from MaeArt p, StkArt n Where  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & VGAlma & "'"
'''    frmReferencia.Label1.Caption = "Artículos"
'''    frmReferencia.show vbmodal
'''    Adodc2.Close
'''    If vGUtil(1) <> "" Then
'''            Text1 = (vGUtil(1))
'''    End If
'''    If Text1 <> "" And Text2 <> "" And Text1 > Text2 Then
'''            MsgBox "Ingrese un codigo menor al fin ", vbOKOnly, "Error"
'''            Exit Sub
'''   End If
'''   If Text1 <> "" Then
'''            Text2.Enabled = True
'''            Text2.SetFocus
'''   End If
'''End Sub
'''
'''Private Sub Text2_DblClick()
'''Dim Adodc2 As New ADODB.Recordset
'''    'Adodc2.Open "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD from MaeArt p, StkArt n Where n.STSKDIS<> 0 and  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & VGAlma & "'  ", cConexCom, adOpenStatic, adLockOptimistic
'''    Adodc2.Open "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD from MaeArt p, StkArt n Where   p.ACODIGO =  n.STCODIGO and n.STALMA = '" & VGAlma & "'  ", cConexCom, adOpenStatic, adLockOptimistic
'''    'frmReferencia.conectar Adodc2, "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD from MaeArt p, StkArt n Where n.STSKDIS<> 0 and  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & VGAlma & "'  "
'''    frmReferencia.Conectar Adodc2, "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD from MaeArt p, StkArt n Where  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & VGAlma & "'  "
'''    frmReferencia.Label1.Caption = "Artículos"
'''    frmReferencia.show vbmodal
'''    Adodc2.Close
'''    If vGUtil(1) <> "" Then
'''            Text2 = (vGUtil(1))
'''    End If
'''    If Text1 <> "" And Text2 <> "" And Text1 > Text2 Then
'''            MsgBox "Ingrese un codigo menor al fin ", vbOKOnly, "Error"
'''            Exit Sub
'''   End If
'''   If Text2 <> "" Then
'''            Text2.Enabled = True
'''            Text2.SetFocus
'''   End If
'''End Sub
'''Private Sub Combo1_Click()
'''rS.MoveFirst
'''rS.Move Combo1.ListIndex
'''almacen = Format(rS(0), "00")
'''End Sub
'''Private Sub Carga_Almacen()
'''Dim rSql As String
'''Dim i As Integer
'''rSql = "Select TAALMA,TADESCRI FROM TabAlm "
'''Set rS = New ADODB.Recordset
'''rS.Open rSql, cConexCom, adOpenStatic
'''Do While Not rS.EOF
'''     Combo1.AddItem (rS(1))
'''     rS.MoveNext
'''     If rS.EOF Then Exit Do
'''Loop
'''rS.MoveFirst
'''For i = 0 To rS.RecordCount - 1
'''  If rS(0) = VGAlma Then
'''    Combo1.ListIndex = i
'''    Exit For
'''  Else
'''    rS.MoveNext
'''  End If
'''Next
'''End Sub
'''Private Sub ValorizaXArticulo(ByVal vCodArt As String, ByVal arMes As String)
'''Dim TCamb As Double
'''Dim Li As Integer
'''Dim nCambio, nSaldo As Double, nCosPro, nCosProUS As Double
'''Dim nPrecio, nPrecioUS, xPrecio As Double, nCantid As Double
'''Dim cMesPro, cMesActu, cMesAnte As String
'''Dim Rsql1 As String
'''Dim nTipCam, cSql1 As String
''''**********Roberto
'''Dim VALMOV, VALANTE, VALMOVUS, VALANTEUS As Double
'''Dim nMes, nYear As Long
'''Dim nSal, nIng, nSaldoInicial As Double
'''Dim dfecha As Date
'''Dim csql As String
'''Dim XNUMDOC As String
'''On Local Error GoTo ERRAR
'''
'''Adodc1.Filter = " Decodigo='" & vCodArt & "'"
'''xPrecio = 0
'''nPrecio = 0: nCantid = 0
'''nPrecioUS = 0
'''nSal = 0: nIng = 0
'''nSaldoInicial = 0
'''Adodc1.MoveFirst
'''nCosProUS = 0: nCosPro = 0
'''nMes = Month(Adodc1("CAFECDOC"))
'''nYear = Year(Adodc1("CAFECDOC"))
'''dfecha = Adodc1("CAFECDOC")
'''
'''
'''Do While Not Adodc1.EOF
'''
'''   If Year(Adodc1("CAFECDOC")) <> nYear Or Month(Adodc1("CAFECDOC")) <> nMes Then
'''
'''      cMesPro = Format(nYear, "0000") & Format(nMes, "00")
'''      If cMesPro = arMes Then
'''         Rsql1 = "INSERT INTO  MoResMes (SMALMA,SMCODIGO,SMMESPRO,SMSALDOINI,SMCANENT,SMCANSAL,SMMNPREUNI,SMUSPREUNI)VALUES('" & almacen & "','" & vCodArt & "','" & cMesPro & "'," & nSaldoInicial & "," & nIng & "," & nSal & "," & nCosPro & "," & nCosProUS & ")"
'''         cConexCom.BeginTrans
'''         cConexCom.Execute Rsql1
'''         cConexCom.CommitTrans
'''      End If
'''
'''      cMesActu = (Format(Year(Adodc1("CAFECDOC")), "0000") & Format(Month(Adodc1("CAFECDOC")), "00"))
'''      nSaldoInicial = nSaldoInicial + (nIng - nSal)
'''      nIng = 0
'''      nSal = 0
'''      cMesAnte = AnioMesSiguiente(cMesPro)
'''      While cMesAnte <> cMesActu
'''            If cMesPro = arMes Then
'''               Rsql1 = "INSERT INTO  MoResMes (SMALMA,SMCODIGO,SMMESPRO,SMSALDOINI,SMCANENT,SMCANSAL,SMMNPREUNI,SMUSPREUNI)VALUES('" & almacen & "','" & vCodArt & "','" & cMesAnte & "'," & nSaldoInicial & "," & nIng & "," & nSal & "," & nCosPro & "," & nCosProUS & ")"
'''               cConexCom.BeginTrans
'''               cConexCom.Execute Rsql1
'''               cConexCom.CommitTran
'''           End If
'''           cMesAnte = AnioMesSiguiente(cMesAnte)
'''      Wend
'''      '*************************************************
'''      dfecha = Adodc1("CAFECDOC")
'''      nMes = Month(Adodc1("CAFECDOC"))
'''      nYear = Year(Adodc1("CAFECDOC"))
'''
'''  Else
'''
'''     '*************************************************
'''      If Adodc1!CATIPCAM = 0 Or Adodc1!CATIPCAM = 1 Then
'''            If UCase(Dir$(cRuta4)) = UCase(VGNameCont & ".MDB") Then
'''               TCamb = Val(Devolver_Dato(3, Format(dfecha, "DD/MM/YYYY"), "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "ME", "TIPOMON_CODIGO"))
'''            Else
'''               If UCase(Dir$(cRuta2)) = "BDCOMUN.MDB" Then
'''                  TCamb = Val(Devolver_Dato(1, Format(dfecha, "DD/MM/YYYY"), "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "ME", "TIPOMON_CODIGO"))
'''               Else
'''                  TCamb = VGTipCamb
'''               End If
'''            End If
'''      Else
'''          TCamb = Adodc1!CATIPCAM
'''      End If
'''      '*************************************************
'''
'''      nCantid = Adodc1("DECANTID")
'''
'''      '*************************************************
'''      '***DOCUMENTOS EN  DOLARES
'''      If cNull(Adodc1!CACODMON) = "ME" Then
'''         nPrecio = Adodc1("DEPRECIO") * TCamb
'''      Else
'''         nPrecio = Adodc1("DEPRECIO")
'''      End If
'''      '*************************************************
'''      '*************************************************
'''      '***DOCUMENTOS EN SOLES
'''      If cNull(Adodc1!CACODMON) = "ME" Then
'''         nPrecioUS = Adodc1("DEPRECIO")
'''      Else
'''         If Round(TCamb, 3) > 0 Then
'''            nPrecioUS = Adodc1("DEPRECIO") / TCamb
'''         Else
'''            nPrecioUS = 0
'''         End If
'''      End If
'''      '*************************************************
'''
'''      If Adodc1("CATIPMOV") = "I" Then
'''         nSaldo = nSaldo + nCantid
'''         VALMOV = nCantid * nPrecio
'''         VALMOVUS = nCantid * nPrecioUS 'valorizacion en dolares
'''         nIng = nIng + nCantid
'''      Else
'''         nSaldo = nSaldo - nCantid
'''         VALMOV = nCantid * nCosPro
'''         VALMOVUS = nCantid * nCosProUS 'valorizacion en dolares
'''         nSal = nSal + nCantid
'''      End If
'''
'''      If Adodc1("CATIPMOV") = "I" Then
'''         If nSaldo <> 0 Then
'''            nCosPro = (VALMOV + VALANTE) / nSaldo
'''            nCosProUS = (VALMOVUS + VALANTEUS) / nSaldo
'''         End If
'''      End If
'''
'''      VALANTE = nCosPro * nSaldo
'''      VALANTEUS = nCosProUS * nSaldo
'''      dfecha = Adodc1("CAFECDOC")
'''
'''      Adodc1.MoveNext
'''   End If
'''
'''
'''Loop
'''
'''
'''     cMesPro = Format(Year(dfecha), "0000") & Format(Month(dfecha), "00")
'''
''''*************************************************
'''     If cMesPro = arMes Then
'''        Rsql1 = "INSERT INTO  MoResMes (SMALMA,SMCODIGO,SMMESPRO,SMSALDOINI,SMCANENT,SMCANSAL,SMMNPREUNI,SMUSPREUNI)VALUES('" & almacen & "','" & vCodArt & "','" & cMesPro & "'," & nSaldoInicial & "," & nIng & "," & nSal & "," & nCosPro & "," & nCosProUS & ")"
'''
'''        cConexCom.BeginTrans
'''        cConexCom.Execute Rsql1
'''        cConexCom.CommitTrans
'''     End If
''' '*************************************************
'''      nSaldoInicial = nSaldoInicial + (nIng - nSal)
'''      cMesActu = AnioMesSiguiente(Format(Year(Now), "0000") & Format(Month(Now), "00"))
'''      nIng = 0
'''      nSal = 0
'''      cMesAnte = AnioMesSiguiente(cMesPro)
''''
'''      While cMesAnte <> cMesActu
'''            If cMesPro = arMes Then
'''               Rsql1 = "INSERT INTO  MoResMes (SMALMA,SMCODIGO,SMMESPRO,SMSALDOINI,SMCANENT,SMCANSAL,SMMNPREUNI,SMUSPREUNI)VALUES('" & almacen & "','" & vCodArt & "','" & cMesAnte & "'," & nSaldoInicial & "," & nIng & "," & nSal & "," & nCosPro & "," & nCosProUS & ")"
'''               cConexCom.BeginTrans
'''               cConexCom.Execute Rsql1
'''               cConexCom.CommitTrans
'''            End If
'''            cMesAnte = AnioMesSiguiente(cMesAnte)
'''      Wend
'''
'''      cConexCom.Execute "Update STKART SET STSKDIS=" & nSaldoInicial + (nIng - nSal) & ",STKPREPRO=" & nCosPro & " WHERE STALMA='" & almacen & "' AND STCODIGO='" & vCodArt & "'"
'''
'''Exit Sub
'''
'''ERRAR:
'''MsgBox Err.Description
'''BarraProc.Visible = False
'''cArticulo.Visible = False
'''Resume
'''End Sub
'''
'''
'''Sub Cmd_RestoreSaldos_Click()
'''Dim cAnoMes As String, cCod As String
'''Dim cSql1 As String, CSQL2 As String
'''Dim nSaldo As Double, nCosPro As Double
'''Dim nPrecio As Double, nCantid As Double
'''Dim nCount, nMaxRec As Integer
'''Dim csql As String
'''Dim cSql22 As String
'''On Error GoTo ErrCarga
'''Dim MaeartRs As New ADODB.Recordset
'''Dim cMesActu, cMesCirr As String
'''Set Adodc1 = New ADODB.Recordset
'''
'''cAnoMes = Format(dFecVal.Year, "0000") & Format(dFecVal.Month, "00")
'''cMesCirr = UltimoCierre
'''
'''If cMesCirr <> "" Then
'''   If cAnoMes <= cMesCirr Then
'''     MsgBox "El Mes Que Usted Selecciono ya Esta Cerrado", vbInformation, "Verifique...!"
'''     Exit Sub
'''   End If
'''
''''   If cAnoMes > AnioMesSiguiente(cMesCirr) Then
''''     MsgBox "El Mes Que Usted Selecciono No Pueder Ser Recalculado" & Chr(10) & "Por Favor Seleccione el Mes Anterior", vbInformation, "Verifique...!"
''''     Exit Sub
''''   End If
'''End If
'''
'''cArticulo.Caption = "Espere Un Momento....! "
'''Frame1.Refresh
'''
'''If (Text1 = "" Or Text2 = "") And opt1.Value = False Then
'''   MsgBox "Debe Indicar un Rango de Articulos...", vbInformation, "Verifique....!"
'''   Exit Sub
'''End If
'''
'''
'''cSql1 = "Select * From MovAlmDet A Inner Join MovAlmCab B On B.CAALMA = A.DEALMA And B.CATD = A.DETD  And B.CANUMDOC = A.DENUMDOC "
'''cSql1 = cSql1 & " Where CAALMA = '" & almacen & "' and not (CATD='GS' And CACODMOV='GF' And CASITGUI='F') "
'''
'''If opt1.Value = True Then
'''   cSql1 = cSql1 & " And CASITGUI<>'A'  and DECODIGO<>'TEXTO' Order By DECODIGO,CAFECDOC,CAHORA"
'''Else
'''   cSql1 = cSql1 & " And CASITGUI<>'A'  and DECODIGO<>'TEXTO' and  (decodigo>='" & Text1 & "' and decodigo<='" & Text2 & "')  Order By DECODIGO,CAFECDOC,CAHORA"
'''End If
'''
'''Set Adodc1 = New ADODB.Recordset
'''Adodc1.Open cSql1, cConexCom, adOpenForwardOnly
'''
'''If Adodc1.EOF Then
'''   MsgBox "No Existe Informnación Registrada en el la Fecha que Usted Indico", vbInformation, "Verifique....!"
'''   Exit Sub
'''End If
'''
'''Label2(0).Visible = True
'''Label2(1).Visible = True
'''BarraProc.Visible = True
'''cArticulo.Visible = True
'''
'''If opt1.Value = True Then
'''   csql = "delete From MORESMES Where SMALMA='" & almacen & "' and SMMESPRO>='" & cAnoMes & "'"
'''Else
'''   csql = "delete From MORESMES Where SMALMA='" & almacen & "' and SMMESPRO>='" & cAnoMes & "' AND SMCODIGO>='" & Text1 & "' AND SMCODIGO<='" & Text2 & "'"
'''End If
'''cConexCom.Execute csql
'''
''''*******************************************************
''''**********'iNICIALIZA A 0 todos los articulos de Stkart (stock de Articulos )
'''If opt1.Value = True Then
'''   csql = "UPDATE STKART SET STSKDIS=0 WHERE STALMA='" & almacen & "'"
'''Else
'''   csql = "UPDATE STKART SET STSKDIS=0 WHERE STALMA='" & almacen & "' and stcodigo>='" & Text1 & "' and stcodigo<='" & Text2 & "'"
'''End If
'''
'''   cConexCom.Execute csql
'''
'''
'''BarraProc.Min = 50
'''Set Adodc2 = New ADODB.Recordset
'''Adodc2.Open "Select * from MORESMES where SMMESPRO ='" & AnioMesAnterior(cAnoMes) & "'", cConexCom, adOpenStatic
'''
'''If opt1.Value = True Then
'''   cSql22 = "Select distinct (Decodigo) From MovAlmDet A Inner Join MovAlmCab B On B.CAALMA = A.DEALMA And B.CATD = A.DETD  And B.CANUMDOC = A.DENUMDOC "
'''   cSql22 = cSql22 & " Where CAALMA = '" & almacen & "' and not (CATD='GS' And CACODMOV='GF' And CASITGUI='F') "
'''   cSql22 = cSql22 & " And CASITGUI<>'A' "
'''Else
'''   cSql22 = "Select distinct (Decodigo) From MovAlmDet A Inner Join MovAlmCab B On B.CAALMA = A.DEALMA And B.CATD = A.DETD  And B.CANUMDOC = A.DENUMDOC "
'''   cSql22 = cSql22 & " Where CAALMA = '" & almacen & "' and not (CATD='GS' And CACODMOV='GF' And CASITGUI='F') "
'''   cSql22 = cSql22 & " And CASITGUI<>'A' and ( Decodigo>='" & Text1 & "' and Decodigo<='" & Text2 & "' ) ORDER BY DECODIGO"
'''End If
'''BarraProc.Min = 10
'''
'''MaeartRs.Open cSql22, cConexCom, adOpenStatic
'''nCount = 0
'''nMaxRec = MaeartRs.RecordCount
'''BarraProc.Max = 100 + nMaxRec
'''BarraProc.Min = 0
'''Frame1.Refresh
'''While Not MaeartRs.EOF
'''    nCount = nCount + 1
'''    BarraProc.Value = nCount
'''    cArticulo.Caption = "Recalculando Saldos : " & Format(nCount, "00000") & "     -     " & Format(nMaxRec, "00000") & " " & Chr(10) & (MaeartRs!DECODIGO)
'''    Me.Refresh
'''    Call ValorizaXArticulo(MaeartRs!DECODIGO, cAnoMes)
'''    MaeartRs.MoveNext
'''Wend
'''
'''    BarraProc.Visible = False
'''    cArticulo.Visible = False
'''    Label2(0).Visible = False
'''    Label2(1).Visible = False
'''
'''MaeartRs.Close
'''Adodc1.Close
'''Adodc2.Close
'''
'''Exit Sub
'''ErrCarga:
'''    MsgBox Err.Description
'''    BarraProc.Visible = False
'''    cArticulo.Visible = False
'''    Label2(0).Visible = False
'''    Label2(1).Visible = False
'''    Resume
'''End Sub



Option Explicit
Dim PCount As Long
Dim cConexAux As ADODB.Connection
Dim Adodc1 As ADODB.Recordset
Dim Adodc2 As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim cRt As String
Dim almacen As String
Dim nTra As Integer

Private Sub cmd_exit_Click()
Unload Me
End Sub

Private Sub Form_Load()
 central Me
 Call Carga_Almacen
 'dFecVal.Value = Format(Now, "dd/mm/yyyy")
End Sub

Private Sub Combo1_Click()
rs.MoveFirst
rs.Move Combo1.ListIndex
almacen = Format(rs(0), "00")
End Sub
Private Sub Carga_Almacen()
Dim RSQL As String
Dim I As Integer
RSQL = "Select TAALMA,TADESCRI FROM TabAlm "
Set rs = New ADODB.Recordset
rs.Open RSQL, cConexCom, adOpenStatic
Do While Not rs.EOF
     Combo1.AddItem (rs(1))
     rs.MoveNext
     If rs.EOF Then Exit Do
Loop
rs.MoveFirst
For I = 0 To rs.RecordCount - 1
  If rs(0) = VGAlma Then
    Combo1.ListIndex = I
    Exit For
  Else
    rs.MoveNext
  End If
Next
End Sub

Sub Cmd_RestoreSaldos_Click()
Dim cSql22 As String
On Error GoTo ErrCarga
Dim MaeartRs As New ADODB.Recordset
Set Adodc1 = New ADODB.Recordset
Dim nCount, nMaxRec As Long
'cAnoMes = Format(dFecVal.Year, "0000") & Format(dFecVal.Month, "00")
'cMesCirr = UltimoCierre


'Label2(0).Visible = True
'Label2(1).Visible = True
BarraProc.Visible = True
cArticulo.Visible = True

cSql22 = "Select afserie,aflote,acodigo from maeart "
MaeartRs.Open cSql22, cConexCom, adOpenStatic
nCount = 0
nMaxRec = MaeartRs.RecordCount
BarraProc.Max = 100 + nMaxRec
BarraProc.Min = 0
Frame1.Refresh
While Not MaeartRs.EOF
    nCount = nCount + 1
    BarraProc.Value = nCount
    cArticulo.Caption = "Recalculando Stock : " & Format(nCount, "00000") & "     -     " & Format(nMaxRec, "00000") & " " & Chr(10) & (MaeartRs!ACODIGO)
    Me.Refresh
    'RMM****************************************************
    If opt1.Value = True And MaeartRs!afserie = "S" Then
       ClsTock.CalculaStockSerie almacen, (MaeartRs!ACODIGO)
       ClsTock.CalculaSaldoNoValorizado VGAlma, MaeartRs!ACODIGO, Now
    Else
       If opt1.Value = False And MaeartRs!afLOTE = "S" Then
          ClsTock.CalculaStockLOTE almacen, (MaeartRs!ACODIGO)
          ClsTock.CalculaSaldoNoValorizado VGAlma, MaeartRs!ACODIGO, Now
       End If
    End If
    '*******************************************************
    MaeartRs.MoveNext
Wend

    BarraProc.Visible = False
    cArticulo.Visible = False
'    Label2(0).Visible = False
'    Label2(1).Visible = False

MaeartRs.Close

Exit Sub
ErrCarga:
    MsgBox Err.Description
    BarraProc.Visible = False
    cArticulo.Visible = False
'    Label2(0).Visible = False
'    Label2(1).Visible = False
    
End Sub
