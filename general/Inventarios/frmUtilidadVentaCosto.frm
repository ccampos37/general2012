VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUtilidadVentaCosto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Utilidades Ventas /  Costos"
   ClientHeight    =   3885
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   5520
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3720
      Left            =   72
      TabIndex        =   0
      Top             =   108
      Width           =   5376
      Begin VB.CommandButton cmdActualiza 
         Caption         =   "&Actualizar"
         Height          =   600
         Left            =   1728
         Picture         =   "frmUtilidadVentaCosto.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2952
         Width           =   900
      End
      Begin VB.CommandButton CmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   624
         Left            =   4140
         Picture         =   "frmUtilidadVentaCosto.frx":025B
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2952
         Width           =   675
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   600
         Left            =   576
         Picture         =   "frmUtilidadVentaCosto.frx":069D
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2952
         Width           =   960
      End
      Begin VB.ComboBox Combo1 
         Height          =   288
         ItemData        =   "frmUtilidadVentaCosto.frx":0ADF
         Left            =   1776
         List            =   "frmUtilidadVentaCosto.frx":0AE1
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   540
         Width           =   3150
      End
      Begin MSComCtl2.DTPicker dFech_ini 
         Height          =   336
         Left            =   2628
         TabIndex        =   1
         Top             =   1224
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   582
         _Version        =   393216
         Format          =   53018625
         CurrentDate     =   37085
      End
      Begin MSComCtl2.DTPicker dFech_fin 
         Height          =   300
         Left            =   2628
         TabIndex        =   2
         Top             =   1800
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   529
         _Version        =   393216
         Format          =   53018625
         CurrentDate     =   37085
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin MSComctlLib.ProgressBar BarraProc 
         Height          =   216
         Left            =   612
         TabIndex        =   9
         Top             =   2628
         Visible         =   0   'False
         Width           =   4212
         _ExtentX        =   7435
         _ExtentY        =   370
         _Version        =   393216
         Appearance      =   1
         Min             =   10
         Max             =   1000
      End
      Begin VB.Label cArticulo 
         Caption         =   "Label1"
         Height          =   252
         Left            =   612
         TabIndex        =   10
         Top             =   2304
         Visible         =   0   'False
         Width           =   4212
      End
      Begin VB.Label Label6 
         Caption         =   "Almacen"
         ForeColor       =   &H80000008&
         Height          =   372
         Left            =   720
         TabIndex        =   6
         Top             =   600
         Width           =   732
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Final"
         Height          =   300
         Left            =   396
         TabIndex        =   4
         Top             =   1800
         Width           =   1812
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Inicial"
         Height          =   300
         Left            =   396
         TabIndex        =   3
         Top             =   1260
         Width           =   1812
      End
   End
End
Attribute VB_Name = "frmUtilidadVentaCosto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cConexAux As ADODB.Connection
Dim adodc1 As ADODB.Recordset
Dim Adodc2 As ADODB.Recordset
Dim Adodc3 As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim cRt As String
Dim almacen As String
Dim almacenAnt  As String
Dim nTra As Integer
Dim nMes, nAnno As Long
Dim Adodc4 As ADODB.Recordset
Dim Adostk As ADODB.Recordset
Dim AdoMes As ADODB.Recordset
Dim rsPrec As New ADODB.Recordset
Dim Total As Double
Dim tipo As String

Private Sub CmdAceptar_Click()
Dim rsql As String
Dim Va1 As String, Va2 As String
Screen.MousePointer = 11
On Error GoTo Err


     CrystalReport1.WindowTitle = "Inv504 -- Control de Inventarios"
     CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv504.rpt"
  
  Dim ccadena As String
  If CrystalReport1.ReportFileName = "" Then
      MsgBox "No hay registros a imprimir", vbInformation, "Aviso"
      Screen.MousePointer = 1
      Exit Sub
  End If
  ccadena = "({COSPROFECH.AUXALMA}='" & almacen & "') and {COSPROFECH.AUXCANT}>0.00001 AND ({COSPROFECH.AUXFECDOC} IN DATE (" & Format(dFech_ini, "yyyy") & "," & Format(dFech_ini, "mm") & "," & Format(dFech_ini, "dd") & ") " & _
      " To DATE (" & Format(dFech_fin, "yyyy") & "," & Format(dFech_fin, "mm") & "," & Format(dFech_fin, "dd") & ")) "

  Ubi_Tab CrystalReport1
  CrystalReport1.DiscardSavedData = True
  CrystalReport1.Destination = crptToWindow
  CrystalReport1.WindowShowPrintBtn = True
  CrystalReport1.WindowShowRefreshBtn = True
  CrystalReport1.WindowShowSearchBtn = True
  CrystalReport1.WindowShowPrintSetupBtn = True
  CrystalReport1.ReplaceSelectionFormula (ccadena)
  CrystalReport1.formulas(0) = "ALMACEN = '" & UCase(Combo1.text) & "'"
  CrystalReport1.formulas(2) = "EMPRESA= '" & UCase(VGparametros.RucEmpresa) & "'"
  CrystalReport1.formulas(3) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
  CrystalReport1.formulas(4) = "MONEY= 'SOLES'"
  CrystalReport1.formulas(5) = "xfechIni= '" & dFech_ini.Value & "'"
  CrystalReport1.formulas(6) = "xfechFin= '" & dFech_fin.Value & "'"
  If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
 Screen.MousePointer = 1
 Exit Sub
Err:
   MsgBox Err.Description, vbInformation, "Aviso"
   Screen.MousePointer = 1
   Exit Sub
End Sub

Private Sub cmdActualiza_Click()
Dim rsql As String
Dim Va1 As String, Va2 As String
Screen.MousePointer = 11
On Error GoTo Err
Carga_RepoVal

   Screen.MousePointer = 1
 Exit Sub
Err:
   MsgBox Err.Description, vbInformation, "Aviso"
   Screen.MousePointer = 1
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Combo1_Click()
rs.MoveFirst
rs.Move Combo1.ListIndex
almacen = Format(rs(0), "00")
End Sub

Private Sub Form_Load()
central Me
Carga_Almacen
dFech_fin = Date
dFech_ini = DateAdd("m", -1, Date)
End Sub
Private Sub Carga_Almacen()
Dim rsql As String
Dim I As Integer
rsql = "Select TAALMA,TADESCRI FROM TabAlm "
Set rs = New ADODB.Recordset
rs.Open rsql, VGCNx, adOpenStatic
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

Private Sub Carga_RepoVal()
Dim rAdo As ADODB.Recordset
Dim cAnoMes As String
Dim cSql1 As String, CSQL2 As String
Dim nSaldo As Double, nCosPro As Double
Dim nPrecio As Double, nCantid As Double
Dim I, nCount As Integer
Dim rsql As String
Dim csql As String
Dim cSql22 As String
Dim Rs2 As New ADODB.Recordset
Dim saldo As Double
On Error GoTo ErrCarga
Set adodc1 = New ADODB.Recordset
Set rAdo = New ADODB.Recordset
Dim MaeartRs As New ADODB.Recordset

'*******************************************************
BarraProc.Visible = True
cArticulo.Visible = True
cArticulo.Caption = "Espere Un Momento....! "
Frame1.Refresh

'*******************************************************
rsql = "SELECT TOP 1 CAFECDOC From MovAlmCab  WHERE CACIERRE = false AND CAALMA = '" & VGAlma & "'  ORDER BY CAFECDOC asc"
Set Rs2 = New ADODB.Recordset
Rs2.Open rsql, VGCNx, adOpenStatic
If Not Rs2.EOF Then
    Rs2.MoveFirst
    nMes = Month(Rs2(0))
    nAnno = Year(Rs2(0))
    If nMes = 13 Then
         nMes = 1
    End If
End If
'*******************************************************
'**Carga todos los Movimientos al respectivo Recordset
'*******************************************************
Set adodc1 = New ADODB.Recordset
cSql1 = "Select * From MovAlmDet A Inner Join MovAlmCab B On B.CAALMA = A.DEALMA And B.CATD = A.DETD  And B.CANUMDOC = A.DENUMDOC "
cSql1 = cSql1 & " Where CAALMA = '" & almacen & "' and not (CATD='GS' And CACODMOV='GF' And CASITGUI='F') "
cSql1 = cSql1 & " And CASITGUI<>'A'  AND FORMAT( YEAR(CAFECDOC),'0000' )+ FORMAT( MONTH(CAFECDOC),'00' ) >='" & Format(nAnno, "0000") + Format(nMes, "00") & "' Order By DECODIGO,CAFECDOC,CAHORA"
adodc1.Open cSql1, VGCNx, adOpenForwardOnly
'*******************************************************
'**Carga todos Codigos de los productos al respectivo Recordset
'*******************************************************
BarraProc.Min = 50
cSql22 = "SELECT STKART.STALMA, STKART.STCODIGO, STKART.STSKDIS FROM STKART WHERE STALMA='" & almacen & "'"
BarraProc.Min = 10
MaeartRs.Open cSql22, VGCNx, adOpenStatic

nCount = 0
BarraProc.Max = 100 + MaeartRs.RecordCount


csql = "DELETE From COSPROFECH WHERE  AUXALMA='" & almacen & "' AND FORMAT( YEAR(AUXFECDOC),'0000' )+ FORMAT( MONTH(AUXFECDOC),'00' ) >='" & Format(nAnno, "0000") + Format(nMes, "00") & "'"
VGCNx.Execute csql

'******************************************************
BarraProc.Min = 50
Set Adodc2 = New ADODB.Recordset
Adodc2.Open "Select * from MORESMES where SMMESPRO >='" & AnioMesAnterior(Format(nAnno, "0000") + Format(nMes, "00")) & "' AND SMALMA='" & almacen & "'", VGCNx, adOpenStatic
'******************************************************

'*******************************************************

BarraProc.Min = 0
While Not MaeartRs.EOF
    nCount = nCount + 1
    BarraProc.Value = nCount
    cArticulo.Caption = "Actualizando Datos : " & nCount & "   " & (MaeartRs!stcodigo)
    Frame1.Refresh
    Call ValorizaXArticulo(MaeartRs!stcodigo, Format(nAnno, "0000") + Format(nMes, "00")) 'Procedimiento que Valoriza por Articulo
    MaeartRs.MoveNext
Wend

    BarraProc.Visible = False
    cArticulo.Visible = False

Exit Sub
ErrCarga:
        MsgBox Err.Description
        If nTra = 1 Then cConexAux.RollbackTrans
        BarraProc.Visible = False
        cArticulo.Visible = False
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub ADOConectar()
cRt = App.Path & "\BdAuxCom.Mdb"
Set cConexAux = New ADODB.Connection
cConexAux.CursorLocation = adUseClient
cConexAux.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & cRt & ";"
cConexAux.Open
End Sub
'****************************************************
'**
'****************************************************
Private Sub ValorizaXArticulo(ByVal vCodArt As String, ByVal cArmes As String)
Dim TCamb As Double
Dim Li As Integer
Dim nCambio, nSaldo As Double, nCosPro, nCosProUS As Double
Dim nPrecio, nPrecioUS, xPrecio As Double, nCantid As Double
Dim cMesPro, cMesActu, cMesAnte, cAnoMes As String
Dim Rsql1 As String
Dim nTipCam, cSql1 As String
'**********Roberto
Dim VALMOV, VALANTE, VALMOVUS, VALANTEUS As Double
Dim nSal, nIng, nSaldoInicial As Double
Dim dfecha As Date
Dim csql As String
Dim XNUMDOC As String
On Local Error GoTo ERRAR

adodc1.Filter = " Decodigo='" & vCodArt & "'"
xPrecio = 0
nPrecio = 0: nCantid = 0
nPrecioUS = 0
nSal = 0: nIng = 0
nSaldoInicial = 0
If Not adodc1.EOF Then adodc1.MoveFirst
nCosProUS = 0: nCosPro = 0

If Not adodc1.EOF Then
   nMes = Month(adodc1("CAFECDOC"))
   nAnno = Year(adodc1("CAFECDOC"))
   dfecha = adodc1("CAFECDOC")
Else
   dfecha = Format(Format(nAnno, "0000") + Format(nMes, "00"), "dd/mm/yyyy")
End If

cAnoMes = cArmes

Adodc2.Filter = "SMCODIGO = '" & vCodArt & "' AND SMMESPRO ='" & AnioMesAnterior(cAnoMes) & "'"
If Not Adodc2.EOF Then
   nSaldo = IIf(IsNull(Adodc2!SMSALDOINI), 0, Adodc2!SMSALDOINI) + (Adodc2!SMCANENT - Adodc2!SMCANSAL)
   nCosPro = Adodc2("SMMNPREUNI")
   nCosProUS = Adodc2("SMUSPREUNI")
   VALANTE = nCosPro * nSaldo
   VALANTEUS = nCosProUS * nSaldo
Else
   nSaldo = 0: nCosPro = 0: nCosProUS = 0
   VALANTE = 0: VALANTEUS = 0
End If


Do While Not adodc1.EOF
    
   If Year(adodc1("CAFECDOC")) <> nAnno Or Month(adodc1("CAFECDOC")) <> nMes Then
  
      cMesPro = Format(nAnno, "0000") & Format(nMes, "00")
      'Rsql1 = "INSERT INTO  MoResMes (SMALMA,SMCODIGO,SMMESPRO,SMSALDOINI,SMCANENT,SMCANSAL,SMMNPREUNI,SMUSPREUNI)VALUES('" & almacen & "','" & vCodArt & "','" & cMesPro & "'," & nSaldoInicial & "," & nIng & "," & nSal & "," & nCosPro & "," & nCosProUS & ")"
      
      'Vgcnx.BeginTrans
      'Vgcnx.Execute Rsql1
      'Vgcnx.CommitTrans
      
      cMesActu = (Format(Year(adodc1("CAFECDOC")), "0000") & Format(Month(adodc1("CAFECDOC")), "00"))
      nSaldoInicial = nSaldoInicial + (nIng - nSal)
      nIng = 0
      nSal = 0
      cMesAnte = AnioMesSiguiente(cMesPro)
      While cMesAnte <> cMesActu
      '      Rsql1 = "INSERT INTO  MoResMes (SMALMA,SMCODIGO,SMMESPRO,SMSALDOINI,SMCANENT,SMCANSAL,SMMNPREUNI,SMUSPREUNI)VALUES('" & almacen & "','" & vCodArt & "','" & cMesAnte & "'," & nSaldoInicial & "," & nIng & "," & nSal & "," & nCosPro & "," & nCosProUS & ")"
      '      Vgcnx.BeginTrans
      '      Vgcnx.Execute Rsql1
      '      Vgcnx.CommitTrans
            cMesAnte = AnioMesSiguiente(cMesAnte)
      Wend
      '*************************************************
      dfecha = adodc1("CAFECDOC")
      nMes = Month(adodc1("CAFECDOC"))
      nAnno = Year(adodc1("CAFECDOC"))
            
  Else
  
     '*************************************************
      If adodc1!DETIPCAM = 0 Or adodc1!DETIPCAM = 1 Then
            If UCase(Dir$(cRuta4)) = UCase(VGNameCont & ".MDB") Then
               TCamb = Val(Devolver_Dato(3, Format(dfecha, "DD/MM/YYYY"), "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "02", "TIPOMON_CODIGO"))
            Else
               If UCase(Dir$(cRuta2)) = "BDCOMUN.MDB" Then
                  TCamb = Val(Devolver_Dato(1, Format(dfecha, "DD/MM/YYYY"), "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "02", "TIPOMON_CODIGO"))
               Else
                  TCamb = adodc1!DETIPCAM
               End If
            End If
      Else
          TCamb = adodc1!DETIPCAM
      End If
      '*************************************************
      
      nCantid = adodc1("DECANTID")
      nPrecio = adodc1("DEPRECIO")
      
      '*************************************************
      '***DOCUMENTOS EN  DOLARES
      If cNull(adodc1!CACODMON) = "02" Then
         nPrecio = adodc1("DEPRECIO") * TCamb
      Else
         nPrecio = adodc1("DEPRECIO")
      End If
      '*************************************************
      '*************************************************
      '***DOCUMENTOS EN SOLES
      If cNull(adodc1!CACODMON) = "02" Then
         nPrecioUS = adodc1("DEPRECIO")
      Else
         If Round(TCamb, 3) <> 0 Then
            nPrecioUS = adodc1("DEPRECIO") / TCamb
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
            nCosPro = (VALMOV + VALANTE) / nSaldo
            nCosProUS = (VALMOVUS + VALANTEUS) / nSaldo
         End If
      End If
     
      VALANTE = nCosPro * nSaldo
      VALANTEUS = nCosProUS * nSaldo
      dfecha = adodc1("CAFECDOC")
     
      If (adodc1!CATD = "FT" Or adodc1!CATD = "BV") Then
         
         If cNull(adodc1!CACODMON) = "02" Then
            xPrecio = PrecioFact(adodc1!CATD, adodc1!CANUMDOC, adodc1!decodigo) * TCamb
         Else
            xPrecio = PrecioFact(adodc1!CATD, adodc1!CANUMDOC, adodc1!decodigo)
         End If
         
         csql = "INSERT INTO COSPROFECH (AUXALMA, AUXTD, AUXNUMDOC, AUXCODART, AUXFECDOC,AUXCANT,AUXPRECIO, AUXPRECOS)VALUES ('" & adodc1!CAALMA & "','" & adodc1!CATD & "','" & adodc1!CANUMDOC & "','" & adodc1!decodigo & "','" & adodc1!CAFECDOC & "'," & adodc1("DECANTID") & "," & xPrecio & "," & nCosPro & ")"
         VGCNx.BeginTrans
         VGCNx.Execute csql
         VGCNx.CommitTrans
      End If
      
      If (adodc1!CATD = "GS") And (adodc1!CATIPGUI = "GV" And adodc1!Casitgui = "F") Then
         If adodc1!CARFTDOC <> "" And adodc1!CARFNDOC <> "" Then
            
            If cNull(adodc1!CACODMON) = "02" Then
               xPrecio = PrecioFact(adodc1!CARFTDOC, adodc1!CARFNDOC, adodc1!decodigo) * TCamb
            Else
               xPrecio = PrecioFact(adodc1!CARFTDOC, adodc1!CARFNDOC, adodc1!decodigo)
            End If
            
            If xPrecio <> -1 Then
               csql = "INSERT INTO COSPROFECH (AUXALMA, AUXTD, AUXNUMDOC, AUXCODART, AUXFECDOC,AUXCANT,AUXPRECIO, AUXPRECOS)VALUES ('" & adodc1!CAALMA & "','" & adodc1!CARFTDOC & "','" & adodc1!CARFTDOC & "','" & adodc1!decodigo & "','" & adodc1!CAFECDOC & "'," & adodc1("DECANTID") & "," & xPrecio & "," & nCosPro & ")"
               VGCNx.BeginTrans
               VGCNx.Execute csql
               VGCNx.CommitTrans
            End If
         End If
      End If

      adodc1.MoveNext
   End If
   
   
Loop


     'cMesPro = Format(Year(dfecha), "0000") & Format(Month(dfecha), "00")

'*************************************************
    ' Rsql1 = "INSERT INTO  MoResMes (SMALMA,SMCODIGO,SMMESPRO,SMSALDOINI,SMCANENT,SMCANSAL,SMMNPREUNI,SMUSPREUNI)VALUES('" & almacen & "','" & vCodArt & "','" & cMesPro & "'," & nSaldoInicial & "," & nIng & "," & nSal & "," & nCosPro & "," & nCosProUS & ")"

     'Vgcnx.BeginTrans
     'Vgcnx.Execute Rsql1
     'Vgcnx.CommitTrans
 '*************************************************
'      nSaldoInicial = nSaldoInicial + (nIng - nSal)
'      cMesActu = AnioMesSiguiente(Format(Year(Now), "0000") & Format(Month(Now), "00"))
'      nIng = 0
'      nSal = 0
'      cMesAnte = AnioMesSiguiente(cMesPro)
''      If nSaldoInicial <> 0 Then
''         MsgBox " no se movio"
''      End If
'      While cMesAnte <> cMesActu
'      '      Rsql1 = "INSERT INTO  MoResMes (SMALMA,SMCODIGO,SMMESPRO,SMSALDOINI,SMCANENT,SMCANSAL,SMMNPREUNI,SMUSPREUNI)VALUES('" & almacen & "','" & vCodArt & "','" & cMesAnte & "'," & nSaldoInicial & "," & nIng & "," & nSal & "," & nCosPro & "," & nCosProUS & ")"
'      '      Vgcnx.BeginTrans
'      '      Vgcnx.Execute Rsql1
'      '      Vgcnx.CommitTrans
'            cMesAnte = AnioMesSiguiente(cMesAnte)
'      Wend

      'Vgcnx.Execute "Update STKART SET STSKDIS=" & nSaldoInicial & ",STKPREPRO=" & nCosPro & " WHERE STALMA='" & almacen & "' AND STCODIGO='" & vCodArt & "'"

Exit Sub

ERRAR:
MsgBox Err.Description
BarraProc.Visible = False
cArticulo.Visible = False

End Sub

Function PrecioFact(ByVal arTd As String, ByVal arNumdoc As String, ByVal arCodi As String) As Double
         Set rsPrec = New ADODB.Recordset
         rsPrec.Open "Select DFPREC_ORI from facdet where dftd='" & arTd & "' and dfnumser+dfnumdoc='" & arNumdoc & "' and dfcodigo='" & arCodi & "'", VGCNx, adOpenForwardOnly, adLockReadOnly
         PrecioFact = IIf(Not rsPrec.EOF, rsPrec!DFPREC_ORI, -1)
         rsPrec.Close
End Function
  

