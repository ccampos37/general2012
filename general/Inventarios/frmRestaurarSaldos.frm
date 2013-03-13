VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRestaurarSaldos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Restaurar Saldos"
   ClientHeight    =   4695
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   5235
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_exit 
      Caption         =   "&Salir"
      Height          =   645
      Left            =   2652
      Picture         =   "frmRestaurarSaldos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4032
      Width           =   735
   End
   Begin VB.CommandButton Cmd_RestoreSaldos 
      Caption         =   "&Aceptar"
      Height          =   645
      Left            =   1872
      Picture         =   "frmRestaurarSaldos.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4032
      Width           =   705
   End
   Begin VB.Frame Frame3 
      Height          =   1524
      Left            =   72
      TabIndex        =   8
      Top             =   1728
      Width           =   2496
      Begin VB.OptionButton opt2 
         Caption         =   "Por Articulo"
         Height          =   228
         Left            =   432
         TabIndex        =   10
         Top             =   936
         Width           =   1416
      End
      Begin VB.OptionButton opt1 
         Caption         =   "Todos"
         Height          =   228
         Left            =   432
         TabIndex        =   9
         Top             =   396
         Value           =   -1  'True
         Width           =   1704
      End
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   1488
      Left            =   2664
      TabIndex        =   3
      Top             =   1728
      Width           =   2496
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   888
         MaxLength       =   20
         TabIndex        =   5
         Top             =   360
         Width           =   1275
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   888
         TabIndex        =   4
         Top             =   912
         Width           =   1275
      End
      Begin VB.Label Label2 
         Caption         =   "Inicio "
         Height          =   252
         Index           =   1
         Left            =   252
         TabIndex        =   7
         Top             =   396
         Width           =   732
      End
      Begin VB.Label Label3 
         Caption         =   "Fin"
         Height          =   252
         Left            =   252
         TabIndex        =   6
         Top             =   936
         Width           =   732
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1668
      Left            =   72
      TabIndex        =   0
      Top             =   36
      Width           =   5088
      Begin VB.ComboBox Combo1 
         Height          =   288
         ItemData        =   "frmRestaurarSaldos.frx":0884
         Left            =   1488
         List            =   "frmRestaurarSaldos.frx":0886
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   396
         Width           =   3150
      End
      Begin MSComCtl2.DTPicker dFecVal 
         Height          =   312
         Left            =   1476
         TabIndex        =   11
         Top             =   1008
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MMMM 'del  ' yyyy"
         Format          =   113639427
         CurrentDate     =   37068
      End
      Begin VB.Label Label1 
         Caption         =   "Mes :"
         Height          =   288
         Left            =   396
         TabIndex        =   12
         Top             =   1068
         Width           =   1068
      End
      Begin VB.Label Label6 
         Caption         =   "Almacen"
         ForeColor       =   &H80000008&
         Height          =   264
         Left            =   396
         TabIndex        =   2
         Top             =   456
         Width           =   732
      End
   End
   Begin MSComctlLib.ProgressBar BarraProc 
      Height          =   192
      Left            =   984
      TabIndex        =   13
      Top             =   3768
      Visible         =   0   'False
      Width           =   3684
      _ExtentX        =   6482
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Min             =   10
      Max             =   1000
   End
   Begin VB.Label cArticulo 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1008
      TabIndex        =   17
      Top             =   3312
      Visible         =   0   'False
      Width           =   3708
   End
   Begin VB.Label Label2 
      Caption         =   "OK"
      Height          =   288
      Index           =   0
      Left            =   252
      TabIndex        =   16
      Top             =   3672
      Visible         =   0   'False
      Width           =   432
   End
End
Attribute VB_Name = "frmRestaurarSaldos"
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
'''    'Adodc2.Open "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD from MaeArt p, StkArt n Where n.STSKDIS<> 0 and  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & VGAlma & "'  ", Vgcnx, adOpenStatic, adLockOptimistic
'''    Adodc2.Open "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD from MaeArt p, StkArt n Where  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & VGAlma & "'  ", Vgcnx, adOpenStatic, adLockOptimistic
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
'''    'Adodc2.Open "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD from MaeArt p, StkArt n Where n.STSKDIS<> 0 and  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & VGAlma & "'  ", Vgcnx, adOpenStatic, adLockOptimistic
'''    Adodc2.Open "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD from MaeArt p, StkArt n Where   p.ACODIGO =  n.STCODIGO and n.STALMA = '" & VGAlma & "'  ", Vgcnx, adOpenStatic, adLockOptimistic
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
'''rS.Open rSql, Vgcnx, adOpenStatic
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
'''         Vgcnx.BeginTrans
'''         Vgcnx.Execute Rsql1
'''         Vgcnx.CommitTrans
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
'''               Vgcnx.BeginTrans
'''               Vgcnx.Execute Rsql1
'''               Vgcnx.CommitTran
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
'''               TCamb = Val(Devolver_Dato(3, Format(dfecha, "DD/MM/YYYY"), "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "02", "TIPOMON_CODIGO"))
'''            Else
'''               If UCase(Dir$(cRuta2)) = "BDCOMUN.MDB" Then
'''                  TCamb = Val(Devolver_Dato(1, Format(dfecha, "DD/MM/YYYY"), "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "02", "TIPOMON_CODIGO"))
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
'''      If cNull(Adodc1!CACODMON) = "02" Then
'''         nPrecio = Adodc1("DEPRECIO") * TCamb
'''      Else
'''         nPrecio = Adodc1("DEPRECIO")
'''      End If
'''      '*************************************************
'''      '*************************************************
'''      '***DOCUMENTOS EN SOLES
'''      If cNull(Adodc1!CACODMON) = "02" Then
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
'''        Vgcnx.BeginTrans
'''        Vgcnx.Execute Rsql1
'''        Vgcnx.CommitTrans
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
'''               Vgcnx.BeginTrans
'''               Vgcnx.Execute Rsql1
'''               Vgcnx.CommitTrans
'''            End If
'''            cMesAnte = AnioMesSiguiente(cMesAnte)
'''      Wend
'''
'''      Vgcnx.Execute "Update STKART SET STSKDIS=" & nSaldoInicial + (nIng - nSal) & ",STKPREPRO=" & nCosPro & " WHERE STALMA='" & almacen & "' AND STCODIGO='" & vCodArt & "'"
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
'''Adodc1.Open cSql1, Vgcnx, adOpenForwardOnly
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
'''Vgcnx.Execute csql
'''
''''*******************************************************
''''**********'iNICIALIZA A 0 todos los articulos de Stkart (stock de Articulos )
'''If opt1.Value = True Then
'''   csql = "UPDATE STKART SET STSKDIS=0 WHERE STALMA='" & almacen & "'"
'''Else
'''   csql = "UPDATE STKART SET STSKDIS=0 WHERE STALMA='" & almacen & "' and stcodigo>='" & Text1 & "' and stcodigo<='" & Text2 & "'"
'''End If
'''
'''   Vgcnx.Execute csql
'''
'''
'''BarraProc.Min = 50
'''Set Adodc2 = New ADODB.Recordset
'''Adodc2.Open "Select * from MORESMES where SMMESPRO ='" & AnioMesAnterior(cAnoMes) & "'", Vgcnx, adOpenStatic
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
'''MaeartRs.Open cSql22, Vgcnx, adOpenStatic
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
Dim adodc1 As ADODB.Recordset
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
 dFecVal.Value = Format(Now, "dd/mm/yyyy")
End Sub

Private Sub opt1_Click()
If Opt2.Value = True Then
    Frame2.Enabled = True
Else
    Frame2.Enabled = False
End If
End Sub

Private Sub opt2_Click()
If Opt2.Value = True Then
    Frame2.Enabled = True
Else
    Frame2.Enabled = False
End If
End Sub

Private Sub Text1_DblClick()
Dim Adodc2 As New ADODB.Recordset

    Adodc2.Open "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD from MaeArt p, StkArt n Where n.STSKDIS<> 0 and  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & VGalma & "'  ", VGCNx, adOpenStatic, adLockOptimistic
    'Adodc2.Open "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD from MaeArt p, StkArt n Where  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & VGAlma & "'  ", Vgcnx, adOpenStatic, adLockOptimistic
    frmReferencia.Conectar Adodc2, "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD from MaeArt p, StkArt n Where n.STSKDIS<> 0 and  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & VGalma & "'"
    frmReferencia.Conectar Adodc2, "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD from MaeArt p, StkArt n Where  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & VGalma & "'"
    frmReferencia.Label1.Caption = "Artículos"
    frmReferencia.Show vbModal
    Adodc2.Close
    If vGUtil(1) <> "" Then
            Text1 = (vGUtil(1))
    End If
    If Text1 <> "" And Text2 <> "" And Text1 > Text2 Then
            MsgBox "Ingrese un codigo menor al fin ", vbOKOnly, "Error"
            Exit Sub
   End If
   If Text1 <> "" Then
            Text2.Enabled = True
            Text2.SetFocus
   End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then Text2_DblClick
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then Text1_DblClick
End Sub


Private Sub Text2_DblClick()
Dim Adodc2 As New ADODB.Recordset
    Adodc2.Open "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD from MaeArt p, StkArt n Where n.STSKDIS<> 0 and  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & VGalma & "'  ", VGCNx, adOpenStatic, adLockOptimistic
    'Adodc2.Open "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD from MaeArt p, StkArt n Where   p.ACODIGO =  n.STCODIGO and n.STALMA = '" & VGAlma & "'  ", Vgcnx, adOpenStatic, adLockOptimistic
    frmReferencia.Conectar Adodc2, "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD from MaeArt p, StkArt n Where n.STSKDIS<> 0 and  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & VGalma & "'  "
    frmReferencia.Conectar Adodc2, "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD from MaeArt p, StkArt n Where  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & VGalma & "'  "
    frmReferencia.Label1.Caption = "Artículos"
    frmReferencia.Show vbModal
    Adodc2.Close
    If vGUtil(1) <> "" Then
            Text2 = (vGUtil(1))
    End If
    If Text1 <> "" And Text2 <> "" And Text1 > Text2 Then
            MsgBox "Ingrese un codigo menor al fin ", vbOKOnly, "Error"
            Exit Sub
   End If
   If Text2 <> "" Then
            Text2.Enabled = True
            Text2.SetFocus
   End If
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
rs.Open RSQL, VGCNx, adOpenStatic
Do While Not rs.EOF
     Combo1.AddItem (rs(1))
     rs.MoveNext
     If rs.EOF Then Exit Do
Loop
rs.MoveFirst
For I = 0 To rs.RecordCount - 1
  If rs(0) = VGalma Then
    Combo1.ListIndex = I
    Exit For
  Else
    rs.MoveNext
  End If
Next
End Sub

Sub Cmd_RestoreSaldos_Click()
Dim cAnoMes As String, cCod As String
Dim cSql22 As String
On Error GoTo ErrCarga
Dim MaeartRs As New ADODB.Recordset
Dim cMesActu, cMesCirr As String
Set adodc1 = New ADODB.Recordset
Dim nCount, nMaxRec As Long
cAnoMes = Format(dFecVal.Year, "0000") & Format(dFecVal.Month, "00")
cMesCirr = UltimoCierre

If cMesCirr <> "" Then
   If cAnoMes <= cMesCirr Then
     MsgBox "El Mes Que Usted Selecciono ya Esta Cerrado", vbInformation, "Verifique...!"
     Exit Sub
   End If
End If
cArticulo.Caption = "Espere Un Momento....! "
Frame1.Refresh

If (Text1 = "" Or Text2 = "") And Opt1.Value = False Then
   MsgBox "Debe Indicar un Rango de Articulos...", vbInformation, "Verifique....!"
   Exit Sub
End If


Label2(0).Visible = True
Label2(1).Visible = True
BarraProc.Visible = True
cArticulo.Visible = True

If Opt1.Value = True Then
   cSql22 = "SELECT STKART.STALMA, STKART.STCODIGO, STKART.STSKDIS FROM STKART WHERE STALMA='" & almacen & "' ORDER BY STCODIGO"
Else
   cSql22 = "SELECT STKART.STALMA, STKART.STCODIGO, STKART.STSKDIS FROM STKART WHERE STALMA='" & almacen & "' AND ( STcodigo>='" & Text1 & "' and STcodigo<='" & Text2 & "' ) ORDER BY STCODIGO"
End If
BarraProc.Min = 10

MaeartRs.Open cSql22, VGCNx, adOpenStatic
nCount = 0
nMaxRec = MaeartRs.RecordCount
BarraProc.Max = 100 + nMaxRec
BarraProc.Min = 0
Frame1.Refresh
While Not MaeartRs.EOF
    nCount = nCount + 1
    BarraProc.Value = nCount
    cArticulo.Caption = "Recalculando Saldos : " & Format(nCount, "00000") & "     -     " & Format(nMaxRec, "00000") & " " & Chr(10) & (MaeartRs!stcodigo)
    Me.Refresh
    'Call ValorizaXArticulo(MaeartRs!DECODIGO, cAnoMes)
    'RMM****************************************************
     clsmovimientos.CalculaSaldoNoValorizado almacen, (MaeartRs!stcodigo), dFecVal.Value
    '*******************************************************
    MaeartRs.MoveNext
Wend

'*********************************************
    clsmovimientos.BorrarServicios almacen, VGCNx
'*********************************************


    BarraProc.Visible = False
    cArticulo.Visible = False
    Label2(0).Visible = False
    Label2(1).Visible = False

MaeartRs.Close

Exit Sub
ErrCarga:
    MsgBox Err.Description
    BarraProc.Visible = False
    cArticulo.Visible = False
    Label2(0).Visible = False
    Label2(1).Visible = False
    
End Sub
