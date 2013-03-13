VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormKardexRevaloriza 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Revalorizar Saldos"
   ClientHeight    =   2520
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5208
   Icon            =   "FormKardexRevaloriza.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   5208
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2355
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   5025
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   600
         Left            =   1800
         Picture         =   "FormKardexRevaloriza.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1590
         Width           =   705
      End
      Begin VB.CommandButton CmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   585
         Left            =   2670
         Picture         =   "FormKardexRevaloriza.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1590
         Width           =   675
      End
      Begin MSComctlLib.ProgressBar BarraProc 
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   1110
         Visible         =   0   'False
         Width           =   4215
         _ExtentX        =   7430
         _ExtentY        =   445
         _Version        =   393216
         Appearance      =   1
         Min             =   10
         Max             =   1000
      End
      Begin VB.ComboBox Combo1 
         Height          =   288
         ItemData        =   "FormKardexRevaloriza.frx":0CC6
         Left            =   1530
         List            =   "FormKardexRevaloriza.frx":0CC8
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   3150
      End
      Begin VB.Label cArticulo 
         Caption         =   "Label1"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   780
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.Label Label6 
         Caption         =   "Almacen"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   300
         Width           =   735
      End
      Begin VB.Label lbltrans 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   3
         Top             =   1770
         Width           =   3225
      End
      Begin VB.Label lbltrans 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1020
         TabIndex        =   2
         Top             =   1890
         Width           =   3225
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   255
      Top             =   2040
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "FormKardexRevaloriza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cConexAux As ADODB.Connection
Dim Adodc1 As ADODB.Recordset
Dim Adodc2 As ADODB.Recordset
Dim Adodc3 As ADODB.Recordset
Dim rS As ADODB.Recordset
Dim cRt As String
Dim almacen As String
Dim almacenAnt  As String
Dim nTra As Integer
Dim Adodc4 As ADODB.Recordset
Dim Adostk As ADODB.Recordset
Dim AdoMes As ADODB.Recordset
Dim rsPrec As New ADODB.Recordset
Dim Total As Double
Dim Tipo As String

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Combo1_Click()
rS.MoveFirst
rS.Move Combo1.ListIndex
almacen = Format(rS(0), "00")
End Sub

Private Sub CmdAceptar_Click()
Dim rSql As String
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

Private Sub Carga_RepoVal()
Dim rAdo As ADODB.Recordset
Dim Aux, cadena As String
Dim cAnoMes As String, cCod As String
Dim cSql1 As String, CSQL2 As String
Dim Dia1 As Integer, Mes1 As Integer
Dim nSaldo As Double, nCosPro As Double
Dim nPrecio As Double, nCantid As Double
Dim i, nCount As Integer
Dim Codi As String
Dim rSql As String
Dim csql As String
Dim cSql22 As String
Dim uSql As String
Dim saldo As Double
On Error GoTo ErrCarga
Set Adodc1 = New ADODB.Recordset
Set rAdo = New ADODB.Recordset
Dim MaeartRs As New ADODB.Recordset

'Indica mes del ultimo cierre
'nMes = Month(Date)
'nAnno = Year(Date)
'rSql = "SELECT TOP 1 CAFECDOC From MovAlmCab  WHERE CACIERRE = false AND CAALMA = '" & VGAlma & "'  ORDER BY CAFECDOC asc"
'Set Rs2 = New ADODB.Recordset
'Rs2.Open rSql, cConexCom, adOpenStatic
'If Not Rs2.EOF Then
'    Rs2.MoveFirst
'    nMes = Month(Rs2(0))
'    nAnno = Year(Rs2(0))
'    If nMes = 13 Then
'         nMes = 1
'    End If
'    MsgBox "Se va a ralizar el cierre del  mes : " & MonthName(nMes), vbInformation, "Aviso"
'End If


csql = "delete  from MORESMES WHERE SMALMA='" & almacen & "'"
cConexCom.Execute csql
'*******************************************************
'**Estas Actualizaciones deben ir en un metodo de Actualizacion....!
'*******************************************************
csql = "UPDATE MOVALMCAB SET CAHORA = FORMAT(VAL(LEFT(CAHORA,2))-12,'00' )+':00:00' " & _
       "WHERE (CATIPMOV='I') AND VAL(LEFT(CAHORA,2)) >= 12 "
cConexCom.Execute csql

csql = "UPDATE MOVALMCAB SET CAHORA = FORMAT(VAL(LEFT(CAHORA,2))+12,'00' )+':00:00'" & _
       "WHERE CATIPMOV='S' AND VAL(LEFT(CAHORA,2)) <= 12"
       
cConexCom.Execute csql
'*******************************************************
'*******************************************************
BarraProc.Visible = True
cArticulo.Visible = True
cArticulo.Caption = "Espere Un Momento....! "
Frame1.Refresh

'*******************************************************
'**Carga todos los Movimientos al respectivo Recordset
'*******************************************************
Set Adodc1 = New ADODB.Recordset
cSql1 = "Select * From MovAlmDet A Inner Join MovAlmCab B On B.CAALMA = A.DEALMA And B.CATD = A.DETD  And B.CANUMDOC = A.DENUMDOC "
cSql1 = cSql1 & " Where CAALMA = '" & almacen & "' and not (CATD='GS' And CACODMOV='GF' And CASITGUI='F') "
cSql1 = cSql1 & " And CASITGUI<>'A'  Order By DECODIGO,CAFECDOC,CAHORA"
Adodc1.Open cSql1, cConexCom, adOpenForwardOnly
'*******************************************************
'**Carga todos Codigos de los productos al respectivo Recordset
'*******************************************************
BarraProc.Min = 50
cSql22 = "Select distinct (Decodigo) From MovAlmDet A Inner Join MovAlmCab B On B.CAALMA = A.DEALMA And B.CATD = A.DETD  And B.CANUMDOC = A.DENUMDOC "
cSql22 = cSql22 & " Where CAALMA = '" & almacen & "' and not (CATD='GS' And CACODMOV='GF' And CASITGUI='F') "
cSql22 = cSql22 & " And CASITGUI<>'A' "
BarraProc.Min = 10
MaeartRs.Open cSql22, cConexCom, adOpenStatic

nCount = 0
BarraProc.Max = 100 + MaeartRs.RecordCount
'*******************************************************
'**********'INICIALIZA A 0 todos los articulos de Stkart (stock de Articulos )
csql = "UPDATE STKART SET STSKDIS=0 WHERE STALMA='" & almacen & "'"
cConexCom.Execute csql
'*******************************************************
BarraProc.Min = 0
While Not MaeartRs.EOF
    nCount = nCount + 1
    BarraProc.Value = nCount
    cArticulo.Caption = "Valorizando Articulo : " & nCount & "   " & (MaeartRs!DECODIGO)
    Frame1.Refresh
    Call ValorizaXArticulo(MaeartRs!DECODIGO) 'Procedimiento que Valoriza por Articulo
    MaeartRs.MoveNext
Wend
'*********************************************
    ClsTock.BorrarServicios almacen, cConexCom
'*********************************************

    BarraProc.Visible = False
    cArticulo.Visible = False

Exit Sub
ErrCarga:
        MsgBox Err.Description
        
        If nTra = 1 Then cConexAux.RollbackTrans
        BarraProc.Visible = False
        cArticulo.Visible = False
End Sub

Private Sub Carga_Almacen()
Dim rSql As String
Dim i As Integer
rSql = "Select TAALMA,TADESCRI FROM TabAlm "
Set rS = New ADODB.Recordset
rS.Open rSql, cConexCom, adOpenStatic
Do While Not rS.EOF
     Combo1.AddItem (rS(1))
     rS.MoveNext
     If rS.EOF Then Exit Do
Loop
rS.MoveFirst
For i = 0 To rS.RecordCount - 1
  If rS(0) = VGAlma Then
    Combo1.ListIndex = i
    Exit For
  Else
    rS.MoveNext
  End If
Next
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub


Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
Dim rSql As String
central Me
Carga_Almacen
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
Private Sub ValorizaXArticulo(ByVal vCodArt As String)
Dim TCamb As Double
Dim Li As Integer
Dim nCambio, nSaldo As Double, nCosPro, nCosProUS As Double
Dim nPrecio, nPrecioUS, xPrecio As Double, nCantid As Double
Dim cMesPro, cMesActu, cMesAnte As String
Dim Rsql1 As String
Dim nTipCam, cSql1 As String
'**********Roberto
Dim VALMOV, VALANTE, VALMOVUS, VALANTEUS As Double
Dim nMes, nYear As Long
Dim nSal, nIng, nSaldoInicial As Double
Dim dfecha As Date
Dim csql As String
Dim XNUMDOC As String
On Local Error GoTo errar

Adodc1.Filter = " Decodigo='" & vCodArt & "'"
xPrecio = 0
nPrecio = 0: nCantid = 0
nPrecioUS = 0
nSal = 0: nIng = 0
nSaldoInicial = 0
Adodc1.MoveFirst
nCosProUS = 0: nCosPro = 0
nMes = Month(Adodc1("CAFECDOC"))
nYear = Year(Adodc1("CAFECDOC"))
dfecha = Adodc1("CAFECDOC")


Do While Not Adodc1.EOF
    
   If Year(Adodc1("CAFECDOC")) <> nYear Or Month(Adodc1("CAFECDOC")) <> nMes Then
  
      cMesPro = Format(nYear, "0000") & Format(nMes, "00")
      Rsql1 = "INSERT INTO  MoResMes (SMALMA,SMCODIGO,SMMESPRO,SMSALDOINI,SMCANENT,SMCANSAL,SMMNPREUNI,SMUSPREUNI)VALUES('" & almacen & "','" & vCodArt & "','" & cMesPro & "'," & nSaldoInicial & "," & nIng & "," & nSal & "," & nCosPro & "," & nCosProUS & ")"
      
      cConexCom.BeginTrans
      cConexCom.Execute Rsql1
      cConexCom.CommitTrans
      
      cMesActu = (Format(Year(Adodc1("CAFECDOC")), "0000") & Format(Month(Adodc1("CAFECDOC")), "00"))
      nSaldoInicial = nSaldoInicial + (nIng - nSal)
      nIng = 0
      nSal = 0
      cMesAnte = AnioMesSiguiente(cMesPro)
      While cMesAnte <> cMesActu
            Rsql1 = "INSERT INTO  MoResMes (SMALMA,SMCODIGO,SMMESPRO,SMSALDOINI,SMCANENT,SMCANSAL,SMMNPREUNI,SMUSPREUNI)VALUES('" & almacen & "','" & vCodArt & "','" & cMesAnte & "'," & nSaldoInicial & "," & nIng & "," & nSal & "," & nCosPro & "," & nCosProUS & ")"
            cConexCom.BeginTrans
            cConexCom.Execute Rsql1
            cConexCom.CommitTrans
            cMesAnte = AnioMesSiguiente(cMesAnte)
      Wend
      '*************************************************
      dfecha = Adodc1("CAFECDOC")
      nMes = Month(Adodc1("CAFECDOC"))
      nYear = Year(Adodc1("CAFECDOC"))
            
  Else
  
     '*************************************************
      If Adodc1!CATIPCAM = 0 Or Adodc1!CATIPCAM = 1 Then
            If UCase(Dir$(cRuta4)) = UCase(VGNameCont & ".MDB") Then
               TCamb = Val(Devolver_Dato(3, Format(dfecha, "DD/MM/YYYY"), "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "ME", "TIPOMON_CODIGO"))
            Else
               If UCase(Dir$(cRuta2)) = "BDCOMUN.MDB" Then
                  TCamb = Val(Devolver_Dato(1, Format(dfecha, "DD/MM/YYYY"), "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "ME", "TIPOMON_CODIGO"))
               Else
                  TCamb = Adodc1!CATIPCAM
               End If
            End If
            
            If TCamb = 0 Then TCamb = Adodc1!DETIPCAM
      Else
          TCamb = Adodc1!DETIPCAM
      End If
      
      nCantid = Adodc1("DECANTID")
      nPrecio = Adodc1("DEPRECIO")

      '*************************************************
      '***DOCUMENTOS EN  DOLARES
      If cNull(Adodc1!CACODMON) = "ME" Then
         nPrecio = Adodc1("DEPRECIO") * TCamb
      Else
         nPrecio = Adodc1("DEPRECIO")
      End If
      '*************************************************
      '*************************************************
      '***DOCUMENTOS EN SOLES
      If cNull(Adodc1!CACODMON) = "ME" Then
         nPrecioUS = Adodc1("DEPRECIO")
      Else
         If Round(TCamb, 3) > 0 Then
            nPrecioUS = Adodc1("DEPRECIO") / TCamb
         Else
            nPrecioUS = 0
         End If
      End If
      '*************************************************
      
      If Adodc1("CATIPMOV") = "I" Then
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
      
      If Adodc1("CATIPMOV") = "I" Then
         If nSaldo <> 0 Then
            nCosPro = (VALMOV + VALANTE) / nSaldo
            nCosProUS = (VALMOVUS + VALANTEUS) / nSaldo
         End If
      End If
     
      VALANTE = nCosPro * nSaldo
      VALANTEUS = nCosProUS * nSaldo
      dfecha = Adodc1("CAFECDOC")
                
      Adodc1.MoveNext
    End If
   
   
Loop


     cMesPro = Format(Year(dfecha), "0000") & Format(Month(dfecha), "00")

'*************************************************
     Rsql1 = "INSERT INTO  MoResMes (SMALMA,SMCODIGO,SMMESPRO,SMSALDOINI,SMCANENT,SMCANSAL,SMMNPREUNI,SMUSPREUNI)VALUES('" & almacen & "','" & vCodArt & "','" & cMesPro & "'," & nSaldoInicial & "," & nIng & "," & nSal & "," & nCosPro & "," & nCosProUS & ")"

     cConexCom.BeginTrans
     cConexCom.Execute Rsql1
     cConexCom.CommitTrans
 '*************************************************
      nSaldoInicial = nSaldoInicial + (nIng - nSal)
      cMesActu = AnioMesSiguiente(Format(Year(Now), "0000") & Format(Month(Now), "00"))
      nIng = 0
      nSal = 0
      cMesAnte = AnioMesSiguiente(cMesPro)
'      If nSaldoInicial <> 0 Then
'         MsgBox " no se movio"
'      End If
      While cMesAnte <= cMesActu
            Rsql1 = "INSERT INTO  MoResMes (SMALMA,SMCODIGO,SMMESPRO,SMSALDOINI,SMCANENT,SMCANSAL,SMMNPREUNI,SMUSPREUNI)VALUES('" & almacen & "','" & vCodArt & "','" & cMesAnte & "'," & nSaldoInicial & "," & nIng & "," & nSal & "," & nCosPro & "," & nCosProUS & ")"
            cConexCom.BeginTrans
            cConexCom.Execute Rsql1
            cConexCom.CommitTrans
            cMesAnte = AnioMesSiguiente(cMesAnte)
      Wend

      If ClsTock.ExisteEnStock(almacen, vCodArt, cConexCom) Then
         cConexCom.Execute "Update STKART SET STSKDIS=" & nSaldoInicial & ",STKPREPRO=" & nCosPro & " WHERE STALMA='" & almacen & "' AND STCODIGO='" & vCodArt & "'"
      Else
         cConexCom.Execute "Insert Into STKART (STALMA,STCODIGO,STKFECULT,STSKDIS,STKPREULT,STKPREPRO)Values('" & almacen & "','" & vCodArt & "','" & dfecha & "'," & nSaldoInicial & ",0," & nCosPro & ")"
      End If

Exit Sub

errar:
MsgBox Err.Description
BarraProc.Visible = False
cArticulo.Visible = False

End Sub

Function PrecioFact(ByVal arTd As String, ByVal arNumdoc As String, ByVal arCodi As String) As Double
         Set rsPrec = New ADODB.Recordset
         rsPrec.Open "Select DFPREC_ORI from facdet where dftd='" & arTd & "' AND  dfnumser+dfnumdoc='" & arNumdoc & "' and dfcodigo='" & arCodi & "'", cConexCom, adOpenForwardOnly, adLockReadOnly
         PrecioFact = IIf(Not rsPrec.EOF, rsPrec!DFPREC_ORI, 0)
         rsPrec.Close
End Function

'''Private Sub Carga2()
'''Dim TCamb As Double
'''Dim Li As Integer
'''Dim cCod As String, Codi As String
'''Dim Aux, cadena As String
'''Dim cAnoMes As String
'''Dim cSql1 As String, CSQL2 As String
'''Dim Dia1 As Integer, Mes1 As Integer
'''Dim nSaldo As Double, nCosPro As Double
'''Dim nPrecio As Double, nCantid As Double
'''Dim cMesPro As String
'''Dim Rsql1 As String
'''Dim nTipCam As String
'''**********Roberto
'''Dim rSAX As New ADODB.Recordset
'''Dim SqlM As String
'''Dim Flag0 As Integer
'''Dim VALMOV, VALANTE As Double
'''Dim nMes, nYear As Long
'''Dim nSal, nIng As Double
'''Dim dfecha As Date
'''On Local Error GoTo ERRAR
'''*************************
'''Mes1 = Month(DTPicker1)
'''cAnoMes = Year(DTPicker1) & Format(Mes1, "00")
'''cCod = "": Codi = "": Li = 0
'''Flag0 = 0
'''nMes = 0
'''Call BorraMorasMes
'''cConexCom.BeginTrans
'''cConexCom.Execute "Delete FROM MoresMes"
'''cConexCom.CommitTrans
'''
'''Adodc3.Filter = "CODART='" & vCodArt & "'"
'''adodc1.MoveFirst
'''Do While Not adodc1.EOF
'''
'''       If (adodc1("CATD") = "GS" And adodc1("CACODMOV") = "GF" And adodc1("CASITGUI") = "F") Then '(adodc1("CATD") = "GS" And adodc1("CACODMOV") = "GV" And adodc1("CASITGUI") = "F") Or
'''            adodc1.MoveNext
'''       Else
'''          If cCod <> adodc1("DECODIGO") Then
'''             If adodc1("DECODIGO") = "8007-AS-E32" Then
'''                MsgBox "HOLA"
'''             End If
'''             nPrecio = 0: nCantid = 0
'''             nSal = 0: nIng = 0
'''             nMes = Month(adodc1("CAFECDOC"))
'''             nYear = Year(adodc1("CAFECDOC"))
'''
'''                nSaldo = 0
'''                nCosPro = 0
'''
'''             VALANTE = nCosPro * nSaldo
'''          End If
'''
'''
'''          nCantid = adodc1("DECANTID")
'''
'''          nPrecio = adodc1("DEPRECIO")
'''
'''          If adodc1("CATIPMOV") = "I" Then
'''             nSaldo = nSaldo + nCantid
'''             VALMOV = nCantid * nPrecio
'''             nIng = nIng + nCantid
'''          Else
'''             nSaldo = nSaldo - nCantid
'''             VALMOV = nCantid * nCosPro
'''             nSal = nSal + nCantid
'''          End If
'''
'''
'''****************************************
'''          If adodc1("CATIPMOV") = "I" Then
'''                If nSaldo <> 0 Then
'''                   nCosPro = (VALMOV + VALANTE) / nSaldo
'''                End If
'''          End If
'''****************************************
'''          VALANTE = nCosPro * nSaldo
'''
'''          CSQL2 = "Insert Into Kardex_Val (COD_ART,FEC_DOC,HOR_DOC,COD_MOV,TIP_TRANSA,NUM_DOC,CAN_ART,PRE_UNIT,COS_PRO,SAL_STOCK,SER_LOT,ING_SAL)  "
'''          CSQL2 = CSQL2 & " Values ('" & adodc1("DECODIGO") & "',#" & Format(adodc1("CAFECDOC"), "mm/dd/yyyy") & "#,'" & adodc1("CAHORA") & "','" & adodc1("CACODMOV") & "',"
'''          CSQL2 = CSQL2 & "'" & adodc1("CATD") & "','" & adodc1("CANUMDOC") & "', " & adodc1("DECANTID") & ","
'''
'''          If adodc1("CATIPMOV") = "I" Then
'''              CSQL2 = CSQL2 & "" & nPrecio & "," & nCosPro & "," & nSaldo & ",'" & IIf(IsNull(adodc1("DESERIE")) Or adodc1("DESERIE") = "", adodc1("DELOTE"), adodc1("DESERIE")) & "','I')"
'''          Else
'''              CSQL2 = CSQL2 & "" & nCosPro & "," & nCosPro & "," & nSaldo & ",'" & IIf(IsNull(adodc1("DESERIE")) Or adodc1("DESERIE") = "", adodc1("DELOTE"), adodc1("DESERIE")) & "','S')"
'''          End If
'''          nTra = 1
'''          cConexAux.BeginTrans
'''          cConexAux.Execute CSQL2
'''          cConexAux.CommitTrans
'''          nTra = 0
'''
'''          dfecha = adodc1("CAFECDOC")
'''          cCod = adodc1("DECODIGO")
'''
'''          adodc1.MoveNext
'''
'''          If Month(adodc1("CAFECDOC") <> nYear Or Month(adodc1("CAFECDOC")) <> nMes) Then
'''
'''             cMesPro = Format(Year(dfecha), "0000") & Format(Month(dfecha), "00")
'''         *************************************************
'''             If UCase(Dir$(cRuta4)) = UCase(VGNameCont & ".MDB") Then
'''                nTipCam = Val(Devolver_Dato(3, Format(dfecha, "DD/MM/YYYY"), "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "ME", "TIPOMON_CODIGO"))
'''             ElseIf UCase(Dir$(cRuta2)) = "BDCOMUN.MDB" Then
'''                    nTipCam = Val(Devolver_Dato(1, Format(dfecha, "DD/MM/YYYY"), "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "ME", "TIPOMON_CODIGO"))
'''             End If
'''
'''             Rsql1 = "INSERT INTO  MoResMes (SMALMA,SMCODIGO,SMMESPRO,SMCANENT,SMCANSAL,SMMNPREUNI,SMUSPREUNI)VALUES('" & almacen & "','" & cCod & "','" & cMesPro & "'," & nIng & "," & nSal & "," & nCosPro & "," & nCosPro / IIf(nTipCam = 0, 1, nTipCam) & ")"
'''
'''
'''             cConexCom.BeginTrans
'''             cConexCom.Execute Rsql1
'''             cConexCom.CommitTrans
'''         *************************************************
'''             nMes = Month(adodc1("CAFECDOC"))
'''             nYear = Year(adodc1("CAFECDOC"))
'''             nIng = 0
'''             nSal = 0
'''          End If
'''       End If
'''
'''       If adodc1.EOF Then Exit Do
'''
'''Loop
'''Exit Sub
'''
'''ERRAR:
'''     MsgBox err.Description
'''
'''     cMesPro = Format(Year(dfecha), "0000") & Format(Month(dfecha), "00")
''' *************************************************
'''     If UCase(Dir$(cRuta4)) = UCase(VGNameCont & ".MDB") Then
'''        nTipCam = Val(Devolver_Dato(3, Format(dfecha, "DD/MM/YYYY"), "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "ME", "TIPOMON_CODIGO"))
'''     ElseIf UCase(Dir$(cRuta2)) = "BDCOMUN.MDB" Then
'''            nTipCam = Val(Devolver_Dato(1, Format(dfecha, "DD/MM/YYYY"), "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "ME", "TIPOMON_CODIGO"))
'''     End If
'''
'''     Rsql1 = "INSERT INTO  MoResMes (SMALMA,SMCODIGO,SMMESPRO,SMCANENT,SMCANSAL,SMMNPREUNI,SMUSPREUNI)VALUES('" & almacen & "','" & cCod & "','" & cMesPro & "'," & nIng & "," & nSal & "," & nCosPro & "," & nCosPro / IIf(nTipCam = 0, 1, nTipCam) & ")"
'''
'''
'''     cConexCom.BeginTrans
'''     cConexCom.Execute Rsql1
'''     cConexCom.CommitTrans
''' *************************************************
'''    Resume
'''End Sub

