VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmKardexRevalorizaMes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Revalorización Mensual"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4260
      Left            =   90
      TabIndex        =   0
      Top             =   -45
      Width           =   6105
      Begin VB.OptionButton Opt2 
         Caption         =   "Cierre Previo"
         Height          =   228
         Left            =   3636
         TabIndex        =   13
         Top             =   2790
         Width           =   1416
      End
      Begin VB.OptionButton Opt1 
         Caption         =   "Cierre Definitivo"
         Height          =   228
         Left            =   1332
         TabIndex        =   12
         Top             =   2745
         Value           =   -1  'True
         Width           =   1848
      End
      Begin MSComctlLib.ProgressBar BarraProc 
         Height          =   225
         Left            =   1350
         TabIndex        =   9
         Top             =   1935
         Visible         =   0   'False
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
         Min             =   10
         Max             =   1000
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FrmKardexRevalorizaMes.frx":0000
         Left            =   1650
         List            =   "FrmKardexRevalorizaMes.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   330
         Width           =   3450
      End
      Begin VB.CommandButton Cmd_Revalorizar 
         Caption         =   "&Aceptar"
         Height          =   645
         Left            =   2280
         Picture         =   "FrmKardexRevalorizaMes.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3300
         Width           =   810
      End
      Begin VB.CommandButton cmd_exit 
         Caption         =   "&Salir"
         Height          =   645
         Left            =   3120
         Picture         =   "FrmKardexRevalorizaMes.frx":0446
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3300
         Width           =   810
      End
      Begin MSComCtl2.DTPicker dFecVal 
         Height          =   312
         Left            =   1650
         TabIndex        =   1
         Top             =   816
         Width           =   3450
         _ExtentX        =   6085
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
         Format          =   108331011
         CurrentDate     =   37068
      End
      Begin MSComctlLib.ProgressBar BarraProc1 
         Height          =   225
         Left            =   1350
         TabIndex        =   7
         Top             =   2295
         Visible         =   0   'False
         Width           =   3900
         _ExtentX        =   6879
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
         Min             =   10
         Max             =   1000
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   210
         Top             =   3165
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin MSComCtl2.DTPicker DFecnuevo 
         Height          =   315
         Left            =   0
         TabIndex        =   14
         Top             =   3840
         Visible         =   0   'False
         Width           =   1050
         _ExtentX        =   1852
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
         Format          =   108331011
         CurrentDate     =   37068
      End
      Begin VB.Label Label2 
         Caption         =   "Observaciones"
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   2325
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label Label2 
         Caption         =   "OK"
         Height          =   285
         Index           =   0
         Left            =   630
         TabIndex        =   10
         Top             =   1935
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label cArticulo 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1350
         TabIndex        =   8
         Top             =   1320
         Visible         =   0   'False
         Width           =   3705
      End
      Begin VB.Label Label6 
         Caption         =   "Almacen :"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   375
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Mes :"
         Height          =   285
         Left            =   720
         TabIndex        =   2
         Top             =   870
         Width           =   1065
      End
   End
End
Attribute VB_Name = "FrmKardexRevalorizaMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PCount As Long
'Dim cConexAux As ADODB.Connection
Dim adodc1 As New ADODB.Recordset
Dim Adodc2 As New ADODB.Recordset
Dim rs As ADODB.Recordset
Dim rsPrec As ADODB.Recordset
Dim cRt As String
Dim almacen, cMesCirr As String
Dim nTra As Integer

Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub cmd_exit_Click()
  Unload Me
End Sub

Private Sub Combo1_Click()
rs.MoveFirst
rs.Move Combo1.ListIndex
almacen = Format(rs(0), "00")
End Sub

Private Sub Cmd_Revalorizar_Click()
On Error GoTo Err
   Dim SQL As String
   Screen.MousePointer = 11
   Carga_RepoVal
   Screen.MousePointer = 1
Exit Sub
Err:
   MsgBox Err.Description, vbInformation, "Aviso"
   Screen.MousePointer = 1
End Sub

Private Sub Carga_RepoVal()
Dim cAnoMes As String, cCod As String
Dim cSql1 As String, CSQL2 As String
Dim nSaldo As Double, nCosPro As Double
Dim nPrecio As Double, nCantid As Double
Dim nCount, nMaxRec As Integer
Dim csql As String
Dim registro As Integer
Dim cSql22, cAnomesNuevo As String
On Error GoTo ErrCarga
Dim MaeartRs As New ADODB.Recordset
Dim Rsmes As New ADODB.Recordset
Dim cMesActu As String
Set adodc1 = New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim RSQL As String
cAnoMes = Format(dFecVal.Year, "0000") & Format(dFecVal.Month, "00")
DFecnuevo = dFecVal + 31
cAnomesNuevo = Format(DFecnuevo.Year, "0000") & Format(DFecnuevo.Month, "00")
Set Rsmes = VGCNx.Execute("select max(cierrmes) as cierremes from al_cierresmensuales where cierralma='" & almacen & "'")
cMesCirr = ESNULO(Rsmes!cierremes, "")
If cMesCirr <> "" Then
      
   If cAnoMes <= cMesCirr Then
     MsgBox "El Mes Que Usted Selecciono ya Esta Cerrado o " & Chr(10) & " No hay Informacion Registrada en la Respectiva Fecha ", vbInformation, "Verifique...!"
     Exit Sub
   End If
   
   If cAnoMes > AnioMesSiguiente(cMesCirr) Then
     MsgBox "El Mes que usted selecciono No Pueder Ser Valorizado" & Chr(10) & "Por Favor Valorize el Mes Anterior", vbInformation, "Verifique...!"
     Exit Sub
   End If
   'rs.Open "Select CIERRMES FROM al_cierresnemsuales WHERE  CIERRALMA='" & almacen & "'", Vgcnx, adOpenStatic, adLockReadOnly
   Set rs = VGCNx.Execute("Select CIERRMES FROM al_cierresmensuales WHERE  CIERRALMA='" & almacen & "'")
   
   If Not rs.EOF Then
   '*****************************************
       Set rs = New ADODB.Recordset
       Set rs = VGCNx.Execute("Select CIERRMES FROM al_cierresmensuales WHERE  CIERRALMA='" & almacen & "' AND CIERRMES='" & cMesCirr & "'")
       If rs.EOF Then
          cMesCirr = AnioMesAnterior(cMesCirr)
          MsgBox "Es Necesario Cerrar el Mes Anterior para el Manejo Correcto de sus Saldos " & Chr(10) & "Por Favor Valorize el Mes Anterior", vbInformation, "Seleccione el Mes Anterior...!"
          rs.Close
          registro = 1
       End If
   End If
   rs.Close
End If

RSQL = "Select  p.ACODIGO, p.ADESCRI, m.CACODMOV ,n.DENUMDOC, m.CANOMPRO, m.CARFTDOC " & _
          "from MaeArt p,MovAlmCab m, MovAlmDet n   " & _
          "where  m.CAALMA ='" & almacen & _
          "' AND  n.DEALMA = m.CAALMA and CATIPMOV='I'  and p.ACODIGO = n.DECODIGO and  m.CASITGUI <> 'A'  and " & _
          " n.DEPRECIO = 0  AND ISNULL(CACIERRE,0)<>1 And  n.DENUMDOC  = m.CANUMDOC and  n.DETD = m.CATD   AND MONTH(CAFECDOC) <= " & dFecVal.Month & " AND YEAR(CAFECDOC) = " & dFecVal.Year & " ORDER BY m.CANUMDOC"

Set Rsmes = New ADODB.Recordset
Set Rsmes = VGCNx.Execute(RSQL)

If Rsmes.RecordCount > 0 And Opt1.Value = True Then
   MsgBox "Debe Valorizar todos sus Articulos Pendientes", vbInformation, mensaje1
   Rsmes.Close
   Screen.MousePointer = 1
   Exit Sub
End If
Rsmes.Close

'*******************************************************

Label2(0).Visible = True
Label2(1).Visible = True

BarraProc.Visible = True
BarraProc1.Visible = True
cArticulo.Visible = True
cArticulo.Caption = "Espere Un Momento....! "
Frame1.Refresh

Set adodc1 = New ADODB.Recordset
cSql1 = "Select * From MovAlmDet A Inner Join MovAlmCab B On B.CAALMA = A.DEALMA And B.CATD = A.DETD  And B.CANUMDOC = A.DENUMDOC "
cSql1 = cSql1 & " Where CAALMA = '" & almacen & "' and not (CACODMOV='GF' And CASITGUI='F') "
cSql1 = cSql1 & " And CASITGUI<>'A' AND cast(year(CAFECDOC) as varchar(4)) + replicate('0',2-len(month(CAFECDOC))) + cast(month(CAFECDOC) as varchar(2)) >='" & Format(dFecVal.Year, "0000") + Format(dFecVal.Month, "00") & "'   order By DECODIGO,CAFECDOC,catipmov,CAHORA"

'Adodc1.Open cSql1, Vgcnx, adOpenStatic, adLockReadOnly

Set adodc1 = VGCNx.Execute(cSql1)

If adodc1.EOF Then
   MsgBox "No Existe Información Registrada en la Fecha que Usted Indico", vbInformation, "Verifique....!"
    BarraProc.Visible = False
    BarraProc1.Visible = False
    cArticulo.Visible = False
    Label2(0).Visible = False
    Label2(1).Visible = False
    Exit Sub
End If


csql = "Delete From al_Kardex_Val"
VGCNx.Execute csql

'******************************************************
cSql22 = "SELECT STKART.STALMA, STKART.STCODIGO, STKART.STSKDIS FROM STKART WHERE STALMA='" & almacen & "'"
BarraProc.Min = 10
'MaeartRs.Open cSql22, Vgcnx, adOpenStatic
Set MaeartRs = VGCNx.Execute(cSql22)
'******************************************************
nCount = 0
nMaxRec = MaeartRs.RecordCount
BarraProc.Max = 100 + nMaxRec
BarraProc1.Max = 100 + nMaxRec
BarraProc.Min = 0
BarraProc1.Min = 0
Frame1.Refresh
While Not MaeartRs.EOF
    nCount = nCount + 1
    BarraProc.Value = nCount
    cArticulo.Caption = "Valorizando Articulo : " & Format(nCount, "00000") & "     -     " & Format(nMaxRec, "00000") & " " & Chr(10) & (MaeartRs!stcodigo)
    Frame1.Refresh
    Call ClsTock.RestauraStock(almacen, MaeartRs!stcodigo, CDate("01/" & Format(dFecVal.Month, "00") & "/" & Format(dFecVal.Year, "0000")))
    'Call ValorizaXArticuloMes(MaeartRs!STCODIGO, cAnoMes)
    MaeartRs.MoveNext
Wend
'*********************************************
    ClsTock.BorrarServicios almacen, VGCNx
'*********************************************
    BarraProc.Visible = False
    BarraProc1.Visible = False
    cArticulo.Visible = False
    Label2(0).Visible = False
    Label2(1).Visible = False
    
    MaeartRs.Close
         
    csql = "Select cod_art from al_Kardex_Val"
    'MaeartRs.Open csql, cConexAux, adOpenStatic, adLockPessimistic
    Set MaeartRs = VGCNx.Execute(csql)
    
    If Not MaeartRs.EOF Then
       Call Reporte
       If MsgBox("En el Proceso de Cierre de Mes se Encontraron ciertas Irregularidades  que Se Detalla en el Reporte" & Chr(10) & "Desea Continuar Con el Cierre ", vbInformation + vbYesNo, Caption) = vbYes Then
            If Opt1.Value = True Then
               VGCNx.BeginTrans
               VGCNx.Execute "INSERT INTO  al_cierresmensuales (CierrMes,CierrFech,CierrOper,Cierralma)VALUES('" & cAnoMes & "'," & Format(Now, "dd/mm/yyyy") & ",'RMAZA','" & almacen & "')"
               VGCNx.CommitTrans
               csql = "UPDATE MovAlmCab set CACIERRE =  TRUE " & _
                      " where  CAALMA = '" & almacen & "'   AND MONTH(CAFECDOC) =" & dFecVal.Month & " AND YEAR(CAFECDOC) = " & dFecVal.Year
               VGCNx.BeginTrans
               VGCNx.Execute csql
               VGCNx.CommitTrans
               cMesCirr = cAnoMes
            Else
               MsgBox "Recuerde que Debe Realizar un Cierre Definitivo...! ", vbInformation, Caption
            End If
       End If
    Else
       If Opt1.Value = True Then
          nTra = 1
          VGCNx.BeginTrans
          VGCNx.Execute "INSERT INTO  al_cierresmensuales (CierrMes,CierrFech,CierrOper,Cierralma) VALUES('" & cAnoMes & "'," & Format(Now, "dd/mm/yyyy") & ",'RMAZA','" & almacen & "')"
          csql = "UPDATE MovAlmCab set CACIERRE =  1 " & _
                 " where  CAALMA = '" & almacen & "'   AND MONTH(CAFECDOC) =" & dFecVal.Month & " AND YEAR(CAFECDOC) = " & dFecVal.Year
          VGCNx.Execute csql
              Screen.MousePointer = 11
         
         'actualizando los saldos
    
          Set VGCommandoSP = New ADODB.Command
          VGCommandoSP.ActiveConnection = VGgeneral
          VGCommandoSP.CommandType = adCmdStoredProc
          VGCommandoSP.CommandText = "al_cierremensual_pro"
          VGCommandoSP.Parameters.Refresh
          With VGCommandoSP
              .Parameters("@base") = VGParamSistem.BDEmpresa
              .Parameters("@almacen") = almacen
              .Parameters("@mesactual") = cAnoMes
              .Parameters("@Mesnuevo") = cAnomesNuevo
              .Execute
          End With
          VGCNx.CommitTrans
          nTra = 0
          cMesCirr = cAnoMes
       Else
          MsgBox "Recuerde que Debe Realizar un Cierre Definitivo...! ", vbInformation, Caption
       End If
    End If
    MaeartRs.Close
    
adodc1.Close

Exit Sub
ErrCarga:
        MsgBox Err.Description
        If nTra = 1 Then VGCNx.RollbackTrans
        BarraProc.Visible = False
        cArticulo.Visible = False
        BarraProc1.Visible = False
        Label2(0).Visible = False
        Label2(1).Visible = False
        Exit Sub
        Resume
End Sub

Private Sub Carga_Almacen()
Dim RSQL As String
Dim I As Integer
Dim X As Integer
RSQL = "Select TAALMA,TADESCRI FROM TabAlm "
Set rs = New ADODB.Recordset
'rs.Open RSQL, Vgcnx, adOpenStatic
Set rs = VGCNx.Execute(RSQL)
Do While Not rs.EOF
     Combo1.AddItem (rs(1))
     rs.MoveNext
     If rs.EOF Then Exit Do
Loop
X = 0
rs.MoveFirst
For I = 0 To rs.RecordCount - 1
  If rs(0) = VGAlma Then
    Combo1.ListIndex = I
    X = 1
    Exit For
  Else
    rs.MoveNext
  End If
Next
If X = 0 Then Combo1.ListIndex = 1
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub


Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
Dim RSQL As String
Dim Rs2 As New ADODB.Recordset
central Me
'Call ADOConectar
Call Carga_Almacen

'dFecVal.Value = Format(Now, "dd/mm/yyyy")

RSQL = "SELECT TOP 1 min(CAFECDOC) From MovAlmCab  WHERE isnull(CACIERRE,0) = 0 AND CAALMA = '" & almacen & "'"
Set Rs2 = New ADODB.Recordset
'Rs2.Open RSQL, Vgcnx, adOpenStatic
Set Rs2 = VGCNx.Execute(RSQL)
If Not Rs2.EOF Then
    Rs2.MoveFirst
    If IsNull(Rs2.Fields(0)) Then
       MsgBox "No hay Informacion que tenga Cierre Pendiente ", vbInformation, "Aviso"
       Exit Sub
       Unload Me
    End If
    
    dFecVal.Value = Format(Rs2(0), "DD/MM/YYYY")
    cMesCirr = AnioMesAnterior(Format(Year(Rs2(0)), "0000") & Format(Month(Rs2(0)), "00"))
Else
    MsgBox "No hay Informacion que tenga Cierre Pendiente ", vbInformation, "Aviso"
    Exit Sub
End If
        
End Sub

'Private Sub ADOConectar()
'cRt = App.Path & "\BdAuxCom.Mdb"
'Set cConexAux = New ADODB.Connection
'cConexAux.CursorLocation = adUseClient
'cConexAux.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & cRt & ";"
'cConexAux.Open
'End Sub

Private Sub ValorizaXArticuloMes(ByVal vCodArt As String, ByVal cArmes As String)
Dim TCamb As Double
Dim Li As Integer
Dim nCambio, nSaldo As Double, nCosPro, nCosProUS As Double
Dim nPrecio, nPrecioUS, xPrecio As Double, nCantid As Double
Dim cMesPro, cMesActu, cMesAnte, cAnoMes As String
Dim Rsql1 As String
Dim nTipCam, cSql1 As String
'**********Roberto
Dim VALMOV, VALANTE, VALMOVUS, VALANTEUS As Double
Dim nMes, nYear As Long
Dim nSal, nIng, nSaldoInicial As Double
Dim dfecha As Date
Dim csql, CSQL2 As String

On Local Error GoTo ERRAR

adodc1.Filter = " Decodigo='" & vCodArt & "'"
xPrecio = 0
nPrecio = 0: nCantid = 0
nPrecioUS = 0
nSal = 0: nIng = 0
nSaldoInicial = 0
nCosProUS = 0: nCosPro = 0

If Not adodc1.EOF Then
   nMes = Month(adodc1("CAFECDOC"))
   nYear = Year(adodc1("CAFECDOC"))
   dfecha = adodc1("CAFECDOC")
Else
   dfecha = CDate("01/01/1500") '****************fecha vacia
End If

'If vCodArt = "10084" Then
'   MsgBox "vCodArt"
'End If

cAnoMes = cArmes

Adodc2.Filter = "SMCODIGO = '" & vCodArt & "' AND SMMESPRO ='" & AnioMesAnterior(cAnoMes) & "' AND SMALMA='" & almacen & "'"

If Not Adodc2.EOF Then
   nSaldo = IIf(IsNull(Adodc2!SMSALDOINI), 0 + (Adodc2!SMCANENT - Adodc2!SMCANSAL), Adodc2!SMSALDOINI) + (Adodc2!SMCANENT - Adodc2!SMCANSAL)
   nSaldoInicial = nSaldo
   nCosPro = Adodc2("SMMNPREUNI")
   nCosProUS = Adodc2("SMUSPREUNI")
   VALANTE = nCosPro * nSaldo
   VALANTEUS = nCosProUS * nSaldo
Else
   nSaldoInicial = 0
   nSaldo = 0: nCosPro = 0: nCosProUS = 0
   VALANTE = 0
   VALANTEUS = 0
End If

Do While Not adodc1.EOF
    
   If Year(adodc1("CAFECDOC")) <> nYear Or Month(adodc1("CAFECDOC")) <> nMes Then
  
      cMesPro = Format(nYear, "0000") & Format(nMes, "00")
      If cArmes = cMesPro Then
         Rsql1 = "INSERT INTO  MoResMes (SMALMA,SMCODIGO,SMMESPRO,SMSALDOINI,SMCANENT,SMCANSAL,SMMNPREUNI,SMUSPREUNI)VALUES('" & almacen & "','" & vCodArt & "','" & cMesPro & "'," & nSaldoInicial & "," & nIng & "," & nSal & "," & nCosPro & "," & nCosProUS & ")"
         VGCNx.BeginTrans
         VGCNx.Execute Rsql1
         VGCNx.CommitTrans
      End If
      cMesActu = (Format(Year(adodc1("CAFECDOC")), "0000") & Format(Month(adodc1("CAFECDOC")), "00"))
      nSaldoInicial = nSaldoInicial + (nIng - nSal)
      nIng = 0
      nSal = 0
      cMesAnte = AnioMesSiguiente(cMesPro)
      While cMesAnte <> cMesActu
         If cArmes = cMesAnte Then
            Rsql1 = "INSERT INTO  MoResMes (SMALMA,SMCODIGO,SMMESPRO,SMSALDOINI,SMCANENT,SMCANSAL,SMMNPREUNI,SMUSPREUNI)VALUES('" & almacen & "','" & vCodArt & "','" & cMesAnte & "'," & nSaldoInicial & "," & nIng & "," & nSal & "," & nCosPro & "," & nCosProUS & ")"
            VGCNx.BeginTrans
            VGCNx.Execute Rsql1
            VGCNx.CommitTrans
         End If
         
            cMesAnte = AnioMesSiguiente(cMesAnte)
      Wend
      '*************************************************
      dfecha = adodc1("CAFECDOC")
      nMes = Month(adodc1("CAFECDOC"))
      nYear = Year(adodc1("CAFECDOC"))
            
   Else
  
      '*************************************************
       If adodc1!CATIPCAM = 0 Or adodc1!CATIPCAM = 1 Then
             If UCase(Dir$(cRuta4)) = UCase(VGNameCont & ".MDB") Then
                TCamb = Val(Devolver_Dato(3, Format(dfecha, "DD/MM/YYYY"), "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "02", "TIPOMON_CODIGO"))
             Else
                If UCase(Dir$(cRuta2)) = "BDCOMUN.MDB" Then
                   TCamb = Val(Devolver_Dato(1, Format(dfecha, "DD/MM/YYYY"), "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "02", "TIPOMON_CODIGO"))
                Else
                   TCamb = VGTipCamb
                End If
             End If
       Else
           TCamb = adodc1!CATIPCAM
       End If
       '*************************************************
       
       nCantid = adodc1("DECANTID")
       
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
          If Round(TCamb, 3) > 0 Then
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
      
     '**************************************
     '** ESTE PROCESO GUARDARA TODOS AQUELLOS PRODUCTOS QUE ESTAN VALORIZANDOSE MAL
     '** SI EN CASO SE ENCONTRARA SALDOS NEGATIVOS EL PROGRAMA LISTARA DICHOS CASOS
     '**************************************
     
     If nSaldo < 0 And cArmes = Format(Year(adodc1("CAFECDOC")), "0000") & Format(Month(adodc1("CAFECDOC")), "00") Then
          PCount = PCount + 1
          BarraProc1.Value = PCount
          CSQL2 = "Insert Into al_Kardex_Val (COD_ART,FEC_DOC,HOR_DOC,COD_MOV,TIP_TRANSA,NUM_DOC,CAN_ART,PRE_UNIT,COS_PRO,SAL_STOCK,SER_LOT,ING_SAL)  "
          CSQL2 = CSQL2 & " Values ('" & adodc1("DECODIGO") & "',#" & Format(adodc1("CAFECDOC"), "mm/dd/yyyy") & "#,'" & adodc1("CAHORA") & "','" & adodc1("CACODMOV") & "',"
          CSQL2 = CSQL2 & "'" & adodc1("CATD") & "','" & adodc1("CANUMDOC") & "', " & adodc1("DECANTID") & ","
          If adodc1("CATIPMOV") = "I" Then
              CSQL2 = CSQL2 & "" & nPrecio & "," & nCosPro & "," & nSaldo & ",'" & IIf(IsNull(adodc1("DESERIE")) Or adodc1("DESERIE") = "", adodc1("DELOTE"), adodc1("DESERIE")) & "','I')"
          Else
              CSQL2 = CSQL2 & "" & nCosPro & "," & nCosPro & "," & nSaldo & ",'" & IIf(IsNull(adodc1("DESERIE")) Or adodc1("DESERIE") = "", adodc1("DELOTE"), adodc1("DESERIE")) & "','S')"
          End If
          nTra = 1
          VGCNx.BeginTrans
          VGCNx.Execute CSQL2
          VGCNx.CommitTrans
          nTra = 0
     End If
     '**************************************
     '**************************************
     VALANTE = nCosPro * nSaldo
     VALANTEUS = nCosProUS * nSaldo
     dfecha = adodc1("CAFECDOC")
      
     adodc1.MoveNext
   End If
   
   
Loop

If IsDate(dfecha) Then
     
     If Not Year(dfecha) = 1500 Then  '*********FECHA vACIA
        cMesPro = Format(Year(dfecha), "0000") & Format(Month(dfecha), "00")
     Else
        cMesPro = cAnoMes
     End If
        '*************************************************
        If cArmes = cMesPro Then
            Rsql1 = "INSERT INTO  MoResMes (SMALMA,SMCODIGO,SMMESPRO,SMSALDOINI,SMCANENT,SMCANSAL,SMMNPREUNI,SMUSPREUNI)VALUES('" & almacen & "','" & vCodArt & "','" & cMesPro & "'," & nSaldoInicial & "," & nIng & "," & nSal & "," & nCosPro & "," & nCosProUS & ")"
            VGCNx.BeginTrans
            VGCNx.Execute Rsql1
            VGCNx.CommitTrans
        End If
        '*************************************************
        nSaldoInicial = nSaldoInicial + (nIng - nSal)
        cMesActu = AnioMesSiguiente(Format(Year(Now), "0000") & Format(Month(Now), "00"))
        nIng = 0
        nSal = 0
        cMesAnte = AnioMesSiguiente(cMesPro)
        
        While cMesAnte <= cMesActu
           If cArmes = cMesAnte Then
              Rsql1 = "INSERT INTO  MoResMes (SMALMA,SMCODIGO,SMMESPRO,SMSALDOINI,SMCANENT,SMCANSAL,SMMNPREUNI,SMUSPREUNI)VALUES('" & almacen & "','" & vCodArt & "','" & cMesAnte & "'," & nSaldoInicial & "," & nIng & "," & nSal & "," & nCosPro & "," & nCosProUS & ")"
              VGCNx.BeginTrans
              VGCNx.Execute Rsql1
              VGCNx.CommitTrans
           End If
           cMesAnte = AnioMesSiguiente(cMesAnte)
        Wend
End If
      
      VGCNx.Execute "Update STKART SET STSKDIS=" & nSaldoInicial + (nIng - nSal) & ",STKPREPRO=" & nCosPro & " WHERE STALMA='" & almacen & "' AND STCODIGO='" & vCodArt & "'"
      
Exit Sub

ERRAR:
MsgBox Err.Description
BarraProc.Visible = False
cArticulo.Visible = False

End Sub

Sub Reporte()
  On Error GoTo Err
  Screen.MousePointer = 11
  CrystalReport1.WindowTitle = "Inv501 -- Control de Inventarios"
  CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv501.rpt"
  
  Ubi_Tab CrystalReport1
  CrystalReport1.DiscardSavedData = True
  CrystalReport1.Destination = crptToWindow
  CrystalReport1.WindowShowPrintBtn = True
  CrystalReport1.WindowShowRefreshBtn = True
  CrystalReport1.WindowShowSearchBtn = True
  CrystalReport1.WindowShowPrintSetupBtn = True
  CrystalReport1.formulas(0) = "ALMACEN = '" & UCase(Combo1.text) & "'"
  CrystalReport1.formulas(1) = "Mes= '" & UCase(Format(dFecVal, "MMMM - yyyy")) & "'"
  CrystalReport1.formulas(2) = "EMPRESA= '" & UCase(VGparametros.RucEmpresa) & "'"
  CrystalReport1.formulas(3) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
  
  If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
  
Screen.MousePointer = 1
Exit Sub
Err:
   MsgBox Err.Description, vbInformation, "Aviso"
   Screen.MousePointer = 1
   
End Sub
Function PrecioFact(ByVal arTd As String, ByVal arNumdoc As String, ByVal arCodi As String) As Double
         Set rsPrec = New ADODB.Recordset
         rsPrec.Open "Select dfprec_ven from facdet where dftd='" & arTd & "' and dfnumser+dfnumdoc='" & arNumdoc & "' and dfcodigo='" & arCodi & "'", VGCNx, adOpenForwardOnly, adLockReadOnly
         PrecioFact = IIf(Not rsPrec.EOF, rsPrec!dfprec_ven, 0)
         rsPrec.Close
End Function
'''
'''Private Sub ValorizaXArticuloMesBAKKKKKKK(ByVal vCodArt As String, ByVal arMesActu As String)
'''Dim TCamb As Double
'''Dim nSaldo As Double, nCosPro, nCosProUS As Double
'''Dim nPrecio, nPrecioUS As Double, nCantid As Double
'''Dim cMesPro, cMesActu, cMesAnte As String
'''Dim cSql1, CSQL2 As String
'''Dim nTipCam As String
'''Dim nFlagCalcu As Boolean
'''**********Roberto
'''Dim VALMOV, VALANTE, VALMOVUS, VALANTEUS As Double
'''Dim nMes, nYear As Long
'''Dim nSal, nIng, nSaldoInicial As Double
'''Dim dfecha As Date
'''On Local Error GoTo ERRAR
'''
'''nPrecio = 0: nPrecioUS = 0: nCantid = 0
'''nCosPro = 0: nCosProUS = 0
'''nSal = 0: nIng = 0
'''nSaldoInicial = 0
'''
'''Adodc2.Filter = " SMCODIGO = '" & vCodArt & "'"
'''adodc1.Filter = "Decodigo='" & vCodArt & "'"
'''
'''If Not adodc1.EOF Then adodc1.MoveFirst
'''
'''
'''If Not adodc1.EOF Then
'''   adodc1.MoveFirst
'''   nMes = Month(adodc1("CAFECDOC"))
'''   nYear = Year(adodc1("CAFECDOC"))
'''   dfecha = adodc1("CAFECDOC")
'''Else
'''   dfecha = Format(dFecVal.Value, "dd/mm/yyyy")
'''   nMes = dFecVal.Month
'''   nYear = dFecVal.Year
'''End If
'''
'''
'''If Not Adodc2.EOF Then
'''   nSaldoInicial = Adodc2!SMSALDOINI + (Adodc2!SMCANENT - Adodc2!SMCANSAL)
'''   nSaldo = Adodc2!SMSALDOINI + (Adodc2!SMCANENT - Adodc2!SMCANSAL)
'''   nCosPro = Adodc2("SMMNPREUNI")
'''   nCosProUS = Adodc2("SMUSPREUNI")
'''Else
'''   nSaldo = 0: nCosPro = 0
'''   nSaldoInicial = 0
'''End If
'''
'''nFlagCalcu = False
'''
'''Do While Not adodc1.EOF
'''
'''   *************************************************
'''    If adodc1!CATIPCAM = 0 Or adodc1!CATIPCAM = 1 Then
'''          If UCase(Dir$(cRuta4)) = UCase(VGNameCont & ".MDB") Then
'''             TCamb = Val(Devolver_Dato(3, Format(dfecha, "DD/MM/YYYY"), "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "02", "TIPOMON_CODIGO"))
'''          Else
'''             If UCase(Dir$(cRuta2)) = "BDCOMUN.MDB" Then
'''                TCamb = Val(Devolver_Dato(1, Format(dfecha, "DD/MM/YYYY"), "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "02", "TIPOMON_CODIGO"))
'''             Else
'''                TCamb = VGTipCamb
'''             End If
'''          End If
'''    Else
'''        TCamb = adodc1!CATIPCAM
'''    End If
'''    *************************************************
'''
'''     nCantid = adodc1("DECANTID")
'''     nPrecio = adodc1("DEPRECIO")
'''
'''     If Round(TCamb, 3) > 0 Then
'''         nPrecioUS = (adodc1("DEPRECIO") / TCamb)
'''     Else
'''         nPrecioUS = 0
'''     End If
'''
'''
'''     If adodc1("CATIPMOV") = "I" Then
'''        nSaldo = nSaldo + nCantid
'''        VALMOV = nCantid * nPrecio
'''        VALMOVUS = nCantid * nPrecioUS
'''        nIng = nIng + nCantid
'''     Else
'''        nSaldo = Round(nSaldo, 4) - Round(nCantid, 4)
'''        VALMOV = nCantid * nCosPro
'''        VALMOVUS = nCantid * nCosProUS
'''        nSal = nSal + nCantid
'''     End If
'''
'''     If adodc1("CATIPMOV") = "I" Then
'''        If nSaldo <> 0 Then
'''           nCosPro = (VALMOV + VALANTE) / nSaldo
'''           nCosProUS = (VALMOVUS + VALANTEUS) / nSaldo
'''        End If
'''    End If
'''
'''    **************************************
'''    ** ESTE PROCESO GUARDARA TODOS AQUELLOS PRODUCTOS QUE ESTAN VALORIZANDOSE MAL
'''    ** SI EN CASO SE ENCONTRARA SALDOS NEGATIVOS EL PROGRAMA LISTARA DICHOS CASOS
'''    **************************************
'''
'''    If nSaldo < 0 Then
'''         PCount = PCount + 5
'''         BarraProc1.Value = PCount
'''         CSQL2 = "Insert Into al_Kardex_Val (COD_ART,FEC_DOC,HOR_DOC,COD_MOV,TIP_TRANSA,NUM_DOC,CAN_ART,PRE_UNIT,COS_PRO,SAL_STOCK,SER_LOT,ING_SAL)  "
'''         CSQL2 = CSQL2 & " Values ('" & adodc1("DECODIGO") & "',#" & Format(adodc1("CAFECDOC"), "mm/dd/yyyy") & "#,'" & adodc1("CAHORA") & "','" & adodc1("CACODMOV") & "',"
'''         CSQL2 = CSQL2 & "'" & adodc1("CATD") & "','" & adodc1("CANUMDOC") & "', " & adodc1("DECANTID") & ","
'''         If adodc1("CATIPMOV") = "I" Then
'''             CSQL2 = CSQL2 & "" & nPrecio & "," & nCosPro & "," & nSaldo & ",'" & IIf(IsNull(adodc1("DESERIE")) Or adodc1("DESERIE") = "", adodc1("DELOTE"), adodc1("DESERIE")) & "','I')"
'''         Else
'''             CSQL2 = CSQL2 & "" & nCosPro & "," & nCosPro & "," & nSaldo & ",'" & IIf(IsNull(adodc1("DESERIE")) Or adodc1("DESERIE") = "", adodc1("DELOTE"), adodc1("DESERIE")) & "','S')"
'''         End If
'''         nTra = 1
'''         cConexAux.BeginTrans
'''         cConexAux.Execute CSQL2
'''         cConexAux.CommitTrans
'''         nTra = 0
'''    End If
'''    **************************************
'''
'''    VALANTE = nCosPro * nSaldo
'''    VALANTEUS = nCosProUS * nSaldo
'''
'''    dfecha = adodc1("CAFECDOC")
'''
'''    adodc1.MoveNext
'''
'''    nFlagCalcu = True
'''Loop
'''
'''     If (Not Adodc2.EOF) Or nFlagCalcu = True Then
'''        cMesPro = Format(Year(dfecha), "0000") & Format(Month(dfecha), "00")
'''        *************************************************
'''        cSql1 = "INSERT INTO  MoResMes (SMALMA,SMCODIGO,SMMESPRO,SMSALDOINI,SMCANENT,SMCANSAL,SMMNPREUNI,SMUSPREUNI)VALUES('" & almacen & "','" & vCodArt & "','" & arMesActu & "'," & nSaldoInicial & "," & nIng & "," & nSal & "," & nCosPro & "," & nCosProUS & ")"
'''        Vgcnx.BeginTrans
'''        Vgcnx.Execute cSql1
'''        Vgcnx.CommitTrans
'''        *************************************************
'''        nFlagCalcu = False
'''     End If
'''
'''
'''       Vgcnx.Execute "Update STKART SET STSKDIS=" & nSaldoInicial & ",STKPREPRO=" & nCosPro & " WHERE STALMA='" & almacen & "' AND STCODIGO='" & vCodArt & "'"
'''
'''Exit Sub
'''
'''ERRAR:
''' MsgBox err.Description
''' BarraProc.Visible = False
''' cArticulo.Visible = False
''' BarraProc1.Visible = False
''' Label2(0).Visible = False
''' Label2(1).Visible = False
'''
''' Resume
'''End Sub

