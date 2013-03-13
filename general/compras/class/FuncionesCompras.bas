Attribute VB_Name = "FuncionesCompras"
Option Explicit
Dim ndiaMes(12) As Integer
Dim cdesMes(12) As String
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Sub CentrarForm(nFormPrin As Form, nFormu As Form)
  nFormu.Left = (nFormPrin.Width - nFormu.Width) / 2
  nFormu.Top = ((nFormPrin.Height - nFormu.Height) / 2) - 600
End Sub
Sub Impresion(cNombreReporte As String)
On Error GoTo X
  With MDIPrincipal.cryRpt
        .Reset
        .Destination = crptToWindow
        .WindowState = crptMaximized
         If Right(VGParamSistem.RutaReport, 1) = "\" Then
            .ReportFileName = VGParamSistem.RutaReport & cNombreReporte
        Else
            .ReportFileName = VGParamSistem.RutaReport & "\" & cNombreReporte
        End If
        If VGsql = 1 Then
           .Connect = "Provider=SQLOLEDB;Password=" & VGParamSistem.PwdGEN & ";Persist Security Info=True;User ID=" & VGParamSistem.UsuarioGEN & ";Initial Catalog=" & VGParamSistem.BDEmpresaGEN & ";SERVER=" & VGParamSistem.ServidorGEN
          Else
           .Connect = VGCadenaReport2
         End If
        .DiscardSavedData = True
        If cNombreReporte = "rptTipoCambio.rpt" Then
           Set VGvardllgen = New dllgeneral.dll_general
           .formulas(0) = "@Mes=" & CInt(VGParamSistem.Mesproceso)
           .formulas(1) = "@Mesanno='" & VGvardllgen.DesMes(VGParamSistem.Mesproceso) & "-" & VGParamSistem.Anoproceso & "'"
           .formulas(2) = "@Anual='" & VGParamSistem.Anoproceso & "'"
        End If
        .Action = 1
  End With
  Exit Sub
X:
  MsgBox "Error inesperado: " & err.Number & "  " & err.Description, vbExclamation
End Sub
Public Sub CargarParametrosCompras()

Dim rsaux As ADODB.Recordset
    
    Set rsaux = New ADODB.Recordset
    rsaux.Open "co_sistema", VGCNx
    If rsaux.RecordCount = 0 Then Exit Sub
    Set VGvardllgen = New dllgeneral.dll_general
    
    VGParametros.monedabase = Trim(rsaux!monedacodigo)
    VGParametros.NomEmpresa = Trim(rsaux!sistemadescripcionempresa)
    VGParametros.direccionempresa = Trim(rsaux!sistemadireccionempresa)
    VGParametros.RucEmpresa = Trim(rsaux!sistemaempresaruc)
    VGParametros.ctascompra = ArmaCriterioComodin(rsaux!sistemactacomp, "cuentacodigo")
    VGParametros.Igv = rsaux!sistemaigv / 100
    
    'Parametros Exclusivos para la generacion de asientos a contabilidad
    
    VGParametros.xLibro = VGvardllgen.ESNULO(rsaux!sistemalibro, "")
    VGParametros.xTipAnal = VGvardllgen.ESNULO(rsaux!sistematipanal, "00")
    VGParametros.xsubasiento = VGvardllgen.ESNULO(rsaux!sistemasubasiento, "00")
    VGParametros.xCtaIGV = VGvardllgen.ESNULO(rsaux!sistemactaIGV, "00")
    VGParametros.xCtaIES = VGvardllgen.ESNULO(rsaux!sistemactaIES, "00")
    VGParametros.xCtaRTA = VGvardllgen.ESNULO(rsaux!sistemactaRTA, "00")
    VGParametros.Auxaut = True ' Se tiene que crear el campo para controlar auxiliar automatico
    VGParametros.xCodPercepcion = VGvardllgen.ESNULO(rsaux!codigopercepcion, "00")  'Para controlar percepciones
    
    'Cargar parametros para pasar a cuentas por cobrar
    
    VGParametros.CpTiplan = VGvardllgen.ESNULO(rsaux!sistematipoplan, "00")
    VGParametros.CpOficina = VGvardllgen.ESNULO(rsaux!sistemaoficina, "00")
    
    VGParametros.xCtaTotal = rsaux!sistemactatotal
    VGParametros.permite_tc = IIf(VGvardllgen.ESNULO(rsaux!permite_tc, 0) = 0, False, True)
    VGParametros.sistemaactivaccostos = IIf(VGvardllgen.ESNULO(rsaux!sistemaactivaccostos, 0) = 0, False, True)
    VGParametros.sistemaasientoenlinea = IIf(VGvardllgen.ESNULO(rsaux!sistemaasientoenlinea, 0) = 0, False, True)
    VGParametros.sistemactrlgastos = IIf(VGvardllgen.ESNULO(rsaux!sistemactrlgastos, 0) = 0, False, True)
    
    If ExisteElem(1, VGCNx, "co_sistema", "sistemamultiempresas") Then
       VGParametros.sistemamultiempresas = IIf(VGvardllgen.ESNULO(rsaux!sistemamultiempresas, 0) = 0, False, True)
    End If
    VGParametros.minimoretencion = IIf(VGvardllgen.ESNULO(rsaux!sistemaminimoretencion, 0) = 0, 99999, rsaux!sistemaminimoretencion)
    VGParametros.sistemabancarizacion = IIf(VGvardllgen.ESNULO(rsaux!bancarizacion, 0) = 0, 0, rsaux!bancarizacion)
    VGParametros.sistemabancarizacion01 = IIf(VGvardllgen.ESNULO(rsaux!minimobancarizacion01, 0) = 0, 9999999, rsaux!minimobancarizacion01)
    VGParametros.sistemabancarizacion02 = IIf(VGvardllgen.ESNULO(rsaux!minimobancarizacion02, 0) = 0, 9999999, rsaux!minimobancarizacion02)
    
    VGParametros.controlaestadosrendicion = IIf(VGvardllgen.ESNULO(rsaux!controlaestadosrendicion, 0) = 0, 0, rsaux!controlaestadosrendicion)
    VGParametros.diasatrazorendicion = IIf(VGvardllgen.ESNULO(rsaux!diasatrazorendicion, 0) = 0, 0, rsaux!diasatrazorendicion)
    VGParametros.diacierrerendicion = IIf(VGvardllgen.ESNULO(rsaux!diacierrerendicion, 0) = 0, 1, rsaux!diacierrerendicion)
    VGParametros.numeracionautomaticalibro = IIf(VGvardllgen.ESNULO(rsaux!numeracionautomaticalibro, 0) = 0, 0, rsaux!numeracionautomaticalibro)
   
    MDIPrincipal.Caption = "Sistema de Provision de Compras - " & Trim(rsaux!sistemadescripcionempresa)
    MDIPrincipal.StatusBar1.Panels(5).Text = "Servidor (" & VGParamSistem.Servidor & ")"
    MDIPrincipal.StatusBar1.Panels(6).Text = "Base de Datos (" & VGParamSistem.BDEmpresa & ")"
    
    Set rsaux = New ADODB.Recordset
    rsaux.Open "select sistemaultimonivel,sistemaultimonivelcostos from  ct_sistema", VGcnxCT, adOpenKeyset, adLockReadOnly
    If rsaux.RecordCount = 0 Then Exit Sub
    VGnumniveles = rsaux!sistemaultimonivel
    VGnumnivcos = ESNULO(rsaux!sistemaultimonivelcostos, 1)
    
    Set rsaux = New ADODB.Recordset
    rsaux.Open "select sistemaultimonivel from  co_sistema", VGCNx, adOpenKeyset, adLockReadOnly
    If rsaux.RecordCount = 0 Then Exit Sub
    VGnumnivgas = rsaux!sistemaultimonivel
    
    Set rsaux = New ADODB.Recordset
    Set rsaux = VGCNx.Execute("select *  from  vt_sistema")
    If rsaux.RecordCount = 0 Then Exit Sub
    VGParamSistem.tipoanaliticocodigo = rsaux!tipoanaliticocodigo
    VGParamSistem.familiaproyectos = rsaux!familiaproyectos
     
    Set rsaux = New ADODB.Recordset
    Set rsaux = VGcnxCT.Execute("select top 1 sistemactaajustedeb,sistemactaajustehab,asientoAutoCCostos,cuentacodigoCostos from  ct_sistema")
    If rsaux.RecordCount = 0 Then Exit Sub
    VGParametros.sistemactaajustedeb = RTrim(rsaux!sistemactaajustedeb)
    VGParametros.sistemactaajustehab = RTrim(rsaux!sistemactaajustehab)
    VGParametros.AsientoAutoxCCostos = ESNULO(RTrim(rsaux!asientoAutoCCostos), 0)
    VGParametros.cuentadeCostos = ESNULO(rsaux!cuentacodigoCostos, "40100")
    
        
   If IsNumeric(VGParamSistem.Anoproceso) And IsNumeric(VGParamSistem.Mesproceso) Then
        SQL = "select * from ct_cierremensual where empresacodigo='" & VGParametros.empresacodigo & "' and " _
        & " anio='" & VGParamSistem.Anoproceso & "' and mes=" & Trim(VGParamSistem.Mesproceso) & " "
        Set rsaux = VGCNx.Execute(SQL)
        If rsaux.RecordCount > 0 Then VGParametros.cierremes = IIf(rsaux!compras = True, True, False)
        If VGtipolicencia = "T" Then
           If VGfechalicencia < VGParamSistem.FechaTrabajo Then
              VGParametros.cierremes = True
              MsgBox ("error en la tabla de tipo de cambio, comunicarse con sistemas  ")
              Exit Sub
           End If
        End If
    End If

    
    
End Sub
Public Function Coversion(MonOrigi As String, VCambio As Double, monto As Double)
'FCP
Dim valor As Double
On Error GoTo errtext
    Coversion = 0
    If MonOrigi = VGParametros.monedabase Then
        valor = monto / VCambio ' Soles ==> a Dolares
      Else
        valor = monto * VCambio ' Dolares ==> a Soles
    End If
    Coversion = Round(valor, 2)
    Exit Function
errtext:
    Coversion = 0
End Function
Public Sub HabilitarDetalle(flag As Boolean, framex As Frame, formx As Form)
'FCP
On Error Resume Next
'If Not flag Then flag = VGParametros.cierremes
framex.Enabled = flag
Dim Control As Control
    For Each Control In formx.Controls
        If UCase(Control.Container.Name) = UCase(framex.Name) Then
            Control.Enabled = flag
End If
    Next
End Sub

Public Sub Parametrogastos()
 Dim rs As ADODB.Recordset
 Dim cuenta As String
 Dim i As Integer
 Dim J As Integer
 Dim num As Integer
 
Set rs = New ADODB.Recordset
Set rs = VGCNx.Execute("SELECT sistemaconfiguragastos FROM co_sistema")
    
If Not (rs.BOF Or rs.EOF) Then
   cuenta = Trim(rs(0))
   For i = 1 To Len(cuenta)
       If Mid(cuenta, i, 1) = "*" Then num = num + 1
   Next
   ReDim VG_gNIVELES(Len(cuenta) - num)
   J = 0
   For i = 1 To Len(cuenta) Step 2
       VG_gNIVELES(J) = Mid(cuenta, i, 1)
       J = J + 1
   Next
   VGnumnivgas = Len(cuenta) - num
End If
Set rs = Nothing
End Sub

Public Sub ClearControlsInframe(framex As Frame, formx As Form)
'FCP
On Error Resume Next
    Dim Control As Control
    For Each Control In formx.Controls
        If UCase(Control.Container.Name) = UCase(framex.Name) Then
            If UCase(Left(Control.Name, 2)) <> "LE" Then
                If TypeOf Control Is TextBox Then Control.Text = ""
                If TypeOf Control Is TextFer.TxFer Then Control.Text = ""
                If TypeOf Control Is Label Then Control.Caption = ""
                'If TypeOf Control Is DTPicker Then Control.Value = Date
            End If
        End If
    Next
End Sub
Public Sub EjecutarLote(RichTextBox1 As RichTextBox, cnx As ADODB.Connection)
'Funcion Creada por fernando cossio
'Ejecuta scrip de lotes generadas en la secuencia de comandos del SQL
Dim pos As Long, ini As Long
Dim i As Integer
Dim cad As String
Dim Cont As Long, longi As Long
Dim conpos As Long, sqlcad As String
    pos = 1
    ini = 1
    longi = Len(RichTextBox1.Text)
    Do While pos <> 0
        pos = InStr(pos + 2, RichTextBox1.Text, "GO", vbTextCompare)
        sqlcad = ""
        If pos + 2 > longi Then Exit Do
        If pos = 0 Then Exit Do
        If Asc(Mid(RichTextBox1.Text, pos - 1, 1)) = 10 And Asc(Mid(RichTextBox1.Text, pos + 2, 1)) = 13 Then
            Cont = Cont + 1
            sqlcad = Mid(RichTextBox1.Text, ini, pos - (ini + 2))
            RichTextBox1.SelStart = pos: RichTextBox1.SelLength = 2
            ini = pos + 2
            cnx.Execute sqlcad
        End If
    Loop
End Sub

Public Function Espunto(ByRef texto As Variant) As Variant
    If Trim(texto) = "." Or Trim(texto) = "-" Then
        Espunto = "0"
      Else
        Espunto = texto
    End If
End Function

Public Sub ModoEditable(flagModo As Boolean, Formu As Form, cNameCtrX As String)
 Dim i As Integer
    Dim Control As Control
    For Each Control In Formu.Controls
       If UCase(Control.Name) <> UCase(cNameCtrX) Then
           If TypeOf Control Is TextBox Then Control.Enabled = flagModo
           If TypeOf Control Is TextFer.TxFer Then Control.Enabled = flagModo
           If TypeOf Control Is CheckBox Then Control.Enabled = flagModo
           If TypeOf Control Is ctrlayuda_f.Ctr_Ayuda Then Control.Enabled = flagModo
       End If
    Next
End Sub

Public Sub LimpiarForm(Formu As Form, cNameCtrX As String)
 Dim i As Integer
    Dim Control As Control
    For Each Control In Formu.Controls
       If UCase(Control.Name) <> UCase(cNameCtrX) Then
           If TypeOf Control Is TextBox Then Control.Text = Empty
           If TypeOf Control Is TextFer.TxFer Then Control.Text = Empty
           If TypeOf Control Is CheckBox Then Control.Value = 0
           If TypeOf Control Is ctrlayuda_f.Ctr_Ayuda Then
              Control.xclave = Empty
              Control.xnombre = Empty
           End If
       End If
    Next
End Sub

Public Function GeneraCodigo(Conex As ADODB.Connection, csql As String, cNumCeros As String) As String
 Dim rsX As ADODB.Recordset
 Set rsX = New ADODB.Recordset
 Set rsX = Conex.Execute(csql)
 GeneraCodigo = Format(Val(IIf(IsNull(rsX(0)), 0, rsX(0))) + 1, cNumCeros)
 Set rsX = Nothing
End Function

Public Function ValidaAsientos() As Boolean
 Dim SQL As String
    SQL = "SELECT ct_asiento.asientocodigo as Código, ct_asiento.asientodescripcion as Descripción "
    SQL = SQL & "FROM ct_asiento "
    SQL = SQL & "WHERE ct_asiento.asientocodigo<>'00'"
   
    Set VGvardllgen = New dllgeneral.dll_general
    If VGvardllgen.VerificaDatoExistente(VGCNx, SQL) <= 0 Then
        ValidaAsientos = False
        MsgBox "Faltan Registrar los Asientos por la Opción correspondiente", vbInformation, "Sistema Contable"
    Else
        ValidaAsientos = True
    End If
    Set VGvardllgen = Nothing
End Function

Public Function ValidaSubAsientos(xCodAsiento As String) As Boolean
 Dim SQL As String
    SQL = "SELECT subasientocodigo FROM ct_subasiento WHERE subasientocodigo<>'00' "
    SQL = SQL & "AND asientocodigo like '" & xCodAsiento & "%'"
    
    Set VGvardllgen = New dllgeneral.dll_general
    If VGvardllgen.VerificaDatoExistente(VGCNx, SQL) <= 0 Then
        ValidaSubAsientos = False
        MsgBox "Faltan registrar los SubAsientos que corresponden al Asiento Nº " & xCodAsiento, vbInformation, "Sistema Contable"
    Else
        ValidaSubAsientos = True
    End If
    Set VGvardllgen = Nothing
End Function

Public Sub CancelaDocumentos()
    On Error GoTo Mayor
    Dim X As Long
    Screen.MousePointer = 11
    VGCNx.BeginTrans
    Set VGCommandoSP = New ADODB.Command
    VGCommandoSP.ActiveConnection = VGGeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "ct_ProcCanc_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@Base") = VGParamSistem.BDEmpresa
        .Parameters("@Ano") = VGParamSistem.Anoproceso
        .Parameters("@Mes") = VGParamSistem.Mesproceso
        .Execute X
    End With
    VGCNx.CommitTrans
    Screen.MousePointer = 1
    MsgBox "Se Cancelo los Documentos Satisfactoriamente  " & Chr(13) & _
           "Items Afectados ", vbInformation
    Exit Sub
Mayor:
    Screen.MousePointer = 1
    VGCNx.RollbackTrans
    MsgBox "No se pudo Cancelar los Documentos " & Chr(13) & err.Description, vbExclamation
End Sub
Private Function ArmaCriterioComodin(cad As String, Campo As String) As String
Dim pos As Integer, cadaux As String, criterio As String
Dim valor As String
    criterio = ""
    Do While True
        pos = InStr(1, cad, "%", vbTextCompare)
        If pos = 0 Then Exit Do
        valor = "'" & Left(cad, pos) & "'"
        cad = Right(cad, (Len(cad) - pos))
        criterio = criterio & Campo & " like " & valor & " or "
    Loop
    ArmaCriterioComodin = Left(criterio, Len(criterio) - 3)
End Function
Public Function FechS(Fecha As Variant, Tipo As TIPFECHA) As Variant
Dim H As Date
Dim fechaAux As Double
On Error GoTo ErrorFecha
   H = CDate(Fecha)
   Select Case Tipo
      Case Sqlf: 'Para transformar al sql
        fechaAux = DateSerial(Year(Fecha), Month(Fecha), Day(Fecha)) - 2
      Case Adof: 'Para transformar al ado Y AL ACCESS
         fechaAux = DateSerial(Year(Fecha), Month(Fecha), Day(Fecha))
   End Select
   FechS = fechaAux
   Exit Function
ErrorFecha:
   Select Case Tipo
      Case Sqlf: FechS = "Null"
      Case Adof: FechS = Null
   End Select
End Function

Public Sub ParametroCuentagastos()
 Dim rs As ADODB.Recordset
 Dim cuenta As String
 Dim i As Integer
 Dim J As Integer
 Dim num As Integer
 
    Set rs = New ADODB.Recordset
    Set rs = VGCNx.Execute("SELECT sistemaconfiguragastos FROM co_sistema")
    
    If Not (rs.BOF Or rs.EOF) Then
        cuenta = Trim(rs(0))
        For i = 1 To Len(cuenta)
            If Mid(cuenta, i, 1) = "*" Then num = num + 1
        Next
        ReDim VG_gNIVELES(Len(cuenta) - num)
        J = 0
        For i = 1 To Len(cuenta) Step 2
            VG_gNIVELES(J) = Mid(cuenta, i, 1)
            J = J + 1
        Next
        VGnumnivgas = Len(cuenta) - num
    End If
    Set rs = Nothing
End Sub

Public Function UltNumeroAuto(Tabla As String, op As String, cnx As ADODB.Connection) As Long
Dim rsaux As ADODB.Recordset
On Error GoTo errornum
    Set rsaux = New ADODB.Recordset
    Select Case op
        Case 1
'            rsaux.Open "SELECT Numx=isnull(IDENT_CURRENT('" & TABLA & "'),0)", cnx, adOpenKeyset, adLockReadOnly
            rsaux.Open "SELECT top 1 Numx=isnull(cabprovinumero,1) from co_sistema ", cnx, adOpenKeyset, adLockReadOnly
    End Select
    If rsaux.EOF Or rsaux.BOF Then
      UltNumeroAuto = 1
      Exit Function
    Else
      UltNumeroAuto = rsaux!Numx
      Set rsaux = New ADODB.Recordset
    End If
    Exit Function
errornum:
    UltNumeroAuto = -1
End Function
Public Sub adicionacamposCO()
On Error GoTo err
If Not ExisteElem(1, VGCNx, "co_cabprovi" & VGParamSistem.Anoproceso, "empresacodigo") Then
        VGCNx.Execute "ALTER TABLE co_cabprovi" & VGParamSistem.Anoproceso & " ADD empresacodigo VARCHAR(2) NULL"
End If
If Not ExisteElem(1, VGCNx, "co_detprovi" & VGParamSistem.Anoproceso, "entidadcodigo") Then
        VGCNx.Execute "ALTER TABLE co_detprovi" & VGParamSistem.Anoproceso & " ADD entidadcodigo VARCHAR(11) NULL"
End If
Exit Sub
err:
Resume Next
End Sub

