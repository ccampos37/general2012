Attribute VB_Name = "ModFunciones"
Option Explicit
Dim ndiaMes(12) As Integer
Dim cdesMes(12) As String
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Enum TIPFECHA
   Sqlf = 1
   Adof = 2
End Enum

Public Function Devolver_Dato(Tipo As Integer, Cod As String, Tabla As String, Campo As String, Fecha As Boolean, CampDev As String, Optional Cod2 As String, Optional Campo2 As String, Optional Cod3 As String, Optional Campo3 As String, Optional Cod4 As Double, Optional Campo4 As String) As String
Dim cSel1 As ADODB.Recordset, cF As String
Set cSel1 = New ADODB.Recordset

If Trim(Campo) <> "" Then
    If Fecha = False Then
        cF = "Select " & CampDev & " from " & Tabla & "  Where " & Campo & " =  '" & Cod & "' "
    Else
        cF = "Select " & CampDev & " from " & Tabla & "  Where " & Campo & " =  #" & Format(Cod, "mm/dd/yyyy") & "#"
    End If
End If
If Trim(Campo2) <> "" Then
    cF = cF & " and " & Campo2 & " = '" & Cod2 & "' "
End If
If Trim(Campo3) <> "" Then
    cF = cF & " and " & Campo3 & " = '" & Cod3 & "' "
End If
If Trim(Campo4) <> "" Then
    cF = cF & " and " & Campo4 & " = '" & Cod4 & "' "
End If
Select Case Tipo
  Case 1 'Bd. Comun
              cSel1.Open cF, VGcnx, adOpenStatic
  Case 2 'Bd. Config
              cSel1.Open cF, VGcnx, adOpenStatic
  Case 3 'Bd. Contabilidad
              cSel1.Open cF, VGcnxCT, adOpenStatic
End Select

If cSel1.RecordCount > 0 Then
     Devolver_Dato = IIf(Not IsNull(cSel1(0)), cSel1(0), "")
Else
     Devolver_Dato = ""
End If
End Function

Public Property Get ComputerName() As Variant
    Dim sName As String
    Dim iRetVal As Long
    Dim ipos As Integer
    sName = Space$(255)
    iRetVal = GetComputerName(sName, 255&)
    If iRetVal = 0 Then
      ComputerName = ""
      Exit Property
    End If
    ipos = InStr(sName, Chr$(0))
    ComputerName = Left$(sName, ipos - 1)
End Property
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
        .ReportFileName = App.Path & "\" & VGParamSistem.carpetareportes & "\" & cNombreReporte
        .LogOnServer "pdssql.dll", VGParamSistem.Servidor, VGParamSistem.BDEmpresa, VGParamSistem.UsuarioCT, VGParamSistem.PWD
        .Connect = vgCADENAREPORT2
        .DiscardSavedData = True
        
        If cNombreReporte = "rptTipoCambio.rpt" Then
           Set VGvardllgen = New dllgeneral.dll_general
           .formulas(0) = "@Mes=" & CInt(VGParamSistem.Mesproceso)
           .formulas(1) = "@Mesanno='" & VGvardllgen.DESMES(VGParamSistem.Mesproceso) & "-" & VGParamSistem.Anoproceso & "'"
           .formulas(2) = "@Anual='" & VGParamSistem.Anoproceso & "'"
        End If
        .Action = 1
  End With
  Exit Sub
X:
  MsgBox "Error inesperado: " & Err.Number & "  " & Err.Description, vbExclamation
End Sub
'Sub Impresion(cNombreReporte As String)
'On Error GoTo X
'  With MDIPrincipal.cryRpt
'        .Reset
'        .Destination = crptToWindow
'        .WindowState = crptMaximized
'        .ReportFileName = App.Path & "\" & VGParamSistem.carpetareportes & "\" & cNombreReporte
'        '.LogOnServer "pdssql.dll", VGParamSistem.Servidor, VGParamSistem.BDEmpresa, "sa", ""
'        .LogOnServer "pdssql.dll", VGParamSistem.Servidor, VGParamSistem.BDEmpresa, VGParamSistem.Usuario, VGParamSistem.Pwd
'        .Connect = vgCADENAREPORT
'        .DiscardSavedData = True
'
'        If cNombreReporte = "rptTipoCambio.rpt" Then
'          Set VGvardllgen = New dllgeneral.dll_general
'          .Formulas(0) = "@Mes=" & CInt(VGParamSistem.Mesproceso)
'          .Formulas(1) = "@Mesanno='" & VGvardllgen.DESMES(VGParamSistem.Mesproceso) & "-" & VGParamSistem.Anoproceso & "'"
'        End If
'
'        .Action = 1
'  End With
'  Exit Sub
'X:
'  MsgBox "Error inesperado: " & Err.Number & "  " & Err.Description, vbExclamation
'End Sub
Sub ImpresionRptProc(cNombreReporte As String, PFormulas(), Param(), Optional orden As String, Optional Titulo As String)
Dim I As Integer
On Error GoTo X
    Screen.MousePointer = 11
    With MDIPrincipal.CryRptProc
        .Reset
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowTitle = Titulo
        Call PropCrystal(MDIPrincipal.CryRptProc)
        .ReportFileName = VGParamSistem.RutaReport & "\" & VGParamSistem.carpetareportes & "\" & cNombreReporte
        .LogOnServer "pdssql.dll", VGParamSistem.ServidorGEN, VGParamSistem.BDEmpresaGEN, VGParamSistem.UsuarioGEN, ""
        .Connect = vgCADENAREPORT2
        .formulas(0) = "@Emp='" & VGParamCompra.NomEmpresa & "'"
        .formulas(1) = "@Ruc='" & VGParamCompra.RucEmpresa & "'"
        If UBound(PFormulas) > 0 Then
            For I = 0 To UBound(PFormulas) - 1
                .formulas(2 + I) = PFormulas(I)
            Next
        End If
        .DiscardSavedData = True
        If UBound(Param) > 0 Then
            For I = 0 To UBound(Param) - 1
                .StoredProcParam(I) = Param(I)
            Next
        End If
        If orden <> "" Then Call CrystOrden(MDIPrincipal.CryRptProc, orden)
        If .Status <> 2 Then .Action = 1
    End With
    Screen.MousePointer = 1
    Exit Sub
X:
  If Err.Number = 9 Then Resume Next
  Screen.MousePointer = 1
  MsgBox "Error inesperado: " & Err.Number & "  " & Err.Description, vbExclamation
End Sub
Sub ImpresionRptbase(cNombreReporte As String, PFormulas(), Param(), Optional orden As String, Optional Titulo As String)
Dim I As Integer
On Error GoTo X
    Screen.MousePointer = 11
    With MDIPrincipal.CryRptProc
        .Reset
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowTitle = Titulo
        Call PropCrystal(MDIPrincipal.CryRptProc)
        .ReportFileName = VGParamSistem.RutaReport & "\" & cNombreReporte
        .LogOnServer "pdssql.dll", VGParamSistem.Servidor, VGcnx.DefaultDatabase, "sa", ""
        .Connect = vgCADENAREPORT2
        .formulas(0) = "@Emp='" & VGParamCompra.NomEmpresa & "'"
        .formulas(1) = "@Ruc='" & VGParamCompra.RucEmpresa & "'"
        If UBound(PFormulas) > 0 Then
            For I = 0 To UBound(PFormulas) - 1
                .formulas(2 + I) = PFormulas(I)
            Next
        End If
        .DiscardSavedData = True
        If UBound(Param) > 0 Then
            For I = 0 To UBound(Param) - 1
                .StoredProcParam(I) = Param(I)
            Next
        End If
        If orden <> "" Then Call CrystOrden(MDIPrincipal.CryRptProc, orden)
        If .Status <> 2 Then .Action = 1
    End With
    Screen.MousePointer = 1
    Exit Sub
X:
  If Err.Number = 9 Then Resume Next
  Screen.MousePointer = 1
  MsgBox "Error inesperado: " & Err.Number & "  " & Err.Description, vbExclamation
End Sub
Private Sub CrystOrden(ByRef cry As CrystalReport, cad As String)
Dim pos As Integer, cadaux As String, I As Integer
Dim Valor As String
    Do While True
        pos = InStr(1, cad, ",", vbTextCompare)
        I = 0
        If pos = 0 Then Exit Do
        Valor = Left(cad, pos - 1)
        cry.SortFields(I) = Valor
        I = I + 1
        cad = Right(cad, (Len(cad) - pos))
    Loop
End Sub
Public Sub CargarParametrosCompras()

Dim rsaux As ADODB.Recordset
    
    Set rsaux = New ADODB.Recordset
    rsaux.Open "co_sistema", VGcnx
    If rsaux.RecordCount = 0 Then Exit Sub
    Set VGvardllgen = New dllgeneral.dll_general
    
    VGParamCompra.monedabase = "01"
    VGParamCompra.NomEmpresa = Trim(rsaux!sistemadescripcionempresa)
    VGParamCompra.direccionempresa = Trim(rsaux!sistemadireccionempresa)
    VGParamCompra.RucEmpresa = Trim(rsaux!sistemaempresaruc)
    VGParamCompra.ctascompra = ArmaCriterioComodin(rsaux!sistemactacomp, "cuentacodigo")
    VGParamCompra.Igv = rsaux!sistemaigv / 100
    
    'Parametros Exclusivos para la generacion de asientos a contabilidad
    
    VGParamCompra.xLibro = VGvardllgen.ESNULO(rsaux!sistemalibro, "")
    VGParamCompra.xTipAnal = VGvardllgen.ESNULO(rsaux!sistematipanal, "00")
    VGParamCompra.xsubasiento = VGvardllgen.ESNULO(rsaux!sistemasubasiento, "00")
    VGParamCompra.xCtaIGV = VGvardllgen.ESNULO(rsaux!sistemactaIGV, "00")
    VGParamCompra.xCtaIES = VGvardllgen.ESNULO(rsaux!sistemactaIES, "00")
    VGParamCompra.xCtaRTA = VGvardllgen.ESNULO(rsaux!sistemactaRTA, "00")
    VGParamCompra.Auxaut = True ' Se tiene que crear el campo para controlar auxiliar automatico
    
    'Cargar parametros para pasar a cuentas por cobrar
    
    VGParamCompra.CpTiplan = VGvardllgen.ESNULO(rsaux!sistematipoplan, "00")
    VGParamCompra.CpOficina = VGvardllgen.ESNULO(rsaux!sistemaoficina, "00")
    
    VGParamCompra.xCtaTotal = rsaux!sistemactatotal
    VGParamCompra.permite_tc = IIf(VGvardllgen.ESNULO(rsaux!permite_tc, 0) = 0, False, True)
    VGParamCompra.sistemaactivaccostos = IIf(VGvardllgen.ESNULO(rsaux!sistemaactivaccostos, 0) = 0, False, True)
    VGParamCompra.sistemaasientoenlinea = IIf(VGvardllgen.ESNULO(rsaux!sistemaasientoenlinea, 0) = 0, False, True)
    VGParamCompra.sistemactrlgastos = IIf(VGvardllgen.ESNULO(rsaux!sistemactrlgastos, 0) = 0, False, True)
    
    VGParamSistem.carpetareportes = "Reportes"
    If ExisteElem(1, VGcnx, "co_sistema", "sistemamultiempresas") Then
       VGParamCompra.sistemamultiempresas = IIf(VGvardllgen.ESNULO(rsaux!sistemamultiempresas, 0) = 0, False, True)
    End If
    MDIPrincipal.Caption = "Sistema de Provision de Compras - " & Trim(rsaux!sistemadescripcionempresa)
    MDIPrincipal.StatusBar1.Panels(5).Text = "Servidor (" & VGParamSistem.Servidor & ")"
    MDIPrincipal.StatusBar1.Panels(6).Text = "Base de Datos (" & VGParamSistem.BDEmpresa & ")"
    
    Set rsaux = New ADODB.Recordset
    rsaux.Open "select sistemaultimonivel,sistemaultimonivelcostos from  ct_sistema", VGcnxCT, adOpenKeyset, adLockReadOnly
    If rsaux.RecordCount = 0 Then Exit Sub
    VGnumniveles = rsaux!sistemaultimonivel
    VGnumnivcos = rsaux!sistemaultimonivelcostos
    
    Set rsaux = New ADODB.Recordset
    rsaux.Open "select sistemaultimonivel from  co_sistema", VGcnx, adOpenKeyset, adLockReadOnly
    If rsaux.RecordCount = 0 Then Exit Sub
    VGnumnivgas = rsaux!sistemaultimonivel
    
    
End Sub
Public Function Coversion(MonOrigi As String, VCambio As Double, monto As Double)
'FCP
Dim Valor As Double
On Error GoTo errtext
    Coversion = 0
    If MonOrigi = VGParamCompra.monedabase Then
        Valor = monto / VCambio ' Soles ==> a Dolares
      Else
        Valor = monto * VCambio ' Dolares ==> a Soles
    End If
    Coversion = Round(Valor, 2)
    Exit Function
errtext:
    Coversion = 0
End Function
Public Sub HabilitarDetalle(FLAG As Boolean, framex As Frame, formx As Form)
'FCP
On Error Resume Next
framex.Enabled = FLAG
Dim Control As Control
    For Each Control In formx.Controls
        If UCase(Control.Container.Name) = UCase(framex.Name) Then
            Control.Enabled = FLAG
        End If
    Next
End Sub

Public Sub Parametrogastos()
 Dim Rs As ADODB.Recordset
 Dim cuenta As String
 Dim I As Integer
 Dim j As Integer
 Dim num As Integer
 
Set Rs = New ADODB.Recordset
Set Rs = VGcnx.Execute("SELECT sistemaconfiguragastos FROM co_sistema")
    
If Not (Rs.BOF Or Rs.EOF) Then
   cuenta = Trim(Rs(0))
   For I = 1 To Len(cuenta)
       If Mid(cuenta, I, 1) = "*" Then num = num + 1
   Next
   ReDim VG_gNIVELES(Len(cuenta) - num)
   j = 0
   For I = 1 To Len(cuenta) Step 2
       VG_gNIVELES(j) = Mid(cuenta, I, 1)
       j = j + 1
   Next
   VGnumnivgas = Len(cuenta) - num
End If
Set Rs = Nothing
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
Dim I As Integer
Dim cad As String
Dim cont As Long, longi As Long
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
            cont = cont + 1
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
 Dim I As Integer
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
 Dim I As Integer
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

Public Sub PropCrystal(ByRef CrystalRpt As CrystalReport)
    CrystalRpt.WindowShowCancelBtn = True
    CrystalRpt.WindowShowCloseBtn = True
    CrystalRpt.WindowShowExportBtn = True
    CrystalRpt.WindowShowGroupTree = True
    CrystalRpt.WindowShowNavigationCtls = True
    CrystalRpt.WindowShowPrintBtn = True
    CrystalRpt.WindowShowPrintSetupBtn = True
    CrystalRpt.WindowShowProgressCtls = True
    CrystalRpt.WindowShowSearchBtn = True
    CrystalRpt.WindowShowZoomCtl = True
End Sub

Public Function ValidaAsientos() As Boolean
 Dim SQL As String
    SQL = "SELECT ct_asiento.asientocodigo as Código, ct_asiento.asientodescripcion as Descripción "
    SQL = SQL & "FROM ct_asiento "
    SQL = SQL & "WHERE ct_asiento.asientocodigo<>'00'"
   
    Set VGvardllgen = New dllgeneral.dll_general
    If VGvardllgen.VerificaDatoExistente(VGcnx, SQL) <= 0 Then
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
    If VGvardllgen.VerificaDatoExistente(VGcnx, SQL) <= 0 Then
        ValidaSubAsientos = False
        MsgBox "Faltan registrar los SubAsientos que corresponden al Asiento Nº " & xCodAsiento, vbInformation, "Sistema Contable"
    Else
        ValidaSubAsientos = True
    End If
    Set VGvardllgen = Nothing
End Function

Public Function XRecuperaTipoCambio(Fecha As String, Tipo As Tipocambio, cnx As ADODB.Connection) As Double
Dim rsaux As ADODB.Recordset
Set rsaux = New ADODB.Recordset
Dim Campo As String
    XRecuperaTipoCambio = 0
    Select Case Tipo
        Case Compra
            Campo = "tipocambiocompra"
        Case Venta
            Campo = "tipocambioventa"
        Case Promedio
            Campo = "tipocambiopromedio"
        Case Else
            Campo = "tipocambioventa"
    End Select
    rsaux.Open "Select Valor=isnull(" & Campo & ",0)  from ct_tipocambio where tipocambiofecha ='" & Fecha & "'", cnx, adOpenKeyset, adLockReadOnly
    If rsaux.RecordCount > 0 Then
        XRecuperaTipoCambio = rsaux!Valor
    End If
End Function
Public Function ExisteSQL(ByVal cnx As ADODB.Connection, ByVal SentenciaSQL As String) As Boolean
On Error GoTo SaliError
    Screen.MousePointer = 11
    ExisteSQL = False
    Dim rsaux As ADODB.Recordset
    Set rsaux = New ADODB.Recordset
    rsaux.Open SentenciaSQL, cnx, adOpenKeyset, adLockReadOnly
    If rsaux.RecordCount > 0 Then
        ExisteSQL = True
    End If
    Screen.MousePointer = 1
    Exit Function
SaliError:
    Screen.MousePointer = 1
    ExisteSQL = False
    MsgBox Err.Description
 '   Resume
End Function
Public Sub CancelaDocumentos()
    On Error GoTo Mayor
    Dim X As Long
    Screen.MousePointer = 11
    VGcnx.BeginTrans
    Set VGCommandoSP = New ADODB.Command
    VGCommandoSP.ActiveConnection = VGcnxMarfice
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "ct_ProcCanc_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@Base") = VGParamSistem.BDEmpresa
        .Parameters("@Ano") = VGParamSistem.Anoproceso
        .Parameters("@Mes") = VGParamSistem.Mesproceso
        .Execute X
    End With
    VGcnx.CommitTrans
    Screen.MousePointer = 1
    MsgBox "Se Cancelo los Documentos Satisfactoriamente  " & Chr(13) & _
           "Items Afectados ", vbInformation
    Exit Sub
Mayor:
    Screen.MousePointer = 1
    VGcnx.RollbackTrans
    MsgBox "No se pudo Cancelar los Documentos " & Chr(13) & Err.Description, vbExclamation
End Sub
Public Function UltNumeroAuto(Tabla As String, OP As String, cnx As ADODB.Connection) As Long
Dim rsaux As ADODB.Recordset
On Error GoTo errornum
    Set rsaux = New ADODB.Recordset
    Select Case OP
        Case 1
            rsaux.Open "SELECT Numx=isnull(IDENT_CURRENT('" & Tabla & "'),0)", cnx, adOpenKeyset, adLockReadOnly
    End Select
    If rsaux.EOF Or rsaux.BOF Then
      UltNumeroAuto = 0
      Exit Function
    End If
     
    If rsaux.RecordCount = 1 Then
        UltNumeroAuto = rsaux!Numx
    Else
        UltNumeroAuto = -1
    End If
    Exit Function
errornum:
    UltNumeroAuto = -1
End Function
Private Function ArmaCriterioComodin(cad As String, Campo As String) As String
Dim pos As Integer, cadaux As String, criterio As String
Dim Valor As String
    criterio = ""
    Do While True
        pos = InStr(1, cad, "%", vbTextCompare)
        If pos = 0 Then Exit Do
        Valor = "'" & Left(cad, pos) & "'"
        cad = Right(cad, (Len(cad) - pos))
        criterio = criterio & Campo & " like " & Valor & " or "
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
 Dim Rs As ADODB.Recordset
 Dim cuenta As String
 Dim I As Integer
 Dim j As Integer
 Dim num As Integer
 
    Set Rs = New ADODB.Recordset
    Set Rs = VGcnx.Execute("SELECT sistemaconfiguragastos FROM co_sistema")
    
    If Not (Rs.BOF Or Rs.EOF) Then
        cuenta = Trim(Rs(0))
        For I = 1 To Len(cuenta)
            If Mid(cuenta, I, 1) = "*" Then num = num + 1
        Next
        ReDim VG_gNIVELES(Len(cuenta) - num)
        j = 0
        For I = 1 To Len(cuenta) Step 2
            VG_gNIVELES(j) = Mid(cuenta, I, 1)
            j = j + 1
        Next
        VGnumnivgas = Len(cuenta) - num
    End If
    Set Rs = Nothing
End Sub

Public Function ExisteElem(Tip As Integer, Cn As ADODB.Connection, Tabla As String, _
        Optional Campo As String) As Boolean
'Funcion que devuelve un valor verdadero si es que encuentra el elemento
'Creado por Fernando Cossio
    Dim SQL As String
    Dim rsaux As New ADODB.Recordset
   '*------------------------------*
   '0 Si Existe la tabla
   '1 Si Existe el Campo
   ExisteElem = False
   Tabla = UCase(Tabla): Campo = UCase(Campo)
On Error GoTo ErrExiste
   SQL = ""
    Select Case Tip
        Case 0:
            SQL = "Select Top 1 * From " & Tabla
        Case 1:
            SQL = "Select Top 1 " & Campo & " From " & Tabla
    End Select
    rsaux.Open SQL, Cn
    ExisteElem = True
    Exit Function
ErrExiste:
    ExisteElem = False
End Function

