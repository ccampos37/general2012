Attribute VB_Name = "ModFunciones"
Option Explicit
Dim ndiaMes(12) As Integer
Dim cdesMes(12) As String
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Sub CentrarForm(nFormPrin As Form, nFormu As Form)
  nFormu.Left = (nFormPrin.Width - nFormu.Width) / 2
  nFormu.Top = ((nFormPrin.Height - nFormu.Height) / 2) - 600
End Sub

'FIXIT: Declare 'aFormulas' con un tipo de datos de enlace en tiempo de compilación        FixIT90210ae-R1672-R1B8ZE
Function fReporte(ByVal sReportname As String, oRs As ADODB.Recordset, ByVal sTitulo As String, Optional aFormulas As Variant) As String
'**************************************
' Ruta y nombre del reporte
' Recordset
' titulo del reporte
' Arreglo de formulas
'******************************//

'FIXIT: Declare 'oApp' and 'oRpt' and 'oRptOptions' and 'oDatabase' and 'oTables' and 'oTable1' and 'oFieldDefns' and 'oFieldDefn' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Dim oApp As Object, oRpt As Object, oRptOptions As Object, oDatabase As Object, oTables As Object, oTable1 As Object, oFieldDefns As Object, oFieldDefn As Object
'FIXIT: Declare 'oPageEngine' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
Dim oPageEngine As Object
Dim iLoopCount As Integer


Dim DirCadena As String

    DirCadena = App.Path & "\Reportes\" & Trim$(sReportname)
    If oRs Is Nothing Then
        GoTo ErrCodigo
    Else
        If oRs.RecordCount = 0 Then
           fReporte = "No existen datos para los parámatros especificados"
           Exit Function
        End If
    End If
    
    'Crea el objeto aplicacion del Crystal Report
    'Set oApp = CreateObject("Crystal.CRPE.Application")
'FIXIT: Declare 'cry' con un tipo de datos de enlace en tiempo de compilación              FixIT90210ae-R1672-R1B8ZE
    Dim cry As Object
    Set cry = frmMantOperacion.Controls.Add("crystal.CrystalReport", "hola")
    Dim cryst As Crystal.CrystalReport
    
    cryst.ReportFileName = App.Path & "\Reportes\rptOperacion.rpt"
    cryst.DiscardSavedData = True
    
    cryst.Action = 1
    'Crea el objeto reporte del Crystal Report
    If IsObject(oRpt) Then
        Set oRpt = Nothing
    End If
    
    Set oRpt = oApp.OpenReport(DirCadena, 1)
    'Desactiva los errores del motor del Crystal Report
    Set oRptOptions = oRpt.Options
    oRptOptions.MorePrintEngineErrorMessages = 0

    'Crear una coleccion de oDatabase el cual referencia a las bases de datos usada en el reporte
    oRpt.DiscardSavedData
    Set oDatabase = oRpt.Database

    'Instantiates a Tables collection which references the Tables of the Database object.
    Set oTables = oDatabase.Tables

    Set oTable1 = oTables.Item(1)
    'Instancia un objeto tabla  el cual referencia a la primera tabla usada en el reporte.

    oTable1.SetPrivateData 3, oRs
    'La linea "SetPrivateData"  le dice al reporte que el origen de datos es el recordset

    'On Error Resume Next
    
    'Asigna los valores que se encuentran en el arreglo a las formulas
    If Not IsMissing(aFormulas) Then
    Set oFieldDefns = oRpt.FormulaFields
'    For iLoopCount = 0 To UBound(aFormulas)
'        sFormulaNombre = Trim$(left(aFormulas(iLoopCount), InStr(1, aFormulas(iLoopCount), "=") - 1))
'        sFormulaValor = Mid$(aFormulas(iLoopCount), InStr(1, aFormulas(iLoopCount), "=") + 1, Len(aFormulas(iLoopCount)))
'        Set oFieldDefn = oFieldDefns.Item(sFormulaNombre)
'        oFieldDefn.Text = sFormulaValor
'    Next
    End If
    oRpt.ReadRecords
    
    If err.Number <> 0 Then
        oRpt.LastErrorString
        
        'Dim rpt As CRPEAuto.Report
        fReporte = "ERROR"
        GoTo ErrCodigo
    Else
        If IsObject(oPageEngine) Then
            Set oPageEngine = Nothing
        End If
        Set oPageEngine = oRpt.PageEngine
    End If
    oRpt.Preview sTitulo
    fReporte = "OK"
    Exit Function

ErrCodigo:
    err.Raise err.Number, "fReporte", "Un error ha ocurrido en el servidor al intentar accesar a los datos del reporte!" & vbCr & err.Description
End Function

Sub Impresion(cNombreReporte As String)
On Error GoTo x
  With MDIPrincipal.cryRpt
        .Reset
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .ReportFileName = VGParamSistem.RutaReport & VGParamSistem.carpetareportes & cNombreReporte
    '    .LogOnServer "pdssql.dll", VGParamSistem.Servidor, VGParamSistem.BDEmpresa, VGParamSistem.UsuarioReporte, VGParamSistem.Pwd
        If VGsql = 1 Then
           .Connect = "Provider=SQLOLEDB;Password=" & VGParamSistem.PwdGEN & ";Persist Security Info=True;User ID=" & VGParamSistem.UsuarioGEN & ";Initial Catalog=" & VGParamSistem.BDEmpresaGEN & ";SERVER=" & VGParamSistem.ServidorGEN
          Else
           .Connect = VGCadenaReport2
         End If
         .DiscardSavedData = True
        
        If cNombreReporte = "rptTipoCambio.rpt" Then
           Set VGvardllgen = New dllgeneral.dll_general
           .Formulas(0) = "@Mes=" & CInt(VGParamSistem.Mesproceso)
           .Formulas(1) = "@Mesanno='" & VGvardllgen.DesMes(VGParamSistem.Mesproceso) & "-" & VGParamSistem.Anoproceso & "'"
           .Formulas(2) = "@Anual='" & VGParamSistem.Anoproceso & "'"
        End If
        .Action = 1
  End With
  Exit Sub
x:
  MsgBox "Error inesperado: " & err.Number & "  " & err.Description, vbExclamation
End Sub

Private Sub CrystOrden(ByRef cry As CrystalReport, cad As String)
Dim pos As Integer, cadaux As String, i As Integer
Dim valor As String
    Do While True
        pos = InStr(1, cad, ",", vbTextCompare)
        i = 0
        If pos = 0 Then Exit Do
        valor = Left(cad, pos - 1)
        cry.SortFields(i) = valor
        i = i + 1
        cad = Right(cad, (Len(cad) - pos))
    Loop
End Sub

Public Sub ParametroCuenta(Index As Integer)
 Dim rs As ADODB.Recordset
 Dim cuenta As String
 Dim costos As String
 Dim i As Integer
 Dim J As Integer
 Dim configuracion As String
 Dim num1 As Integer
 Dim num2 As Integer
Select Case Index
  Case 0
    Set rs = New ADODB.Recordset
    Set rs = VGCNx.Execute("SELECT sistemaconfiguracuenta,sistemaconfiguracentrocostos FROM ct_sistema")
    
    If Not (rs.BOF Or rs.EOF) Then
        cuenta = IIf(IsNull(rs(0)), "", Trim$(rs(0)))
        costos = IIf(IsNull(rs(1)), "", Trim$(rs(1)))
        Set rs = Nothing
    Else
        configuracion = Trim$(rs(1))
    End If
  Case 1
        cuenta = strvalor
        costos = strvalor1
End Select
For i = 1 To Len(costos)
      If Mid$(costos, i, 1) = "*" Then num2 = num2 + 1
Next
ReDim VG_cNIVELES(Len(costos) - num2)
J = 0
For i = 1 To Len(costos) Step 2
    VG_cNIVELES(J) = Mid$(costos, i, 1)
    J = J + 1
Next
VGnumnivelescentrocosto = Len(costos) - num2

' ****
                                
For i = 1 To Len(cuenta)
    If Mid$(cuenta, i, 1) = "*" Then num1 = num1 + 1
 Next
 ReDim VG_aNIVELES(Len(cuenta) - num1)
 J = 0
 For i = 1 To Len(cuenta) Step 2
     VG_aNIVELES(J) = Mid$(cuenta, i, 1)
     J = J + 1
 Next
 VGnumnivelescuenta = Len(cuenta) - num1

End Sub

Public Sub CargarParametrosContabilidad()
'Fernando Cossio Peralta
Dim RSAUX As ADODB.Recordset
    Set RSAUX = New ADODB.Recordset
    Set RSAUX = VGCNx.Execute(" Select top 1 * from ct_sistema")
    If RSAUX.RecordCount = 0 Then Exit Sub
    VGParametros.monedabase = RSAUX!monedacodigo
    VGParametros.IGV = Trim$(RSAUX!sistemavalorigv)
    VGParametros.CuadreAsiento = RSAUX!sistemaestcuadreasiento
    VGParametros.impresionalta = RSAUX!sistematipoimpresion
    VGParametros.sistemamonista = RSAUX!sistemamonista
   VGParametros.sistemactaajustedeb = RTrim$(RSAUX!sistemactaajustedeb)
    VGParametros.sistemactaajustehab = RTrim$(RSAUX!sistemactaajustehab)
    VGParametros.AsientoAutoxCCostos = ESNULO(RTrim(RSAUX!asientoAutoCCostos), 0)
    VGParametros.cuentadeCostos = RTrim(ESNULO(RSAUX!cuentacodigoCostos, 0))


''    If Not VGParametros.impresionalta Then
''        VGParamSistem.carpetareportes = "ReportesMatricial"
''      Else
''        VGParamSistem.carpetareportes = "Reportes"
''    End If

    VGParametros.asientocodigo = RSAUX!sistemaasientocodigo
    VGParametros.subasientocodigo = RSAUX!sistemasubasientocodigo
    
   If IsNumeric(VGParamSistem.Anoproceso) And IsNumeric(VGParamSistem.Mesproceso) Then
        SQL = "select * from ct_cierremensual where empresacodigo='" & VGParametros.empresacodigo & "' and " _
        & " anio='" & VGParamSistem.Anoproceso & "' and mes=" & Trim(VGParamSistem.Mesproceso) & " "
        Set RSAUX = VGCNx.Execute(SQL)
        If RSAUX.RecordCount > 0 Then VGParametros.cierremes = IIf(RSAUX!Contabilidad = True, True, False)
    End If
    
      Call ParametroCuenta(0)
     
'    MDIPrincipal.Caption = "Sistema de Contabilidad - " & Trim$(rsaux!sistemadescripcionempresa)
    MDIPrincipal.StatusBar1.Panels(5).Text = "Servidor (" & VGParamSistem.Servidor & ")"
    MDIPrincipal.StatusBar1.Panels(6).Text = "Base de Datos (" & VGParamSistem.BDEmpresa & ")"
    
End Sub
'FIXIT: Declare 'Coversion' con un tipo de datos de enlace en tiempo de compilación        FixIT90210ae-R1672-R1B8ZE
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
Public Sub HabilitarDetalle(flag As Boolean, framex As Frame)
'FCP
On Error Resume Next
framex.Enabled = flag
Dim Control As Control
    For Each Control In frmantcomprobantes.Controls
'FIXIT: 'Container.Name' no es una propiedad del objeto genérico 'Control' en Visual Basic .NET. Para obtener acceso a 'Container.Name', declare 'Control' utilizando su tipo real en lugar de 'Control'     FixIT90210ae-R1460-RCFE85
        If UCase$(Control.Container.Name) = UCase$(framex.Name) Then
            Control.Enabled = flag
        End If
    Next
    If VGMonSubAsiento = "" Or VGMonSubAsiento = "00" Then
        frmantcomprobantes.CtrAyu_Moneda.Enabled = True
      Else
 '       frmantcomprobantes.CtrAyu_Moneda.Enabled = False
    End If
End Sub
Public Sub ClearControlsInframe(framex As Frame)
'FCP
On Error Resume Next
    Dim Control As Control
    For Each Control In frmantcomprobantes.Controls
'FIXIT: 'Container.Name' no es una propiedad del objeto genérico 'Control' en Visual Basic .NET. Para obtener acceso a 'Container.Name', declare 'Control' utilizando su tipo real en lugar de 'Control'     FixIT90210ae-R1460-RCFE85
        If UCase$(Control.Container.Name) = UCase$(framex.Name) Then
            If UCase$(Left(Control.Name, 2)) <> "LE" Then
                If TypeOf Control Is TextBox Then Control.Text = ""
                If TypeOf Control Is TextFer.TxFer Then Control.Text = ""
'FIXIT: 'Caption' no es una propiedad del objeto genérico 'Control' en Visual Basic .NET. Para obtener acceso a 'Caption', declare 'Control' utilizando su tipo real en lugar de 'Control'     FixIT90210ae-R1460-RCFE85
                If TypeOf Control Is Label Then Control.Caption = ""
'FIXIT: 'Value' no es una propiedad del objeto genérico 'Control' en Visual Basic .NET. Para obtener acceso a 'Value', declare 'Control' utilizando su tipo real en lugar de 'Control'     FixIT90210ae-R1460-RCFE85
                If TypeOf Control Is DTPicker Then Control.Value = frmantcomprobantes.DTPFechaContab.Value
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
    On Error GoTo error
    longi = Len(RichTextBox1.Text)
    Do While pos <> 0
        pos = InStr(pos + 2, RichTextBox1.Text, "GO", vbTextCompare)
        sqlcad = ""
        If pos + 2 > longi Then Exit Do
        If pos = 0 Then Exit Do
        If Asc(Mid$(RichTextBox1.Text, pos - 1, 1)) = 10 And Asc(Mid$(RichTextBox1.Text, pos + 2, 1)) = 13 Then
            Cont = Cont + 1
            sqlcad = Mid$(RichTextBox1.Text, ini, pos - (ini + 2))
            RichTextBox1.SelStart = pos: RichTextBox1.SelLength = 2
            ini = pos + 2
            cnx.Execute sqlcad
        End If
    Loop
error:
  MsgBox "Error inesperado: " & err.Number & "  " & err.Description, vbExclamation
  
  Exit Sub
  Resume
End Sub

'FIXIT: Declare 'Espunto' and 'texto' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Public Function Espunto(ByRef texto As Variant) As Variant
    If Trim$(texto) = "." Then
        Espunto = "0"
      Else
        Espunto = texto
    End If
End Function

Public Sub ModoEditable(flagModo As Boolean, Formu As Form, cNameCtrX As String)
 Dim i As Integer
    Dim Control As Control
    For Each Control In Formu.Controls
       If UCase$(Control.Name) <> UCase$(cNameCtrX) Then
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
       If UCase$(Control.Name) <> UCase$(cNameCtrX) Then
           If TypeOf Control Is TextBox Then Control.Text = Empty
           If TypeOf Control Is TextFer.TxFer Then Control.Text = Empty
'FIXIT: 'Value' no es una propiedad del objeto genérico 'Control' en Visual Basic .NET. Para obtener acceso a 'Value', declare 'Control' utilizando su tipo real en lugar de 'Control'     FixIT90210ae-R1460-RCFE85
           If TypeOf Control Is CheckBox Then Control.Value = 0
           If TypeOf Control Is ctrlayuda_f.Ctr_Ayuda Then
'FIXIT: 'xclave' no es una propiedad del objeto genérico 'Control' en Visual Basic .NET. Para obtener acceso a 'xclave', declare 'Control' utilizando su tipo real en lugar de 'Control'     FixIT90210ae-R1460-RCFE85
              Control.xclave = Empty
'FIXIT: 'xnombre' no es una propiedad del objeto genérico 'Control' en Visual Basic .NET. Para obtener acceso a 'xnombre', declare 'Control' utilizando su tipo real en lugar de 'Control'     FixIT90210ae-R1460-RCFE85
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
    Dim x As Long
    Dim NombrePC As String
    Randomize   'Inicializa el generador de números aleatorios.
'FIXIT: Reemplazar la función 'Str' con la función 'Str$'.                                 FixIT90210ae-R9757-R1B8ZE
    NombrePC = Trim$(Str(CLng(Rnd * 10000000)))
    Screen.MousePointer = 11
    VGCNx.BeginTrans
    Set VGCommandoSP = New ADODB.Command
    VGCommandoSP.ActiveConnection = VGGeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "ct_ProcCanc_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@BaseConta") = VGParamSistem.BDEmpresa
        .Parameters("@empresa") = VGParametros.empresacodigo
        .Parameters("@Anno") = VGParamSistem.Anoproceso
        .Parameters("@Mes") = VGParamSistem.Mesproceso
        .Parameters("@NombrePC") = NombrePC
        .Execute x
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

