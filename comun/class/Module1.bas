Attribute VB_Name = "Module1"
Private Sub CrystOrden(ByRef cry As CrystalReport, cad As String)
Dim pos As Integer, cadaux As String, i As Integer
Dim Valor As String
    Do While True
        pos = InStr(1, cad, ",", vbTextCompare)
        i = 0
        If pos = 0 Then Exit Do
        Valor = Left(cad, pos - 1)
        cry.SortFields(i) = Valor
        i = i + 1
        cad = Right(cad, (Len(cad) - pos))
    Loop
End Sub

Sub ImpresionRptbase(cNombreReporte As String, PFormulas(), Param(), Optional ORDEN As String, Optional titulo As String)
Dim i As Integer
On Error GoTo X
    Screen.MousePointer = 11
    With MDIPrincipal.CryRptProc
        .Reset
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowTitle = titulo
        Call PropCrystal(MDIPrincipal.CryRptProc)
        .ReportFileName = VGParamSistem.RutaReport & "\" & cNombreReporte
        If VGsql = 1 Then
           .Connect = "Provider=SQLOLEDB;Password=" & VGParamSistem.PwdGEN & ";Persist Security Info=True;User ID=" & VGParamSistem.UsuarioGEN & ";Initial Catalog=" & VGParamSistem.BDEmpresaGEN & ";SERVER=" & VGParamSistem.ServidorGEN
          Else
           .Connect = VGCadenaReport2
 
        End If
           
        .Formulas(0) = "@Emp='" & VGParametros.NomEmpresa & "'"
        .Formulas(1) = "@Ruc='" & VGParametros.RucEmpresa & "'"
        If UBound(PFormulas) > 0 Then
            For i = 0 To UBound(PFormulas) - 1
                .Formulas(2 + i) = PFormulas(i)
            Next
        End If
        .DiscardSavedData = True
        If UBound(Param) > 0 Then
            For i = 0 To UBound(Param) - 1
                .StoredProcParam(i) = Param(i)
            Next
        End If
        If ORDEN <> "" Then Call CrystOrden(MDIPrincipal.CryRptProc, ORDEN)
        If .Status <> 2 Then .Action = 1
    End With
    Screen.MousePointer = 1
    Exit Sub
X:
  If Err.Number = 9 Then Resume Next
  Screen.MousePointer = 1
  MsgBox "Error inesperado: " & Err.Number & "  " & Err.Description, vbExclamation
End Sub
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
    CrystalRpt.Destination = crptToWindow
    CrystalRpt.WindowState = crptMaximized
  
End Sub

Sub ImpresionRpt_SubRpt_Proc(cNombreReporte As String, PFormulas(), Param(), cNombreSubRpt As String, Optional ORDEN As String, Optional titulo As String)
Dim strBuscar As New dll_apis
Dim i As Integer
On Error GoTo X
    Screen.MousePointer = 11
    With MDIPrincipal.CryRptProc
        .Reset
        .WindowTitle = titulo
        Call PropCrystal(MDIPrincipal.CryRptProc)
        If Right(VGParamSistem.RutaReport, 1) <> "\" Then VGParamSistem.RutaReport = VGParamSistem.RutaReport & "\"
        .ReportFileName = VGParamSistem.RutaReport + cNombreReporte
        
        If VGsql = 1 Then
           .Connect = "Provider=SQLOLEDB;Password=" & VGParamSistem.PwdGEN & ";Persist Security Info=True;User ID=" & VGParamSistem.UsuarioGEN & ";Initial Catalog=" & VGParamSistem.BDEmpresaGEN & ";SERVER=" & VGParamSistem.ServidorGEN
          Else
           .Connect = VGCadenaReport2

        End If
           
        .Formulas(0) = "@Empresa='" & VGParametros.NomEmpresa & "'"
        .Formulas(1) = "@Ruc='" & VGParametros.RucEmpresa & "'"
        If UBound(PFormulas) > 0 Then
            For i = 0 To UBound(PFormulas) - 1
                .Formulas(2 + i) = PFormulas(i)
            Next
        End If
        .DiscardSavedData = True
        If UBound(Param) > 0 Then
            For i = 0 To UBound(Param) - 1
                .StoredProcParam(i) = Param(i)
            Next
        End If
   '     .DiscardSavedData = True
        '***Para el SubReporte
        .SubreportToChange = cNombreSubRpt
        If VGsql = 1 Then
           .Connect = "Provider=SQLOLEDB;Password=" & VGParamSistem.PwdGEN & ";Persist Security Info=True;User ID=" & VGParamSistem.UsuarioGEN & ";Initial Catalog=" & VGParamSistem.BDEmpresaGEN & ";SERVER=" & VGParamSistem.ServidorGEN
          Else
           .Connect = VGCadenaReport2

        End If

        If UBound(Param) > 0 Then
            For i = 0 To UBound(Param) - 1
                .StoredProcParam(i) = Param(i)
            Next
        End If
        If ORDEN <> "" Then Call CrystOrden(MDIPrincipal.CryRptProc, ORDEN)
        If .Status <> 2 Then .Action = 1
    End With
    Screen.MousePointer = 1
    Exit Sub
X:
  If Err.Number = 9 Then Resume Next
  Screen.MousePointer = 1
  MsgBox "Error inesperado: " & Err.Number & "  " & Err.Description, vbExclamation
End Sub

Public Function DateSQL(ByVal Fecha As String) As String
    'On Error GoTo ERR
    If IsNull(Fecha) Then Exit Function
        Select Case VGformatofecha
            Case "DMY"
            DateSQL = "'" & Format(Fecha, "dd/mm/yyyy") & "'"
            Case "MDY"
            DateSQL = "'" & Format(Fecha, "mm/dd/yyyy") & "'"
        End Select
'ERR:
 '    DateSQL = "'" & Day(FECHA) & "/" & Month(FECHA) & "/" & Year(FECHA) & "'"
End Function


'Public Sub Init_ControlDataGrid(EsteGrid As DataGrid)
' With EsteGrid
'  .AllowAddNew = False
'  .AllowDelete = False
'  .AllowUpdate = False
'  .AllowRowSizing = False
'  .TabAction = dbgControlNavigation
'  .MarqueeStyle = dbgHighlightRow
 ' .Font =
' End With
'End Sub

Public Function Devolver_Dato(tipo As Integer, Cod As String, Tabla As String, Campo As String, Fecha As Boolean, CampDev As String, Optional Cod2 As String, Optional Campo2 As String, Optional Cod3 As String, Optional Campo3 As String, Optional Cod4 As Double, Optional Campo4 As String) As String
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
Select Case tipo
  Case 1 'Bd. Comun
              cSel1.Open cF, VGCNx, adOpenStatic
  Case 2 'Bd. Config
              cSel1.Open cF, VGConfig, adOpenStatic
  Case 3 'Bd. Contabilidad
              cSel1.Open cF, VGcnxCT, adOpenStatic
End Select

If cSel1.RecordCount > 0 Then
     Devolver_Dato = IIf(Not IsNull(cSel1(0)), cSel1(0), "")
Else
     Devolver_Dato = ""
End If
End Function

