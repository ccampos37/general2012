Attribute VB_Name = "Proceds"
Declare Function GetPrivateProfileString Lib "kernel32" Alias _
        "GetPrivateProfileStringA" (ByVal lpAplicationName _
        As String, ByVal lpKeyName As Any, ByVal lpDefault _
        As String, ByVal lpReturnedString As String, ByVal nSize _
        As Long, ByVal lpFileName As String) As Long
Public Enum TIPFECHA
   Sqlf = 1
   Adof = 2
End Enum

Public Function sGetIni(sIniFile As String, sSection As String, sKey _
                        As String, sDefault As String) As String
 Dim sTemp As String * 256
 Dim nLength As Integer
 sTemp = Space$(256)
 nLength = GetPrivateProfileString(sSection, sKey, sDefault, sTemp, _
           255, sIniFile)
 sGetIni = Left$(sTemp, nLength)
End Function
Public Function Validar_RUC(xRuc As String) As Boolean
    Validar_RUC = True
    Exit Function
 Dim FLAG As Boolean
 Dim TAB_VAL(1 To 7) As Integer
 Dim NX As Integer, NY As Integer, NR As Integer, I As Integer
 Dim CadNR As String
 
 TAB_VAL(1) = 2
 TAB_VAL(2) = 7
 TAB_VAL(3) = 6
 TAB_VAL(4) = 5
 TAB_VAL(5) = 4
 TAB_VAL(6) = 3
 TAB_VAL(7) = 2
 FLAG = True
 xRuc = Trim(xRuc)
 
 If xRuc <> " " Then
  If xRuc <> "00000002" Then
     If Len(RTrim(xRuc)) < 8 Then
         MsgBox "Número de R.U.C. no tiene 8 dígitos", vbExclamation, "Ingreso de Datos"
         FLAG = flase
      Else
         NX = 0
         NR = 0
         NY = 0
         CadNR = ""
         For I = 1 To 7
             NX = NX + Val(Mid(xRuc, I, 1)) * TAB_VAL(I)
         Next I
         NY = NX \ 11
         NR = 11 - (NX - (NY * 11))
         CadNR = Trim(String(10 - Len(Str(NR)) + 1, "0")) & Trim(Str(NR))
         If Mid(CadNR, 10, 1) = Mid(xRuc, 8, 1) Then
            FLAG = True
         Else
            MsgBox "Número de R.U.C. invalido", vbExclamation, "Ingreso de Datos"
            FLAG = False
         End If
      End If
   Else
      MsgBox "Anexo emite Liquidaciones de compra", vbExclamation, "Ingreso de Datos"
   End If
 End If
 Validar_RUC = FLAG
End Function

Public Function DarCarnetSeg(ByVal FechaNac As Date, ByVal ApePat As String, ByVal ApeMat As String, ByVal NOMBRE As String, ByVal Sexo As Byte)
    Dim xCad As String, xTmp As String, X As Byte, ArrNum As String
    Dim I As Integer, txN As String, txC As String, txA As Integer
    ArrNum = "A1B2C3D4E5F6G7H8I9J1K2L3M4N5O6P7Q8R9S2T3U4V5W6X7Y8Z9"
    ApePat = UCase(ApePat)
    ApeMat = UCase(ApeMat)
    NOMBRE = UCase(NOMBRE)
    xCad = Format(Year(FechaNac) - 1900, "00") & Format(Month(FechaNac), "00") & Format(Day(FechaNac), "00") & Sexo
    If Len(ApePat) < 4 Then xTmp = Right(ApePat, 1) Else xTmp = Mid(ApePat, 4, 1)
    If xTmp = " " Then MsgBox "Tarea con resultados poco fiables: 0454. El apellido paterno puede ser compuesto", vbCritical
    xCad = xCad & Left(ApePat, 1) & xTmp
    If Len(ApeMat) < 4 Then xTmp = Right(ApeMat, 1) Else xTmp = Mid(ApeMat, 4, 1)
    If xTmp = " " Then MsgBox "Tarea con resultados poco fiables: 0455. El apellido materno puede ser compuesto", vbCritical
    xCad = xCad & Left(ApeMat, 1) & xTmp & Left(NOMBRE, 1) & "00"
    xTmp = Left(xCad, 7)
    For X = 8 To 12
        xTmp = xTmp & Mid(ArrNum, InStr(ArrNum, Mid(xCad, X, 1)) + 1, 1)
    Next
    txC = ""
    For X = 1 To 12
        I = Val(Mid(xTmp, X, 1)) * IIf(X Mod 2 = 0, 1, 2)
        txN = Trim(Str(I))
        If Len(txN) = 2 Then
            I = Val(Left(txN, 1)) + Val(Right(txN, 1))
            txN = Trim(Str(I))
        End If
        txC = txC & txN
    Next
    txA = 0
    For X = 1 To 12
        txA = txA + Val(Mid(txC, X, 1))
    Next
    I = Trim(Str(10 - IIf(txA Mod 10 = 0, 10, txA Mod 10)))
    xCad = xCad & I
    If Len(xCad) <> 15 Then
        Beep
        MsgBox "La operación no se ha efectuado correctamente, revise los datos que corresponden al apellido paterno, materno y el nombre, así como la fecha de nacimiento y el sexo", vbInformation
        xCad = "*** Error ***"
    End If
    DarCarnetSeg = xCad
End Function

Public Sub Init_ControlDataGrid(EsteGrid As DataGrid)
 With EsteGrid
  .AllowAddNew = False
  .AllowDelete = False
  .AllowUpdate = False
  .AllowRowSizing = False
  .TabAction = dbgControlNavigation
  .MarqueeStyle = dbgHighlightRow
 End With
End Sub

Public Function DateSQL(ByVal FECHA As Date) As String
    'On Error GoTo ERR
    If IsNull(FECHA) Then Exit Function
        Select Case REGSISTEMA.FORMATOFECHA
            Case "DMY"
            DateSQL = "'" & Format(FECHA, "dd/mm/yyyy") & "'"
            Case "MDY"
            DateSQL = "'" & Format(FECHA, "mm/dd/yyyy") & "'"
        End Select
'ERR:
 '    DateSQL = "'" & Day(FECHA) & "/" & Month(FECHA) & "/" & Year(FECHA) & "'"
End Function

Public Function FechaMMAAAA(ByVal MesAnno As String) As Date
    FechaMMAAAA = CDate("01/" & Left(MesAnno, 2) & "/" & Right(MesAnno, 4))
End Function

Function ExisteTabla(ByVal NOMBRETabla As String) As Boolean
    Dim RsTbls As New ADODB.Recordset
    Set RsTbls = DBSYSTEM.OpenSchema(adSchemaTables)
    RsTbls.FIND "[Table_Name]='" & UCase(NOMBRETabla) & "'"
    If RsTbls.EOF Then ExisteTabla = False Else ExisteTabla = True
End Function
Function ExisteTabla2(ByVal NOMBRETabla As String) As Boolean
    Dim RsTbls As New ADODB.Recordset
    Set RsTbls = DBADMINPER.OpenSchema(adSchemaTables)
    NOMBRETabla = Replace(NOMBRETabla, "[", ""): NOMBRETabla = Replace(NOMBRETabla, "]", "")
    RsTbls.FIND "[Table_Name]='" & UCase(Trim(NOMBRETabla)) & "'"
    If RsTbls.EOF Then ExisteTabla2 = False Else ExisteTabla2 = True
End Function

Function ExisteTablaAux(ByVal NOMBRETabla As String) As Boolean
'    Dim RsTbls As New ADODB.Recordset
'    Set RsTbls = DBAUXCOM.OpenSchema(adSchemaTables)
'    NOMBRETabla = Replace(NOMBRETabla, "[", ""): NOMBRETabla = Replace(NOMBRETabla, "]", "")
'    RsTbls.FIND "[Table_Name]='" & UCase(Trim(NOMBRETabla)) & "'"
'    If RsTbls.EOF Then ExisteTablaAux = False Else ExisteTablaAux = True
'
    Dim RsTbls As New ADODB.Recordset
    Dim FLAG As Boolean
On Error GoTo handler
    Set RsTbls = New ADODB.Recordset
    RsTbls.Open "SELECT COUNT(*) FROM " & UCase(Trim(NOMBRETabla)), DBSYSTEM, adOpenKeyset, adLockReadOnly
    If FLAG Then GoTo handler2
    ExisteTablaAux = True
    Exit Function
GO:
    Set RsTbls = New ADODB.Recordset
    RsTbls.Open "SELECT COUNT(*) FROM " & UCase(Trim(NOMBRETabla)), DBAUXCOM, adOpenKeyset, adLockReadOnly
    If FLAG Then
        ExisteTablaAux = False
      Else
        ExisteTablaAux = True
    End If
    Exit Function
handler:
   FLAG = True
   Resume Next
handler2:
   FLAG = False
   GoTo GO
End Function

Function ExisteTablaSQL(ByVal NOMBRETabla As String, ByVal xConec As ADODB.Connection) As Boolean
    Dim RsTbls As New ADODB.Recordset
    Set RsTbls = xConec.OpenSchema(adSchemaTables)
    NOMBRETabla = Replace(NOMBRETabla, "[", ""): NOMBRETabla = Replace(NOMBRETabla, "]", "")
    RsTbls.FIND "[Table_Name]='" & UCase(Trim(NOMBRETabla)) & "'"
    If RsTbls.EOF Then ExisteTablaSQL = False Else ExisteTablaSQL = True
End Function

Public Sub ActivarTools(Reg_Act As REGWIN)
    'Procedimiento para Usar en <Form>.Activate
    'Controla barra de Herramientas por c/ventana
    MDIPrincipal.Toolbar1.Buttons("Nuevo").Enabled = Reg_Act.NUEVO
    MDIPrincipal.Toolbar1.Buttons("Editar").Enabled = Reg_Act.EDITAR
    MDIPrincipal.Toolbar1.Buttons("Eliminar").Enabled = Reg_Act.ELIMINAR
    MDIPrincipal.Toolbar1.Buttons("Imprimir").Enabled = Reg_Act.IMPRIMIR
    MDIPrincipal.Toolbar1.Buttons("Preliminar").Enabled = Reg_Act.PRELIMINAR
    MDIPrincipal.Toolbar1.Buttons("Buscar").Enabled = Reg_Act.BUSCAR
    MDIPrincipal.Toolbar1.Buttons("Filtrar").Enabled = Reg_Act.FILTRAR
    MDIPrincipal.BarraEstado.Panels(2) = Screen.ActiveForm.Tag & " "
    With MDIPrincipal
        .Men02_06.Enabled = Reg_Act.NUEVO
        .Men02_07.Enabled = Reg_Act.ELIMINAR
        .Men02_05.Enabled = Reg_Act.EDITAR
        .Men02_08.Enabled = Reg_Act.BUSCAR
    End With
End Sub
Public Function BusCad(Caracter As String, CADENA As String) As Long
'Busca un caracter en la cadena especificada y me devuelve la cantidad de repeticiones
'del caracter encontrado
    Dim I As Long, ACUM As Long
    ACUM = 0
    For I = 1 To Len(CADENA)
        If Mid(CADENA, I, 1) = Caracter Then
            ACUM = ACUM + 1
        End If
    Next
    BusCad = ACUM
End Function
Public Function Getcad(Caracter As String, Numcad As Integer, CADENA As String) As String
'Devuelve una cadena hasta el numero de repeticiones del caracter especificado
    Dim I As Integer
    Dim ACUM As Integer
    ACUM = 0
    For I = 1 To Len(CADENA)
        If Mid(CADENA, I, 1) = Caracter Then
           ACUM = ACUM + 1
        End If
        If ACUM = Numcad Then Exit Function
        Getcad = Getcad + Mid(CADENA, I, 1)
    Next
End Function
Public Sub SetPosition(ByVal Formulario As Form)
    SaveSetting App.CompanyName, "DYB" & Formulario.Name, "Top", Formulario.TOP
    SaveSetting App.CompanyName, "DYB" & Formulario.Name, "Left", Formulario.Left
End Sub
Public Sub GetPosition(ByRef Formulario As Form)
    Formulario.TOP = GetSetting(App.CompanyName, "DYB" & Formulario.Name, "Top", 0)
    Formulario.Left = GetSetting(App.CompanyName, "DYB" & Formulario.Name, "Left", 0)
End Sub
Public Function ExisteCampo(CAMPO As String, TABLA As String, ByVal Conexion As ADODB.Connection) As Boolean
    On Error GoTo ErrNoHayChocherita
    ExisteCampo = False
    If CAMPO = "" Then Exit Function
    Dim RSAUX As New ADODB.Recordset
    RSAUX.Open "SELECT TOP 1 " & UCase(CAMPO) & " FROM " & UCase(TABLA), Conexion, adOpenStatic, adLockReadOnly
    ExisteCampo = True
    Set RSAUX = Nothing
    Exit Function
ErrNoHayChocherita:
    ExisteCampo = False
    Exit Function
End Function

Public Function DevuelveValor(CadenaSQL As String, ConexionADO As ADODB.Connection) As Variant
    On Error GoTo ErrDevolver
    Dim RSAUX As New ADODB.Recordset
    RSAUX.Open CadenaSQL, ConexionADO, adOpenStatic, adLockReadOnly
    If RSAUX.EOF Or RSAUX.RecordCount = 0 Then
        Exit Function
    Else
        DevuelveValor = IIf(IsNull(RSAUX(0).Value), 0, RSAUX(0).Value)
    End If
    Set RSAUX = Nothing
ErrDevolver:
    Exit Function
End Function
Public Function GetValor(CadenaSQL As String, ConexionADO As ADODB.Connection) As Variant
'Funcion Modificada Por Fernando cossio
 GetValor = Null
 On Error GoTo ErrDevolver
    Dim RSAUX As New ADODB.Recordset
    RSAUX.Open CadenaSQL, ConexionADO, adOpenStatic, adLockReadOnly
    If RSAUX.EOF Or RSAUX.RecordCount = 0 Then
        Exit Function
    Else
        GetValor = RSAUX(0).Value
    End If
    Set RSAUX = Nothing
ErrDevolver:
    Exit Function
    MsgBox ERR.Description
End Function
Public Sub CreaTempCostos(Conexion As ADODB.Connection, TABLA As String, CAMPO As String)
    Dim SQL As String
    Dim RSTEMP As New ADODB.Recordset
    Dim Niveles As Long, I As Integer, NivCC As Long
    Dim CC As String
    Screen.MousePointer = 11
    RSTEMP.Open TABLA, DBSYSTEM, adOpenKeyset, adLockReadOnly
    If RSTEMP.RecordCount = 0 Then Exit Sub
    If ExisteTablaAux(" [##TEMPCOSTOS" & VGL_COMPUTER & "] ") Then
        Conexion.Execute "DROP TABLE  [##TEMPCOSTOS" & VGL_COMPUTER & "] "
    End If
    Niveles = CantNivCC(CAMPO, TABLA, DBSYSTEM)
    Conexion.Execute "CREATE TABLE  [##TEMPCOSTOS" & VGL_COMPUTER & "] (" & CAMPO & " VarChar(10))"
    For I = 1 To 5
        Conexion.Execute "Alter Table  [##TEMPCOSTOS" & VGL_COMPUTER & "]  Add NIVEL" & Trim(Str(I)) & " VarChar(10) "
    Next
    RSTEMP.MoveFirst
    Dim RsTempNiv As New ADODB.Recordset
    RsTempNiv.Open " [##TEMPCOSTOS" & VGL_COMPUTER & "] ", Conexion, adOpenDynamic, adLockOptimistic
    Do While Not RSTEMP.EOF
        NivCC = BusCad(".", RSTEMP.Fields(CAMPO).Value) + 1
        RsTempNiv.AddNew
        RsTempNiv.Fields(CAMPO).Value = RSTEMP.Fields(CAMPO).Value
        For I = 1 To NivCC
            RsTempNiv.Fields("Nivel" & Trim(Str(I))).Value = _
            Getcad(".", I, RSTEMP.Fields(CAMPO).Value)
        Next
        RsTempNiv.Update
        RSTEMP.MoveNext
    Loop
    Screen.MousePointer = 1
End Sub
Public Function CantNivCC(CAMPO As String, TABLA As String, Conexion As ADODB.Connection) As Long
    Dim Max As Long, NumNiv As Long
    Dim RsCosto As New ADODB.Recordset
    CantNivCC = 0
    Max = 0
    RsCosto.Open TABLA, Conexion, adOpenKeyset, adLockReadOnly
    If RsCosto.RecordCount = 0 Then Exit Function
    CantNivCC = 1
    Do While Not RsCosto.EOF
        NumNiv = BusCad(".", RsCosto.Fields(CAMPO).Value)
        If Max < NumNiv Then Max = NumNiv
        RsCosto.MoveNext
    Loop
    CantNivCC = Max + 1
End Function

'--------------------------------------
' NUMLET - Convierte Numeros a Letras
' RETORMA LET CON 160 BYTES
'--------------------------------------
Public Function NUMLET(NUM As Double) As String
Dim cLET As String
Dim cWork As String
Dim cUNIDAD As String
Dim cDECENA As String
Dim cCENTENA As String
Dim nMODULUS As Integer
Dim NI As Integer
Dim nK As Integer
Dim Lit1 As String
Dim Lit2 As String
Dim Lit3 As String
Dim Lit4 As String
Dim Lit5 As String
Lit1 = "Uno    Doc    Trec   Cuatroc  Quin   Seisc  Setec  Ochoc  Novec  "
Lit2 = "Diez     Veinte   Treinta  Cuarenta CincuentaSesenta  Setenta  Ochenta  Noventa  "
Lit3 = "Once      Doce      Trece     Catorce   Quince    Dieciseis DiecisieteDieciocho Diecinueve"
Lit4 = "Uno   Dos   Tres  CuatroCinco Seis  Siete Ocho  Nueve "
Lit5 = "Millon    Billon    Trillon   CuatrillonQuintillon"
'Proceso Input = Num , Output = Let

cLET = ""
If NUM > 0.99 Then
    'Separa los Enteros en una Cadena de Caracteres
     If InStr(1, Trim(Str(NUM)), ".", 0) > 0 Then
        cWork = Mid(Trim(Str(NUM)), 1, InStr(1, Trim(Str(NUM)), ".", 0) - 1)
     Else
        cWork = Str(NUM)
     End If
     nMODULUS = Int(Len(Trim(cWork)) / 3)
     nMODULUS = Len(Trim(cWork)) - (nMODULUS * 3)
     
     If nMODULUS > 0 Then
        cWork = String(3 - nMODULUS, "0") & Trim(cWork)
     End If
     
     nK = (Len(Trim(cWork)) / 3) - 1
    'Procesa de Mil en Mil
     NI = 1
     Do While NI < Len(Trim(cWork)) - 1
        cCENTENA = Mid(Trim(cWork), NI, 1)
        cDECENA = Mid(Trim(cWork), NI + 1, 1)
        cUNIDAD = Mid(Trim(cWork), NI + 2, 1)
        'Centenas
        If cCENTENA <> "0" Then
            If cCENTENA = "1" Then
                cLET = cLET & "Cien "
                If cDECENA <> "0" Or cUNIDAD <> "0" Then
                    cLET = Mid(cLET, 1, (Len(cLET) - 1)) & "to "
                End If
            Else
                cLET = cLET & Trim(Mid(Lit1, ((Val(cCENTENA) - 1) * 7) + 1, 7)) & "ientos "
            End If
        End If
        'Decenas
        If cDECENA <> "0" Then
            If cDECENA = "1" And cUNIDAD <> "0" Then
                If ((Val(cUNIDAD) - 1) * 10) + 1 > 0 Then cLET = cLET + Trim(Mid(Lit3, ((Val(cUNIDAD) - 1) * 10) + 1, 10))
            Else
                If ((Val(cDECENA) - 1) * 9) + 1 > 0 Then cLET = cLET + Trim(Mid(Lit2, ((Val(cDECENA) - 1) * 9) + 1, 9))
            End If
        End If
        'Unidades
        If cUNIDAD <> "0" Then
            If cDECENA > "1" Then
                cLET = Mid(cLET, 1, (Len(cLET) - 1)) & "i"
                If ((Val(cUNIDAD) - 1) * 6) + 1 > 0 Then cLET = cLET + LCase(Trim(Mid(Lit4, ((Val(cUNIDAD) - 1) * 6) + 1, 6)))
            Else
                If cDECENA < "1" Then
                    If ((Val(cUNIDAD) - 1) * 6) + 1 > 0 Then cLET = cLET + Trim(Mid(Lit4, ((Val(cUNIDAD) - 1) * 6) + 1, 6))
                End If
            End If
        End If
        cLET = cLET & " "
        'Pone Miles o Millones
        If nK > 0 Then
            If cCENTENA & cDECENA & cUNIDAD = "001" Then
                cLET = Mid(cLET, 1, Len(cLET) - 2) & " "
            End If
            nMODULUS = Int(nK / 2)
            nMODULUS = nK - (nMODULUS * 2)
            If nMODULUS = 0 Then
                cLET = cLET + Trim(Mid(Lit5, (((nK / 2) - 1) * 10) + 1, 10))
                If cCENTENA & cDECENA & cUNIDAD = "001" Or NUM > 1999999 Then
                    cLET = cLET & "es "
                Else
                    cLET = cLET & " "
                End If
            Else
                If cCENTENA & cDECENA & cUNIDAD > "000" Then
                    cLET = cLET & "Mil "
                End If
            End If
            nK = nK - 1
        End If
        NI = NI + 3
    Loop
    cLET = cLET & "con "
End If
If InStr(1, Trim(Str(NUM)), ".", 0) > 0 Then
    cLET = cLET + Mid(Trim(Str(NUM)), InStr(1, Trim(Str(NUM)), ".", 0) + 1, 2) & "/100" & " "
Else
    cLET = cLET + "00/100" & " "
End If
NUMLET = cLET
End Function
Public Function FechS(FECHA As Variant, TIPO As TIPFECHA) As Variant
Dim H As Date
Dim fechaAux As Double
On Error GoTo ErrorFecha
   H = CDate(FECHA)
   Select Case TIPO
      Case Sqlf: 'Para transformar al sql
         fechaAux = DateSerial(Year(FECHA), Month(FECHA), Day(FECHA)) - 2
      Case Adof: 'Para transformar al ado Y AL ACCESS
         fechaAux = DateSerial(Year(FECHA), Month(FECHA), Day(FECHA))
   End Select
   FechS = fechaAux
   Exit Function
ErrorFecha:
   Select Case TIPO
      Case Sqlf: FechS = "Null"
      Case Adof: FechS = Null
   End Select
End Function
Public Function ESNULO(EXPRESION As Variant, VALOR As Variant) As Variant
On Error GoTo errfun
   If IsNull(EXPRESION) Or Trim(EXPRESION) = Empty Then
      ESNULO = VALOR
     Else: ESNULO = EXPRESION
   End If
   Exit Function
errfun:
   ESNULO = 0
End Function
Public Function RestringeCaracter(ByVal KEY As Integer, ByVal CAD As String) As Integer
    If InStr(CAD, Chr(KEY)) > 0 Then
        If KEY <> 44 Then RestringeCaracter = 0 Else RestringeCaracter = KEY
      Else
        RestringeCaracter = KEY
    End If
End Function

Public Function DESMES(ByVal MES As Integer) As String
Select Case MES
    Case 1
        DESMES = "ENERO"
    Case 2
        DESMES = "FEBRERO"
    Case 3
        DESMES = "MARZO"
    Case 4
        DESMES = "ABRIL"
    Case 5
        DESMES = "MAYO"
    Case 6
        DESMES = "JUNIO"
    Case 7
        DESMES = "JULIO"
    Case 8
        DESMES = "AGOSTO"
    Case 9
        DESMES = "SETIEMBRE"
    Case 10
        DESMES = "OCTUBRE"
    Case 11
        DESMES = "NOVIEMBRE"
    Case 12
        DESMES = "DICIEMBRE"
End Select

End Function
