Attribute VB_Name = "Module1"
Option Explicit
Public VGSeleccion As Integer    ' modificar o adicionar. descripcion en el formulario registro
Public VGRegEnt  As Integer         ' registro de entrada o salida
Public VGSoles As Boolean           'indica si se trabaja con S/. o $
Public VGtransp As Boolean          'indica el trasnportista ,true para llamar del manteniminento,false  de g. remision
Public VGForm   As Integer            'indica el formulario en uso
Public VGForm1   As Integer          ' indica el formulario de procedencia para la ayuda
Public VGtipocreacion As Integer   'Para el modificar
Public VGabrev As String              ' codigo de unidad
Public VGcrea As Boolean
Public VGAlma  As String              'Codigo del almacen
Public VGval As Boolean               'Indicca si es valorizado
Public VGCOMP As String * 3         'Codigo de la compannia
Public VGLongCodigo As Integer    'Inddica la long de codigo de un articulo
Public VGActualizar As Boolean    'Para el caso de modificar  y restaurar informacion  en caso de no modificar
Public VGElimina As Boolean          'Para el caso de utilizar el formulario de eliminar y anular
Public VGAyuClie As Boolean
Public VGGuiaSal As Boolean        'Para el caso del form de Guia de salida en que puede crear o  modificar parcial
Public VGRuta As String                 'ruta de la base de datos
Public VGTipCamb As Double        'Tipo de Cambios
Public VGCodMon As String * 2      'Tipo de Moneda
Public VGWrk As Workspace
Public VGBaseDatos As Database
Public VGSALIR As Boolean
Public VGEstadomodi As Boolean    'Estado de modificacion
Public VGUsuario As String              'el usuario de la aplicacion
Public VGValnuevo  As Boolean           'Para doc valorizados
Public VGUsua  As String
Public VGPass  As String
Public VGNemp  As String                   'Nombre de la empresa
Public VGRclie As Boolean
Public cAnexo As String
Public VGIASA As String                     'Codigo de la empresa IASA, aplicacion personalizada
Public vGAdmLog As Boolean             'Login del Administrador
Public VGNameCont  As String             'Nombre de contabilidad
Public VGContTra As String                 'NombredBD de trasacciones de Contabilidad
Public VGAutomatico As Boolean         'Indica si la numeracion no es editable
Public VGRepKxVal As Integer
Public VGcc As Integer     'Indica si el reporte espor centro de costo o autorizado
'Public rsstock As Recordset

Public mensaje1 As String
Public cConexConf As ADODB.Connection    'BdConfigFac
Public cConexCom As ADODB.Connection   'BdComun
Public cConexCont As ADODB.Connection  'BdContabilidad

 Public cRutPath As String
 Public cRutP As String
 Public cRuta6 As String
 Public cNomBd  As String
 Public cNomBd5 As String
 Public Const cNomBd6 As String = "BdWenco.Mdb"
 Public cNomBd4 As String
 Public cNomBd2 As String
 Public cRuta5 As String
 Public cRuta2 As String                    'Contiene la ruta de Bd, incluyend el nombre  ***********
 Public cRuta3 As String
 Public cRuta4 As String
 Public sName As String

Public VGNomAlm As String              ' Nombre del almacen
Public RUTA As String                        'Indica solo la carpeta donde se ha instalado

Public Const NUMMAGICO As Integer = 5

'Variables globales para Administradores
Public VGTEMP As String
Public VGADM_CODIDO As String
Public VGADM_NOMBRE As String
Public VGADM_PASSWORD As String
Public VGAdmLogin As Boolean
'Variables globales para Empresas

Public VGUSU_CODIGO As String
Public VGUSU_PASSWORD As String
Public VGcod As String                          'Se utiliza para las consultas
Public vGUtil(4) As String                        'Se para los pases de ayuda
Public arrayserie()   As String                'Ingreso masivo de serie
Public Enum TIPFECHA
   Sqlf = 1
   Adof = 2
End Enum


'********************
' CL : Debe ser fijo       Compras Locales
' CE :Comprar al extranjero


Public Sub central(f As Form)
f.Left = (Screen.Width - f.Width) / 2
 f.Top = (Screen.Height - f.Height) / 2
End Sub

Public Sub Enfoque(Obj As Object)
  Obj.SelStart = 0
  Obj.SelLength = Len(Obj)
End Sub

Public Sub Init_ControlDBGrid(EsteGrid As DBGrid)
 With EsteGrid
  .AllowAddNew = False
  .AllowDelete = False
  .AllowUpdate = False
  .AllowRowSizing = False
  .TabAction = dbgControlNavigation
  .MarqueeStyle = dbgHighlightRow
 End With
End Sub

Public Sub Init_ControlDataGrid(EsteGrid As DataGrid)
 With EsteGrid
  .AllowAddNew = False
  .AllowDelete = False
  .AllowUpdate = False
  .AllowRowSizing = False
  .TabAction = dbgControlNavigation
  .MarqueeStyle = dbgHighlightRow
 ' .Font =
 End With
End Sub

Public Function Validar_RUC(xRuc As String) As Boolean
 Dim flag As Boolean
 Dim TAB_VAL(1 To 7) As Integer
 Dim nX As Integer, NY As Integer, NR As Integer, i As Integer
 Dim CadNR As String
 
' TAB_VAL(1) = 2
' TAB_VAL(2) = 7
' TAB_VAL(3) = 6
' TAB_VAL(4) = 5
' TAB_VAL(5) = 4
' TAB_VAL(6) = 3
' TAB_VAL(7) = 2
 flag = True
 xRuc = Trim(xRuc)
 
' If xRuc <> " " Then
  'If xRuc <> "00000002" Then
     If Len(RTrim(xRuc)) < 11 Then
         MsgBox "Número de R.U.C. no tiene 11 dígitos", vbExclamation, "Ingreso de Datos"
         flag = False
      Else
'         nX = 0
'         NR = 0
'         NY = 0
'         CadNR = ""
'         For i = 1 To 7
'             nX = nX + Val(Mid(xRuc, i, 1)) * TAB_VAL(i)
'         Next i
'         NY = nX \ 11
'         NR = 11 - (nX - (NY * 11))
'         CadNR = Trim(String(10 - Len(Str(NR)) + 1, "0")) & Trim(Str(NR))
'         If Mid(CadNR, 10, 1) = Mid(xRuc, 8, 1) Then
'            flag = True
''         Else
'            MsgBox "Número de R.U.C. invalido", vbExclamation, "Ingreso de Datos"
'            flag = False
'         End If
      End If
'   Else
'      MsgBox "Anexo emite Liquidaciones de compra", vbExclamation, "Ingreso de Datos"
 '  End If
 'End If
 Validar_RUC = flag
End Function
Public Sub AlinearAyuda(f As Form)
f.Left = FormPrincipal.Left + FormPrincipal.Width - f.Width
' f.Top = FormPrincipal.Height - FormPrincipal.ScaleHeight
f.Top = (Screen.Height - f.Height) / 2
End Sub

Public Sub AlinearFrm(f As Form)
 f.Left = FormPrincipal.Left + 50
 f.Top = FormPrincipal.Top + 50
End Sub

'función que encripta una cadena
Public Function CODIFICA(cadena As String, Valor As Integer) As String
 Dim ciclo As Integer, posic As Integer, ult_sal As Integer
 Dim carac As String, cadena_cod As String, cad As String
 Dim utl_sal As String
 posic = 0: ult_sal = 0
 carac = "": cadena_cod = "": cad = ""
 cadena = UCase(Trim(cadena))
 For ciclo = 1 To Len(cadena)
         carac = Mid(cadena, ciclo, 1)
         If (ciclo Mod 2) = 0 Then
            carac = UCase(carac)
        Else
            carac = LCase(carac)
        End If
        cadena_cod = cadena_cod & carac
 Next ciclo

 For ciclo = 1 To Len(cadena_cod)
     posic = ciclo Mod 7
    carac = Mid(cadena_cod, ciclo, 1)
    Select Case posic
    Case 0:
             carac = Chr(Asc(carac) * 2)
    Case 1:
            carac = Chr(Asc(carac) - Valor)
    Case 2:
            carac = Chr(Asc(carac) - (ciclo * 2))
            ult_sal = Asc(carac)
    Case 3:
            If ult_sal > 10 Then ult_sal = ult_sal - (Int(ult_sal / 10) * 10)
                carac = Chr(Asc(carac) - Valor + ult_sal)
    Case 4:
            carac = Chr(Asc(carac) - ciclo)
             utl_sal = Asc(carac)
    Case 5:
            If ult_sal > 10 Then ult_sal = ult_sal - (Int(ult_sal / 10) * 10)
                carac = Chr(Asc(carac) - Valor + ult_sal)
            End Select
            cad = cad + carac
 Next ciclo
 CODIFICA = cad
End Function

'función que desencripta una cadena
Public Function DECODIFICA(cadena As String, Valor As Integer) As String
 Dim ciclo As Integer, posic As Integer, val_n As Integer, val_an As Integer
 Dim carac As String, cad As String
 cadena = Trim(cadena)
 cad = ""
 val_n = 0: val_an = 0
 For ciclo = 1 To Len(cadena)
  carac = Mid(cadena, ciclo, 1)
  posic = ciclo Mod 7
  Select Case posic
  Case 0:
         val_n = Asc(carac) / 2
  Case 1:
         val_n = Asc(carac) + Valor
  Case 2:
         val_n = Asc(carac) + (ciclo * 2)
         val_an = Asc(carac)
  Case 3:
         If val_an > 10 Then val_an = val_an - (Int(val_an / 10) * 10)
         val_n = Asc(carac) + Valor - val_an
  Case 4:
         val_n = Asc(carac) + ciclo
  Case 5:
         If val_an > 10 Then val_an = val_an - (Int(val_an / 10) * 10)
         val_n = Asc(carac) + Valor - val_an
  Case 6:
        val_n = Asc(carac)
  End Select
  cad = cad + Chr(val_n)
 Next ciclo
 DECODIFICA = UCase(cad)
End Function

Function ValidFecha(vText As String) As String
Dim cTxtNew As String, ncnt As Integer
Dim cTxt As String, cTxtDig As String

cTxtDig = "": cTxtNew = ""
For ncnt = 1 To Len(vText)
      cTxt = Mid(vText, ncnt, 1)
      If cTxt = "/" Then
         cTxtNew = cTxtNew & Str(Val(cTxtDig)) & "/"
         cTxtDig = ""
      Else
         If cTxt <> "_" Then cTxtDig = cTxtDig & cTxt
      End If
Next
If cTxtDig <> "" Then cTxtNew = cTxtNew & Str(Val(cTxtDig))

If IsDate(cTxtNew) Then
   ValidFecha = Format(CDate(cTxtNew), "dd/mm/yyyy")
End If
End Function

Function ValidHora(vText As String) As String
Dim cTxtNew As String

cTxtNew = "01/01/74 " & vText
If IsDate(cTxtNew) Then
   ValidHora = Format(CDate(cTxtNew), "hh:mm")
Else
    ValidHora = "00:00"
End If
End Function

Function FValidFec(vText As String) As Boolean
Dim cTxtNew As String, ncnt As Integer
Dim cTxt As String, cTxtDig As String

If Day(vText) = Null Then
   FValidFec = False
Else
  If IsNull(Day(CDate(vText))) Then
     FValidFec = False
    Exit Function
  End If
  FValidFec = True
End If
'If IsDate(cTxtNew) Then
'   ValidFecha = Format(CDate(cTxtNew), "dd/mm/yyyy")
'Else
'
'End If
End Function

Function DesMes(nMes As String) As String
Dim DescriMes As String

Select Case nMes
   Case "01"
               DescriMes = "ENERO"
   Case "02"
               DescriMes = "FEBRERO  "
   Case "03"
               DescriMes = "MARZO"
   Case "04"
               DescriMes = "ABRIL"
    Case "05"
               DescriMes = "MAYO "
    Case "06"
               DescriMes = "JUNIO "
    Case "07"
               DescriMes = "JULIO "
    Case "08"
               DescriMes = "AGOSTO "
    Case "09"
               DescriMes = "SETIEMBRE "
    Case "10"
               DescriMes = "OCTUBRE "
    Case "11"
               DescriMes = "NOVIEMBRE "
    Case "12"
               DescriMes = "DICIEMBRE "
End Select

DesMes = DescriMes
End Function

Public Function codigo(cCod As String) As Boolean  ' Codigo del ARTICULO
Dim cSelC As ADODB.Recordset, csql As String
If Trim(cCod) = "" Then
    codigo = False
    Exit Function
End If
csql = "Select ACODIGO from MaeART where ACODIGO = '" & SupCadSQL(Trim(cCod)) & "'"
Set cSelC = New ADODB.Recordset
cSelC.Open csql, cConexCom, adOpenStatic
If cSelC.RecordCount > 0 Then
    codigo = False: cSelC.Close
    Exit Function
End If
codigo = True: cSelC.Close
End Function

Public Function Codigo2(cCod As String) As Boolean  'Codigo del FABRICANTE
Dim cSelC As ADODB.Recordset, csql As String
If Trim(cCod) = "" Then
    MsgBox "Falta Codigo", vbInformation, "Mensaje"
    Codigo2 = False
    Exit Function
End If
csql = "Select ACODIGO2 from MaeART where ACODIGO2 = '" & SupCadSQL(Trim(cCod)) & "'"
Set cSelC = New ADODB.Recordset
cSelC.Open csql, cConexCom, adOpenStatic
If cSelC.RecordCount > 0 Then
    Codigo2 = False: cSelC.Close
    Exit Function
End If
Codigo2 = True: cSelC.Close
End Function

Public Function CodigoC(cCod As String) As Boolean     'Codigo del Cliente
Dim cSelC As ADODB.Recordset, csql As String
If Trim(cCod) = "" Then
    MsgBox "Falta Codigo", vbInformation, "Mensaje"
    CodigoC = False
    Exit Function
End If
csql = "Select ccodcli from Maecli where ccodcli = '" & SupCadSQL(cCod) & "'"
Set cSelC = New ADODB.Recordset
cSelC.Open csql, cConexCom, adOpenStatic
If cSelC.RecordCount > 0 Then
    CodigoC = False: cSelC.Close
    Exit Function
End If
CodigoC = True: cSelC.Close
End Function

Public Function Existe(tipo As Integer, Cod As String, Tabla As String, Campo As String, Fecha As Boolean, Optional Cod2 As String, Optional cCampo2 As String, Optional Cod3 As String, Optional cCampo3 As String, Optional Cod4 As Boolean, Optional cCampo4 As String, Optional Cod5 As String, Optional cCampo5 As String) As Boolean
Dim cSel1 As ADODB.Recordset, cSL As String
Set cSel1 = New ADODB.Recordset

 If Fecha Then
        cSL = "Select * from " & Tabla & "  Where " & Campo & " =  #" & Format(Cod, "mm/dd/yyyy") & "#"
 Else
       If UCase(Tabla) = "PUNTO_VENTA" Then
                cSL = "Select * from " & Tabla & "  Where " & Campo & " =  '" & Cod & "'"
       Else
                cSL = "Select * from " & Tabla & "  Where " & Campo & " =  '" & Cod & "'"
       End If
       If Trim(Cod2) <> "" Then
            cSL = cSL & " And  " & cCampo2 & " =  '" & SupCadSQL(Cod2) & "'"
       End If
       If Trim(Cod3) <> "" Then
            cSL = cSL & " And  " & cCampo3 & " =  '" & SupCadSQL(Cod3) & "'"
       End If
       If Trim(cCampo4) <> "" Then
            If Cod4 = True Then
                cSL = cSL & " And  " & cCampo4
            Else
                cSL = cSL & " And  " & Not cCampo4
            End If
        End If
        If Trim(Cod5) <> "" Then
            cSL = cSL & " And  " & cCampo5 & " =  '" & Cod5 & "'"
        End If
 End If
 
Select Case tipo
Case 1 'Bd. Comun
            cSel1.Open cSL, cConexCom, adOpenStatic
Case 2 'Bd. Config
            cSel1.Open cSL, cConexConf, adOpenStatic
Case 3 'Bd. Contab
            cSel1.Open cSL, cConexCont, adOpenStatic
End Select

If cSel1.RecordCount > 0 Then
     Existe = True
Else
     Existe = False
End If
'csel1.Close
End Function
Public Sub Ubi_Tab(oT As CrystalReport)
Dim nI As Integer, nN As Integer
nN = oT.RetrieveDataFiles
For nI = 0 To nN
    If InStr(UCase(oT.DataFiles(nI)), "BDCOMUN") > 0 Then
        oT.DataFiles(nI) = cRuta2
    
    ElseIf InStr(UCase(oT.DataFiles(nI)), "BDAUXCOM") > 0 Then
        oT.DataFiles(nI) = App.Path & "\BDAUXCOM.MDB"
        
    ElseIf InStr(UCase(oT.DataFiles(nI)), "BDWENCO") > 0 Then     'Configuración
        oT.DataFiles(nI) = sName & "\BDWENCO.MDB"
    ElseIf InStr(UCase(oT.DataFiles(nI)), VGNameCont) > 0 Then
        If UCase(Dir$(cRuta4)) = UCase(VGNameCont & ".MDB") Then
             oT.DataFiles(nI) = cRutPath & "\" & VGNameCont & ".MDB"
        Else
             oT.DataFiles(nI) = cRuta2
        End If
    End If
Next nI
End Sub

'Posicionar la barra en el DataGrid
Public Function Pos_Dato(Adc As ADODB.Recordset) As Integer
Dim nN As Integer

Adc.MoveNext
If Not Adc.EOF Then
          nN = Adc.Bookmark - 1
Else
    Adc.MovePrevious
    Adc.MovePrevious
    If Not Adc.BOF Then
          nN = Adc.Bookmark
    End If
End If

Pos_Dato = nN
End Function

Public Function Pos_Dato1(Adc1 As Recordset, cCampo As String) As String
Dim cCodigo As String
Dim cCodigo1 As String
        
    cCodigo = Adc1(cCampo)
    Adc1.Delete
    Adc1.MoveNext
    If Adc1.EOF Then
       Adc1.MoveFirst
       If Adc1.BOF Then
       Else
         cCodigo1 = Adc1(cCampo)
       End If
    Else
      cCodigo1 = Adc1(cCampo)
    End If
 Pos_Dato1 = cCodigo1
End Function

Public Function fFam(cFam As String) As String        'FAMILIA
Dim cSqlA As String, cSelA As ADODB.Recordset
If Trim(cFam) = "" Then
    fFam = ""
    Exit Function
End If
cSqlA = "Select * FROM FAMILIA WHERE FAM_CODIGO = '" & Trim(cFam) & "' ORDER BY FAM_CODIGO "
Set cSelA = New ADODB.Recordset
cSelA.Open cSqlA, cConexCom, adOpenStatic
If cSelA.RecordCount = 0 Then
    fFam = "": cSelA.Close
    Exit Function
Else
    fFam = cSelA("FAM_NOMBRE")
End If
cSelA.Close
End Function

Public Function Val_Ayu(cAyu As String, cCodayu As String) As String
Dim cSqlA As String, cSelA As ADODB.Recordset

If Trim(cAyu) = "" Then
    Val_Ayu = ""
    Exit Function
End If

cSqlA = "Select * FROM TABAYU WHERE TCOD='" & cCodayu & "' And tClave = '" & Trim(cAyu) & "' ORDER BY TCLAVE "
Set cSelA = New ADODB.Recordset
cSelA.Open cSqlA, cConexCom, adOpenStatic
If cSelA.RecordCount = 0 Then
    Val_Ayu = "": cSelA.Close
    Exit Function
Else
    Val_Ayu = cSelA("tdescri")
End If
cSelA.Close
End Function

Public Function fPre(cPre As String) As String 'Precio
Dim cSqlA As String, cSelA As ADODB.Recordset
If Trim(cPre) = "" Then
    fPre = ""
    Exit Function
End If
cSqlA = "SELECT Cod_LisPre,Des_LisPre FROM TIPO_PRECIO where Cod_LisPre= '" & Trim(cPre) & "' ORDER BY Cod_LisPre"
Set cSelA = New ADODB.Recordset
cSelA.Open cSqlA, cConexCom, adOpenStatic
If cSelA.RecordCount = 0 Then
    fPre = "": cSelA.Close
    Exit Function
Else
    fPre = cSelA("Des_LisPre")
End If
cSelA.Close
End Function


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
            cSel1.Open cF, cConexCom, adOpenStatic
Case 2 'Bd. Config
            cSel1.Open cF, cConexConf, adOpenStatic
Case 3 'Bd. Contabilidad
            cSel1.Open cF, cConexCont, adOpenStatic
End Select

If cSel1.RecordCount > 0 Then
     Devolver_Dato = IIf(Not IsNull(cSel1(0)), cSel1(0), "")
Else
     Devolver_Dato = ""
End If
End Function

Public Function fDis(cDis As String) As String 'Distrito
Dim cSqlA As String, cSelA As ADODB.Recordset
If Trim(cDis) = "" Then
    fDis = ""
    Exit Function
End If
cSqlA = "Select * FROM TABAYU WHERE TCOD='13' And tClave = '" & Trim(cDis) & "' ORDER BY TCLAVE "
Set cSelA = New ADODB.Recordset
cSelA.Open cSqlA, cConexCom, adOpenStatic
If cSelA.RecordCount = 0 Then
    fDis = "": cSelA.Close
    Exit Function
Else
    fDis = cSelA("tdescri")
End If
cSelA.Close
End Function

Public Function fGir(cGir As String) As String 'Giro
Dim cSqlA As String, cSelA As ADODB.Recordset
If Trim(cGir) = "" Then
    fGir = ""
    Exit Function
End If
cSqlA = "Select * FROM TABAYU WHERE TCOD='62' And tClave = '" & Trim(cGir) & "' ORDER BY TCLAVE "
Set cSelA = New ADODB.Recordset
cSelA.Open cSqlA, cConexCom, adOpenStatic
If cSelA.RecordCount = 0 Then
    fGir = "": cSelA.Close
    Exit Function
Else
    fGir = cSelA("tdescri")
End If
cSelA.Close
End Function

Public Function NumPto(cKey As Integer) As Boolean
If (cKey < 48 Or cKey > 57) And cKey <> 46 And cKey <> 13 And cKey <> 8 Then
    NumPto = False
Else
    NumPto = True
End If
End Function

'Numeros sin pto. decimal
Public Function NumSpto(cKey As Integer) As Boolean
If (cKey < 48 Or cKey > 57) And cKey <> 13 And cKey <> 8 Then
    NumSpto = False
Else
    NumSpto = True
End If
End Function
Public Function fEqui(cEqui As String, cUni) As String       'UNIDAD DE EQUIVALENCIA
Dim cSqlA As String, cSelA As ADODB.Recordset
If Trim(cEqui) = "" Then
    fEqui = ""
    Exit Function
End If
cSqlA = "Select * FROM TABEQUI WHERE EQUNIEQUI = '" & Trim(cEqui) & "' AND EQUNIPRI = '" & Trim(cUni) & "' ORDER BY EQUNIEQUI"
Set cSelA = New ADODB.Recordset
cSelA.Open cSqlA, cConexCom, adOpenStatic
If cSelA.RecordCount = 0 Then
    fEqui = "": cSelA.Close
    Exit Function
Else
    fEqui = "WWWW"
End If
cSelA.Close
End Function

Public Function Last_Day(mes As Integer, Aa As Integer) As Integer
Dim dia As Integer
Last_Day = 0
 If mes > 0 And mes < 13 Then
  If Aa > 1000 Then
    Select Case mes
     Case 1, 3, 5, 7, 8, 10, 12:
        dia = 31
     Case 4, 6, 9, 11:
        dia = 30
     Case 2:
        If (Aa Mod 4) = 0 Then
         dia = 29
        Else
         dia = 28
        End If
    End Select
    Last_Day = dia
  End If
End If
End Function

'*************************************************
'Elimina de ( ' ) de una Cadena
'para Grabarla en una instrucción SQL
'*************************************************
Public Function SupCadSQL(S As String) As String
 Dim Aux As String
 If Not IsNull(S) Then
     Aux = Replace(S, "'", "''")
 End If
 SupCadSQL = Aux
 
End Function

Public Function prove(txt As TextBox) As String
 Dim rS As Recordset
 Dim rSql As String
   rSql = "select PRVCNOMBRE FROM maeprov where PRVCCODIGO= '" & txt & "'" '
   
   Set rS = VGBaseDatos.OpenRecordset(rSql, dbOpenSnapshot)
   If Not rS.EOF Then
      prove = rS(0)
   Else
     MsgBox "El codigo del proveedor no existe !", vbExclamation, "Error"
     prove = ""
  End If
  rS.Close
End Function

Public Function DateSQL(ByVal Fecha As Date) As String
    DateSQL = "#" & Month(Fecha) & "/" & Day(Fecha) & "/" & Year(Fecha) & "#"
End Function

Public Sub HabilitarMenu_Usuarios(Optional tipo As String, Optional Emp As String, Optional Nivel As String)
Dim ADOMenu As ADODB.Recordset
Dim ADOUsuA As ADODB.Recordset
Dim nArr() As Boolean
Dim nNum As Integer
Dim nUn As Integer

Set ADOMenu = New ADODB.Recordset
Set ADOUsuA = New ADODB.Recordset
If Not vGAdmLog Then
  ADOUsuA.Open "Select * From Men_Usu_Inv Where USU_CODIGO = '" & tipo & "' and EMP_CODIGO = '" & Emp & "' order by MEN_CODIGO", cConexConf, adOpenStatic
End If
ADOMenu.Open "Select * From Menu_Inv Order by Men_Codigo", cConexConf, adOpenStatic

nNum = 67  'Contiene la cantidad de opciones en el Menu (Si hay cambios aumentar o disminuir en número)

If nNum > ADOMenu.RecordCount Or nNum < ADOMenu.RecordCount Then
   ' MsgBox "La cantidad de opciones registradas en el programa no es igual a las de la tabla", vbInformation, "Verificar"
    'ADOUsuA.Close: ADOMenu.Close: Exit Sub
End If

ReDim nArr(1 To ADOMenu.RecordCount, 1 To 2)

For nUn = 1 To ADOMenu.RecordCount
            nArr(nUn, 1) = False
             nArr(nUn, 2) = False
Next nUn

If Nivel = "A1" Then   'administrador solo conf archivo salir
        nUn = 1
        Do While Not ADOMenu.EOF
                If ADOMenu("Men_Codigo") = "01" Or ADOMenu("Men_Codigo") = "0109" Or ADOMenu("Men_Codigo") = "07" Or ADOMenu("Men_Codigo") = "0701" _
                                    Or ADOMenu("Men_Codigo") = "070101" Or ADOMenu("Men_Codigo") = "070102" Or ADOMenu("Men_Codigo") = "070103" Or ADOMenu("Men_Codigo") = "0703" Then
                                            nArr(nUn, 1) = True
                                            nArr(nUn, 2) = ADOMenu("Men_Visible")
                Else
                                            nArr(nUn, 1) = False
                                            nArr(nUn, 2) = ADOMenu("Men_Visible")
                End If
                nUn = nUn + 1
                ADOMenu.MoveNext
                If ADOMenu.EOF Then Exit Do
        Loop
        ADOMenu.Close
        If ADOUsuA.State <> 0 Then
          ADOUsuA.Close
       End If
ElseIf Nivel = "A2" Or Nivel = "M" Then   'adminnistrador todas las opciones
        nUn = 1
        ADOMenu.MoveFirst
        Do While Not ADOMenu.EOF
                 nArr(nUn, 1) = True
                 nArr(nUn, 2) = ADOMenu("Men_Visible")
                 ADOMenu.MoveNext
                 If ADOMenu.EOF Then Exit Do
                 nUn = nUn + 1
        Loop
Else
        nUn = 1
        If ADOMenu.RecordCount > 0 Then
                Do While Not ADOMenu.EOF
                            If ADOUsuA.RecordCount > 0 Then
                                    ADOUsuA.MoveFirst
                                    ADOUsuA.Filter = "Men_Codigo = '" & ADOMenu("Men_Codigo") & "'"
                                    If Not ADOUsuA.EOF Then
                                            nArr(nUn, 1) = ADOUsuA("Men_Hab")
                                            nArr(nUn, 2) = ADOMenu("Men_Visible")
                                    Else
                                            nArr(nUn, 1) = False
                                            nArr(nUn, 2) = ADOMenu("Men_Visible")
                                    End If
                                    ADOUsuA.Filter = ""
                            Else
                                    nArr(nUn, 1) = False
                                    nArr(nUn, 2) = ADOMenu("Men_Visible")
                            End If
                            ADOMenu.MoveNext
                            If ADOMenu.EOF Then Exit Do
                            nUn = nUn + 1
                Loop
                
        Else
                For nUn = 1 To ADOMenu.RecordCount
                        nArr(nUn, 1) = False
                        nArr(nUn, 2) = ADOMenu("Men_Visible")
            Next nUn
        End If
        ADOMenu.Close
        If ADOUsuA.State <> 0 Then
           ADOUsuA.Close
       End If
End If

FormPrincipal.mant.Enabled = nArr(1, 1)
FormPrincipal.mant.Visible = nArr(1, 2)
FormPrincipal.Men_ManArt.Enabled = nArr(2, 1)
FormPrincipal.Men_ManArt.Visible = nArr(2, 2)
FormPrincipal.Men_mnulogistica.Enabled = nArr(3, 1)
FormPrincipal.Men_mnulogistica.Visible = nArr(3, 2)
FormPrincipal.Men_MantPro.Enabled = nArr(4, 1)
FormPrincipal.Men_MantPro.Visible = nArr(4, 2)
FormPrincipal.Men_MantClie.Enabled = nArr(5, 1)
FormPrincipal.Men_MantClie.Visible = nArr(5, 2)
FormPrincipal.Men_mnu_alma.Enabled = nArr(6, 1)
FormPrincipal.Men_mnu_alma.Visible = nArr(6, 2)
FormPrincipal.Men_ManTra.Enabled = nArr(7, 1)
FormPrincipal.Men_ManTra.Visible = nArr(7, 2)
FormPrincipal.Men_mnucasillero.Enabled = nArr(8, 1)
FormPrincipal.Men_mnucasillero.Visible = nArr(8, 2)
FormPrincipal.Men_ManAyu.Enabled = nArr(9, 1)
FormPrincipal.Men_ManAyu.Visible = nArr(9, 2)
FormPrincipal.mnu_unidades_02.Enabled = nArr(10, 1)
FormPrincipal.mnu_unidades_02.Visible = nArr(10, 2)
FormPrincipal.Men_ayuFam_03.Enabled = nArr(11, 1)
FormPrincipal.Men_ayuFam_03.Visible = nArr(11, 2)
FormPrincipal.mnu_auto_05.Enabled = nArr(12, 1)
FormPrincipal.mnu_auto_05.Visible = nArr(12, 2)
'FormPrincipal.mnu_salir.Enabled = nArr(13, 1)
'FormPrincipal.mnu_salir.Visible = nArr(13, 2)
FormPrincipal.Men_mnutransn.Enabled = nArr(14, 1)
FormPrincipal.Men_mnutransn.Visible = nArr(14, 2)
FormPrincipal.mnulistado.Enabled = nArr(15, 1)
FormPrincipal.mnulistado.Visible = nArr(15, 2)
FormPrincipal.mnu_catarticulo.Enabled = nArr(16, 1)
FormPrincipal.mnu_catarticulo.Visible = nArr(16, 2)
FormPrincipal.mnu_catclirente.Enabled = nArr(17, 1)
FormPrincipal.mnu_catclirente.Visible = nArr(17, 2)
FormPrincipal.mnu_catproveed.Enabled = nArr(18, 1)
FormPrincipal.mnu_catproveed.Visible = nArr(18, 2)
FormPrincipal.Tra.Enabled = nArr(19, 1)
FormPrincipal.Tra.Visible = nArr(19, 2)
FormPrincipal.Men_TraRegEnt.Enabled = nArr(20, 1)
FormPrincipal.Men_TraRegEnt.Visible = nArr(20, 2)
FormPrincipal.Men_TraRegSal.Enabled = nArr(21, 1)
FormPrincipal.Men_TraRegSal.Visible = nArr(21, 2)
FormPrincipal.Men_mnGui.Enabled = nArr(22, 1)
FormPrincipal.Men_mnGui.Visible = nArr(22, 2)
FormPrincipal.mnu_repIASA.Enabled = nArr(23, 1)
FormPrincipal.mnu_repIASA.Visible = nArr(23, 2)
FormPrincipal.mnu_recepcion.Enabled = nArr(24, 1)
FormPrincipal.mnu_recepcion.Visible = nArr(24, 2)
FormPrincipal.mnucons.Enabled = nArr(25, 1)
FormPrincipal.mnucons.Visible = nArr(25, 2)
FormPrincipal.mnu_stkArt1.Enabled = nArr(26, 1)
FormPrincipal.mnu_stkArt1.Visible = nArr(26, 2)
FormPrincipal.mnu_conValArtPend.Enabled = nArr(27, 1)
FormPrincipal.mnu_conValArtPend.Visible = nArr(27, 2)
FormPrincipal.mnu_provart.Enabled = nArr(28, 1)
FormPrincipal.mnu_provart.Visible = nArr(28, 2)
FormPrincipal.mnu_docvalorizado.Enabled = nArr(29, 1)
FormPrincipal.mnu_docvalorizado.Visible = nArr(29, 2)
FormPrincipal.mnu_movart.Enabled = nArr(30, 1)
FormPrincipal.mnu_movart.Visible = nArr(30, 2)
FormPrincipal.mnurep.Enabled = nArr(31, 1)
FormPrincipal.mnurep.Visible = nArr(31, 2)
FormPrincipal.Men_RepAlm.Enabled = nArr(32, 1)
FormPrincipal.Men_RepAlm.Visible = nArr(32, 2)
FormPrincipal.Men_AlmStock_01.Enabled = nArr(33, 1)
FormPrincipal.Men_AlmKar_02.Enabled = nArr(34, 1)
FormPrincipal.Men_AlmKar_02.Visible = nArr(34, 2)
FormPrincipal.Men_InvMovKar_03.Enabled = False                     'nArr(35, 1)
FormPrincipal.Men_InvMovKar_03.Visible = False      ' nArr(35, 2)
FormPrincipal.mnu_artxven_05.Enabled = nArr(36, 1)
FormPrincipal.mnu_artxven_05.Visible = nArr(36, 2)
FormPrincipal.mnu_artven_06.Enabled = nArr(37, 1)
FormPrincipal.mnu_artven_06.Visible = nArr(37, 2)
FormPrincipal.Men_RepVal.Enabled = nArr(38, 1)
FormPrincipal.Men_RepVal.Visible = nArr(38, 2)
FormPrincipal.Men_InvKarVal_01.Enabled = nArr(39, 1)
FormPrincipal.Men_InvKarVal_01.Visible = nArr(39, 2)
FormPrincipal.mnu_valxdoc_03.Enabled = nArr(40, 1)
FormPrincipal.mnu_valxdoc_03.Visible = nArr(40, 2)
FormPrincipal.mnu_rentabilidad.Enabled = nArr(41, 1)
FormPrincipal.mnu_rentabilidad.Visible = nArr(41, 2)
FormPrincipal.mnu_repo.Enabled = nArr(42, 1)
FormPrincipal.mnu_repo.Visible = nArr(42, 2)
FormPrincipal.mnu_rotación.Enabled = nArr(43, 1)
FormPrincipal.mnu_rotación.Visible = nArr(43, 2)
FormPrincipal.mnu_guiaIngIasa.Enabled = nArr(44, 1)
FormPrincipal.mnu_guiaIngIasa.Visible = nArr(44, 2)
FormPrincipal.mnu_catart.Enabled = nArr(45, 1)
FormPrincipal.mnu_catart.Visible = nArr(45, 2)
FormPrincipal.Pro.Enabled = nArr(46, 1)
FormPrincipal.Pro.Visible = nArr(46, 2)
FormPrincipal.Men_ProVal.Enabled = nArr(47, 1)
FormPrincipal.Men_ProVal.Visible = nArr(47, 2)
FormPrincipal.Men_ProCieMen_01.Enabled = nArr(48, 1)
FormPrincipal.Men_ProCieMen_01.Visible = nArr(48, 2)
FormPrincipal.Men_ProEsp.Enabled = nArr(49, 1)
FormPrincipal.Men_ProEsp.Visible = nArr(49, 2)
FormPrincipal.Men_EspInvFis_01.Enabled = nArr(50, 1)
FormPrincipal.Men_EspInvFis_01.Visible = nArr(50, 2)

FormPrincipal.mnu_Asiento_02.Enabled = nArr(51, 1)  '' Asiento
FormPrincipal.mnu_Asiento_02.Visible = nArr(51, 2)

FormPrincipal.Men_GuiRem.Enabled = nArr(52, 1)
FormPrincipal.Men_GuiRem.Visible = nArr(52, 2)
FormPrincipal.Men_GuiEli_01.Enabled = nArr(53, 1)
FormPrincipal.Men_GuiEli_01.Visible = nArr(53, 2)
FormPrincipal.Men_GuiDev_02.Enabled = nArr(54, 1)
FormPrincipal.Men_GuiDev_02.Visible = nArr(54, 2)
FormPrincipal.Men_GuiDoc.Enabled = nArr(55, 1)
FormPrincipal.Men_GuiDoc.Visible = nArr(55, 2)
FormPrincipal.Men_CocMod_01.Enabled = nArr(56, 1)
FormPrincipal.Men_CocMod_01.Visible = nArr(56, 2)
FormPrincipal.Men_EliDoc_02.Enabled = nArr(57, 1)
FormPrincipal.Men_EliDoc_02.Visible = nArr(57, 2)
FormPrincipal.mnu_ajuste.Enabled = nArr(58, 1)
FormPrincipal.mnu_ajuste.Visible = nArr(58, 2)
FormPrincipal.Men_TraVal.Enabled = nArr(59, 1)
FormPrincipal.Men_TraVal.Visible = nArr(59, 2)
FormPrincipal.Men_TraCor.Enabled = nArr(60, 1)
FormPrincipal.Men_TraCor.Visible = nArr(60, 2)
'FormPrincipal.sis.Enabled = nArr(60, 1)
'FormPrincipal.sis.Visible = nArr(60, 2)
'FormPrincipal.Men_SisCrea.Enabled = nArr(61, 1)
'FormPrincipal.Men_SisCrea.Visible = nArr(61, 2)
'FormPrincipal.mnu_Emp_01.Enabled = nArr(62, 1)
'FormPrincipal.mnu_Emp_01.Visible = nArr(62, 2)
'FormPrincipal.Men_SisUsu_02.Enabled = nArr(63, 1)
'FormPrincipal.Men_SisUsu_02.Visible = nArr(63, 2)
'FormPrincipal.Men_SisAdminis_03.Enabled = nArr(64, 1)
'FormPrincipal.Men_SisAdminis_03.Visible = nArr(64, 2)
FormPrincipal.Men_SisPar.Enabled = nArr(65, 1)
FormPrincipal.Men_SisPar.Visible = nArr(65, 2)
'FormPrincipal.Men_SisCam.Enabled = nArr(66, 1)
'FormPrincipal.Men_SisCam.Visible = nArr(66, 2)
'FormPrincipal.Men_SisAl.Enabled = nArr(67, 1)
'FormPrincipal.Men_SisAl.Visible = nArr(67, 2)
End Sub

Public Sub Verificar_Sistema()
Dim Seguridad As ProcSistema        'Definición de la Variables del Sistema
Set Seguridad = New ProcSistema
If Not App.PrevInstance Then        'Comprobar instancia previa
    If Not Seguridad.PrVerifSeg Then End 'COMPROBACIÓN DEL SISTEMA (Copia)
    'Si es False  - Error
    'Si es True   - Pasa
Else
  'Solo una instancia
  End
End If
End Sub
Public Function FechS(Fecha As Variant, tipo As TIPFECHA) As Variant
Dim H As Date
Dim fechaAux As Double
On Error GoTo ErrorFecha
   H = CDate(Fecha)
   Select Case tipo
      Case Sqlf: 'Para transformar al sql
         fechaAux = DateSerial(Year(Fecha), Month(Fecha), Day(Fecha)) - 2
      Case Adof: 'Para transformar al ado
         fechaAux = DateSerial(Year(Fecha), Month(Fecha), Day(Fecha))
   End Select
   FechS = fechaAux
   Exit Function
ErrorFecha:
   Select Case tipo
      Case Sqlf: FechS = "Null"
      Case Adof: FechS = Null
   End Select
End Function
Public Function ExisteElem(Tip As Integer, Cn As ADODB.Connection, Tabla As String, _
        Optional Campo As String) As Boolean
'Funcion que devuelve un valor verdadero si es que encuentra el elemento
'Creado por Fernando Cossio
    Dim SQL As String
    Dim RsAux As New ADODB.Recordset
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
    RsAux.Open SQL, Cn
    ExisteElem = True
    Exit Function
ErrExiste:
    ExisteElem = False
End Function

Function ExisteTabla(ByVal nombretabla As String, ByVal ADOConnection As ADODB.Connection) As Boolean
    Dim RsTbls As New ADODB.Recordset
    Set RsTbls = ADOConnection.OpenSchema(adSchemaTables)
    RsTbls.Find "[Table_Name]='" & nombretabla & "'"
    If RsTbls.EOF Then ExisteTabla = False Else ExisteTabla = True
End Function

Public Function AGREGARBASE(codigo As String) As Boolean
On Error GoTo errfil
    AGREGARBASE = True
    Screen.MousePointer = 11
    If UCase(Dir$(sName & "Data\" & codigo & "\" & "BdComun.mdb")) <> "BDCOMUN.MDB" Then
        FileCopy sName & "Bdplanti.mdk", sName & "Data\" & codigo & "\" & "BDComun.mdb"
        FileCopy sName & "BdTransf.mdk", sName & "Data\" & codigo & "\" & "BDTransf.mdb"
        MsgBox "Proceso Terminado", vbInformation, "Información"
    Else
        MsgBox "Proceso ya ha sido Generado", vbInformation, "Información"
    End If
    Screen.MousePointer = 1
Exit Function
errfil:
    AGREGARBASE = False
    Select Case err.Number
        Case 53
            MsgBox "No se encontraron las plantillas(BDPLANTI ó BDTRANSF) en la ruta especificada en archivo Invetarios.Ini" & Chr(13) & _
                   "Ruta especificada: " & sName
        Case 76
            MsgBox "No se encuentra la carpeta """ & codigo & """ de la empresa especificada" & Chr(13) & _
                   "en la ruta: " & sName & "DATA\"
        Case Else
           MsgBox err.Description
    End Select
End Function
Public Sub ActualizaBD()
 Dim SQL As String
  If Not ExisteElem(0, cConexCom, "KITS") Then
      SQL = " Create Table KITS (CODART Text(20),CODKIT Text(20), " & _
      "  CANART double)"
      cConexCom.Execute SQL
  End If
  If Not ExisteElem(1, cConexCom, "MAEART", "AMARCA") Then
      cConexCom.Execute "ALTER TABLE MAEART ADD COLUMN   AMARCA  TEXT(20)" '
  End If
  If Not ExisteElem(1, cConexCom, "MAEART", "ACOLOR") Then
      cConexCom.Execute "ALTER TABLE  MAEART  ADD COLUMN   ACOLOR  TEXT(20)" '
  End If
   If Not ExisteElem(0, cConexCom, "MAECOLOR") Then
        SQL = " Create Table MAECOLOR (COD_COLOR Text(20),DESCRI_COLOR Text(20), " & _
        " CONSTRAINT Clave PRIMARY KEY (COD_COLOR))"
        cConexCom.Execute SQL
  End If
  If Not ExisteElem(0, cConexCom, "MAEMARCA") Then
        SQL = " Create Table MAEMARCA (COD_MARCA Text(20),DESCRI_MARCA Text(20), " & _
        " CONSTRAINT Clave PRIMARY KEY (COD_MARCA))"
        cConexCom.Execute SQL
  End If
'  If Not ExisteElem(1, cConexCom, "FAMILIA", "FAM_COMPRA") Then
'      Conexion.Execute "ALTER TABLE " & familia & " ADD COLUMN  " & FAM_COMPRA & " TEXT(20)" '
'  End If
  
End Sub
