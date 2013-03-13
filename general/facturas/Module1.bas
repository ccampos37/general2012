Attribute VB_Name = "Module1"
Option Explicit
Public VGSeleccion As Integer    ' modificar o adicionar. descripcion en el formulario registro
'Fernando: 04/09/2001:
Public VGRUCEMP As String        ' Ruc de la empresa
'***
Public VGServer As String
Public VGBase As String
Public VGBUsuario As String
Public VGPassw As String

Public VGServer2 As String
Public VGBase2 As String
Public VGBUsuario2 As String
Public VGPassw2 As String

Public VGDIRE As String
Public VGBase3 As String

Public cn As New ADODB.Connection
Public GPunto As String * 2  'punto de venta

'*********

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
'Public VGWrk As Workspace
'Public VGBaseDatos As Database
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
Public VG_FecTrab As Date
Public VGcod As String                          'Se utiliza para las consultas
Public vGUtil(4) As String                        'Se para los pases de ayuda
Public arrayserie()   As String                'Ingreso masivo de serie
Public Enum TIPFECHA
   Sqlf = 1
   Adof = 2
End Enum
'RMM**************************************
Public ClsTock As New ClasStock
Public ClsTDoc As New ClasDocumento
Public VGTip_Alma As String
Public VGLadrillera As Boolean
Dim cConexAux As New ADODB.Connection


'Variables de Ventana de Ayuda
Public nAyuda As String
Public nDetalle As String

Public g_ptoventa As String

Public Const g_TipoSol = "01"
Public Const g_TipoDolar = "02"





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
 Dim nX As Integer, NY As Integer, NR As Integer, I As Integer
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
f.Left = FrmPrincipal.Left + FrmPrincipal.Width - f.Width
' f.Top = FrmPrincipal.Height - FrmPrincipal.ScaleHeight
f.Top = (Screen.Height - f.Height) / 2
End Sub

Public Sub AlinearFrm(f As Form)
 f.Left = FrmPrincipal.Left + 50
 f.Top = FrmPrincipal.Top + 50
End Sub

'función que encripta una cadena
Public Function CODIFICA(cadena As String, VALOR As Integer) As String
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
            carac = Chr(Asc(carac) - VALOR)
    Case 2:
            carac = Chr(Asc(carac) - (ciclo * 2))
            ult_sal = Asc(carac)
    Case 3:
            If ult_sal > 10 Then ult_sal = ult_sal - (Int(ult_sal / 10) * 10)
                carac = Chr(Asc(carac) - VALOR + ult_sal)
    Case 4:
            carac = Chr(Asc(carac) - ciclo)
             utl_sal = Asc(carac)
    Case 5:
            If ult_sal > 10 Then ult_sal = ult_sal - (Int(ult_sal / 10) * 10)
                carac = Chr(Asc(carac) - VALOR + ult_sal)
            End Select
            cad = cad + carac
 Next ciclo
 CODIFICA = cad
End Function

'función que desencripta una cadena
Public Function DECODIFICA(cadena As String, VALOR As Integer) As String
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
         val_n = Asc(carac) + VALOR
  Case 2:
         val_n = Asc(carac) + (ciclo * 2)
         val_an = Asc(carac)
  Case 3:
         If val_an > 10 Then val_an = val_an - (Int(val_an / 10) * 10)
         val_n = Asc(carac) + VALOR - val_an
  Case 4:
         val_n = Asc(carac) + ciclo
  Case 5:
         If val_an > 10 Then val_an = val_an - (Int(val_an / 10) * 10)
         val_n = Asc(carac) + VALOR - val_an
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
csql = "Select ACODIGO,adescri from MaeART where ACODIGO = '" & SupCadSQL(Trim(cCod)) & "'"
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
        cSL = "Select * from " & Tabla & "  Where " & Campo & " =  '" & Cod & "'"
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
'nN = oT.RetrieveDataFiles
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

Public Function fMost(cMost As String) As String                        'corregir
Dim cSqlA As String, cSelA As ADODB.Recordset
If Trim(cMost) = "" Then
    fMost = ""
    Exit Function
End If
cSqlA = "SELECT f.FAM_NOMBRE, l.LIN_NOMBRE"
cSqlA = cSqlA & " FROM MAEART m INNER JOIN (FAMILIA f INNER JOIN LINEAS l ON f.FAM_CODIGO=l.FAM_CODIGO) ON (m.AFAMILIA = f.FAM_CODIGO) AND (m.AMODELO=l.LIN_CODIGO)"
cSqlA = cSqlA & " WHERE m.ACODIGO='7895'"
Set cSelA = New ADODB.Recordset
cSelA.Open cSqlA, cConexCom, adOpenStatic
If cSelA.RecordCount = 0 Then
    fMost = "": cSelA.Close
    Exit Function
Else
    fMost = cSelA("f.FAM_NOMBRE") + ", " + cSelA("l.LIN_NOMBRE") + ", " + cSelA("g.GRU_NOMBRE") + ", " + cSelA("m.ACOLOR") + ", " + cSelA("m.AMARCA")
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
 Dim rs As New ADODB.Recordset
 Dim rsql As String
   rsql = "select PRVCNOMBRE FROM maeprov where PRVCCODIGO= '" & txt & "'" '
   
   'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
   Set rs = cConexCom.Execute(rsql)
   If Not rs.EOF Then
      prove = rs(0)
   Else
     MsgBox "El codigo del proveedor no existe !", vbExclamation, "Error"
     prove = ""
  End If
  rs.Close
End Function

Public Function DateSQL(ByVal Fecha As Date) As String
    DateSQL = "'" & Month(Fecha) & " / " & Day(Fecha) & " / " & Year(Fecha) & "'"
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

FrmPrincipal.mant.Enabled = nArr(1, 1)
FrmPrincipal.mant.Visible = nArr(1, 2)
FrmPrincipal.Men_ManArt.Enabled = nArr(2, 1)
FrmPrincipal.Men_ManArt.Visible = nArr(2, 2)
FrmPrincipal.Men_mnulogistica.Enabled = nArr(3, 1)
FrmPrincipal.Men_mnulogistica.Visible = nArr(3, 2)
FrmPrincipal.Men_MantPro.Enabled = nArr(4, 1)
FrmPrincipal.Men_MantPro.Visible = nArr(4, 2)
FrmPrincipal.Men_MantClie.Enabled = nArr(5, 1)
FrmPrincipal.Men_MantClie.Visible = nArr(5, 2)
FrmPrincipal.Men_mnu_alma.Enabled = nArr(6, 1)
FrmPrincipal.Men_mnu_alma.Visible = nArr(6, 2)
FrmPrincipal.Men_ManTra.Enabled = nArr(7, 1)
FrmPrincipal.Men_ManTra.Visible = nArr(7, 2)
FrmPrincipal.Men_mnucasillero.Enabled = nArr(8, 1)
FrmPrincipal.Men_mnucasillero.Visible = nArr(8, 2)
FrmPrincipal.Men_ManAyu.Enabled = nArr(9, 1)
FrmPrincipal.Men_ManAyu.Visible = nArr(9, 2)
FrmPrincipal.mnu_unidades_02.Enabled = nArr(10, 1)
FrmPrincipal.mnu_unidades_02.Visible = nArr(10, 2)
FrmPrincipal.Men_ayuFam_03.Enabled = nArr(11, 1)
FrmPrincipal.Men_ayuFam_03.Visible = nArr(11, 2)
FrmPrincipal.mnu_auto_05.Enabled = nArr(12, 1)
FrmPrincipal.mnu_auto_05.Visible = nArr(12, 2)
FrmPrincipal.Men_mnutransn.Enabled = nArr(14, 1)
FrmPrincipal.Men_mnutransn.Visible = nArr(14, 2)
FrmPrincipal.mnucons.Enabled = nArr(25, 1)
FrmPrincipal.mnucons.Visible = nArr(25, 2)
FrmPrincipal.mnu_stkArt1.Enabled = nArr(26, 1)
FrmPrincipal.mnu_stkArt1.Visible = nArr(26, 2)
FrmPrincipal.mnu_conValArtPend.Enabled = nArr(27, 1)
FrmPrincipal.mnu_conValArtPend.Visible = nArr(27, 2)
FrmPrincipal.mnu_provart.Enabled = nArr(28, 1)
FrmPrincipal.mnu_provart.Visible = nArr(28, 2)
FrmPrincipal.mnu_docvalorizado.Enabled = nArr(29, 1)
FrmPrincipal.mnu_docvalorizado.Visible = nArr(29, 2)
FrmPrincipal.mnu_movart.Enabled = nArr(30, 1)
FrmPrincipal.mnu_movart.Visible = nArr(30, 2)
FrmPrincipal.mnurep.Enabled = nArr(31, 1)
FrmPrincipal.mnurep.Visible = nArr(31, 2)
FrmPrincipal.Men_RepAlm.Enabled = nArr(32, 1)
FrmPrincipal.Men_RepAlm.Visible = nArr(32, 2)
FrmPrincipal.Men_AlmStock_01.Enabled = nArr(33, 1)
FrmPrincipal.Men_AlmKar_02.Enabled = nArr(34, 1)
FrmPrincipal.Men_AlmKar_02.Visible = nArr(34, 2)
FrmPrincipal.Men_InvMovKar_03.Enabled = False                     'nArr(35, 1)
FrmPrincipal.Men_InvMovKar_03.Visible = False      ' nArr(35, 2)
FrmPrincipal.mnu_artven_06.Enabled = nArr(37, 1)
FrmPrincipal.mnu_artven_06.Visible = nArr(37, 2)
FrmPrincipal.Men_RepVal.Enabled = nArr(38, 1)
FrmPrincipal.Men_RepVal.Visible = nArr(38, 2)
FrmPrincipal.Men_InvKarVal_01.Enabled = nArr(39, 1)
FrmPrincipal.Men_InvKarVal_01.Visible = nArr(39, 2)
FrmPrincipal.mnu_valxdoc_03.Enabled = nArr(40, 1)
FrmPrincipal.mnu_valxdoc_03.Visible = nArr(40, 2)
FrmPrincipal.Pro.Enabled = nArr(46, 1)
FrmPrincipal.Pro.Visible = nArr(46, 2)
FrmPrincipal.Men_ProVal.Enabled = nArr(47, 1)
FrmPrincipal.Men_ProVal.Visible = nArr(47, 2)
FrmPrincipal.Men_ProCieMen_01.Enabled = nArr(48, 1)
FrmPrincipal.Men_ProCieMen_01.Visible = nArr(48, 2)
FrmPrincipal.Men_ProEsp.Enabled = nArr(49, 1)
FrmPrincipal.Men_ProEsp.Visible = nArr(49, 2)
FrmPrincipal.Men_EspInvFis_01.Enabled = nArr(50, 1)
FrmPrincipal.Men_EspInvFis_01.Visible = nArr(50, 2)

FrmPrincipal.mnu_Asiento_02.Enabled = nArr(51, 1)  '' Asiento
FrmPrincipal.mnu_Asiento_02.Visible = nArr(51, 2)

FrmPrincipal.Men_GuiRem.Enabled = nArr(52, 1)
FrmPrincipal.Men_GuiRem.Visible = nArr(52, 2)
FrmPrincipal.Men_GuiEli_01.Enabled = nArr(53, 1)
FrmPrincipal.Men_GuiEli_01.Visible = nArr(53, 2)
FrmPrincipal.Men_GuiDev_02.Enabled = nArr(54, 1)
FrmPrincipal.Men_GuiDev_02.Visible = nArr(54, 2)
FrmPrincipal.Men_GuiDoc.Enabled = nArr(55, 1)
FrmPrincipal.Men_GuiDoc.Visible = nArr(55, 2)
FrmPrincipal.Men_CocMod_01.Enabled = nArr(56, 1)
FrmPrincipal.Men_CocMod_01.Visible = nArr(56, 2)
FrmPrincipal.Men_EliDoc_02.Enabled = nArr(57, 1)
FrmPrincipal.Men_EliDoc_02.Visible = nArr(57, 2)
FrmPrincipal.mnu_ajuste.Enabled = nArr(58, 1)
FrmPrincipal.mnu_ajuste.Visible = nArr(58, 2)
FrmPrincipal.Men_TraVal.Enabled = nArr(59, 1)
FrmPrincipal.Men_TraVal.Visible = nArr(59, 2)
FrmPrincipal.Men_TraCor.Enabled = nArr(60, 1)
FrmPrincipal.Men_TraCor.Visible = nArr(60, 2)
FrmPrincipal.Men_SisPar.Enabled = nArr(65, 1)
FrmPrincipal.Men_SisPar.Visible = nArr(65, 2)

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
Public Function ExisteElem(Tip As Integer, cn As ADODB.Connection, Tabla As String, _
        Optional Campo As String) As Boolean
'Funcion que devuelve un valor verdadero si es que encuentra el elemento
'Creado por Fernando Cossio
    Dim SQL As String
    Dim RSAUX As New ADODB.Recordset
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
    RSAUX.Open SQL, cn
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
    Select Case Err.Number
        Case 53
            MsgBox "No se encontraron las plantillas(BDPLANTI ó BDTRANSF) en la ruta especificada en archivo Invetarios.Ini" & Chr(13) & _
                   "Ruta especificada: " & sName
        Case 76
            MsgBox "No se encuentra la carpeta """ & codigo & """ de la empresa especificada" & Chr(13) & _
                   "en la ruta: " & sName & "DATA\"
        Case Else
           MsgBox Err.Description
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
 '*****************************************************************
 '*** ULTIMA ACTUALIZACION 28/06/2001    ROBERTO M.M.
 '*****************************************************************
   If Not ExisteElem(0, cConexCom, "CIERRMESVALOR") Then
        SQL = " Create Table CIERRMESVALOR (CIERRMES Text(6),CIERRALMA Text(2),CIERRFECH DATETIME, CIERROPER TEXT(15) , " & _
        " CONSTRAINT Clave PRIMARY KEY (CIERRMES))"
        cConexCom.Execute SQL
  Else
       If Not ExisteElem(1, cConexCom, "CIERRMESVALOR", "CIERRALMA") Then
           cConexCom.Execute "ALTER TABLE  CIERRMESVALOR  ADD COLUMN  CIERRALMA text(2) " '
       End If
  End If
  
  If Not ExisteElem(1, cConexCom, "MORESMES", "SMSALDOINI") Then
      cConexCom.Execute "ALTER TABLE  MORESMES  ADD COLUMN  SMSALDOINI  DOUBLE " '
  End If
   
  If Not ExisteElem(0, cConexCom, "COSPROFECH") Then
        SQL = " Create Table COSPROFECH ( AUXALMA Text(3),AUXTD Text(3),AUXNUMDOC Text(10),AUXCODART Text(20) ,AUXFECDOC DATETIME,AUXCANT DOUBLE,AUXPRECIO DOUBLE,AUXPRECOS DOUBLE   )" '(AUXTD , AUXNUMDOC , AUXCODART , AUXFECDOC )
        cConexCom.Execute SQL
  End If
  
  
  If Not ExisteElem(1, cConexCom, "KARDEXAUX", "TIPDOCRF") Then
     cConexCom.Execute "ALTER TABLE  KARDEXAUX  ADD COLUMN  TIPDOCRF text(2) " '
  End If
  
  If Not ExisteElem(1, cConexCom, "KARDEXAUX", "NUMDOCRF") Then
     cConexCom.Execute "ALTER TABLE  KARDEXAUX  ADD COLUMN  NUMDOCRF text(10) " '
  End If
  
  If Not ExisteElem(1, cConexCom, "KARDEXAUX", "NOMREFE") Then
     cConexCom.Execute "ALTER TABLE  KARDEXAUX  ADD COLUMN  NOMREFE text(50) " '
  End If
  
  Call ADOConectar
  
  If Not ExisteElem(1, cConexAux, "al_Kardex_Val", "ING_SAL") Then
     cConexAux.Execute "ALTER TABLE  al_Kardex_Val  ADD COLUMN  ING_SAL TEXT(20) " '
  End If
  
  cConexAux.Close
  
  If Not ExisteElem(0, cConexCom, "InveFisiCab") Then
        SQL = " Create Table InveFisiCab ( AUXNUMINVE TEXT(10), AUXALMA Text(3),AUXFECH DATETIME ,AUXRESPON TEXT(15),AUXOBSER TEXT(255)" & _
        ", CONSTRAINT Clave PRIMARY KEY ( AUXALMA,AUXNUMINVE )  )"
        cConexCom.Execute SQL
  Else
       If Not ExisteElem(1, cConexCom, "InveFisiCab", "AUXESTADO") Then
          cConexCom.Execute "ALTER TABLE  InveFisiCab  ADD COLUMN  AUXESTADO TEXT(2) " '*****INDICA SI YA SE INGRESO O ESTA PENDIENTE EN INVENTARIO FISICO
       End If
  End If

  
  If Not ExisteElem(0, cConexCom, "InveFisiDet") Then
        SQL = " Create Table InveFisiDet ( AUXNUMINVE TEXT(10), AUXALMA Text(3), AUXFAMIL Text(8),AUXCODART Text(20) ,AUXSTOCK DOUBLE,AUXINGR DOUBLE,AUXDIFE DOUBLE " & _
        ", CONSTRAINT Clave PRIMARY KEY ( AUXALMA , AUXNUMINVE,AUXCODART )  )"
        cConexCom.Execute SQL
  Else
       If Not ExisteElem(1, cConexCom, "InveFisiDet", "AUXFAMIL") Then
          cConexCom.Execute "ALTER TABLE  InveFisiDet  ADD COLUMN  AUXFAMIL Text(8) " '*****INDICA SI YA SE INGRESO O ESTA PENDIENTE EN INVENTARIO FISICO
       End If
        
  End If

  If Not ExisteElem(1, cConexCom, "CONFIGURACION", "conf_codigoIng") Then
     cConexCom.Execute "ALTER TABLE CONFIGURACION  ADD COLUMN  conf_codigoIng Text(8) "
  End If
  
  
 '*****************************************************************
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
  
   If ExisteIndice(cRuta2, "InveFisiDet", "clave") Then
     EliminaIndice cRuta2, "InveFisiDet", "clave"
     CreaIndice cRuta2, "InveFisiDet", "PRIMARYKEY", True, "AUXNUMINVE", "AUXALMA", "AUXCODART"
  End If
     
  
'  If Not ExisteElem(1, cConexCom, "FAMILIA", "FAM_COMPRA") Then
'      Conexion.Execute "ALTER TABLE " & familia & " ADD COLUMN  " & FAM_COMPRA & " TEXT(20)" '
'  End If
 
 'Dim Dtb As Database
' Dim Tdf As TableDef
' Dim Campo As Field

 'Set Dtb = OpenDatabase(cRuta2)
' Set Tdf = Dtb.TableDefs("MAEART")
' Tdf.Fields("AMARCA").AllowZeroLength = True
' Tdf.Fields("ACOLOR").AllowZeroLength = True
'
' 'Set Dtb = OpenDatabase(cRuta2)
' Set Tdf = Dtb.TableDefs("InveFisiCab")
' Tdf.Fields("AUXOBSER").AllowZeroLength = True
'
' Set Dtb = OpenDatabase(cRuta2)
' Set Tdf = Dtb.TableDefs("InveFisiDet")
' Tdf.Fields("AUXFAMIL").AllowZeroLength = True
'
' Dtb.Close
' Set Dtb = Nothing
' Exit Sub
  
End Sub

Public Function EncMes(FF As Date) As Boolean       'CORREGIR
 Dim rs As ADODB.Recordset
 Dim rsql As String
    If Month(FF) < 12 Then
        rsql = "SELECT CACIERRE FROM MOVALMCAB WHERE CAFECDOC >= " & Format(Month(FF), "00") & "/01/" & Format(Year(FF), "0000") & " AND CAFECDOC < " & Format(Val(Month(FF)) + 1, "00") & "/01/" & (Format(Year(FF), "0000")) & ""
    Else
        rsql = "SELECT CACIERRE FROM MOVALMCAB WHERE CAFECDOC >= " & Format(Month(FF), "00") & "/01/" & Format(Year(FF), "0000") & " AND CAFECDOC < 01/01/" & (Format(Val(Year(FF)) + 1, "0000")) & ""
    End If
    Set rs = New ADODB.Recordset
    rs.Open rsql, cConexCom, adOpenStatic, adLockReadOnly
    If rs.EOF Then EncMes = False: rs.Close:  Exit Function
    If rs(0) = True Then EncMes = True
    rs.Close
End Function
'*************** FUNCIONES ADICIONADAS A 28 JUNIO DEL 2001
'*************** ROBERTO M.M.
Function AnioMesAnterior(ByVal arAnioMes As String) As String
Dim LMes, LAnio As String
   If Val(Mid(arAnioMes, 5, 2)) = 1 Then
      LAnio = Val(Left(arAnioMes, 4)) - 1
      LMes = 12
   Else
      LAnio = Val(Left(arAnioMes, 4))
      LMes = Val(Mid(arAnioMes, 5, 2)) - 1
   End If
   AnioMesAnterior = Format(LAnio, "0000") & Format(LMes, "00")
End Function

Function AnioMesSiguiente(ByVal arAnioMes As String) As String
Dim LMes, LAnio As String
   If Val(Mid(arAnioMes, 5, 2)) = 12 Then
      LAnio = Val(Left(arAnioMes, 4)) + 1
      LMes = 1
   Else
      LAnio = Val(Left(arAnioMes, 4))
      LMes = Val(Mid(arAnioMes, 5, 2)) + 1
   End If
   AnioMesSiguiente = Format(LAnio, "#000") & Format(LMes, "0#")
End Function
'**************************************************************
'**************************************************************
Function DevolverTCambio(ByVal arDate As Date) As Double
On Error Resume Next
          If UCase(Dir$(cRuta4)) = UCase(cNomBd4) Then
              DevolverTCambio = Val(Devolver_Dato(3, CDate(arDate), "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "ME", "TIPOMON_CODIGO"))
          ElseIf UCase(Dir$(cRuta2)) = "BDCOMUN.MDB" Then
              DevolverTCambio = Val(Devolver_Dato(1, CDate(arDate), "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "ME", "TIPOMON_CODIGO"))
          End If
End Function


Function UltimoCierre() As String
Dim rs As New ADODB.Recordset
rs.Open "Select max(CierrMes) as Tot From CierrMesValor ", cConexCom, adOpenForwardOnly, adLockReadOnly
If Not rs.EOF Then
   If Not (IsNull(rs!tot) Or rs!tot = "") Then
      UltimoCierre = rs!tot
   Else
      UltimoCierre = "" 'IIf(IsNull(rs!Min) Or rs!Min = "", "", AnioMesAnterior(Format(Year(rs!Min), "0000") & Format(Month(rs!Min), "00")))
   End If
End If
rs.Close
End Function

Private Sub ADOConectar()
Dim cRt As String
cRt = App.Path & "\BdAuxCom.Mdb"
Set cConexAux = New ADODB.Connection
cConexAux.CursorLocation = adUseClient
cConexAux.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & cRt & ";"
cConexAux.Open
End Sub

Function cNull(ByVal arNulo As Variant) As Variant
         cNull = IIf(IsNull(arNulo), "", arNulo)
End Function


Public Function ExisteIndice(BaseDeDatos As String, nombretabla As String, NombreIndice As String) As Boolean
'**********************************************
'*                                            *
'*   Verifica la existencia de un índice      *
'*   Funcion creada por Julio Calderón        *
'*                                            *
'**********************************************

'   ExisteIndice = False
    
'   'Dim dbs As DAO.Database
'   Dim Tdf As DAO.TableDef
'   Dim idxBucle As DAO.Index
'
'   Set dbs = OpenDatabase(BaseDeDatos)
'   Set Tdf = dbs(nombretabla)
'   With Tdf
'      ' Recorre la colección Indexes de la tabla.
'      For Each idxBucle In .Indexes
'         If StrConv(idxBucle.name, vbUpperCase) = StrConv(NombreIndice, vbUpperCase) Then
'           ExisteIndice = True
'         End If
'      Next idxBucle
'   End With
'   dbs.Close
End Function

Public Sub CreaIndice(BaseDeDatos As String, _
                      nombretabla As String, _
                      NombreIndice As String, _
                      ClavePrimaria As Boolean, _
                      CampoIndice1 As String, _
                      Optional CampoIndice2 As String, _
                      Optional CampoIndice3 As String, _
                      Optional CampoIndice4 As String)
'************************************************
'*                                              *
'*   Crea un índice compuesto de hasta 4 campos *
'*   Funcion creada por Julio Calderón          *
'*                                              *
'************************************************
   
'   Dim dbs As DAO.Database
'   Dim Tdf As DAO.TableDef
'   Dim NuevoIndice As DAO.Index
'   Dim idxNombre As DAO.Index
'   Dim idxBucle As DAO.Index
'
'   Set dbs = OpenDatabase(BaseDeDatos)
'   Set Tdf = dbs(nombretabla)
'   With Tdf
'      ' Primero crea objeto Index, crea y agrega los
'      ' objetos Field al objeto Index y después agrega
'      ' el objeto Index a la colección Indexes de TableDef.
'      Set NuevoIndice = .CreateIndex(NombreIndice)
'      With NuevoIndice
'         .Fields.Append .CreateField(CampoIndice1)
'         If CampoIndice2 <> "" Then
'           .Fields.Append .CreateField(CampoIndice2)
'         End If
'         If CampoIndice3 <> "" Then
'           .Fields.Append .CreateField(CampoIndice3)
'         End If
'         If CampoIndice4 <> "" Then
'           .Fields.Append .CreateField(CampoIndice4)
'         End If
'      End With
'
'      NuevoIndice.Primary = ClavePrimaria
'
'      .Indexes.Append NuevoIndice
'      ' Actualiza la colección para que pueda tener
'      ' acceso a los objetos Index nuevos.
'      .Indexes.Refresh
'   End With
'   dbs.Close
End Sub

Public Sub EliminaIndice(ByVal BaseDeDatos As String, ByVal nombretabla As String, ByVal NombreIndice As String)
'**********************************************
'*                                            *
'*   Elimina un índice                        *
'*   Sub creada por Julio Calderón            *
'*                                            *
'**********************************************

'   Dim dbs As DAO.Database
'   Dim Tdf As DAO.TableDef
'   Dim idxBucle As DAO.Index
'
'   Set dbs = OpenDatabase(BaseDeDatos)
'   Set Tdf = dbs(nombretabla)
'
'   With Tdf
'   'Recorre la colección Indexes de la tabla.
'     For Each idxBucle In .Indexes
'         If StrConv(idxBucle.name, vbUpperCase) = StrConv(NombreIndice, vbUpperCase) Then
'          .Indexes.Delete (NombreIndice)
'          Exit For
'         End If
'     Next idxBucle
'   End With
'   dbs.Close
End Sub


Function UltimoCierreFech(ByVal xFecha As Date) As Date
Dim Rs2 As New ADODB.Recordset
Dim rsql As String
Dim CIERRE, Anio, mes As String
Dim lFecha As Date
CIERRE = UltimoCierre
If CIERRE = "" Then
   UltimoCierreFech = CDate(Format(xFecha, "dd/MM/yyyy"))
   Exit Function
Else
   Anio = Left(CIERRE, 4)
   mes = Mid(CIERRE, 5, 2)
   lFecha = CDate("01/" & mes & "/" & Anio)
   If xFecha <= Fin_MES(lFecha) Then
       UltimoCierreFech = CDate(Format(Fin_MES(lFecha) + 1, "dd/MM/yyyy"))
    Else
       UltimoCierreFech = CDate(Format(xFecha, "dd/MM/yyyy"))
    End If
End If
       
End Function

Sub Enteros_Positivos(k As Integer, t As TextBox)
    If k = 8 Then Exit Sub
    If k < 48 Or k > 57 Then
        k = 0
    End If
End Sub

Sub Reales_Positivos(k As Integer, t As TextBox)
Dim t1 As String
    k = Asc(UCase(Chr(k)))
    If k = 8 Then Exit Sub
    If k <> 45 And k <> 44 And k <> 32 And k <> 69 And k <> 43 Then
        t1 = Left(t, t.SelStart)
        t1 = t1 & Chr(k) & Right(t, Len(t) - Len(t1))
        If IsNumeric(t1) Then Exit Sub
    End If
    k = 0
    
End Sub

Public Function DateSQL2000(ByVal Fecha As Variant) As Variant
    If Len(Trim(Fecha)) > 0 And IsDate(Fecha) Then
      ' DateSQL2000 = "'" & Month(Fecha) & "/" & Day(Fecha) & "/" & Year(Fecha) & "'"
        DateSQL2000 = "'" & Fecha & "'"
    Else
       'DateSQL2000 = "'" & Month(0) & "/" & Day(0) & "/" & Year(0) & "'"
       'DateSQL2000 = "'" & "00/00/0000" & "'"
       DateSQL2000 = "Null"
    End If
End Function

Function nNull(ByVal arNulo As Variant) As Variant
         nNull = IIf(IsNull(arNulo), 0, arNulo)
End Function

'*****************************************************************
'EL ULTIMO INGRESO CUYO PRECIO SEA MAYOR QUE CERO  RMM 09/06/2001
'*****************************************************************
Function UltimoPrecio(ByVal arCodigo As String, ByVal moneda As String) As Double
 Dim rs As New ADODB.Recordset
 Dim rsql As String
 rsql = "SELECT DEprecio,CACODMON,CATIPCAM FROM MovAlmDet AS A INNER JOIN MovAlmCab AS B ON (B.CAALMA = A.DEALMA) AND " & _
         "(B.CATD = A.DETD) AND (B.CANUMDOC = A.DENUMDOC) WHERE CAALMA = '" & VGAlma & "'  And CASITGUI<>'A' and " & _
         "catipmov='I'  and decodigo='" & arCodigo & "' and  a.deprecio<>0 AND  cafecdoc= ( SELECT max(cafecdoc) FROM MovAlmDet AS A " & _
         "INNER JOIN MovAlmCab AS B ON (B.CAALMA = A.DEALMA) AND (B.CATD = A.DETD) AND (B.CANUMDOC = " & _
         "A.DENUMDOC) WHERE CAALMA = '" & VGAlma & "'  And CASITGUI<>'A'  and a.deprecio<>0 and catipmov='I' and decodigo='" & arCodigo & "') "
 rs.Open rsql, cConexCom, adOpenStatic
 UltimoPrecio = 0
 If Not rs.EOF Then
    If rs(1) = "ME" And moneda = "MN" And nNull(rs(2)) <> 0 Then UltimoPrecio = Format(rs(0) * rs(2), "###,##0.00")
    If rs(1) = "ME" And moneda = "ME" And nNull(rs(2)) <> 0 Then UltimoPrecio = Format(rs(0), "###,##0.00")
    If rs(1) = "MN" And moneda = "ME" And nNull(rs(2)) <> 0 Then UltimoPrecio = Format(rs(0) / rs(2), "###,##0.00")
    If rs(1) = "MN" And moneda = "MN" And nNull(rs(2)) <> 0 Then UltimoPrecio = Format(rs(0), "###,##0.00")
 Else
     UltimoPrecio = 0
 End If
 
 rs.Close
End Function



Function Fin_MES(ByVal Afech As Date) As Date
Dim mes, Anio, lastday As String
mes = Format(Month(Afech), "0#")
Anio = CStr(Year(Afech))
lastday = "31/" & mes & "/" + Anio

If IsDate(lastday) Then
   Fin_MES = Format(lastday, "dd/mm/yyyy")
   Exit Function
Else
   lastday = "30/" & mes & "/" + Anio
   If IsDate(lastday) Then
      Fin_MES = Format(lastday, "dd/mm/yyyy")
      Exit Function
   Else
       lastday = "29/" & mes & "/" + Anio
       If IsDate(lastday) Then
          Fin_MES = Format(lastday, "dd/mm/yyyy")
          Exit Function
       Else
           lastday = "28/" & mes & "/" + Anio
           If IsDate(lastday) Then
              Fin_MES = Format(lastday, "dd/mm/yyyy")
              Exit Function
           Else
              MsgBox "Existe errores en la function Fin_mes...!"
           End If
       End If
   End If
End If
End Function

Function Ini_MES(ByVal Afech As Date) As Date
Dim mes, Anio As String
mes = Format(Month(Afech), "0#")
Anio = CStr(Year(Afech))
Ini_MES = CDate("01/" & mes & "/" + Anio)
End Function

Sub Tabula(ByVal key As Long)
    If key = 13 Then SendKeys "{tab}"
End Sub


Function FechMask(ByVal arFech As Variant) As Variant
If IsNull(arFech) Then
   FechMask = "__/__/____"
   Exit Function
End If
If Year(arFech) < 1901 Or Not IsDate(arFech) Then
   FechMask = "__/__/____"
Else
   FechMask = arFech
End If

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
Public Sub ModiFieldDef(ByVal sDataBase As String, ByVal sTable As String, ByVal sField As String, _
 Optional ByVal sType As String, _
 Optional ByVal Decimales As Variant, _
 Optional ByVal AllowZeroLen As Variant, _
 Optional ByVal FRequired As Variant, _
 Optional ByVal DefVal As Variant)
 
End Sub
Public Sub DEMORA(TIEMPO As Double)
'Fernando: 06/08/2001:
Dim HORA As Double
    Screen.MousePointer = 11
    HORA = Time()
    Do While Format(TimeSerial(0, 0, TIEMPO) + HORA, "HH:MM:SS") <> Format(Time(), "HH:MM:SS")
       ' DoEvents
    Loop
    Screen.MousePointer = 1
End Sub

Public Function ConectarAux() As ADODB.Connection
'Fernando: 06/08/2001:
Dim cRt As String
    cRt = App.Path & "\BdAuxCom.Mdb"
    Set ConectarAux = New ADODB.Connection
    ConectarAux.CursorLocation = adUseClient
    ConectarAux.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & cRt & ";"
    ConectarAux.Open
End Function

