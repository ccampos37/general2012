Attribute VB_Name = "Funcionesventas"

Public Function FechS(Fecha As Variant, tipo As TIPFECHA) As Variant
Dim h As Date
Dim fechaAux As Double
On Error GoTo ErrorFecha
   h = CDate(Fecha)
   Select Case tipo
      Case Sqlf: 'Para transformar al sql
         fechaAux = DateSerial(Year(Fecha), Month(Fecha), Day(Fecha)) - 2
      Case Adof: 'Para transformar al ado Y AL ACCESS
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

Public Function TraeTipoCambio(vfecha As Date, vcon As ADODB.Connection) As String
    Dim rsbuscn As New ADODB.Recordset
    
    Set rsbuscn = vcon.Execute("select * from ct_tipocambio where tipocambiofecha='" & vfecha & "'")
    If rsbuscn.RecordCount > 0 Then
        TraeTipoCambio = rsbuscn!tipocambioventa
    Else
        TraeTipoCambio = "1.00"
    End If
    Set rsbuscn = Nothing

End Function


Public Function TraePrecio(vlista As String, vArti As String, vcon As ADODB.Connection, vAlma As String) As Double
    Dim rsbuscn As New ADODB.Recordset
    
    Set rsbuscn = vcon.Execute("select * from listapre" & Trim(vlista) & " where productocodigo='" & vArti & "'")
    If rsbuscn.RecordCount > 0 Then
        TraePrecio = IIf(IsNull(rsbuscn!productoprecvta), 1, rsbuscn!productoprecvta)
    Else
        TraePrecio = 1
    End If
    Set rsbuscn = Nothing

End Function


Public Sub CargarTipoVentas(xcombo As ComboBox, xtipo)
  
  Select Case xtipo
    Case 1     '--condicion documento--
     xcombo.Clear
     xcombo.AddItem "0-Activo"
     xcombo.AddItem "1-Anulado"
     xcombo.ListIndex = 0
   Case 2   '--tipodocumento --
     xcombo.Clear
     xcombo.AddItem g_tipobol & "-Boleta"
     xcombo.AddItem g_tipofac & "-Factura"
     xcombo.AddItem g_tipoguia & "-B.O."
     xcombo.AddItem g_tipoticketB & "-Ticket Boleta"
     xcombo.AddItem g_tipoticketf & "-Ticket factura"
     xcombo.ListIndex = 0
   Case 3   '---estado
     xcombo.Clear
     xcombo.AddItem "S-SI"
     xcombo.AddItem "N-NO"
     xcombo.ListIndex = 0
   Case 4  '-- Tipo persona
     xcombo.Clear
     xcombo.AddItem "1-NATURAL"
     xcombo.AddItem "2-JURIDICA"
     xcombo.ListIndex = 0
   Case 5  '-tipo pais
     xcombo.Clear
     xcombo.AddItem "1-PERUANA"
     xcombo.AddItem "2-EXTRANJERA"
     xcombo.ListIndex = 0
   Case 6   '--todos los tipodocumentos --
     xcombo.Clear
     xcombo.AddItem g_tipobol & "-Boleta"
     xcombo.AddItem g_tipofac & "-Factura"
     xcombo.AddItem g_tipoguia & "-B.O."
     xcombo.AddItem g_tipoped & "-Pedido"
     xcombo.AddItem g_tipoticketB & "-Ticket Boleta"
     xcombo.AddItem g_tipoticketf & "-Ticket factura"
     xcombo.AddItem "07-Ticket factura"
     xcombo.ListIndex = 0
     
  End Select
End Sub


Public Sub Main()
   

   Call Configurar_Conexiones
   Call adicionarcampos
   'Call Cargar_Parametros_Funcionales
   FrmIngreso.Show
  
   
   
End Sub

Public Sub Configurar_Conexiones()
  
  On Error GoTo nerror
   ' *****Archivo de Punto de Venta*******
   Dim VGdllApi As New dll_apisgen.dll_apis
   Dim strbase As String
   Dim struser As String
   Dim strpass As String
   Dim strserver As String
   
   Dim strconecta As String
 
    VGcomputer = UCase(ComputerName)
    VGsql = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "SQL", "?")
    VGsql = IIf(VGsql = "?", 0, VGsql)
    
    VGformatofecha = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "FECHASQL", "?")
    VGformatofecha = IIf(VGformatofecha = "?", "DMY", VGformatofecha)
    
    VGParametros.empresacodigo = "01"
    
    
      VGParamSistem.BDEmpresaGEN = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "BDDATOS", "?")
    VGParamSistem.ServidorGEN = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "SERVIDOR", "?")
    VGParamSistem.UsuarioGEN = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "USUARIO", "?")
    VGParamSistem.PwdGEN = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "PASSW", "?")
        
    
    'Conexion de Compras
    VGParamSistem.BDEmpresa = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "BDDATOS", "?")
    VGParamSistem.Servidor = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "SERVIDOR", "?")
    VGParamSistem.Usuario = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "USUARIO", "?")
    VGParamSistem.PWD = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "PASSW", "?")
    vgorden = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "ORDEN", "?")
   
    'VGParamSistem.ServidorCONF = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "SERVIDOR", "?")
   
   
    VGParamSistem.BDempresaCONF = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "BDDATOSCONF", "?")
    If VGParamSistem.BDempresaCONF = "?" Then VGParamSistem.BDempresaCONF = "bdwenco"
   
' reportes
   
   VGParamSistem.Rutareport = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "REPORTES", "VENTAS", "?")
    VGParamSistem.carpetareportes = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "CARPETAREPORTES", "?")
        
    'Conexion de Contabilidad
    VGParamSistem.BDEmpresaCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "BDDATOS", "?")
    If VGParamSistem.BDEmpresaCT = "" Then
       VGParamSistem.BDEmpresaCT = VGParamSistem.BDEmpresa
       VGParamSistem.ServidorCT = VGParamSistem.Servidor
       VGParamSistem.UsuarioCT = VGParamSistem.Usuario
       VGParamSistem.PwdCT = VGParamSistem.PWD
     Else
       VGParamSistem.BDEmpresaCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "BDDATOS", "?")
       VGParamSistem.ServidorCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "SERVIDOR", "?")
       VGParamSistem.UsuarioCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "USUARIO", "?")
       VGParamSistem.PwdCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "PASSW", "?")
   End If
    If VGParamSistem.BDEmpresa = "?" Or VGParamSistem.Servidor = "?" Then
        MsgBox "No se ha Configurado bien los parametros BDDATOS y SERVIDOR en el archivo " & Chr(13) & _
               App.Path & "\MARFICE.INI"
    End If
    If VGParamSistem.Rutareport = "" Or VGParamSistem.Rutareport = "?" Then
       VGParamSistem.Rutareport = App.Path
       VGParamSistem.carpetareportes = "Reportes"
    End If
       
    'Establecer Cadena de Conexión de Reportes
    
    VGCadenaReport2 = "DSN=jckconsultores;DSQ=" & VGParamSistem.BDEmpresaGEN & ";UID=" & VGParamSistem.UsuarioGEN & ";PWD=" & VGParamSistem.PwdGEN & ""
        
    'Establecer Conexiones
    Set VGGeneral = New ADODB.Connection
    VGGeneral.CursorLocation = adUseClient
    VGGeneral.CommandTimeout = 0
    VGGeneral.ConnectionTimeout = 0
    VGGeneral.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.UsuarioGEN & ";Password=" & VGParamSistem.PwdGEN & ";Initial Catalog=" & VGParamSistem.BDEmpresaGEN & ";Data Source=" & VGParamSistem.Servidor
    VGGeneral.Open
   
   'Call AdjuntarServ(VGGeneral, VGParamSistem.Servidor)
   
   'Conexion de Compras el Principal
    Set VGCNx = New ADODB.Connection
    VGCNx.CursorLocation = adUseClient
    VGCNx.CommandTimeout = 0
    VGCNx.ConnectionTimeout = 0
    VGCNx.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.Usuario & ";Password=" & VGParamSistem.PWD & ";Initial Catalog=" & VGParamSistem.BDEmpresa & ";Data Source=" & VGParamSistem.Servidor
    VGCNx.Open
    
   
   'Conexion de Cofiguracion
    Set VGConfig = New ADODB.Connection
    VGConfig.CursorLocation = adUseClient
    VGConfig.CommandTimeout = 0
    VGConfig.ConnectionTimeout = 0
    VGConfig.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.Usuario & ";Password=" & VGParamSistem.PWD & ";Initial Catalog=" & VGParamSistem.BDempresaCONF & ";Data Source=" & VGParamSistem.Servidor
    VGConfig.Open
  

      
   Exit Sub
nerror:
   If Err Then
       MsgBox "Comunicarse con Sistemas " & Chr(13) & Chr(10) & Err.Number & "-" & Err.Description
       Err = 0
       End
   End If

End Sub

Public Function MostrarFormVentasVentas(pVentana As Form, pPos As String)
   pVentana.Icon = LoadPicture(App.Path & "\factu.ico")
   
   If pPos = "C" Then
     pVentana.Left = ((Screen.Width - pVentana.Width) / 2) - 350
     pVentana.Top = ((Screen.Height - pVentana.Height) / 2) - 350
   ElseIf pPos = "I" Then
      pVentana.Left = 300
      pVentana.Top = 300
   ElseIf pPos = "M" And pVentana.Visible = False Then
      pVentana.Caption = pVentana.Caption & "  " & VGParamSistem.BDEmpresa
      pVentana.Width = Screen.Width
   ElseIf pPos = "C1" Then
     pVentana.Left = ((Screen.Width - pVentana.Width) / 2) - 350
     pVentana.Top = ((Screen.Height - pVentana.Height) / 2) - 350
     Exit Function
   ElseIf pPos = "C2" Then
     pVentana.Left = ((Screen.Width - pVentana.Width) / 2) - 350
     pVentana.Top = ((Screen.Height - pVentana.Height) / 2) - 350
     Exit Function
   End If
   If pPos = "M" Then
      pVentana.StatusBar1.Panels(1).Text = "EMPRESA: " & VGParametros.nomempresa
      pVentana.StatusBar1.Panels(2).Text = "PTO. VENTA: " & g_ptoventa
      pVentana.StatusBar1.Panels(2).Alignment = sbrLeft
   Else
 '     pVentana.StatusBar1.Panels(1).Text = "FORMATO : " & Escadena(pVentana.Caption)
 '     pVentana.StatusBar1.Panels(2).Text = "USUARIO: " & g_usuario
 '     pVentana.StatusBar1.Panels(2).Alignment = sbrLeft
   End If
'   pVentana.StatusBar1.Panels(1).Alignment = sbrLeft
'   pVentana.StatusBar1.Panels(3).Text = "FECHA :" & Format(Date, "dd/mm/yyyy")
'   pVentana.StatusBar1.Panels(3).Alignment = sbrRight
'   pVentana.StatusBar1.Panels(4).Text = "HORA :" & Format(Time, "hh:mm:ss")
'   pVentana.StatusBar1.Panels(4).Alignment = sbrRight
 
End Function



Public Function aImpresora(wFile)
  Dim wbat, wcade As String
  Dim X As Long
  Dim rnrofile As Double
  
  On Error GoTo nerror
  rnrofile = CInt((90 * (Rnd(10) + 1)))
  wbat = "c:\printer" & CStr(Val(Right(wFile, 5))) & ".bat"
  Open wbat For Output As #rnrofile
  Print #rnrofile, "@echo off"
  Print #rnrofile, "Type " & wFile & " >" & Left(Printer.Port, Len(Printer.Port) - 1)
  Print #rnrofile, "cls"
  Print #rnrofile, "exit"
  Close #rnrofile
  wcade = "start /m " & Trim(wbat)
  X = Shell(wcade, vbHide)
  DoEvents
  
nerror:
   If Err Then
      MsgBox "Error: " & Err.Number & "-" & Err.Description, vbCritical, "AVISO"
      Err = 0
   End If

End Function

