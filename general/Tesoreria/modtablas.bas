Attribute VB_Name = "Modtablas"
Public Sub adicionacampos()
Dim SQL As String
Dim rsql As New Recordset
If Not ExisteElem(1, VGcnx, "te_cabecerarecibos", "empresacodigo") Then
        VGcnx.Execute "ALTER TABLE te_cabecerarecibos ADD empresacodigo VARCHAR(2) NULL"
End If
If Not ExisteElem(1, VGcnx, "te_cabecerarecibos", "cabcomprobnumero") Then
        VGcnx.Execute "ALTER TABLE te_cabecerarecibos ADD cabcomprobnumero INT NULL"
End If
If Not ExisteElem(1, VGcnx, "te_detallerecibos", "entidadcodigo") Then
        VGcnx.Execute "ALTER TABLE te_detallerecibos ADD entidadcodigo VARCHAR(11) NULL"
End If
If Not ExisteElem(1, VGcnx, "te_detallerecibos", "centrocostocodigo") Then
        VGcnx.Execute "ALTER TABLE te_detallerecibos ADD centrocostocodigo VARCHAR(10) NULL"
End If
If Not ExisteElem(1, VGcnx, "te_conceptocaja", "conceptosiccosto") Then
        VGcnx.Execute "ALTER TABLE te_conceptocaja ADD conceptosiccosto VARCHAR(1) NULL"
End If
If Not ExisteElem(1, VGcnx, "te_parametroempresa", "empresalistaestadoclientes") Then
        VGcnx.Execute "ALTER TABLE te_parametroempresa ADD empresalistaestadoclientes VARCHAR(1) NULL"
End If
If Not ExisteElem(1, VGcnx, "te_parametroempresa", "empresalistaestadoproveedor") Then
        VGcnx.Execute "ALTER TABLE te_parametroempresa ADD empresalistaestadoproveedor VARCHAR(1) NULL"
End If
If ExisteElem(1, VGcnx, "cp_tipodocumento", "tdocumentonumerador") Then
        VGcnx.Execute "ALTER TABLE cp_tipodocumento ALTER COLUMN tdocumentonumerador VARCHAR(11) NULL"
End If

If Not ExisteElem(0, VGcnx, "te_rendiciones") Then
   SQL = " Create Table te_rendiciones (oficinacodigo VarChar(3),"
   SQL = SQL & "codigocaja varchar(2),monedacodigo VarChar(2),"
   SQL = SQL & "rendicionnumero VarChar(6),"
   SQL = SQL & "rendicionsaldoinicial float,"
   SQL = SQL & "rendicioningresos float,"
   SQL = SQL & "rendicionegresos float,"
   SQL = SQL & "rendicionsaldofinal float,"
   SQL = SQL & "rendicionfecha datetime,"
   SQL = SQL & "usuariocodigo varchar(8),fechaact datetime "
   SQL = SQL & " CONSTRAINT PK_te_rendiciones "
   SQL = SQL & " PRIMARY KEY (oficinacodigo,monedacodigo,rendicionnumero))  "
   VGcnx.Execute SQL
End If
If Not ExisteElem(1, VGcnx, "te_detallerecibos", "rendicionnumero") Then
        VGcnx.Execute "ALTER TABLE te_detallerecibos ADD rendicionnumero VARCHAR(6) NULL"
End If
If Not ExisteElem(1, VGcnx, "te_codigocaja", "cajarendiciones") Then
        VGcnx.Execute "ALTER TABLE te_codigocaja ADD cajarendiciones bit NULL"
End If
If Not ExisteElem(1, VGcnx, "te_codigocaja", "rendicionnumero01") Then
        VGcnx.Execute "ALTER TABLE te_codigocaja ADD rendicionnumero01 VARCHAR(6) NULL"
End If
If Not ExisteElem(1, VGcnx, "te_codigocaja", "rendicionnumero02") Then
        VGcnx.Execute "ALTER TABLE te_codigocaja ADD rendicionnumero02 VARCHAR(6) NULL"
End If
If ExisteElem(1, VGcnx, "te_cuentabancos", "cbanco_numero") Then
        VGcnx.Execute "ALTER TABLE te_cuentabancos ALTER COLUMN cbanco_numero VARCHAR(20) NULL"
End If
If Not ExisteElem(1, VGcnx, "cp_cargo", "cargoemiteretencion") Then
        VGcnx.Execute "ALTER TABLE cp_cargo ADD cargoemiteretencion bit NULL"
End If
If Not ExisteElem(1, VGcnx, "cp_tipodocumento", "documentoretencion") Then
        VGcnx.Execute "ALTER TABLE cp_tipodocumento ADD documentoretencion varchar(1) NULL"
End If
If Not ExisteElem(1, VGcnx, "te_detallerecibos", "detalle_no_saldos") Then
        VGcnx.Execute "ALTER TABLE te_detallerecibos ADD detalle_no_saldos varchar(1) NULL"
End If
If Not ExisteElem(1, VGcnx, "te_parametroempresa", "empresaretencion") Then
        VGcnx.Execute "ALTER TABLE te_parametroempresa ADD empresaretencion VARCHAR(1) NULL"
End If
If Not ExisteElem(1, VGcnx, "te_parametroempresa", "empresacodigoretencion") Then
        VGcnx.Execute "ALTER TABLE te_parametroempresa ADD empresacodigoretencion VARCHAR(2) NULL"
End If
If Not ExisteElem(1, VGcnx, "te_parametroempresa", "porcentajeretencion") Then
        VGcnx.Execute "ALTER TABLE te_parametroempresa ADD porcentajeretencion  float  NULL"
End If
If Not ExisteElem(1, VGcnx, "co_sistema", "sistemaminimoretencion") Then
        VGcnx.Execute "ALTER TABLE co_sistema ADD sistemaminimoretencion INTEGER NULL"
End If
If Not ExisteElem(1, VGcnx, "te_parametroempresa", "empresacodigodetraccion") Then
        VGcnx.Execute "ALTER TABLE te_parametroempresa ADD empresacodigodetraccion VARCHAR(2) NULL"
End If
If Not ExisteElem(1, VGcnx, "cp_cargo", "empresacodigo") Then
        VGcnx.Execute "ALTER TABLE cp_cargo ADD empresacodigo varchar(2) NULL"
End If
If Not ExisteElem(1, VGcnx, "co_sistema", "Bancarizacion") Then
        VGcnx.Execute "ALTER TABLE co_sistema ADD Bancarizacion bit NULL"
End If
If Not ExisteElem(1, VGcnx, "co_sistema", "MinimoBancarizacion01") Then
        VGcnx.Execute "ALTER TABLE co_sistema ADD MinimoBancarizacion01 float NULL"
End If
If Not ExisteElem(1, VGcnx, "co_sistema", "MinimoBancarizacion02") Then
        VGcnx.Execute "ALTER TABLE co_sistema ADD MinimoBancarizacion02 float NULL"
End If
End Sub


