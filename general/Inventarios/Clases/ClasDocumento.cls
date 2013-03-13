VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClasDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************Clase Documento ROBERTO MAZA MILLA
Public Sub AnuloCabezera(ByVal AlmaActu As String, ByVal Tdoc As String, ByVal NumDoc As String, ByVal cn As ADODB.Connection)
    csql = "UPDATE movalmcab SET CASITGUI='A' where  CAALMA ='" & AlmaActu & "' AND CATD = '" & Tdoc & "' AND CANUMDOC ='" & NumDoc & "' AND CACIERRE=FALSE AND CASITGUI='V'"
    VGCNx.Execute csql
End Sub

'***************Clase Documento ROBERTO MAZA MILLA
Public Sub EliminoCabezera(ByVal AlmaActu As String, ByVal Tdoc As String, ByVal NumDoc As String, ByVal cn As ADODB.Connection)
    csql = "Delete from movalmcab where  CAALMA ='" & AlmaActu & "' AND CATD = '" & Tdoc & "' AND CANUMDOC ='" & NumDoc & "' AND CACIERRE=FALSE AND (CASITGUI='V' or CASITGUI='A')"
    VGCNx.Execute csql
End Sub

'***************Clase Documento ROBERTO MAZA MILLA
Public Sub EliminoDetalle(ByVal AlmaActu As String, ByVal Tdoc As String, ByVal NumDoc As String, ByVal cn As ADODB.Connection)
    csql = "Delete from movalmdet where  DEALMA ='" & AlmaActu & "' AND DETD = '" & Tdoc & "' AND DENUMDOC ='" & NumDoc & "'"
    VGCNx.Execute csql
End Sub

Public Function EsImportacion(ByVal AlmaActu As String, ByVal Tdoc As String, ByVal NumDoc As String, ByVal cn As ADODB.Connection) As String 'INGRESO POR IMPORTACION
    Dim rs As New ADODB.Recordset
    csql = "Select CANROIMP from MOVALMCAB where  CAALMA ='" & AlmaActu & "' AND CATD = '" & Tdoc & "' AND CANUMDOC ='" & NumDoc & "' AND CACIERRE=0 "
    rs.Open csql, VGCNx, adOpenStatic, adLockReadOnly
    If Not rs.EOF Then
       If cNull(rs!CANROIMP) <> "" Then
          EsImportacion = rs!CANROIMP
       Else
          EsImportacion = ""
       End If
    End If
    rs.Close
End Function

Public Function TienePedido(ByVal AlmaActu As String, ByVal Tdoc As String, ByVal NumDoc As String, ByVal cn As ADODB.Connection) As String 'INGRESO POR IMPORTACION
    Dim rs As New ADODB.Recordset
    csql = "Select CANROPED from MOVALMCAB where  CAALMA ='" & AlmaActu & "' AND CATD = '" & Tdoc & "' AND CANUMDOC ='" & NumDoc & "' AND CACIERRE=0 "
    rs.Open csql, VGCNx, adOpenStatic, adLockReadOnly
    If Not rs.EOF Then
       If cNull(rs!CANROped) <> "" Then
          TienePedido = rs!CANROped
       Else
          TienePedido = ""
       End If
    End If
    rs.Close
End Function


Public Function NroGuiaproveedor(ByVal asAlma As String, ByVal asNumdoc As String, ByVal cn As ADODB.Connection)
    Dim rs As New ADODB.Recordset
    csql = "Select CARFNDOC from MOVALMCAB where  CAALMA ='" & asAlma & "' AND CATD = 'NI' AND CANUMDOC ='" & asNumdoc & "' AND CACIERRE=0 AND CARFTDOC='GC'"
    rs.Open csql, VGCNx, adOpenStatic, adLockReadOnly
    If Not rs.EOF Then
       If cNull(rs!CARFNDOC) <> "" Then
          NroGuiaproveedor = Trim(rs!CARFNDOC)
       Else
          NroGuiaproveedor = ""
       End If
    End If
    rs.Close
End Function

Public Sub RestauraDocImport(ByVal NumIMPORT As String, ByVal arCod As String, ByVal CANTID As Double, ByVal arItem As String, ByVal cn As ADODB.Connection)
    csql = "UPDATE IMPORD SET NCANTENTR=NCANTENTR-" & CANTID & ",NCANSALDO=NCANSALDO+" & CANTID & "  where  CNUMERO='" & NumIMPORT & "' AND CCODARTIC='" & arCod & "' AND CITEM='" & Format(Val(arItem), "000") & "'"
    VGCNx.Execute csql
End Sub

Public Sub DefineEstado(ByVal NumIMPORT As String, ByVal cn As ADODB.Connection)
    Dim cEstado As String
    cEstado = "03"
    csql = "UPDATE IMPORC SET CSITUACION ='" & cEstado & "' where  CNUMERO='" & NumIMPORT & "'"
    VGCNx.Execute csql
End Sub

Public Sub CorrigueNumeracion(ByVal AlmaActu As String, ByVal cn As ADODB.Connection)
Dim rs As New ADODB.Recordset
rs.Open "Select max(canumdoc) AS NUM from MOVALMCAB Where CATD='NI' AND CAALMA='" & AlmaActu & "'", VGCNx, adOpenStatic, adLockReadOnly
If Not rs.EOF Then
   VGCNx.Execute "Update TabAlm set TANUMENT= '" & IIf(IsNull(rs!num), 0, rs!num) & "' where TAALMA='" & AlmaActu & "' "
End If
rs.Close

Set rs = New ADODB.Recordset
rs.Open "Select max(canumdoc) AS NUM from MOVALMCAB Where CATD='NS' AND CAALMA='" & AlmaActu & "'", VGCNx, adOpenStatic, adLockReadOnly
If Not rs.EOF Then
   VGCNx.Execute "Update TabAlm set TANUMSAL= '" & IIf(IsNull(rs!num), 0, rs!num) & "' where TAALMA='" & AlmaActu & "' "
End If
rs.Close

End Sub

Public Sub VerificaTodoDocu()
'Dim RS As New ADODB.Recordset
'Dim Rs2 As New ADODB.Recordset
'RS.Open "Select CAALMA+CATD+CANUMDOC AS NUM from MOVALMCAB ", VGcnx, adOpenStatic, adLockReadOnly
'While Not RS.EOF
'   Set Rs2 = New ADODB.Recordset
'   Rs2.Open "SELECT DISTICNCT DEALMA+DETD+DENUMDOC FROM MOVALMDET WHERE DEALMA+DETD+DENUMDOC='" & RS!num & "'"
'   RS.MoveNext
'Wend
'RS.Close
End Sub

Public Sub RestauraDespachoFactura(ByVal NumFactu As String, ByVal arCod As String, ByVal CANTID As Double, ByVal cn As ADODB.Connection)
    csql = "UPDATE FACDET SET dfcantadep=dfcantadep+" & CANTID & "  where  DFTD+DFNUMSER+DFNUMDOC='" & NumFactu & "' AND DFCODIGO='" & arCod & "'"
    VGCNx.Execute csql
End Sub

Public Function EsDespachoFactura(ByVal AlmaActu As String, ByVal Tdoc As String, ByVal NumDoc As String, ByVal cn As ADODB.Connection) As String
    Dim rs As New ADODB.Recordset
    csql = "Select CARFTDOC,CARFNDOC from MOVALMCAB where  CAALMA ='" & AlmaActu & "' AND CATD = '" & Tdoc & "' AND CANUMDOC ='" & NumDoc & "' AND CACIERRE=0 AND (CARFTDOC='FT' OR CARFTDOC='BV' )"
    rs.Open csql, VGCNx, adOpenStatic, adLockReadOnly
    If Not rs.EOF Then
       If cNull(rs!CARFNDOC) <> "" Then
          EsDespachoFactura = Trim(rs!CARFTDOC) + Trim(rs!CARFNDOC)
       Else
          EsDespachoFactura = ""
       End If
    End If
    rs.Close
End Function

Public Sub RestaurarPedido(ByVal Numped As String, ByVal arCod As String, ByVal CANTID As Double, ByVal cn As ADODB.Connection)
Dim X As Integer
    csql = "UPDATE PEDDET SET DFSALDO=DFSALDO+" & CANTID & "  where  DFNUMPED='" & Numped & "' AND DFCODIGO='" & arCod & "'"
    VGCNx.Execute csql
    'Julio: 27/10/2001:
    csql = "UPDATE PEDCAB SET CFESTADO='V' WHERE  CFNUMPED='" & Numped & "'"
    VGCNx.Execute csql, X
    'Fin Julio: 27/10/2001
End Sub

Public Function ExisteGuiaProvee(ByVal asGuia As String, ByVal asProvee As String, ByVal cn As ADODB.Connection) As Boolean
    Dim rs As New ADODB.Recordset
    csql = "Select CCNUMSER+CCNUMDOC AS NUMDOC From COMGUICAB where  CCCODPRO ='" & asProvee & "' AND CCNUMSER+CCNUMDOC='" & asGuia & "'"
    
    rs.Open csql, VGCNx, adOpenStatic, adLockReadOnly
    If Not rs.EOF Then
       ExisteGuiaProvee = True
    Else
       ExisteGuiaProvee = False
    End If
    rs.Close
End Function

Public Sub EliminoGuiaProveeCompra(ByVal asGuia As String, ByVal asProvee As String, ByVal cn As ADODB.Connection)
    Dim X As Long
    X = 0
    csql = "Delete from COMGUICAB Where CCNUMSER+CCNUMDOC='" & asGuia & "' AND CCCODPRO = '" & asProvee & "' AND  CCESTADO<>'F'"
    VGCNx.Execute csql, X
    
    If X > 0 Then
       csql = "Delete from COMGUIDET Where DCNUMSER+DCNUMDOC='" & asGuia & "' AND DCCODPRO = '" & asProvee & "'"
       VGCNx.Execute csql
    End If
End Sub



Public Function TipoTransaccion(ByVal cAlma As String, ByVal Tdoc As String, ByVal NumDoc As String, ByVal cn As ADODB.Connection) As String
Dim rs As New ADODB.Recordset
rs.Open "SELECT CACODMOV FROM MOVALMCAB WHERE CATD='" & Tdoc & "' AND CANUMDOC='" & NumDoc & "'", VGCNx, adOpenStatic, adLockReadOnly
If Not rs.EOF Then
   TipoTransaccion = rs!cacodmov
Else
   TipoTransaccion = ""
End If
rs.Close
End Function
