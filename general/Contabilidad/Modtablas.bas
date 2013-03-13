Attribute VB_Name = "Modtablas"
Public Sub adicionacampos()
Dim SQL As String
Dim rsql As New Recordset
If Not ExisteElem(1, VGcnx, "ct_sistema", "sistemaconfiguracentrocostos") Then
        VGcnx.Execute "ALTER TABLE ct_sistema ADD sistemaconfiguracentrocostos VARCHAR(20) NULL"
End If
If Not ExisteElem(1, VGcnx, "ct_sistema", "sistemaultimonivelcostos") Then
        VGcnx.Execute "ALTER TABLE ct_sistema ADD sistemaultimonivelcostos INTEGER NULL"
End If
If Not ExisteElem(1, VGcnx, "ct_centrocosto", "centrocostonivel") Then
   VGcnx.Execute "Alter table ct_centrocosto ADD centrocostonivel integer NULL"
End If
End Sub


