Attribute VB_Name = "ACTUALIZACION"
Sub ACTUALIZACION2001()
  


  If Not ExisteElem(0, cConexCom, "CIERRMESVALOR") Then
        sql = " Create Table CIERRMESVALOR (CIERRMES Text(6),CIERRFECH DATETIME, CIERROPER TEXT(15) , " & _
        " CONSTRAINT Clave PRIMARY KEY (CIERRMES))"
        cConexCom.Execute sql
  End If


End Sub
