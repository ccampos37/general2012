Attribute VB_Name = "ModTablas"

Public Function ExisteElem(Tip As Integer, VGcnx As ADODB.Connection, Tabla As String, _
        Optional Campo As String) As Boolean
'Funcion que devuelve un valor verdadero si es que encuentra el elemento
'Creado por Fernando Cossio
    Dim SQL As String
    Dim Rsaux As New ADODB.Recordset
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
    Rsaux.Open SQL, VGcnx
    ExisteElem = True
    Exit Function
ErrExiste:
    ExisteElem = False
End Function


