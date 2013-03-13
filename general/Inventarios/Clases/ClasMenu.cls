VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public TablaUsu As String
Public TabaMenuDet As String
Public TablaMenu As String
Public Conexion As New ADODB.Connection
Dim Contr As Control
Public MDIMenu As Form
Public Sub CrearTablasSeguridad()
    Dim SQL As String
    'Crear Tabla Usuarios
    If Not ExisteElem(0, Conexion, TablaUsu) Then
        SQL = "" & _
        " Create Table " & TablaUsu & " ( " & _
        " Usu_Codigo Text(5)  , Emp_Codigo Text(15) , Usu_PassWord Text(8), " & _
        " Usu_Nivel Text(1), usu_Nombre Text(32), " & _
        " CONSTRAINT Clave " & _
        " PRIMARY KEY (Usu_Codigo, Emp_Codigo)) "
        Conexion.Execute SQL
    End If
    'Crea Tabla Detalle de Menu Usuarios
    If Not ExisteElem(0, Conexion, TabaMenuDet) Then
        SQL = "" & _
        " Create Table " & TabaMenuDet & " ( " & _
        " Usu_Codigo Text(5)  , Emp_Codigo Text(15) , Men_Codigo Text(10), " & _
        " Men_Hab bit, " & _
        " CONSTRAINT Clave " & _
        " PRIMARY KEY (Usu_Codigo, Emp_Codigo,Men_Codigo)) "
        Conexion.Execute SQL
    End If
End Sub
Public Sub CrearTablaMenu()
    Dim SQL As String
    Dim Contr As Control
    Dim RSAUX As New ADODB.Recordset
    Dim ContNiv() As Integer
    Dim Niv As Integer, I As Integer, j As Integer
    Dim CodigoMenu As String
    '*------------------------------------------------------------------*
    If Not ExisteElem(0, Conexion, TablaMenu) Then
        SQL = " Create Table " & TablaMenu & " (Men_Codigo Text(10),Men_Descri Text(64), " & _
        " Men_CodEdit Text(48),  " & _
        " CONSTRAINT Clave PRIMARY KEY (Men_Codigo))"
        Conexion.Execute SQL
      Else:
        Conexion.Execute "Delete From " & TablaMenu
    End If
    'Redimensionando los Niveles
    ReDim ContNiv(MaxNivel + 1) 'MaxNivel : Funcion que determina el maximo nivel en el menu
    '*Armando El Menu
    RSAUX.Open TablaMenu, Conexion, adOpenDynamic, adLockOptimistic
    For Each Contr In MDIMenu.Controls
        If TypeOf Contr Is Menu Then
            If Contr.Visible And Trim(Contr.Caption) <> "-" And UCase(Trim(Contr.Caption)) <> "SALIR" Then
                Niv = NumChar(Contr.Name, "_")
                For I = Niv To Niv
                    ContNiv(I) = ContNiv(I) + 1
                    For j = Niv + 1 To UBound(ContNiv) - 1
                        ContNiv(j) = 0
                    Next
                Next
                RSAUX.AddNew
                CodigoMenu = ""
                For I = 0 To UBound(ContNiv) - 1
                    CodigoMenu = CodigoMenu & IIf(ContNiv(I) > 0, Format(ContNiv(I), "00"), "")
                Next
                RSAUX!tipodesistema = vgtipo
                RSAUX!Men_CodIGO = Trim(CodigoMenu)
                RSAUX!Men_Descri = DevCad("&", Contr.Caption)
                RSAUX!Men_CodEdit = Trim(Contr.Name)
                
                RSAUX!Men_Visible = -1
                
                RSAUX.Update
            End If
        End If
    Next
End Sub
Private Function MaxNivel() As Integer
    Dim Cont As Integer, Aux As Integer
    MaxNivel = 0
    '*---------------------------------*
    For Each Contr In MDIMenu.Controls
        If TypeOf Contr Is Menu Then
            If Contr.Visible And Trim(Contr.Caption) <> "-" And UCase(Trim(Contr.Caption)) <> "SALIR" Then
                Aux = NumChar(Contr.Name, "_")
                If MaxNivel < Aux Then MaxNivel = Aux
            End If
        End If
    Next
End Function
Public Function ExisteElem(Tip As Integer, VGCNx As ADODB.Connection, Tabla As String, _
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
    RSAUX.Open SQL, VGCNx
    ExisteElem = True
    Exit Function
ErrExiste:
    ExisteElem = False
End Function

Public Function CODIFICA(CADENA As String, valor As Integer) As String
    Dim ciclo As Integer, posic As Integer, ult_sal As Integer
    Dim carac As String, cadena_cod As String, cad As String
    posic = 0: ult_sal = 0
    carac = "": cadena_cod = "": cad = ""
    CADENA = UCase(Trim(CADENA))
    For ciclo = 1 To Len(CADENA)
     carac = Mid(CADENA, ciclo, 1)
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
            carac = Chr(Asc(carac) - valor)
     Case 2:
            carac = Chr(Asc(carac) - (ciclo * 2))
            ult_sal = Asc(carac)
     Case 3:
            If ult_sal > 10 Then ult_sal = ult_sal - (Int(ult_sal / 10) * 10)
            carac = Chr(Asc(carac) - valor + ult_sal)
     Case 4:
            carac = Chr(Asc(carac) - ciclo)
            utl_sal = Asc(carac)
     Case 5:
            If ult_sal > 10 Then ult_sal = ult_sal - (Int(ult_sal / 10) * 10)
            carac = Chr(Asc(carac) - valor + ult_sal)
     End Select
     cad = cad + carac
    Next ciclo
    CODIFICA = cad
End Function
'función que desencripta una cadena
Public Function DECODIFICA(CADENA As String, valor As Integer) As String
    Dim ciclo As Integer, posic As Integer, val_n As Integer, val_an As Integer
    Dim carac As String, cad As String
    CADENA = Trim(CADENA)
    cad = ""
    val_n = 0: val_an = 0
    For ciclo = 1 To Len(CADENA)
     carac = Mid(CADENA, ciclo, 1)
     posic = ciclo Mod 7
     Select Case posic
     Case 0:
            val_n = Asc(carac) / 2
     Case 1:
            val_n = Asc(carac) + valor
     Case 2:
            val_n = Asc(carac) + (ciclo * 2)
            val_an = Asc(carac)
     Case 3:
            If val_an > 10 Then val_an = val_an - (Int(val_an / 10) * 10)
            val_n = Asc(carac) + valor - val_an
     Case 4:
            val_n = Asc(carac) + ciclo
     Case 5:
            If val_an > 10 Then val_an = val_an - (Int(val_an / 10) * 10)
            val_n = Asc(carac) + valor - val_an
     Case 6:
           val_n = Asc(carac)
     End Select
     cad = cad + Chr(val_n)
    Next ciclo
    DECODIFICA = UCase(cad)
End Function
Public Function NumChar(CADENA As String, Caracter As String) As Integer
'Funcion Creada por Fernando Cossio
    Dim I As Integer, Cont As Integer
    Cont = 0
    For I = 1 To Len(CADENA)
        If UCase(Caracter) = Mid(CADENA, I, 1) Then
            Cont = Cont + 1
        End If
    Next
    NumChar = Cont
End Function
Public Function DevCad(CaraterNo As String, CADENA As String)
    'Devuelve toda la cadena sin el caracter espcificado en el parametro
    'Funcion creada por fernando cossio
    Dim I As Integer, Cont As String
    acum = ""
    For I = 1 To Len(CADENA)
        If UCase(CaraterNo) <> Mid(CADENA, I, 1) Then
            acum = acum + Mid(CADENA, I, 1)
        End If
    Next
    DevCad = acum
End Function
Public Sub HabilitarMenuNom(Usuario As String, Empresa As String)
'Procedimiento creado por Fernando Cossio
    Dim RsMenUsu As New ADODB.Recordset
    Dim SQL As String
    SQL = " SELECT MEN_USU_INV.USU_CODIGO, MEN_USU_INV.EMP_CODIGO, " & _
          " MENU_INV.Men_CodEdit, MEN_USU_INV.MEN_HAB " & _
          " From " & TabaMenuDet & " As MEN_USU_INV, " & TablaMenu & " as Menu_INV " & _
          " WHERE (((MEN_USU_INV.MEN_CODIGO)=[Menu_INV].[Men_Codigo])) And  " & _
          " MEN_USU_INV.USU_CODIGO='" & Usuario & "' And MEN_USU_INV.EMP_CODIGO='" & _
          Empresa & "'"
    RsMenUsu.Open SQL, Conexion, adOpenKeyset, adLockReadOnly
    Dim n As String
    Do While Not RsMenUsu.EOF
        For Each Contr In MDIMenu.Controls
            If TypeOf Contr Is Menu Then
                n = Contr.Name
                If UCase(n) = UCase(RsMenUsu!Men_CodEdit) Then
                     Contr.Enabled = IIf(RsMenUsu!MEN_HAB, True, False)
                     Exit For
                End If
            End If
        Next
        RsMenUsu.MoveNext
    Loop
End Sub
Public Sub MenuTrue()
'Procedimiento creado por Fernando Cossio
    For Each Contr In MDIMenu.Controls
        If TypeOf Contr Is Menu Then
            Contr.Enabled = True
        End If
    Next
End Sub
