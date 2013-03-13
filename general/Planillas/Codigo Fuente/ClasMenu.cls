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
Public MDIMenu As MDIForm
Public Sub CrearTablasSeguridad()
    Dim SQL As String
    'Crear Tabla Usuarios
    If Not ExisteElem(0, Conexion, UCase(TablaUsu)) Then
        SQL = "" & _
        " Create Table " & UCase(TablaUsu) & " ( " & _
        " USU_CODIGO VarChar(5)  , EMP_CODIGO VarChar(15) , USU_PASSWORD VarChar(8), " & _
        " USU_NIVEL VarChar(1), USU_NOMBRE VarChar(32), " & _
        " CONSTRAINT CLAVE " & _
        " PRIMARY KEY (USU_CODIGO, EMP_CODIGO)) "
        Conexion.Execute SQL
    End If
    'Crea Tabla Detalle de Menu Usuarios
    If Not ExisteElem(0, Conexion, UCase(TabaMenuDet)) Then
        SQL = "" & _
        " Create Table " & UCase(TabaMenuDet) & " ( " & _
        " USU_CODIGO VarChar(5)  , EMP_CODIGO VarChar(15) , MEN_CODIGO VarChar(10), " & _
        " MEN_HAB bit, " & _
        " CONSTRAINT CLAVE " & _
        " PRIMARY KEY (USU_CODIGO, EMP_CODIGO,MEN_CODIGO)) "
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
    If Not ExisteElem(0, Conexion, UCase(TablaMenu)) Then
        SQL = " Create Table " & UCase(TablaMenu) & " (MEN_CODIGO VarChar(10),MEN_DESCRI VarChar(64), " & _
        " MEN_CODEDIT VarChar(48),  " & _
        " CONSTRAINT CLAVE PRIMARY KEY (MEN_CODIGO))"
        Conexion.Execute SQL
      Else:
        Conexion.Execute "Delete From " & UCase(TablaMenu)
    End If
    'Redimensionando los Niveles
    ReDim ContNiv(MaxNivel + 1) 'MaxNivel : Funcion que determina el maximo nivel en el menu
    '*Armando El Menu
    RSAUX.Open UCase(TablaMenu), Conexion, adOpenDynamic, adLockOptimistic
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
                RSAUX!Men_CODIGO = Trim(CodigoMenu)
                RSAUX!Men_Descri = DevCad("&", Contr.Caption)
                RSAUX!MEN_CODEDIT = Trim(Contr.Name)
                RSAUX.Update
            End If
        End If
    Next
End Sub
Private Function MaxNivel() As Integer
    Dim CONT As Integer, Aux As Integer
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
Public Function ExisteElem(Tip As Integer, Cn As ADODB.Connection, TABLA As String, _
        Optional CAMPO As String) As Boolean
'Funcion que devuelve un valor verdadero si es que encuentra el elemento
'Creado por Fernando Cossio
    Dim SQL As String
    Dim RSAUX As New ADODB.Recordset
   '*------------------------------*
   '0 Si Existe la tabla
   '1 Si Existe el Campo
   ExisteElem = False
   TABLA = UCase(TABLA): CAMPO = UCase(CAMPO)
On Error GoTo ErrExiste
   SQL = ""
    Select Case Tip
        Case 0:
            SQL = "Select Top 1 * From " & TABLA
        Case 1:
            SQL = "Select Top 1 " & CAMPO & " From " & TABLA
    End Select
    RSAUX.Open SQL, Cn
    ExisteElem = True
    Exit Function
ErrExiste:
    ExisteElem = False
End Function

Public Function CODIFICA(CADENA As String, VALOR As Integer) As String
    Dim ciclo As Integer, posic As Integer, ult_sal As Integer
    Dim carac As String, cadena_cod As String, CAD As String
    posic = 0: ult_sal = 0
    carac = "": cadena_cod = "": CAD = ""
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
     CAD = CAD + carac
    Next ciclo
    CODIFICA = CAD
End Function
'función que desencripta una cadena
Public Function DECODIFICA(CADENA As String, VALOR As Integer) As String
    Dim ciclo As Integer, posic As Integer, val_n As Integer, val_an As Integer
    Dim carac As String, CAD As String
    CADENA = Trim(CADENA)
    CAD = ""
    val_n = 0: val_an = 0
    For ciclo = 1 To Len(CADENA)
     carac = Mid(CADENA, ciclo, 1)
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
     CAD = CAD + Chr(val_n)
    Next ciclo
    DECODIFICA = UCase(CAD)
End Function
Public Function NumChar(CADENA As String, Caracter As String) As Integer
'Funcion Creada por Fernando Cossio
    Dim I As Integer, CONT As Integer
    CONT = 0
    For I = 1 To Len(CADENA)
        If UCase(Caracter) = Mid(CADENA, I, 1) Then
            CONT = CONT + 1
        End If
    Next
    NumChar = CONT
End Function
Public Function DevCad(CaraterNo As String, CADENA As String)
    'Devuelve toda la cadena sin el caracter espcificado en el parametro
    'Funcion creada por fernando cossio
    Dim I As Integer, CONT As String
    ACUM = ""
    For I = 1 To Len(CADENA)
        If UCase(CaraterNo) <> Mid(CADENA, I, 1) Then
            ACUM = ACUM + Mid(CADENA, I, 1)
        End If
    Next
    DevCad = ACUM
End Function
Public Sub HabilitarMenuNom(USUARIO As String, EMPRESA As String)
'Procedimiento creado por Fernando Cossio
    Dim RsMenUsu As New ADODB.Recordset
    Dim SQL As String
    TabaMenuDet = "USUARIODET"
    TablaMenu = "MENU"
    SQL = " SELECT MEN_USU_COBR.USU_CODIGO, MEN_USU_COBR.EMP_CODIGO, " & _
          " MENU_COBR.MEN_CODEDIT, MEN_USU_COBR.MEN_HAB " & _
          " FROM " & UCase(TabaMenuDet) & " As MEN_USU_COBR, " & UCase(TablaMenu) & " AS MENU_COBR " & _
          " WHERE (((MEN_USU_COBR.MEN_CODIGO)=[MENU_COBR].[MEN_CODIGO])) AND  " & _
          " MEN_USU_COBR.USU_CODIGO='" & USUARIO & "' And MEN_USU_COBR.EMP_CODIGO='" & _
          EMPRESA & "'"
    RsMenUsu.Open SQL, DBSTARPLAN, adOpenKeyset, adLockReadOnly
    Dim N As String
    Do While Not RsMenUsu.EOF
        For Each Contr In MDIPrincipal.Controls
            If TypeOf Contr Is Menu Then
                N = Contr.Name
                If UCase(N) = UCase(RsMenUsu!MEN_CODEDIT) Then
                     Contr.Enabled = IIf(RsMenUsu!MEN_HAB, True, False)
                     Exit For
                End If
            End If
        Next
        RsMenUsu.MoveNext
    Loop
End Sub
Public Sub MenuTrue()
    Exit Sub
    'Procedimiento creado por Fernando Cossio
    For Each Contr In MDIMenu.Controls
        If TypeOf Contr Is Menu Then
            Contr.Enabled = True
        End If
    Next
End Sub



