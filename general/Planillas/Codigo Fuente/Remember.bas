Attribute VB_Name = "Remember"

'-------------------
'IN '" & RegSistema.PathEmpresa & "\Planilla.mdb'"
'IN '" & App.Path & "\BDAuxCom.mdb'
'If GetSetting(App.CompanyName, "Planillas", "Nando", "No") <> "Hola" Then

Sub ESCRIBIR_WENPLAEXP()
    
    Dim CAD As String
    CAD = App.PATH & "\WENPLAEXP.INI"

    VGL_DATE = "DMY"
    Call PROCSIS.PrWriteIni(CAD, "BOOT", "SERVIDOR", VGL_SERVER)
    Call PROCSIS.PrWriteIni(CAD, "BOOT", "BASEACTUAL", REGSISTEMA.BASESQL)
    Call PROCSIS.PrWriteIni(CAD, "BOOT", "DATE", VGL_DATE)
    Call PROCSIS.PrWriteIni(CAD, "BOOT", "BASESTARPLAN", VGL_BASE)
    VGL_USUARIO = "SOPORTE"
    VGL_LOGON = "SOPORTE"
    Call PROCSIS.PrWriteIni(CAD, "BOOT", "USUARIO", VGL_USUARIO)
    Call PROCSIS.PrWriteIni(CAD, "BOOT", "LOGON", VGL_LOGON)
    Call PROCSIS.PrWriteIni(CAD, "BOOT", "NOMEMP", REGSISTEMA.EMPRESA)
    
End Sub
Public Sub DisplayarError(oError As ErrObject, Optional bNoMostrarMensaje As Boolean)
'*******************************Objetivo:Guardar el Error en un *.txt y displayar el error**********
'***********Creado            :por Fernando Cossio Peralta                           **********
'***********Fecha de Creacion :25/09/2001                                                 **********
Dim nFile As Double, sFechahor As String * 20, sProyecto As String * 20, sFormActi As String * 20
Dim CodError As String * 40, sNomError As String * 300, sFileHelp As String * 40
    If bNoMostrarMensaje = False Then MsgBox oError.Description, vbExclamation, App.TITLE
    If Dir$(App.PATH & "\Errores.err") = "" Then
        Open App.PATH & "\Errores.err" For Append As #1
         sFechahor = Left("Fecha_Hora" & String(20, " "), 20)
         sProyecto = Left("Titulo_Aplicacion" & String(20, " "), 20)
         sFormActi = Left("Nombre_formulario" & String(20, " "), 20)
         sCodError = Left("Codigo_Error" & String(40, " "), 40)
         sNomError = Left("Descripcion_Error" & String(300, " "), 300)
         sFileHelp = Left("Archivo_ayuda" & String(40, " "), 40)
        Print #1, sFechahor & sProyecto & sFormActi & sCodError & sNomError & sFileHelp
    Else
        Open App.PATH & "\Errores.err" For Append As #1
    End If
    sFechahor = Left(Date & ":" & Time & String(20, " "), 20)
    sProyecto = Left(App.TITLE & String(20, " "), 20)
    sFormActi = Left(Screen.ActiveForm.Name & String(20, " "), 20)
    sCodError = Left(oError.Number & String(40, " "), 40)
    sNomError = Left(oError.Description & String(300, " "), 300)
    sFileHelp = Left(oError.HelpFile & String(40, " "), 40)
    Print #1, sFechahor & sProyecto & sFormActi & sCodError & sNomError & sFileHelp
    Close #1
End Sub

