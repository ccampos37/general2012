VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Cmdaceptar 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim wCabe(40)
Dim X As New ADODB.Recordset


Private Sub Command2_Click()
Unload Me
End Sub

Private Sub cmdAceptar_Click()

SQL = " select * from vt_pedido where empresacodigo='" & VGParametros.empresacodigo & "' "
SQL = SQL & " and pedidofechasunat>='01/05/2009' order by pedidofechasunat "
Set X = VGCNx.Execute(SQL)
X.MoveFirst
Do While Not X.EOF
 wCabe(2) = X!pedidonumero
 wCabe(17) = X!vendedorcodigo
 wCabe(16) = X!clientecodigo
 wCabe(9) = X!pedidomoneda
 wCabe(10) = X!pedidotipcambio
 wCabe(34) = X!pedidofechasunat
 wCabe(40) = X!modovtacodigo
 Call actualizatesoreria1
 X.MoveNext
Loop
End Sub
Private Sub actualizatesoreria1()
Dim xx1 As New ADODB.Recordset
SQL = " select * from vt_modoventa where modovtacodigo='" & wCabe(40) & "'"
Set xx1 = VGCNx.Execute(SQL)
If xx1!modovtaactctacte = 0 Then
    Dim RSAUX As New ADODB.Recordset
    Dim VGCommandoSP  As New ADODB.Command
    Dim adll As dllgeneral.dll_general
    'Elimar los Detalle antes de grabar
    VGCommandoSP.ActiveConnection = VGgeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "vt_formadepago_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@empresa") = VGParametros.empresacodigo
        .Parameters("@pedido") = wCabe(2)
        Set RSAUX = .Execute
    End With
    If RSAUX.RecordCount() > 0 Then
       RSAUX.MoveFirst
       Do While Not RSAUX.EOF
             Call grabatesoria1(RSAUX)
             RSAUX.MoveNext
       Loop
    End If
End If
End Sub
Private Sub grabatesoria1(ByVal rs As Recordset)
Dim Text1 As String
Dim acmd As New ADODB.Command
Dim rb As New ADODB.Recordset
Dim ingresacargo As Integer
Dim xabono, xzona, xmone, xcuenta, xtipo As String
Dim xnumplan, ximpsol, xtcam, xnumpag As Double
On Error GoTo error1
xtcam = 1
VGCNx.BeginTrans
    'Actualizamos el numerador de tipo de ingreso
Set rb = VGCNx.Execute("select * from te_parametroempresa ")
    If rb.RecordCount > 0 Then
         Text1 = Right("0000000000" & Trim(CStr(CDbl(IIf(IsNull(rb!empresanumeingreso + 1) Or Len(Trim(rb!empresanumeingreso)) = 0, 1, rb!empresanumeingreso + 1)))), 6)
         VGCNx.Execute "Update te_parametroempresa Set empresanumeingreso='" & Right("0000000000" & Trim(CStr(Val(Text1))), 6) & "' "
     End If
rb.Close
Set rb = Nothing
VGCNx.CommitTrans
VGCNx.BeginTrans
    Set acmd.ActiveConnection = VGgeneral
    acmd.CommandType = adCmdStoredProc
    acmd.CommandText = "te_abonadocumento_pro"
    acmd.CommandTimeout = 0
    acmd.Prepared = True
    With acmd
        .Parameters("@base") = VGCNx.DefaultDatabase
        .Parameters("@tipo") = "1"
        .Parameters("@numrecibo") = Escadena(Text1)
        .Parameters("@estadoreg") = ""
        .Parameters("@controlctacte") = "1"
        .Parameters("@vendedorcodigo") = wCabe(17)
        .Parameters("@cajacodigo") = RTrim(rs!banco)
        .Parameters("@clientecodigo") = wCabe(16)
        .Parameters("@descripcion") = ""
        If rs.Fields(0) = "C" Then
           .Parameters("@operacion") = "10"
        Else
           .Parameters("@operacion") = "11"
        End If
        .Parameters("@monedacodigo") = wCabe(9)
        .Parameters("@ingsal") = "I"
        .Parameters("@tipocambio") = wCabe(10)
        .Parameters("@totsoles") = Round(IIf(wCabe(9) = "01", rs!importe, rs!importe * wCabe(10)), 2)
        .Parameters("@totdolares") = Round(IIf(wCabe(9) = "02", rs!importe, rs!importe / wCabe(10)), 2)
        .Parameters("@fechadocumento") = wCabe(34)
        .Parameters("@empresa") = VGParametros.empresacodigo
        .Parameters("@observa") = ""
        .Parameters("@transferauto") = ""
        .Parameters("@numreciboegreso") = ""
        .Parameters("@usuario") = g_usuario
        .Parameters("@cabprovinumero") = wCabe(2)
        .Parameters("@fechaact") = Now
     End With
     acmd.Execute
     Set acmd = Nothing
             Set rb = VGCNx.Execute("select * from cc_tipodocumento where tdocumentocodigo='" & rs.Fields(1) & "'")
             xzona = "01": xmone = g_TipoSol: xnumpag = 1
             If rb.RecordCount > 0 Then
                xabono = rb!tdocumentotipo
                xtipo = IIf(IsNull(rb!tdocumentopermiteaplica), Null, rb!tdocumentopermiteaplica)
                If rs.Fields(7) = g_TipoSol Then
                   xcuenta = "" & Trim(rb!tdocumentocuentasoles)
                Else
                   xcuenta = "" & Trim(rb!tdocumentocuentadolares)
                End If
             Else
                xabono = "": xcuenta = "": xtipo = ""
             End If
             rb.Close
             Set rb = Nothing
        
             ' Registramos datos en Tesoreria
             Set acmd.ActiveConnection = VGgeneral
             acmd.CommandType = adCmdStoredProc
             acmd.CommandText = "te_abonadetalledocumento_pro"
             acmd.CommandTimeout = 0
             acmd.Prepared = True
             With acmd
                 .Parameters("@base") = VGCNx.DefaultDatabase
                 .Parameters("@tipo") = "1"
                 .Parameters("@numrecibo") = Text1
                 .Parameters("@estadoreg") = ""
                 .Parameters("@item") = "1"
                 .Parameters("@emisioncheque") = rs.Fields(0)
                 .Parameters("@tipodocconcepto") = rs.Fields(1)
                 .Parameters("@numdocumento") = rs.Fields(2)
                 .Parameters("@carabo") = xabono
                 .Parameters("@formacan") = rs.Fields(3)
                 .Parameters("@tdqc") = rs.Fields(4)
                 .Parameters("@ndqc") = Trim(rs.Fields(6))
                 .Parameters("@tipocajabanco") = rs.Fields(0)
                 .Parameters("@cajabanco") = RTrim(rs!banco)
                 .Parameters("@numctacte") = Escadena(rs.Fields(10))    'numero de cuenta corriente con tamaño 30
                 .Parameters("@adicionactacte") = "C"
                 .Parameters("@monedadocumento") = xmone
                 .Parameters("@monedacancela") = Escadena(rs.Fields(7))
                 .Parameters("@importesoles") = CDbl(IIf(rs.Fields(7) = g_TipoSol, rs.Fields(8), (rs.Fields(8) * xtcam)))
                 .Parameters("@importedolares") = CDbl(IIf(rs.Fields(7) = g_TipoSol, (rs.Fields(8) / xtcam), rs.Fields(8)))
                 .Parameters("@contabledisponi") = 0      'sale de empresas
                 .Parameters("@fechacancela") = rs.Fields(9)
                 .Parameters("@observacion") = Escadena(rs.Fields(11))
                 .Parameters("@usuario") = g_usuario
                 .Parameters("@fechaact") = Now
             End With
             acmd.Execute
             Set acmd = Nothing
             DoEvents
    
    
VGCNx.CommitTrans
Call GeneraAsientoEnlineaTesor1(CDate(wCabe(34)), VGParametros.empresacodigo, "I", Escadena(Text1), 1, "''''", Left(wCabe(9), 2), "C", "E")
' MsgBox "Los datos han sido grabados satisfactoriamente...!!!", vbInformation, MsgTitle
    Exit Sub
error1:
  MsgBox "No se pudo Grabar " & Err.Description & " - " & Err.Number, vbInformation, Caption
  VGCNx.RollbackTrans

  Exit Sub
  Resume

  
 End Sub
Public Sub GeneraAsientoEnlineaTesor1(Fecha As Date, Empresa As String, m_opcion As String, Nrecibo As String, op As Integer, comprobconta As String, monedacodigo As String, cajabanco As String, m_tipovoucher As String)
Dim rsparimpo As ADODB.Recordset
Dim numerror As Integer
Dim Comando As ADODB.Command
numerror = 0
On Error GoTo Proceso

   VGCNx.BeginTrans

Set rsparimpo = New ADODB.Recordset

rsparimpo.Open "Select * From  ct_importartesoreria Where tipooperacion ='" & UCase(m_opcion) & "' ", VGcnxCT, adOpenKeyset, adLockReadOnly
If rsparimpo.RecordCount() > 0 Then

   Set Comando = New ADODB.Command
   With Comando
        .CommandType = adCmdStoredProc
        .CommandText = "te_GeneraAsientosTesoreriaLinea_pro"
        .CommandTimeout = 0
        .ActiveConnection = VGgeneral
        .Parameters.Refresh
        .Parameters("@BaseConta") = VGcnxCT.DefaultDatabase
        .Parameters("@BaseVenta") = VGCNx.DefaultDatabase
        .Parameters("@empresa") = Empresa
        .Parameters("@Asiento") = rsparimpo!asiento
        .Parameters("@SubAsiento") = rsparimpo!SubAsiento
        .Parameters("@Libro") = rsparimpo!Libro
         
        .Parameters("@Mes") = Format(Month(Fecha), "00")
        .Parameters("@Ano") = Year(Fecha)
            
        .Parameters("@tipanal") = "002"
        .Parameters("@Compu") = VGcomputer
        .Parameters("@Usuario") = VGParamSistem.Usuario
        .Parameters("@TipoMov") = Trim(UCase(m_tipovoucher))
        .Parameters("@Nrecibo") = Nrecibo
        .Parameters("@op") = op
        .Parameters("@comprobconta") = comprobconta
        .Parameters("@ajustehaber") = VGParametros.sistemactaajustehab
        .Parameters("@ajustedebe") = VGParametros.sistemactaajustedeb
        .Execute
   End With
   If numerror = 0 Then
        VGCNx.CommitTrans
        Screen.MousePointer = 1
 '       MsgBox "La Contabilizacion se Realizo Satisfactoriamente", vbInformation, "Sistema de Tesoreria"
   End If
End If
Exit Sub
Proceso:
   numerror = 1
   Screen.MousePointer = 1
    MsgBox Err.Description
    VGCNx.RollbackTrans
   Exit Sub
   Resume
End Sub





