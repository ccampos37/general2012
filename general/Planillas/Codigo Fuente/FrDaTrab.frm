VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrDaTrab 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccion de Datos"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   7035
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "<< Anterior"
      Height          =   360
      Left            =   4500
      TabIndex        =   9
      Top             =   870
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Siguiente >>"
      Height          =   360
      Left            =   5745
      TabIndex        =   8
      Top             =   870
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Seleccionar Todos"
      Height          =   300
      Left            =   30
      TabIndex        =   6
      Top             =   6450
      Width           =   2145
   End
   Begin VB.CommandButton cmSelecc 
      Caption         =   "Terminar"
      Height          =   360
      Left            =   5670
      TabIndex        =   2
      Top             =   6420
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<< Atras"
      Height          =   360
      Left            =   4410
      TabIndex        =   5
      Top             =   6420
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Eliminar los Actuales"
      Height          =   195
      Left            =   30
      TabIndex        =   3
      Top             =   7125
      Visible         =   0   'False
      Width           =   2100
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   2730
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrDaTrab.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LEmpresas 
      Height          =   4935
      Left            =   -15
      TabIndex        =   0
      Top             =   1305
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   8705
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   7832
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label xEmpresa 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   660
      TabIndex        =   7
      Top             =   465
      Width           =   585
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   615
      TabIndex        =   4
      Top             =   345
      Width           =   4125
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccion de Trabajadores a Trasladar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   660
      TabIndex        =   1
      Top             =   210
      Width           =   3315
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   60
      Picture         =   "FrDaTrab.frx":015C
      Top             =   195
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000002&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   825
      Left            =   0
      Top             =   0
      Width           =   6960
   End
End
Attribute VB_Name = "FrDaTrab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS_AUX As New ADODB.Recordset
Dim CNX_AUX2 As New ADODB.Connection
Private Sub CMSELECC_CLICK()
    On Error Resume Next
    Dim RS As New ADODB.Recordset, RSCOMPARA As New ADODB.Recordset, RSAUX As New ADODB.Recordset
    Dim RSAUXTRAB As New ADODB.Recordset, RSAUXTRAB2 As New ADODB.Recordset
    Dim CNX_AUX As ADODB.Connection
    Dim X As Long
    Dim Y As Integer
    Y = 0
    Dim FLAG As Integer
    Dim FILATXT As String
    Set CNX_AUX = New ADODB.Connection
        CNX_AUX.CommandTimeout = 100
        CNX_AUX.ConnectionString = 100
        CNX_AUX.CursorLocation = adUseClient
        CNX_AUX.Provider = "MICROSOFT.JET.OLEDB.3.51"
        CNX_AUX.ConnectionString = "DATA SOURCE=" & ModPlan.TRAS.RUTABASE
        CNX_AUX.Open
            For X = 1 To Me.LEmpresas.ListItems.Count
                If LEmpresas.ListItems(X).Checked = True Then
                        'PREGUNTA SI EL TRABAJADOR YA PERTENCE A LA EMPRESA
                        SQLTIPOTRAB = "SELECT *  FROM TRABAJADORES WHERE CODTRAB='" & Trim$(LEmpresas.ListItems(X).Text) & "'"
                        RSAUXTRAB.Open SQLTIPOTRAB, CNX_AUX, adOpenKeyset, adLockOptimistic
                        If RSAUXTRAB.RecordCount Then 'SI NO EXISTE EL CODIGO
                            ModPlan.DATOTRABAJADOR.CodigoTrab = Trim$(LEmpresas.ListItems(X).Text)
                        Else
                            'INGRESAR EL CODIGO
                            iNGdATO.Caption = "CODIGO DEL TRABAJADOR"
                            iNGdATO.Label1.Caption = "CODIGO"
                            LlamaFrm = 1
                            iNGdATO.Show 1
                        End If
                            ModPlan.DATOTRABAJADOR.DOCUMENTO = RSAUXTRAB.Fields(5)
                            ModPlan.DATOTRABAJADOR.TIPODOCUMENTO = RSAUXTRAB.Fields(4)
                            ModPlan.DATOTRABAJADOR.ApePat = RSAUXTRAB.Fields(1)
                            ModPlan.DATOTRABAJADOR.ApeMat = RSAUXTRAB.Fields(2)
                            ModPlan.DATOTRABAJADOR.NOMBRE = RSAUXTRAB.Fields(3)
                            ModPlan.DATOTRABAJADOR.FechaNac = RSAUXTRAB.Fields(6)
                            ModPlan.DATOTRABAJADOR.ESTADOCIVIL = RSAUXTRAB.Fields(7)
                            ModPlan.DATOTRABAJADOR.UBIGEO = RSAUXTRAB.Fields(8)
                            ModPlan.DATOTRABAJADOR.DIRECCION = RSAUXTRAB.Fields(9)
                            ModPlan.DATOTRABAJADOR.TELEFONO = RSAUXTRAB.Fields(10)
                            ModPlan.DATOTRABAJADOR.Sexo = RSAUXTRAB.Fields(11)
                            ModPlan.DATOTRABAJADOR.FECHAING = RSAUXTRAB.Fields(13)
                            ModPlan.DATOTRABAJADOR.SITUACION = RSAUXTRAB.Fields(14)
                            ModPlan.DATOTRABAJADOR.CARGO = RSAUXTRAB.Fields(18)
                            ModPlan.DATOTRABAJADOR.BASICO = RSAUXTRAB.Fields(23)
                            ModPlan.DATOTRABAJADOR.NFICHA = RSAUXTRAB.Fields(24)
                            ModPlan.DATOTRABAJADOR.CSEGURO = RSAUXTRAB.Fields(25)
                            ModPlan.DATOTRABAJADOR.DEVENGUE = RSAUXTRAB.Fields(28)
                            ModPlan.DATOTRABAJADOR.SALUDVIDA = RSAUXTRAB.Fields(29)
                            ModPlan.DATOTRABAJADOR.ASIGFAM = RSAUXTRAB.Fields(30)
                            ModPlan.DATOTRABAJADOR.FECHACESE = RSAUXTRAB.Fields(31)
                            ModPlan.DATOTRABAJADOR.ESTADOINTERNO = RSAUXTRAB.Fields(32)
                            ModPlan.DATOTRABAJADOR.CODIGOALTERNO = RSAUXTRAB.Fields(33)
                                    SQLTIPOTRAB = "SELECT *  FROM TRABAJADORES WHERE DOCIDEN='" & ModPlan.DATOTRABAJADOR.DOCUMENTO & "' AND TIPODOC='" & ModPlan.DATOTRABAJADOR.TIPODOCUMENTO & "'"
                                    RSAUXTRAB2.Open SQLTIPOTRAB, DBSYSTEM, adOpenKeyset, adLockOptimistic
                                    If Not RSAUXTRAB2.RecordCount Then 'SI NO EXISTE EL TRABAJADOR
                                        If MsgBox("SELECCIONE EL TIPO DE TRABAJADOR CON QUE PASARA EL TRABAJADOR " & LEmpresas.ListItems(X).ListSubItems(1).Text & ", DESEA CONTINUAR?", vbYesNo, "INFORMACION DE INTEGRIDAD") = vbYes Then
                                            SQLTIPOTRAB = "SELECT *  FROM TIPOSTRAB"
                                            RSAUX.Open SQLTIPOTRAB, DBSYSTEM, adOpenKeyset, adLockOptimistic
                                                frmComun.CONECTAR RSAUX
                                                frmComun.Show 1
                                                If VGUTIL(1) <> "" Then
                                                    ModPlan.DATOTRABAJADOR.TIPOTRAB = VGUTIL(1)
                                                End If
                                                RSAUX.Close
                                                If MsgBox("SELECCIONE EL BANCO CON QUE PASARA EL TRABAJADOR " & LEmpresas.ListItems(X).ListSubItems(1).Text & ", DESEA CONTINUAR?", vbYesNo, "INFORMACION DE INTEGRIDAD") = vbYes Then
                                                SQLTIPOTRAB = "SELECT *  FROM BANCOS"
                                                RSAUX.Open SQLTIPOTRAB, DBSYSTEM, adOpenKeyset, adLockOptimistic
                                                    frmComun.CONECTAR RSAUX
                                                    frmComun.Show 1
                                                    If VGUTIL(1) <> "" Then
                                                        ModPlan.DATOTRABAJADOR.BANCO = VGUTIL(1)
                                                        iNGdATO.Caption = "CTA BANCO"
                                                        iNGdATO.Label1.Caption = "CTA BANCO"
                                                        LlamaFrm = 2
                                                        iNGdATO.Show 1
                                                    End If
                                                    RSAUX.Close
                                                    If MsgBox("SELECCIONE EL AREA DE TRABAJO CON QUE PASARA EL TRABAJADOR " & LEmpresas.ListItems(X).ListSubItems(1).Text & ", DESEA CONTINUAR?", vbYesNo, "INFORMACION DE INTEGRIDAD") = vbYes Then
                                                        SQLTIPOTRAB = "SELECT *  FROM AREASTRAB"
                                                        RSAUX.Open SQLTIPOTRAB, DBSYSTEM, adOpenKeyset, adLockOptimistic
                                                            frmComun.CONECTAR RSAUX
                                                            frmComun.Show 1
                                                            If VGUTIL(1) <> "" Then
                                                                ModPlan.DATOTRABAJADOR.AREA = VGUTIL(1)
                                                            End If
                                                            RSAUX.Close
                                                            If MsgBox("SELECCIONE EL CENTRO DE COSTO CON QUE PASARA EL TRABAJADOR " & LEmpresas.ListItems(X).ListSubItems(1).Text & ", DESEA CONTINUAR?", vbYesNo, "INFORMACION DE INTEGRIDAD") = vbYes Then
                                                                SQLTIPOTRAB = "SELECT *  FROM CCOSTOS"
                                                                RSAUX.Open SQLTIPOTRAB, DBSYSTEM, adOpenKeyset, adLockOptimistic
                                                                    frmComun.CONECTAR RSAUX
                                                                    frmComun.Show 1
                                                                    If VGUTIL(1) <> "" Then
                                                                        ModPlan.DATOTRABAJADOR.CCosto = VGUTIL(1)
                                                                    End If
                                                                    iNGdATO.Caption = "DEPARTAMENTO DE TRABAJO"
                                                                    iNGdATO.Label1.Caption = "DEPTO"
                                                                    LlamaFrm = 5
                                                                    iNGdATO.Show 1
                                                                    RSAUX.Close
                                                                    If MsgBox("SELECCIONE EL BANCO CTS CON QUE PASARA EL TRABAJADOR " & LEmpresas.ListItems(X).ListSubItems(1).Text & ", DESEA CONTINUAR?", vbYesNo, "INFORMACION DE INTEGRIDAD") = vbYes Then
                                                                        SQLTIPOTRAB = "SELECT *  FROM BANCOS"
                                                                        RSAUX.Open SQLTIPOTRAB, DBSYSTEM, adOpenKeyset, adLockOptimistic
                                                                            frmComun.CONECTAR RSAUX
                                                                            frmComun.Show 1
                                                                                    If VGUTIL(1) <> "" Then
                                                                                        ModPlan.DATOTRABAJADOR.BANCOCTS = VGUTIL(1)
                                                                                        iNGdATO.Caption = "CTA CTE DEL TRABAJADOR"
                                                                                        iNGdATO.Label1.Caption = "CTA CTE"
                                                                                        LlamaFrm = 3
                                                                                        iNGdATO.Show 1
                                                                                            RSAUX.Close
                                                                                            If MsgBox("SELECCIONE EL AFP CON QUE PASARA EL TRABAJADOR " & LEmpresas.ListItems(X).ListSubItems(1).Text & ", DESEA CONTINUAR?", vbYesNo, "INFORMACION DE INTEGRIDAD") = vbYes Then
                                                                                                SQLTIPOTRAB = "SELECT *  FROM AFPS"
                                                                                                RSAUX.Open SQLTIPOTRAB, DBSYSTEM, adOpenKeyset, adLockOptimistic
                                                                                                    frmComun.CONECTAR RSAUX
                                                                                                    frmComun.Show 1
                                                                                                    If VGUTIL(1) <> "" Then
                                                                                                        ModPlan.DATOTRABAJADOR.AFP = VGUTIL(1)
                                                                                                        iNGdATO.Caption = "CUSPP DEL TRABAJADOR"
                                                                                                        iNGdATO.Label1.Caption = "CUSPP"
                                                                                                        LlamaFrm = 4
                                                                                                        iNGdATO.Show 1
                                                                                                        RSAUX.Close
                                                                                                        If MsgBox("SELECCIONE EL CENTRO DE RIESGO QUE AFECTARA AL TRABAJADOR " & LEmpresas.ListItems(X).ListSubItems(1).Text & ", DESEA CONTINUAR?", vbYesNo, "INFORMACION DE INTEGRIDAD") = vbYes Then
                                                                                                            SQLTIPOTRAB = "SELECT *  FROM CENTROSAR"
                                                                                                            RSAUX.Open SQLTIPOTRAB, DBSYSTEM, adOpenKeyset, adLockOptimistic
                                                                                                            frmComun.CONECTAR RSAUX
                                                                                                            frmComun.Show 1
                                                                                                            If VGUTIL(1) <> "" Then
                                                                                                                ModPlan.DATOTRABAJADOR.CENTROAR = VGUTIL(1)
                                                                                                                'GRABACION
                                                                                                                        FECHANACI = DateSQL(ModPlan.DATOTRABAJADOR.FechaNac)
                                                                                                                        INGRESO = DateSQL(ModPlan.DATOTRABAJADOR.FECHAING)
                                                                                                                        DEVENGUEMES = DateSQL(ModPlan.DATOTRABAJADOR.DEVENGUE)
                                                                                                                        CESE = DateSQL(ModPlan.DATOTRABAJADOR.FECHACESE)
                                                                                                                        SQLEXEC = "INSERT INTO TRABAJADORES (CODTRAB, APEPAT, APEMAT, NOMBRE, TIPDOC, DOCIDEN, FECHANAC, ESTADOCIVIL, UBIGEO, DIRECCIÓN, TELEFONO, SEXO, TIPOTRAB, FECHAING, SITUACIÓN, AREA, CCOSTO, DEPARTAMENTO, CARGO, CTABANCO, BANCO, CTACTS, BANCOCTS, BASICO, NUMFICHA, CARNETSEG, FONDOPENS, CUSPP, MESDEVENGUE, ESSALUDVIDA, ASIGFAM, FECHACESE, ESTADOINTERNO, CODIGOALT, CODSCTR) " & _
                                                                                                                       " VALUES ('" & ModPlan.DATOTRABAJADOR.CodigoTrab & "','" & ModPlan.DATOTRABAJADOR.ApePat & "','" & ModPlan.DATOTRABAJADOR.ApeMat & "','" & ModPlan.DATOTRABAJADOR.NOMBRE & "','" & ModPlan.DATOTRABAJADOR.TIPODOCUMENTO & "','" & ModPlan.DATOTRABAJADOR.DOCUMENTO & "'," & IIf(Len(Trim(FECHANACI)) > 0, FECHANACI, "NULL") & "," & ModPlan.DATOTRABAJADOR.ESTADOCIVIL & ",'" & ModPlan.DATOTRABAJADOR.UBIGEO & "','" & ModPlan.DATOTRABAJADOR.DIRECCION & "','" & _
                                                                                                                       ModPlan.DATOTRABAJADOR.TELEFONO & "'," & ModPlan.DATOTRABAJADOR.Sexo & ",'" & ModPlan.DATOTRABAJADOR.TIPOTRAB & "'," & IIf(Len(Trim(INGRESO)) > 0, INGRESO, "NULL") & ",'" & ModPlan.DATOTRABAJADOR.SITUACION & "','" & ModPlan.DATOTRABAJADOR.AREA & "','" & ModPlan.DATOTRABAJADOR.CCosto & "','" & ModPlan.DATOTRABAJADOR.DEPARTAMENTO & "','" & ModPlan.DATOTRABAJADOR.CARGO & "','" & ModPlan.DATOTRABAJADOR.CTABANCO & "','" & ModPlan.DATOTRABAJADOR.BANCO & "','" & ModPlan.DATOTRABAJADOR.CTACTE & "','" & ModPlan.DATOTRABAJADOR.BANCOCTS & "'," & ModPlan.DATOTRABAJADOR.BASICO & ",'" & ModPlan.DATOTRABAJADOR.NFICHA & "','" & ModPlan.DATOTRABAJADOR.CSEGURO & "','" & _
                                                                                                                       ModPlan.DATOTRABAJADOR.AFP & "','" & ModPlan.DATOTRABAJADOR.CUSPP & "'," & IIf(Len(Trim(DEVENGUEMES)), DEVENGUEMES, "NULL") & "," & IIf(ModPlan.DATOTRABAJADOR.SALUDVIDA = "TRUE", 1, 0) & "," & ModPlan.DATOTRABAJADOR.ASIGFAM & "," & IIf(Len(Trim(CESE)) > 0, CESE, "NULL") & "," & ModPlan.DATOTRABAJADOR.ESTADOINTERNO & ",'" & ModPlan.DATOTRABAJADOR.CODIGOALTERNO & "','" & ModPlan.DATOTRABAJADOR.CENTROAR & "')"
                                                                                                                       DBSYSTEM.Execute SQLEXEC
                                                                                                                        Y = Y + 1
                                                                                                            End If
                                                                                                        End If
                                                                                                    End If
                                                                                    End If
                                                                            
                                                                    Else
                                                                        MsgBox "SE ABORTO EL TRASLADO DEL TRABAJADOR " & Me.LEmpresas.SelectedItem.SubItems(1)
                                                                    End If
                                                            Else
                                                                MsgBox "SE ABORTO EL TRASLADO DEL TRABAJADOR " & Me.LEmpresas.SelectedItem.SubItems(1)
                                                            End If
                                                    Else
                                                        MsgBox "SE ABORTO EL TRASLADO DEL TRABAJADOR " & Me.LEmpresas.SelectedItem.SubItems(1)
                                                    End If
                                            Else
                                                MsgBox "SE ABORTO EL TRASLADO DEL TRABAJADOR " & Me.LEmpresas.SelectedItem.SubItems(1)
                                            End If
                                        Else
                                            MsgBox "SE ABORTO EL TRASLADO DEL TRABAJADOR " & Me.LEmpresas.SelectedItem.SubItems(1)
                                        End If
                                        RSAUXTRAB2.Close
                                    Else
                                        MsgBox "SE ABORTO EL TRASLADO DEL TRABAJADOR " & Me.LEmpresas.SelectedItem.SubItems(1)
                                    End If
                                    RSAUXTRAB.Close
                        End If
                End If
            Next X
        If Y > 0 Then
            MsgBox "LA OPERACIÓN SE REALIZÓ SATISFACTORIAMENTE, " & CStr(Y) & " REGISTROS AFECTADOS", vbInformation, "CONFIRMACIÓN"
        End If
End Sub

Private Sub COMMAND1_CLICK()
On Error Resume Next
    Unload Me
    frEmpTr.Show 1
End Sub

Private Sub COMMAND2_CLICK()
    If MAXIMOQUIEBRE > QQUIEBRE Then
        QQUIEBRE = QQUIEBRE + 1
                    RS_AUX.Open "SELECT * FROM " & ModPlan.TRAS.TABLA & " WHERE QUIEBRE =" & QQUIEBRE & "  ORDER BY " & ModPlan.TRAS.ORDENADO, DBSYSTEM, adOpenStatic, adLockOptimistic
                    RS_AUX.Requery
                    If RS_AUX.RecordCount Then
                            FrDaTrab.LEmpresas.ListItems.Clear
                                  Do While Not RS_AUX.EOF
                                        Set XITEM = FrDaTrab.LEmpresas.ListItems.Add(, "R" & RS_AUX.Fields(0), RS_AUX.Fields(0), , 1)
                                        XITEM.SubItems(1) = RS_AUX.Fields(1)
                                        RS_AUX.MoveNext
                                    Loop
                    End If
                    RS_AUX.Close
    End If
End Sub

Private Sub COMMAND3_Click()
    If QQUIEBRE > 1 Then
        QQUIEBRE = QQUIEBRE - 1
                        RS_AUX.Open "SELECT * FROM " & ModPlan.TRAS.TABLA & " WHERE QUIEBRE =" & QQUIEBRE & "  ORDER BY " & ModPlan.TRAS.ORDENADO, DBSYSTEM, adOpenStatic, adLockOptimistic
                        RS_AUX.Requery
                        If RS_AUX.RecordCount Then
                                FrDaTrab.LEmpresas.ListItems.Clear
                                      Do While Not RS_AUX.EOF
                                            Set XITEM = FrDaTrab.LEmpresas.ListItems.Add(, "R" & RS_AUX.Fields(0), RS_AUX.Fields(0), , 1)
                                            XITEM.SubItems(1) = RS_AUX.Fields(1)
                                            RS_AUX.MoveNext
                                        Loop
                        End If
                        RS_AUX.Close
    End If
End Sub
Private Sub CHECK2_CLICK()
On Error Resume Next
If Check2.Value = 1 Then
    For X = 1 To Me.LEmpresas.ListItems.Count
        LEmpresas.ListItems(X).Checked = True
    Next
Else
    For X = 1 To Me.LEmpresas.ListItems.Count
        LEmpresas.ListItems(X).Checked = False
    Next
End If
End Sub
Private Sub FORM_LOAD()
    xEmpresa.Caption = ModPlan.TRAS.EMPRESA
End Sub


