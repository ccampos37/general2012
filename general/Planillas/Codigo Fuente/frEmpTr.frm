VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frEmpTr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selección de Empresas"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4905
   Icon            =   "frEmpTr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   4905
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frEmpTr.frx":0442
      Left            =   1290
      List            =   "frEmpTr.frx":0444
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3945
      Width           =   2160
   End
   Begin VB.CommandButton cmEditar 
      Caption         =   "Editar"
      Height          =   360
      Left            =   405
      TabIndex        =   4
      Top             =   4740
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.CommandButton cmNueva 
      Caption         =   "&Nueva"
      Height          =   360
      Left            =   -75
      TabIndex        =   3
      Top             =   4740
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.CommandButton cmSelecc 
      Caption         =   "Siguiente >>"
      Default         =   -1  'True
      Height          =   360
      Left            =   3585
      TabIndex        =   2
      Top             =   3915
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   285
      Top             =   3000
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
            Picture         =   "frEmpTr.frx":0446
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LEmpresas 
      Height          =   2955
      Left            =   60
      TabIndex        =   1
      Top             =   885
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5212
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID Empresa"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Ruta"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Base"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccionar Tabla"
      Height          =   375
      Left            =   105
      TabIndex        =   6
      Top             =   3870
      Width           =   1050
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frEmpTr.frx":0898
      Top             =   210
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selección de Empresas para traslados"
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
      Left            =   645
      TabIndex        =   0
      Top             =   465
      Width           =   3255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000002&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   825
      Left            =   60
      Top             =   15
      Width           =   4830
   End
End
Attribute VB_Name = "frEmpTr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsEmp As New ADODB.Recordset
Dim VCARGA As Boolean
Dim XITEM As ListItem
Dim RSTIPDOC As ADODB.Recordset, RSAFP As ADODB.Recordset, RSSCTR As ADODB.Recordset
Dim RSUBIGEO As ADODB.Recordset, RSBANCO As ADODB.Recordset
Dim RSTIPTRAB As ADODB.Recordset, RSCCOSTO As ADODB.Recordset
Private Sub CMSELECC_CLICK()
'ON ERROR RESUME NEXT
Dim XSW As Integer
Dim CNX_AUX As ADODB.Connection
Dim RS_AUX As ADODB.Recordset
Dim Q As Integer
Dim SQLINSERT  As String
Dim Codigo As String, NOMBRES As String, QUIEBRE As Integer
Set CNX_AUX = New ADODB.Connection
Set RS_AUX = New ADODB.Recordset

    If LEmpresas.ListItems.Count = 0 Then Exit Sub
    
        CNX_AUX.CommandTimeout = 100
        CNX_AUX.ConnectionString = 100
        CNX_AUX.CursorLocation = adUseClient
        'CNX_AUX.Provider = "MICROSOFT.JET.OLEDB.3.51"
        CNX_AUX.ConnectionString = "PROVIDER=SQLOLEDB.1;PERSIST SECURITY INFO=FALSE;USER ID=SOPORTE;PASSWORD=SOPORTE;INITIAL CATALOG=" & LEmpresas.SelectedItem.SubItems(3) & ";DATA SOURCE=" & Mid(VGL_SERVER, 2, Len(VGL_SERVER) - 2)
        CNX_AUX.Open
        ModPlan.TRAS.RUTABASE = LEmpresas.SelectedItem.SubItems(3)
        ModPlan.TRAS.EMPRESA = LEmpresas.SelectedItem.SubItems(1)
        ModPlan.TRAS.ESCADENA = True
        XSW = 0
        IDENTITY = False
        If VGLFRM = 1 Then
                Select Case Combo1.ListIndex
                Case 0
                    ModPlan.TRAS.TABLA = "CCOSTOS"
                    ModPlan.TRAS.ORDENADO = "NOMBRE"
                    ModPlan.TRAS.FIELDK = "CODCCOSTO"
                Case 1
                    ModPlan.TRAS.TABLA = "AREASTRAB"
                    ModPlan.TRAS.ORDENADO = "NOMBRE"
                    ModPlan.TRAS.FIELDK = "CODCCOSTO"
                Case 2
                    ModPlan.TRAS.TABLA = "AFPS"
                    ModPlan.TRAS.ORDENADO = "NOMBRE"
                    ModPlan.TRAS.FIELDK = "CODAFP"
                Case 3
                    ModPlan.TRAS.TABLA = "CONCEPTOS"
                    ModPlan.TRAS.ORDENADO = "NOMBRE"
                    ModPlan.TRAS.FIELDK = "CODIGO"
                Case 4
                    ModPlan.TRAS.TABLA = "BANCOS"
                    ModPlan.TRAS.ORDENADO = "NOMBRE"
                    ModPlan.TRAS.FIELDK = "CODBANCO"
                Case 5
                    ModPlan.TRAS.TABLA = "DOCUMENTOS"
                    ModPlan.TRAS.ORDENADO = "DESCRIP"
                    ModPlan.TRAS.FIELDK = "TIPDOC"
                Case 6
                    ModPlan.TRAS.TABLA = "COLUMPL"
                    ModPlan.TRAS.ORDENADO = "NOMBRE"
                    ModPlan.TRAS.FIELDK = "CODIGO"
                Case 7
                    IDENTITY = True
                    ModPlan.TRAS.ESCADENA = False
                    ModPlan.TRAS.TABLA = "FORMULASCTS"
                    ModPlan.TRAS.ORDENADO = "NOMBRE"
                    ModPlan.TRAS.FIELDK = "CODIGO"
                Case 8
                    IDENTITY = True
                    ModPlan.TRAS.ESCADENA = False
                    ModPlan.TRAS.TABLA = "FORMULASGRATI"
                    ModPlan.TRAS.ORDENADO = "NOMBRE"
                    ModPlan.TRAS.FIELDK = "CODIGO"
                Case 9
                    IDENTITY = True
                    ModPlan.TRAS.ESCADENA = False
                    ModPlan.TRAS.TABLA = "FORMULASVAC"
                    ModPlan.TRAS.ORDENADO = "NOMBRE"
                    ModPlan.TRAS.FIELDK = "CODIGO"
                Case 10
                    IDENTITY = True
                    ModPlan.TRAS.ESCADENA = False
                    ModPlan.TRAS.TABLA = "FORMULASUTIL"
                    ModPlan.TRAS.ORDENADO = "NOMBRE"
                    ModPlan.TRAS.FIELDK = "CODIGO"
                End Select
        Else
            Select Case Combo1.ListIndex
                Case 0
                    XSW = 1
                    ModPlan.TRAS.TABLA = "_TMPTRQ"
                    ModPlan.TRAS.ORDENADO = "CODTRAB"
                    ModPlan.TRAS.FIELDK = "CODTRAB"
                End Select
        End If
        
        If VGLFRM = 1 Then
            RS_AUX.Open "SELECT * FROM " & ModPlan.TRAS.TABLA & " ORDER BY " & ModPlan.TRAS.ORDENADO, CNX_AUX, adOpenStatic, adLockOptimistic
            RS_AUX.Requery
            FrDaTr.LEmpresas.ListItems.Clear
            Do While Not RS_AUX.EOF
                Set XITEM = FrDaTr.LEmpresas.ListItems.Add(, "R" & RS_AUX.Fields(0), RS_AUX.Fields(0), , 1)
                    XITEM.SubItems(1) = RS_AUX.Fields(1)
                    RS_AUX.MoveNext
            Loop
            RS_AUX.Close
            CNX_AUX.Close
            Unload Me
            FrDaTr.Show 1
        Else ' SI ES TRABAJADOR
            Dim RSTRAB As New ADODB.Recordset
            RSTRAB.Open "SELECT CODTRAB, NOMBRES FROM VWTRABAJ", CNX_AUX, adOpenKeyset, adLockReadOnly
            If RSTRAB.EOF Or RSTRAB.RecordCount = 0 Then
                MsgBox "NO SE HAN ENCONTRADO REGISTRO DE TRABAJADORES", vbCritical
                Set RSTRAB = Nothing
                Exit Sub
            End If
            frmComun.CONECTAR RSTRAB
            frmComun.Show 1
                If VGUTIL(1) <> "" Then
                    xTrab = RSTRAB!CODTRAB
                    XTRABNOMBRE = RSTRAB!CODTRAB & " : " & RSTRAB!NOMBRES
                End If
                Set RSTRAB = Nothing
                frTrab.xCodTrab.Text = xTrab
                'AQUI VIENE EL SELECT DEL EMPLEADO
                RSTRAB.Open "SELECT * FROM TRABAJADORES WHERE CODTRAB='" & Trim$(xTrab) & "'", CNX_AUX, adOpenKeyset, adLockReadOnly
                If RSTRAB.RecordCount Then
                    'SE LLENA LA DATA DEL EMPLEADO
                    On Error GoTo ERRVACIAR
                        With RSTRAB
                                    frTrab.xCodTrab.Text = !CODTRAB
                                    frTrab.xApePat.Text = Trim(!ApePat)
                                    frTrab.xApeMat.Text = Trim(!ApeMat)
                                    frTrab.xNombre.Text = Trim(!NOMBRE)
                                    frTrab.xTipDoc.Tag = !TIPODOC
                                    frTrab.xDocIden.Text = "" & !DOCIDEN
                                    frTrab.xFechaNac.Value = !FechaNac
                                    frTrab.xEstadoCivil.ListIndex = !ESTADOCIVIL
                                    frTrab.xNoCalculo.Value = IIf(!NOCALCULO = -1, 1, 0)
                                    frTrab.xUbigeo.Tag = "" & !UBIGEO
                                    frTrab.xDireccion.Text = Trim("" & !DIRECCIÓN)
                                    frTrab.xTelefono.Text = Trim("" & !TELEFONO)
                                    frTrab.xSexo.ListIndex = !Sexo
                                    frTrab.xTipoTrab.Tag = !TIPOTRAB
                                    frTrab.xFechaIng.Value = !FECHAING
                                    frTrab.xSituacion.ListIndex = !SITUACIÓN
                                    frTrab.xDepartamento.Tag = !AREA
                                    frTrab.xCCosto.Tag = !CCosto
                                    frTrab.Xcargo.Text = "" & !CARGO
                                    frTrab.xCtaBanco.Text = "" & !CTABANCO
                                    frTrab.xBanco.Tag = !BANCO
                                    frTrab.xCtaCTS.Text = "" & !CTACTS
                                    frTrab.xBancoCTS.Tag = !BANCOCTS
                                    frTrab.xBasico.Text = Format(!BASICO, "0.00")
                                    frTrab.xNumFicha.Text = "" & !NUMFICHA
                                    frTrab.xCarnetSeg.Text = "" & !CARNETSEG
                                    frTrab.xFondoPens.Tag = !FONDOPENS
                                    frTrab.xCuspp.Text = "" & !CUSPP
                                    frTrab.xMesDevengue.Value = "" & !MESDEVENGUE
                                    frTrab.xFechaIAFP.Value = "" & !FECHAIAFP
                                    frTrab.xEsSaludVida.Value = IIf(!ESSALUDVIDA, 1, 0)
                                    frTrab.xAsigFam.Text = Format(!ASIGFAM, "0.00")
                                    frTrab.xFechaCese.Value = Trim("" & !FECHACESE)
                                    frTrab.xCodAlt.Text = Trim("" & !CODIGOALT)
                                    frTrab.xCodCTR.Tag = !CODSCTR
                                    frTrab.xRucEPS.Text = Trim("" & !RUCEPS)
                                    frTrab.xContrato.ListIndex = 0 + !TIPOCONTRATO
                                        If !TIPOCONTRATO = 1 Then
                                                If Not IsNull(!FECHATERMINO) Then frTrab.xFechaTermino.Value = !FECHATERMINO
                                        End If
                                    frTrab.xOpcion01.Value = !OPCION01
                                    frTrab.xOpcion02.Value = !OPCION02
                                    frTrab.xNoPDT.Value = !NOPDT
                                    frTrab.xOpcionA.Text = !OPCIONA
                                    frTrab.xOpcionB.Text = !OPCIONB
                
                                            frTrab.xFoto.Picture = LoadPicture(REGSISTEMA.PATHFOTOS & "\" & frTrab.xCodTrab.Text & ".FTE")
                                            If frTrab.xFoto.Picture <> 0 Then
                                                frTrab.xFoto.Tag = REGSISTEMA.PATHFOTOS & "\" & frTrab.xCodTrab.Text & ".FTE"
                                            End If
                        End With
                        
                        Dim RSAUX2 As New ADODB.Recordset
                                RSAUX2.Open "SELECT CODCCOSTO, NOMBRE FROM AREASTRAB WHERE CODCCOSTO='" & frTrab.xDepartamento.Tag & "'", DBSYSTEM, adOpenStatic
                                
                                If RSAUX2.EOF Then
                                    MsgBox "NO SE ENCUENTRA EL AREA DE TRABAJO DEL TRABAJADOR, SELECCIONAR OTRO", vbCritical
                                Else
                                    frTrab.xDepartamento.Text = RSAUX2!CODCCOSTO & " : " & RSAUX2!NOMBRE
                                End If
                                
                                Set RSAUX2 = Nothing
                                RSTIPDOC.FIND "TIPDOC='" & frTrab.xTipDoc.Tag & "'"
                                
                                If RSTIPDOC.EOF Then
                                        MsgBox "EL TIPO DE DOCUMENTO AL QUE SE REFERIA EL REGISTRO YA NO EXISTE, SELECCIONE OTRO", vbCritical
                                        frTrab.xTipDoc.Tag = ""
                                Else
                                        frTrab.xTipDoc.Text = RSTIPDOC!TIPDOC & " :  " & RSTIPDOC!DESCRIP
                                End If
                                
                                RSUBIGEO.FIND "CODIGO='" & frTrab.xUbigeo.Tag & "'"
                                
                                If RSUBIGEO.EOF Then
                                    MsgBox "EL CÓDIGO DE UBICACIÓN GEOGRÁFICA YA NO EXISTE, SELECCIONE OTRO", vbCritical
                                    frTrab.xUbigeo.Tag = ""
                                Else
                                    frTrab.xUbigeo.Text = RSUBIGEO!Codigo & " : " & RSUBIGEO!LUGAR
                                End If
                    
                                RSTIPTRAB.MoveFirst
                                RSTIPTRAB.FIND "TIPTRAB='" & frTrab.xTipoTrab.Tag & "'"
                                
                                If RSTIPTRAB.EOF Then
                                            MsgBox "EL TIPO DE TRABAJADOR AL QUE SE REFIERE EL REGISTRO ACTUAL YA NO EXISTE, SELECCIONE OTRO", vbCritical
                                            frTrab.xTipoTrab.Tag = ""
                                Else
                                            frTrab.xTipoTrab.Text = RSTIPTRAB!TIPTRAB & " :  " & RSTIPTRAB!DESCRIP
                                End If
                    
                                RSCCOSTO.FIND "CODCCOSTO='" & frTrab.xCCosto.Tag & "'"
                                
                                If RSCCOSTO.EOF Then
                                        MsgBox "EL CENTRO DE COSTO AL QUE SE REFIERE YA NO EXISTE, SELECCIONE OTRO", vbCritical
                                        frTrab.xCCosto.Tag = ""
                                Else
                                        frTrab.xCCosto.Text = RSCCOSTO!CODCCOSTO & " :  " & RSCCOSTO!NOMBRE
                                End If
                    
                                RSBANCO.FIND "CODBANCO='" & frTrab.xBanco.Tag & "'"
                                If RSBANCO.EOF Then
                                    MsgBox "EL CÓDIGO DEL BANCO DE LA CUENTA BANCARIA DE DEPÓSITO DE REMUNERACIONES NO EXISTE, SELECCIONE OTRO", vbCritical
                                    frTrab.xBanco.Tag = ""
                                Else
                                    frTrab.xBanco.Text = RSBANCO!CODBANCO & " :  " & RSBANCO!NOMBRE
                                End If
                                
                                RSBANCO.MoveFirst
                                RSBANCO.FIND "CODBANCO='" & frTrab.xBancoCTS.Tag & "'"
                                If RSBANCO.EOF Then
                                    MsgBox "EL CÓDIGO DEL BANCO DE LA CUENTA BANCARIA DE DEPÓSITO DE CTS NO EXISTE, SELECCIONE OTRO", vbCritical
                                    frTrab.xBancoCTS.Tag = ""
                                Else
                                    frTrab.xBancoCTS.Text = RSBANCO!CODBANCO & " :  " & RSBANCO!NOMBRE
                                End If
                                
                                RSAFP.FIND "CODAFP='" & frTrab.xFondoPens.Tag & "'"
                                If RSAFP.EOF Then
                                    MsgBox "LA ADMINISTRADORA DE FONDO DE PENSIONES NO EXISTE, SELECCIONE OTRA", vbCritical
                                    frTrab.xFondoPens.Tag = ""
                                Else
                                    frTrab.xFondoPens.Text = RSAFP!CODAFP & " :  " & RSAFP!NOMBRE
                                End If
                                
                                RSSCTR.FIND "CODCAR='" & frTrab.xCodCTR.Tag & "'"
                                If RSSCTR.EOF Then
                                    MsgBox "EL REGISTRO REFERENCIADO DEL CENTRO DE ALTO RIESGO - SCTR YA NO EXISTE, SELECCIONE OTRO", vbCritical
                                    frTrab.xCodCTR.Tag = ""
                                Else
                                    frTrab.xCodCTR.Text = RSSCTR!CODCAR & " :  " & RSSCTR!NOMBRE
                                End If
                End If
                Unload Me
        End If
Exit Sub
ERRVACIAR:
Resume Next
End Sub

Private Sub Form_Activate()
Combo1.ListIndex = 0
End Sub

Private Sub Form_Load()
    CargaEmp
    VCARGA = True
    'CREA LA TABLA DE ACCESOS Y USUARIOS
    If VGLFRM = 1 Then
        frEmpTr.Combo1.AddItem "CENTROS DE COSTO"
        frEmpTr.Combo1.AddItem "AREAS DE TRABAJO"
        frEmpTr.Combo1.AddItem "AFPS"
        frEmpTr.Combo1.AddItem "CONCEPTOS"
        frEmpTr.Combo1.AddItem "BANCOS"
        frEmpTr.Combo1.AddItem "DOCUMENTOS"
        frEmpTr.Combo1.AddItem "COLUMNAS DE PLANILLA"
        frEmpTr.Combo1.AddItem "FORMULA CTS"
        frEmpTr.Combo1.AddItem "FORMULA DE GRATIFICACION"
        frEmpTr.Combo1.AddItem "FORMULA DE VACACIONES"
        frEmpTr.Combo1.AddItem "FORMULA SUTIL"
    Else
        Combo1.AddItem "TRABAJADORES"
    End If
    Combo1.ListIndex = 0
End Sub

Public Sub CargaEmp()
On Error Resume Next
    Set RsEmp = Nothing
    RsEmp.Open "SELECT * FROM EMPRESAS ORDER BY NOMBRE", DBSTARPLAN, adOpenStatic, adLockOptimistic
    RsEmp.Requery
    LEmpresas.ListItems.Clear
    Do While Not RsEmp.EOF
        Set XITEM = LEmpresas.ListItems.Add(, "R" & RsEmp!RUC, RsEmp!RUC, , 1)
        XITEM.SubItems(1) = RsEmp!NOMBRE
        XITEM.SubItems(2) = RsEmp!DIRMASTER
        XITEM.SubItems(3) = RsEmp!DIRALMACEN
        RsEmp.MoveNext
    Loop
    RsEmp.Close
End Sub

Private Sub LEMPRESAS_DBLCLICK()
On Error Resume Next
    CMSELECC_CLICK
End Sub

