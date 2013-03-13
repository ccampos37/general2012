VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "CONTROLAYUDA.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FrmRepSaldoCliente 
   Caption         =   "Saldos por Cliente"
   ClientHeight    =   3675
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   ScaleHeight     =   3675
   ScaleWidth      =   8685
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport oCrystalReport 
      Left            =   360
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_Cta 
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   360
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   661
      XcodMaxLongitud =   20
      xcodwith        =   2000
      NomTabla        =   "cc_tipodocumento"
      TituloAyuda     =   "Ayuda de Cuentas"
      ListaCampos     =   "tdocumentocodigo(1),tdocumentodescripcion(1),tdocumentocuentasoles(1),tdocumentocuentadolares(1)"
      XcodCampo       =   "tdocumentocuentasoles"
      XListCampo      =   "tdocumentocuentadolares"
      ListaCamposDescrip=   "Codigo,Descripcion,Cuenta S/,Cuenta $"
      ListaCamposText =   "tdocumentocodigo,tdocumentodescripcion,tdocumentocuentasoles,tdocumentocuentadolares"
      Requerido       =   0   'False
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_Cliente 
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   840
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   661
      XcodMaxLongitud =   11
      xcodwith        =   800
      NomTabla        =   "vt_cliente"
      TituloAyuda     =   "Ayuda de Clientes"
      ListaCampos     =   "clientecodigo(1),clienterazonsocial(1),clienteruc(1),clientedireccion(1),clientedistrito(1),clientesuspendido(1)"
      XcodCampo       =   "clientecodigo"
      XListCampo      =   "clienterazonsocial"
      ListaCamposDescrip=   "Codigo,RazonSocial,RUC,Direccion,Distrito,Suspendido"
      ListaCamposText =   "clientecodigo,clienterazonsocial,clienteruc,clientedireccion,clientedistrito,clientesuspendido"
      Requerido       =   0   'False
   End
   Begin VB.CheckBox chk 
      Height          =   375
      Index           =   2
      Left            =   7680
      TabIndex        =   4
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   4440
      TabIndex        =   6
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   2760
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3600
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1920
      Width           =   3090
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   4395
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1920
      Visible         =   0   'False
      Width           =   1305
   End
   Begin MSComCtl2.DTPicker DTHasta 
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   1320
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   24510465
      CurrentDate     =   37518
   End
   Begin VB.Label lbl 
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   720
      TabIndex        =   12
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label lbl 
      Caption         =   "Cuenta Contable"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   720
      TabIndex        =   11
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label lbl 
      Caption         =   "Incluido Letra x Abonar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   5040
      TabIndex        =   10
      Top             =   3360
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label lbl 
      Caption         =   "Hasta la Fecha"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   9
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label lbl 
      Caption         =   "Moneda"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   720
      TabIndex        =   8
      Top             =   1920
      Width           =   1215
   End
End
Attribute VB_Name = "FrmRepSaldoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim index_combo As Integer
''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim busca As New dll_apisgen.dll_apis
Dim adll As New dllgeneral.dll_general

Private Sub chk_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
     If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmdAceptar_Click(Index As Integer)
Dim X As String
On Error GoTo Errores
                 
 Screen.MousePointer = 11
  
 With oCrystalReport
        .Reset
        .ReportFileName = RutaRepProc & "RepccSaldoxCliente.rpt"
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        .LogOnServer "pdssql.dll", _
         busca.LeerIni(App.Path & "\Camtex.ini", "Bgeneral", "dserver", ""), _
         busca.LeerIni(App.Path & "\Camtex.ini", "Bgeneral", "dbase", ""), _
         busca.LeerIni(App.Path & "\Camtex.ini", "Bgeneral", "duser", ""), _
         busca.LeerIni(App.Path & "\Camtex.ini", "Bgeneral", "dpass", "")
        .Connect = _
        "DSN=" & busca.LeerIni(App.Path & "\Camtex.ini", "Bgeneral", "dserver", "") & ";" & _
        "DSQ=" & busca.LeerIni(App.Path & "\Camtex.ini", "Bgeneral", "dbase", "") & ";" & _
        "UID=" & busca.LeerIni(App.Path & "\Camtex.ini", "Bgeneral", "duser", "") & ";" & _
        "PWD=" & busca.LeerIni(App.Path & "\Camtex.ini", "Bgeneral", "dpass", "")
        '''''''''''''''''''''''''''''''''''''''''''''''''''
        .DiscardSavedData = True
        .Formulas(0) = "Empresa='" & Right(Trim(g_DetalleEmpresa), Len(Trim(g_DetalleEmpresa)) - 3) & "'"
        If Trim(Ctr_Cta.xclave) <> "" Then
            .Formulas(1) = "Cuenta='" & Ctr_Cta.xnombre & "'"
        Else
            .Formulas(1) = "Cuenta='TODOS'"
        End If
        If Trim(Ctr_Cliente.xclave) <> "" Then
            .Formulas(2) = "Cliente='" & Ctr_Cliente.xnombre & "'"
        Else
            .Formulas(2) = "Cliente='TODOS'"
        End If
        .Formulas(3) = "Hasta='" & DTHasta & "'"
        If Combo1.ListIndex <> -1 Then
            .Formulas(4) = "Moneda='" & Combo1.Text & "'"
        Else
            .Formulas(4) = "Moneda='TODOS'"
        End If
        
        .StoredProcParam(0) = busca.LeerIni(App.Path & "\Camtex.ini", "Bventas", "dbase", "")
        If Trim(txt(1)) <> "" Then
            If Trim(txt(1)) = "01" Then
                .StoredProcParam(1) = IIf(Trim(Ctr_Cta.xclave) = "", "%", Trim(Ctr_Cta.xclave))
            ElseIf Trim(txt(1)) = "02" Then
                .StoredProcParam(1) = IIf(Trim(Ctr_Cta.xnombre) = "", "%", Trim(Ctr_Cta.xnombre))
            End If
        Else
            .StoredProcParam(1) = IIf(Trim(Ctr_Cta.xclave) = "", "%", Trim(Ctr_Cta.xclave))
        End If
        .StoredProcParam(2) = IIf(Trim(Ctr_Cliente.xclave) = "", "%", Trim(Ctr_Cliente.xclave))
        .StoredProcParam(3) = IIf(Trim(txt(1)) = "", "%", Trim(txt(1)))
        .StoredProcParam(4) = Format(DTHasta, "DD/MM/YYYY")
        .StoredProcParam(5) = IIf(chk(2).Value = 0, False, True)
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintSetupBtn = True
        .WindowShowExportBtn = True
        .WindowShowZoomCtl = True
        .WindowShowNavigationCtls = True
        .WindowShowPrintBtn = True
        .WindowTitle = "Saldos por Cliente"
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        .SubreportToChange = "RepccSubSaldoxCliente.rpt"
        .LogOnServer "pdssql.dll", _
         busca.LeerIni(App.Path & "\Camtex.ini", "Bgeneral", "dserver", ""), _
         busca.LeerIni(App.Path & "\Camtex.ini", "Bgeneral", "dbase", ""), _
         busca.LeerIni(App.Path & "\Camtex.ini", "Bgeneral", "duser", ""), _
         busca.LeerIni(App.Path & "\Camtex.ini", "Bgeneral", "dpass", "")
        .Connect = _
        "DSN=" & busca.LeerIni(App.Path & "\Camtex.ini", "Bgeneral", "dserver", "") & ";" & _
        "DSQ=" & busca.LeerIni(App.Path & "\Camtex.ini", "Bgeneral", "dbase", "") & ";" & _
        "UID=" & busca.LeerIni(App.Path & "\Camtex.ini", "Bgeneral", "duser", "") & ";" & _
        "PWD=" & busca.LeerIni(App.Path & "\Camtex.ini", "Bgeneral", "dpass", "")
        '''''''''''''''''''''''''''''''''''''''''''''''''''
        .Action = 1
        
  End With
  
Screen.MousePointer = 1

Exit Sub
Errores:
  Screen.MousePointer = 1
  MsgBox "Error inesperado: " & Err.Number & "  " & Err.Description, vbExclamation
  Err = 0
  Exit Sub
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 13 Then SendKeys "{TAB}"
End Sub


Private Sub Ctr_Cliente_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 13 Then SendKeys "{TAB}"
End Sub


Private Sub Ctr_Cta_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub DTHasta_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    MostrarForm Me, "C2"
    Call adll.llenacombo(Combo1, "select monedacodigo,monedadescripcion from gr_moneda", cn)
    DTHasta = Date
    Call Ctr_Cliente.conexion(cn)
    Call Ctr_Cta.conexion(cn)
End Sub

Private Sub Combo1_Click()
  If Combo1.ListCount > 0 Then
     txt(1) = adll.ComboDato(Combo1.Text)
  Else
     txt(1) = ""
  End If
End Sub
