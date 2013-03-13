VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "CONTROLAYUDA.OCX"
Begin VB.Form RptctactexCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuenta Corriente por Cliente"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6420
   Icon            =   "RptctactexCliente.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   6420
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3105
      TabIndex        =   6
      Top             =   3165
      Width           =   1380
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1485
      TabIndex        =   5
      Top             =   3165
      Width           =   1440
   End
   Begin VB.Frame Frame1 
      Height          =   2730
      Left            =   105
      TabIndex        =   3
      Top             =   120
      Width           =   6180
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         Left            =   1665
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1785
         Width           =   1605
      End
      Begin VB.ComboBox cboResumen 
         Height          =   315
         Left            =   1665
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1395
         Visible         =   0   'False
         Width           =   1605
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Cliente 
         Height          =   300
         Left            =   1665
         TabIndex        =   10
         Top             =   255
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   529
         XcodMaxLongitud =   0
         xcodwith        =   900
         NomTabla        =   "vt_cliente"
         ListaCampos     =   "clientecodigo(1),clienterazonsocial(1)"
         XcodCampo       =   "clientecodigo"
         XListCampo      =   "clienterazonsocial"
         ListaCamposDescrip=   "Código,Razón_Social"
         ListaCamposText =   "clientecodigo,clienterazonsocial"
         Requerido       =   0   'False
      End
      Begin MSComCtl2.DTPicker DTP_FechaInicio 
         Height          =   330
         Left            =   1665
         TabIndex        =   11
         Top             =   645
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Format          =   24641537
         CurrentDate     =   37586
      End
      Begin MSComCtl2.DTPicker DTP_FechaFin 
         Height          =   330
         Left            =   1665
         TabIndex        =   12
         Top             =   1020
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Format          =   24641537
         CurrentDate     =   37586
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Zona 
         Height          =   300
         Left            =   1665
         TabIndex        =   14
         Top             =   2115
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   529
         XcodMaxLongitud =   0
         xcodwith        =   900
         NomTabla        =   "vt_zona"
         ListaCampos     =   "zonacodigo(1),zonadescripcion(1)"
         XcodCampo       =   "zonacodigo"
         XListCampo      =   "zonadescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "zonacodigo,zonadescripcion"
         Requerido       =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "Zona"
         Height          =   255
         Index           =   3
         Left            =   390
         TabIndex        =   13
         Top             =   2160
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Desde la Fecha"
         Height          =   255
         Index           =   0
         Left            =   390
         TabIndex        =   7
         Top             =   705
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         Height          =   255
         Index           =   1
         Left            =   390
         TabIndex        =   0
         Top             =   330
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta la Fecha"
         Height          =   255
         Index           =   2
         Left            =   390
         TabIndex        =   1
         Top             =   1080
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Con Resumen"
         Height          =   255
         Index           =   4
         Left            =   390
         TabIndex        =   2
         Top             =   1455
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Moneda"
         Height          =   255
         Index           =   5
         Left            =   390
         TabIndex        =   4
         Top             =   1845
         Width           =   1185
      End
   End
End
Attribute VB_Name = "RptctactexCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adll As New dllgeneral.dll_general

Private Sub Form_Load()
   MostrarForm Me, "C2"
   Call CargarTipo(cboResumen, 3)
   cboMoneda.Clear
   cboMoneda.AddItem g_TipoSol & "-Soles"
   cboMoneda.AddItem g_TipoDolar & "-Dolares"
   cboMoneda.AddItem "03-Ambos"
   cboMoneda.ListIndex = 2
   DTP_FechaInicio.Value = "01/" & Format(Month(Now), "00") & "/" & Year(Date)
   DTP_FechaFin.Value = Format(Date, "dd/mm/yyyy")
   Ctr_Cliente.conexion cn
   Ctr_Zona.conexion cn
End Sub

Private Sub cmdAceptar_Click()
   Call Imprimir
End Sub

Private Sub cmdCancelar_Click()
   Unload Me
End Sub

Sub Imprimir()
Dim arrform(1) As Variant, arrparm(8) As Variant
Dim NombreRep As String, CadOrden As String
Dim NombrePC As String
Dim mon As String
    Randomize   ' Inicializa el generador de números aleatorios.
    NombrePC = Trim(Str(CLng(Rnd * 10000000)))
    arrparm(0) = cn.DefaultDatabase
    arrparm(1) = NombrePC
    arrparm(2) = Format(DTP_FechaInicio.Value, "dd/mm/yyyy")
    arrparm(3) = Format(DTP_FechaFin.Value, "dd/mm/yyyy")
    arrparm(4) = Format(DTP_FechaInicio.Value - 1, "dd/mm/yyyy")
    If cboMoneda.ListIndex = 2 Then
      arrparm(5) = "%"
    Else
      arrparm(5) = Format(cboMoneda.ListIndex + 1, "00")
    End If
    arrparm(6) = IIf(Ctr_Cliente.xclave = Empty, "%", Trim(Ctr_Cliente.xclave))
    arrparm(7) = IIf(Ctr_Zona.xclave = Empty, "%", Trim(Ctr_Zona.xclave))
    arrform(0) = "RangoFecha='" & "DEL " & Format(DTP_FechaInicio.Value, "dd/mm/yyyy") & " AL " & Format(DTP_FechaFin.Value, "dd/mm/yyyy") & "'"
    NombreRep = "RepccCtaCtexCliente.rpt"
    CadOrden = ""
    Call ImpresionRptProc(NombreRep, arrform, arrparm, CadOrden, "Cuenta Corriente por Cliente")
End Sub

Sub ImpresionRptProc(cNombreReporte As String, PFormulas(), Param(), Optional orden As String, Optional Titulo As String)
Dim strBuscar As New dll_apis
Dim I As Integer
On Error GoTo X
    Screen.MousePointer = 11
    With FrmMenu.oCrystalReport
        .Reset
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowTitle = Titulo
        .ReportFileName = RutaRepProc & cNombreReporte
        .LogOnServer "pdssql.dll", _
         strBuscar.LeerIni(App.Path & "\Camtex.ini", "Bgeneral", "dserver", ""), _
         strBuscar.LeerIni(App.Path & "\Camtex.ini", "Bgeneral", "dbase", ""), _
         strBuscar.LeerIni(App.Path & "\Camtex.ini", "Bgeneral", "duser", ""), _
         strBuscar.LeerIni(App.Path & "\Camtex.ini", "Bgeneral", "dpass", "")
        .Connect = _
        "DSN=" & strBuscar.LeerIni(App.Path & "\Camtex.ini", "Bgeneral", "dserver", "") & ";" & _
        "DSQ=" & strBuscar.LeerIni(App.Path & "\Camtex.ini", "Bgeneral", "dbase", "") & ";" & _
        "UID=" & strBuscar.LeerIni(App.Path & "\Camtex.ini", "Bgeneral", "duser", "") & ";" & _
        "PWD=" & strBuscar.LeerIni(App.Path & "\Camtex.ini", "Bgeneral", "dpass", "")
        Call PropCrystal(FrmMenu.oCrystalReport)
        .Formulas(0) = "@Empresa='" & g_DetalleEmpresa & "'"
        .Formulas(1) = "@Ruc='" & "20293847038" & "'"
        If UBound(PFormulas) > 0 Then
            For I = 0 To UBound(PFormulas) - 1
                .Formulas(2 + I) = PFormulas(I)
            Next
        End If
        .DiscardSavedData = True
        If UBound(Param) > 0 Then
            For I = 0 To UBound(Param) - 1
                .StoredProcParam(I) = Param(I)
            Next
        End If
        If .Status <> 2 Then .Action = 1
    End With
    Screen.MousePointer = 1
    Exit Sub
X:
  If Err.Number = 9 Then Resume Next
  Screen.MousePointer = 1
  MsgBox "Error inesperado: " & Err.Number & "  " & Err.Description, vbExclamation
End Sub
