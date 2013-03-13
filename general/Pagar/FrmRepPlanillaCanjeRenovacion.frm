VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmRepPlanillaCanjeRenovacion 
   Caption         =   "Planilla de Canje"
   ClientHeight    =   2760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2760
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   390
      Left            =   1845
      TabIndex        =   9
      Top             =   2205
      Width           =   1170
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   390
      Left            =   3330
      TabIndex        =   8
      Top             =   2205
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      Height          =   1890
      Left            =   90
      TabIndex        =   4
      Top             =   60
      Width           =   6165
      Begin MSComCtl2.DTPicker DTP_FechaFin 
         Height          =   315
         Left            =   1830
         TabIndex        =   1
         Top             =   675
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   27262977
         CurrentDate     =   37518
      End
      Begin MSComCtl2.DTPicker DTP_FechaInicio 
         Height          =   315
         Left            =   1830
         TabIndex        =   0
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   27262977
         CurrentDate     =   37518
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Oficina 
         Height          =   375
         Left            =   1830
         TabIndex        =   2
         Top             =   1065
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   661
         XcodMaxLongitud =   3
         xcodwith        =   500
         NomTabla        =   "cp_oficina"
         TituloAyuda     =   "Ayuda de Oficinas"
         ListaCampos     =   "vendedorcodigo(1),vendedornombres(1),vendedorruc(1),vendedordireccion(1),vendedortelefono(1)"
         XcodCampo       =   "vendedorcodigo"
         XListCampo      =   "vendedornombres"
         ListaCamposDescrip=   "Codigo,Nombres,Ruc,Direccion,Telefono"
         ListaCamposText =   "vendedorcodigo,vendedornombres,vendedorruc,vendedordireccion,vendedortelefono"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Proveedor 
         Height          =   375
         Left            =   1815
         TabIndex        =   3
         Top             =   1425
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   661
         XcodMaxLongitud =   0
         xcodwith        =   900
         NomTabla        =   "cp_proveedor"
         ListaCampos     =   "clientecodigo(1),clienterazonsocial(1)"
         XcodCampo       =   "clientecodigo"
         XListCampo      =   "clienterazonsocial"
         ListaCamposDescrip=   "Código,Razón_Social"
         ListaCamposText =   "clientecodigo,clienterazonsocial"
         Requerido       =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "Proveedor"
         Height          =   330
         Left            =   705
         TabIndex        =   10
         Top             =   1455
         Width           =   1080
      End
      Begin VB.Label lbl 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   795
         TabIndex        =   7
         Top             =   285
         Width           =   855
      End
      Begin VB.Label lbl 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   750
         TabIndex        =   6
         Top             =   675
         Width           =   975
      End
      Begin VB.Label lbl 
         Caption         =   "Oficina"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   720
         TabIndex        =   5
         Top             =   1095
         Width           =   945
      End
   End
   Begin Crystal.CrystalReport oCrystalReport 
      Left            =   105
      Top             =   2130
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "FrmRepPlanillaCanjeRenovacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_Opcion As String
Dim adll As New dllgeneral.dll_general
Dim cTitulo As String

Private Sub Form_Load()
   MostrarForm Me, "C2"
   DTP_FechaInicio.Value = "01/" & Format(Month(Now), "00") & "/" & Year(Date)
   DTP_FechaFin.Value = Format(Date, "dd/mm/yyyy")
   Select Case m_Opcion
      Case "1": cTitulo = "CANJE"
      Case "2": cTitulo = "RENOVACION"
   End Select
   Me.Caption = "Planilla de " & cTitulo
   Ctr_Oficina.conexion VGCNx
   Ctr_Proveedor.conexion VGCNx
End Sub

Private Sub cmdAceptar_Click()
   Call Imprimir
End Sub

Private Sub cmdCancelar_Click()
   Unload Me
End Sub

Sub Imprimir()
Dim arrform(5) As Variant, arrparm(6) As Variant
Dim NombreRep As String, CadOrden As String
Dim NombreSubRep As String
Dim mon As String
   arrparm(0) = VGCNx.DefaultDatabase
   arrparm(1) = Format(DTP_FechaInicio.Value, "dd/mm/yyyy")
   arrparm(2) = Format(DTP_FechaFin.Value, "dd/mm/yyyy")
   arrparm(3) = IIf(Ctr_Oficina.xclave = Empty, "%", Trim$(Ctr_Oficina.xclave))
   arrparm(4) = IIf(Ctr_Proveedor.xclave = Empty, "%", Trim$(Ctr_Proveedor.xclave))
   arrparm(5) = m_Opcion
   arrform(0) = "@Titulo='" & cTitulo & "'"
   arrform(1) = "Empresa='" & g_DetalleEmpresa & "'"
   arrform(2) = "Desde='" & Format(DTP_FechaInicio.Value, "dd/mm/yyyy") & "'"
   arrform(3) = "Hasta='" & Format(DTP_FechaFin.Value, "dd/mm/yyyy") & "'"
   arrform(4) = "Oficina='" & IIf(Ctr_Oficina.xclave = Empty, "Todos", Trim$(Ctr_Oficina.xclave)) & "'"
   NombreRep = "RepcpPlanCanjeRenovacion.rpt"
   NombreSubRep = "RepcpSubPlanCanjeRenovacion.rpt"
   CadOrden = ""
   Call ImpresionRptProc(NombreRep, arrform, arrparm, NombreSubRep, CadOrden, "Planilla de " & cTitulo)
End Sub

Sub ImpresionRptProc(cNombreReporte As String, PFormulas(), Param(), Optional cNombreSubReporte As String, Optional orden As String, Optional Titulo As String)
Dim strBuscar As New dll_apis
Dim i As Integer
On Error GoTo X
    Screen.MousePointer = 11
    With MDIPrincipal.CryRptProc
        .Reset
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowTitle = Titulo
        .ReportFileName = VGparamsistem.RutaReport & cNombreReporte
        .LogOnServer "pdssql.dll", _
         strBuscar.LeerIni(App.Path & "\Marfice.ini", "BDGeneral", "dserver", ""), _
         strBuscar.LeerIni(App.Path & "\Marfice.ini", "BDGeneral", "dbase", ""), _
         strBuscar.LeerIni(App.Path & "\Marfice.ini", "BDGeneral", "duser", ""), _
         strBuscar.LeerIni(App.Path & "\Marfice.ini", "BDGeneral", "dpass", "")
        .Connect = _
        "DSN=" & strBuscar.LeerIni(App.Path & "\Marfice.ini", "BDGeneral", "dserver", "") & ";" & _
        "DSQ=" & strBuscar.LeerIni(App.Path & "\Marfice.ini", "BDGeneral", "dbase", "") & ";" & _
        "UID=" & strBuscar.LeerIni(App.Path & "\Marfice.ini", "BDGeneral", "duser", "") & ";" & _
        "PWD=" & strBuscar.LeerIni(App.Path & "\Marfice.ini", "BDGeneral", "dpass", "")
        Call PropCrystal(MDIPrincipal.CryRptProc)
        
        If UBound(PFormulas) > 0 Then
            For i = 0 To UBound(PFormulas) - 1
                .Formulas(2 + i) = PFormulas(i)
            Next
        End If
        .DiscardSavedData = True
        If UBound(Param) > 0 Then
            For i = 0 To UBound(Param) - 1
                .StoredProcParam(i) = Param(i)
            Next
        End If
        
        If cNombreSubReporte <> Empty Then
          .SubreportToChange = cNombreSubReporte
          .LogOnServer "pdssql.dll", _
           strBuscar.LeerIni(App.Path & "\Marfice.ini", "BDGeneral", "dserver", ""), _
           strBuscar.LeerIni(App.Path & "\Marfice.ini", "BDGeneral", "dbase", ""), _
           strBuscar.LeerIni(App.Path & "\Marfice.ini", "BDGeneral", "duser", ""), _
           strBuscar.LeerIni(App.Path & "\Marfice.ini", "BDGeneral", "dpass", "")
          .Connect = _
          "DSN=" & strBuscar.LeerIni(App.Path & "\Marfice.ini", "BDGeneral", "dserver", "") & ";" & _
          "DSQ=" & strBuscar.LeerIni(App.Path & "\Marfice.ini", "BDGeneral", "dbase", "") & ";" & _
          "UID=" & strBuscar.LeerIni(App.Path & "\Marfice.ini", "BDGeneral", "duser", "") & ";" & _
          "PWD=" & strBuscar.LeerIni(App.Path & "\Marfice.ini", "BDGeneral", "dpass", "")
          If UBound(Param) > 0 Then
              For i = 0 To UBound(Param) - 1
                  .StoredProcParam(i) = Param(i)
              Next
          End If
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

Property Let Opcion(valor As String)
  m_Opcion = valor
End Property
