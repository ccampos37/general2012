VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "CONTROLAYUDA.OCX"
Begin VB.Form RptDocumentosxPagar_2 
   Caption         =   "Documentos Pendientes"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   6060
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Height          =   3165
      Left            =   45
      TabIndex        =   6
      Top             =   15
      Width           =   5940
      Begin VB.Frame Frame2 
         Height          =   1335
         Left            =   150
         TabIndex        =   12
         Top             =   3585
         Visible         =   0   'False
         Width           =   5145
         Begin VB.OptionButton Option1 
            Caption         =   "Relacion x Vendedor"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   15
            Top             =   570
            Width           =   1935
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Todos Movimientos"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   14
            Top             =   270
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Relacion x Banco"
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   13
            Top             =   870
            Width           =   1935
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda1 
            Height          =   315
            Left            =   2160
            TabIndex        =   16
            Top             =   540
            Width           =   2925
            _ExtentX        =   5159
            _ExtentY        =   556
            Enabled         =   0   'False
            XcodMaxLongitud =   0
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda2 
            Height          =   315
            Left            =   2160
            TabIndex        =   17
            Top             =   900
            Width           =   2925
            _ExtentX        =   5159
            _ExtentY        =   556
            Enabled         =   0   'False
            XcodMaxLongitud =   0
         End
      End
      Begin VB.Frame Frame1 
         Height          =   885
         Left            =   135
         TabIndex        =   10
         Top             =   195
         Width           =   2910
         Begin MSComCtl2.DTPicker DTP_FechaRef 
            Height          =   345
            Left            =   1350
            TabIndex        =   0
            Top             =   420
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   609
            _Version        =   393216
            Format          =   24510465
            CurrentDate     =   37588
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha :"
            Height          =   255
            Index           =   1
            Left            =   150
            TabIndex        =   11
            Top             =   480
            Width           =   990
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1650
         Left            =   135
         TabIndex        =   7
         Top             =   1200
         Width           =   5700
         Begin VB.ComboBox cboMoneda 
            Height          =   315
            Left            =   1140
            TabIndex        =   3
            Text            =   "cboMoneda"
            Top             =   1110
            Width           =   2325
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Cliente 
            Height          =   315
            Left            =   1170
            TabIndex        =   1
            Top             =   315
            Width           =   4425
            _ExtentX        =   7805
            _ExtentY        =   556
            XcodMaxLongitud =   0
            xcodwith        =   900
            NomTabla        =   "cp_proveedor"
            ListaCampos     =   "clientecodigo(1),clienterazonsocial(1)"
            XcodCampo       =   "clientecodigo"
            XListCampo      =   "clienterazonsocial"
            ListaCamposDescrip=   "C�digo,Raz�n_Social"
            ListaCamposText =   "clientecodigo,clienterazonsocial"
            Requerido       =   0   'False
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Doc 
            Height          =   300
            Left            =   1155
            TabIndex        =   2
            Top             =   720
            Width           =   4440
            _ExtentX        =   7832
            _ExtentY        =   529
            XcodMaxLongitud =   0
            xcodwith        =   900
            NomTabla        =   "cp_tipodocumento"
            ListaCampos     =   "tdocumentocodigo(1),tdocumentodescripcion(1)"
            XcodCampo       =   "tdocumentocodigo"
            XListCampo      =   "tdocumentodescripcion"
            ListaCamposDescrip=   "C�digo,Descripci�n"
            ListaCamposText =   "tdocumentocodigo,tdocumentodescripcion"
            Requerido       =   0   'False
         End
         Begin VB.Label Label4 
            Caption         =   "Moneda"
            Height          =   225
            Left            =   135
            TabIndex        =   18
            Top             =   1155
            Width           =   645
         End
         Begin VB.Label Label6 
            Caption         =   "Documento"
            Height          =   255
            Left            =   150
            TabIndex        =   9
            Top             =   780
            Width           =   945
         End
         Begin VB.Label Label3 
            Caption         =   "Cliente"
            Height          =   225
            Left            =   135
            TabIndex        =   8
            Top             =   360
            Width           =   525
         End
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3188
      TabIndex        =   5
      Top             =   3495
      Width           =   1260
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1613
      TabIndex        =   4
      Top             =   3495
      Width           =   1260
   End
End
Attribute VB_Name = "RptDocumentosxPagar_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adll As New dllgeneral.dll_general
Dim aRango(5) As Integer

Private Sub Form_Load()
   MostrarForm Me, "C2"
   Ctr_Cliente.conexion cn
   Ctr_Doc.conexion cn
   cboMoneda.Clear
   cboMoneda.AddItem g_TipoSol & "-Soles"
   cboMoneda.AddItem g_TipoDolar & "-Dolares"
   cboMoneda.ListIndex = 1
   DTP_FechaRef.Value = Format(Date, "dd/mm/yyyy")
End Sub

Private Sub cmdAceptar_Click()
  Call Imprimir
End Sub

Private Sub cmdCancelar_Click()
 Unload Me
End Sub
Sub Imprimir()
Dim arrform(2) As Variant, arrparm(8) As Variant
Dim NombreRep As String, CadOrden As String
Dim NombrePC As String
Dim ValorRango As String
Dim I As Integer
    ValorRango = "7*15*30*45*60*"
    Randomize   ' Inicializa el generador de n�meros aleatorios.
    NombrePC = Trim(Str(CLng(Rnd * 10000000)))
    arrparm(0) = cn.DefaultDatabase
    arrparm(1) = NombrePC
    arrparm(2) = Format(DTP_FechaRef.Value, "dd/mm/yyyy")
    arrparm(3) = ValorRango
    arrparm(4) = IIf(Ctr_Cliente.xclave = Empty, "%", Trim(Ctr_Cliente.xclave))
    arrparm(5) = Format(cboMoneda.ListIndex + 1, "00")
    arrparm(6) = cnconta.DefaultDatabase
    arrparm(7) = IIf(Ctr_Doc.xclave = Empty, "%", Trim(Ctr_Doc.xclave))
    CadOrden = ""
    arrform(0) = "@Fecha='" & Format(DTP_FechaRef.Value, "dd/mm/yyyy") & "'"
    arrform(1) = "@Moneda='" & Mid(cboMoneda.Text, InStr(1, cboMoneda.Text, "-", vbTextCompare) + 1, Len(cboMoneda.Text) - InStr(1, cboMoneda.Text, "-", vbTextCompare)) & "'"
    
    NombreRep = "RepcpDocumentosPagar.rpt"
    Call ImpresionRptProc(VG_String, arrform, arrparm, CadOrden, "Documentos Pendientes")
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
