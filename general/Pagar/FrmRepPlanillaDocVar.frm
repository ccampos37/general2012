VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmRepPlanillaDocVar 
   Caption         =   "Planilla Documentos Varios"
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2910
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   1717
      TabIndex        =   3
      Top             =   2220
      Width           =   1170
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   3247
      TabIndex        =   4
      Top             =   2220
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccionar Criterios"
      Height          =   1800
      Left            =   15
      TabIndex        =   5
      Top             =   75
      Width           =   6105
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Proveedor 
         Height          =   345
         Left            =   1305
         TabIndex        =   0
         Top             =   360
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   609
         XcodMaxLongitud =   0
         xcodwith        =   700
         NomTabla        =   "cp_proveedor"
         ListaCampos     =   "clientecodigo(1),clienterazonsocial(1)"
         XcodCampo       =   "clientecodigo"
         XListCampo      =   "clienterazonsocial"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "clientecodigo,clienterazonsocial"
         Requerido       =   0   'False
      End
      Begin MSComCtl2.DTPicker DTHasta 
         Height          =   330
         Left            =   1305
         TabIndex        =   2
         Top             =   1260
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
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
         Format          =   98762753
         CurrentDate     =   37518
      End
      Begin MSComCtl2.DTPicker DTDesde 
         Height          =   330
         Left            =   1305
         TabIndex        =   1
         Top             =   795
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
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
         Format          =   98762753
         CurrentDate     =   37518
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
         Left            =   465
         TabIndex        =   8
         Top             =   1320
         Width           =   825
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
         Left            =   465
         TabIndex        =   7
         Top             =   870
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Proveedor"
         Height          =   330
         Left            =   465
         TabIndex        =   6
         Top             =   405
         Width           =   945
      End
   End
   Begin Crystal.CrystalReport oCrystalReport 
      Left            =   90
      Top             =   1980
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "FrmRepPlanillaDocVar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim busca As New dll_apisgen.dll_apis

Private Sub Form_Load()
   MostrarForm Me, "C2"
   DTDesde = Date
   DTHasta = Date
   Ctr_Proveedor.conexion VGCNx
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
   Unload Me
End Sub

Private Sub cmdAceptar_Click(Index As Integer)
   Call Imprimir
End Sub

Sub Imprimir()
Dim arrform(2) As Variant, arrparm(5) As Variant
Dim NombreRep As String, CadOrden As String
Dim NombreSubRep As String
Dim ValorRango As String
Dim tipo As Integer
tipo = 1
Dim i As Integer
   arrform(0) = "Desde='" & Format(DTDesde.Value, "dd/mm/yyyy") & "'"
   arrform(1) = "Hasta='" & Format(DTHasta.Value, "dd/mm/yyyy") & "'"
   
   arrparm(0) = VGCNx.DefaultDatabase
   arrparm(1) = Format(DTDesde.Value, "dd/mm/yyyy")
   arrparm(2) = Format(DTHasta.Value, "dd/mm/yyyy")
   arrparm(3) = IIf(Ctr_Proveedor.xclave = Empty, "%", Trim$(Ctr_Proveedor.xclave))
   arrparm(4) = tipo
   
   NombreRep = "cp_PlanDocVarios.rpt"
   NombreSubRep = "cp_SubPlanDocVarios.rpt"
   CadOrden = ""
   Call ImpresionRpt_SubRpt_Proc(NombreRep, arrform, arrparm, NombreSubRep, CadOrden, "Planilla Documentos Varios")
End Sub

Sub ImpresionRptProc(cNombreReporte As String, PFormulas(), Param(), cNombreSubRpt As String, Optional orden As String, Optional Titulo As String)
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
                .Formulas(i) = PFormulas(i)
            Next
        End If
        .DiscardSavedData = True
        If UBound(Param) > 0 Then
            For i = 0 To UBound(Param) - 1
                .StoredProcParam(i) = Param(i)
            Next
        End If
        '***Para el SubReporte
        .SubreportToChange = cNombreSubRpt
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
        
        If .Status <> 2 Then .Action = 1
    End With
    Screen.MousePointer = 1
    Exit Sub
X:
  If Err.Number = 9 Then Resume Next
  Screen.MousePointer = 1
  MsgBox "Error inesperado: " & Err.Number & "  " & Err.Description, vbExclamation
End Sub

Private Sub DTDesde_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub DTHasta_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
