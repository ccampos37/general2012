VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmRepVtasxFact 
   Caption         =   "Ventas por Factura"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   ScaleHeight     =   2580
   ScaleWidth      =   5670
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport oCrystalReport 
      Left            =   360
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker DTHasta 
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   1200
      Width           =   1515
      _ExtentX        =   2672
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
      Format          =   74907649
      CurrentDate     =   37518
   End
   Begin MSComCtl2.DTPicker DTDesde 
      Height          =   405
      Left            =   2280
      TabIndex        =   1
      Top             =   720
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   714
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
      Format          =   74907649
      CurrentDate     =   37518
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
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   2865
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
      Left            =   2880
      TabIndex        =   4
      Top             =   1800
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
      Left            =   1320
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
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
      Height          =   300
      Index           =   0
      Left            =   3000
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label lbl 
      Caption         =   "Hasta"
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
      Left            =   1200
      TabIndex        =   8
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lbl 
      Caption         =   "Desde"
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
      Left            =   1200
      TabIndex        =   7
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "Almacen"
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
      Index           =   0
      Left            =   960
      TabIndex        =   6
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "FrmRepVtasxFact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adll As New dllgeneral.dll_general
''''''''''''''''''''''''''''''''''''''''''''''''''''
'Agregar:
Dim busca As New dll_apisgen.dll_apis
''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub cmdAceptar_Click()
On Error GoTo Errores
   If DtDesde > DtHasta Then
       MsgBox "Fecha Desde debe ser mayor a Fecha Hasta", vbInformation, "AVISO"
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   With oCrystalReport
        .Reset
        .ReportFileName = VGParamSistem.Rutareport & "RepvtVtasxFactura.rpt"
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        If VGsql = 1 Then
           .Connect = "Provider=SQLOLEDB;Password=" & VGParamSistem.PwdGEN & ";Persist Security Info=True;User ID=" & VGParamSistem.UsuarioGEN & ";Initial Catalog=" & VGParamSistem.BDEmpresaGEN & ";SERVER=" & VGParamSistem.ServidorGEN
          Else
           .Connect = VGcadenareport2
        End If

        .DiscardSavedData = True
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintSetupBtn = True
        .WindowShowExportBtn = True
        .WindowShowZoomCtl = True
        .WindowShowNavigationCtls = True
        .WindowShowPrintBtn = True
        .WindowTitle = "Ventas por Factura"
        .formulas(0) = "Empresa='" & VGParametros.nomempresa & "'"
        .formulas(1) = "Desde='" & DtDesde & "'"
        .formulas(2) = "Hasta='" & DtHasta & "'"
        If Combo1.ListIndex <> -1 Then
            .formulas(3) = "Almacen='" & Combo1.Text & "'"
        Else
            .formulas(3) = "Almacen='TODOS'"
        End If
        .StoredProcParam(0) = busca.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "BDDATOS", "")
        .StoredProcParam(1) = IIf(Len(Trim(txt(0))) = 0, "%", Trim(txt(0)))
        .StoredProcParam(2) = DtDesde
        .StoredProcParam(3) = DtHasta

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

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Combo1_Click()
   If Combo1.ListCount > 0 Then
      txt(0) = adll.ComboDato(Combo1.Text)
   Else
      txt(0) = ""
   End If
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub DtDesde_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub DtHasta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    MostrarForm Me, "C2"
    Call adll.llenacombo(Combo1, "select taalma,tadescri from tabalm where empresacodigo='" & VGParametros.empresacodigo & "' and puntovtacodigo='" & VGParametros.puntovta & "'", VGCNx)
    DtDesde = Date
    DtHasta = Date
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
     If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
