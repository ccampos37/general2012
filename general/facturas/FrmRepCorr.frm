VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmRepCorr 
   Caption         =   "Numeración por Documentos"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   5490
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport oCrystalReport 
      Left            =   240
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
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
      Left            =   3240
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   1305
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
      Left            =   2325
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   2610
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
      Left            =   2640
      TabIndex        =   4
      Top             =   2280
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
      Left            =   1080
      TabIndex        =   3
      Top             =   2280
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTHasta 
      Height          =   405
      Left            =   2280
      TabIndex        =   2
      Top             =   1440
      Width           =   1575
      _ExtentX        =   2778
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
      Format          =   40960001
      CurrentDate     =   37518
   End
   Begin MSComCtl2.DTPicker DTDesde 
      Height          =   405
      Left            =   2280
      TabIndex        =   1
      Top             =   840
      Width           =   1575
      _ExtentX        =   2778
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
      Format          =   40960001
      CurrentDate     =   37518
   End
   Begin VB.Label lbl 
      Caption         =   "Punto de Venta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   360
      TabIndex        =   8
      Top             =   240
      Width           =   1815
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
      Left            =   1080
      TabIndex        =   6
      Top             =   960
      Width           =   855
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
      Left            =   1155
      TabIndex        =   5
      Top             =   1560
      Width           =   975
   End
End
Attribute VB_Name = "FrmRepCorr"
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

Private Sub cmdAceptar_Click(Index As Integer)
On Error GoTo Errores

 If DTDesde > DtHasta Then
     MsgBox "Fecha Desde debe ser mayor a Fecha Hasta", vbInformation, "AVISO"
     Exit Sub
 End If
 
 Screen.MousePointer = 11
 
 With oCrystalReport
        .Reset
        .ReportFileName = VGParamSistem.Rutareport & "RepvtCorrelativos.rpt"
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
       If VGsql = 1 Then
           .Connect = "Provider=SQLOLEDB;Password=" & VGParamSistem.PwdGEN & ";Persist Security Info=True;User ID=" & VGParamSistem.UsuarioGEN & ";Initial Catalog=" & VGParamSistem.BDEmpresaGEN & ";SERVER=" & VGParamSistem.ServidorGEN
          Else
           .Connect = VGcadenareport2
           '.LogOnServer "pdssql.dll", "", VGParamSistem.BDEmpresaGEN, VGParamSistem.UsuarioGEN, VGParamSistem.PwdGEN
        End If

'        .Connect = _
'        "DSN=" & busca.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "SERVIDOR", "") & ";" & _
'        "DSQ=" & busca.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "BDDATOS", "") & ";" & _
'        "UID=" & busca.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "USUARIO", "") & ";" & _
'        "PWD=" & busca.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "PASSW", "")
        '''''''''''''''''''''''''''''''''''''''''''''''''''
        .DiscardSavedData = True
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintSetupBtn = True
        .WindowShowExportBtn = True
        .WindowShowZoomCtl = True
        .WindowShowNavigationCtls = True
        .WindowShowPrintBtn = True
        .WindowTitle = "Numeración por Documentos"
        .formulas(0) = "Empresa='" & VGParametros.nomempresa & "'"
        .formulas(1) = "Desde='" & DTDesde & "'"
        .formulas(2) = "Hasta='" & DtHasta & "'"
        If Combo1.ListIndex <> -1 Then
            .formulas(3) = "PuntoVenta='" & Combo1.Text & "'"
        Else
            .formulas(3) = "PuntoVenta='TODOS'"
        End If
        .StoredProcParam(0) = busca.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "BDDATOS", "")
        .StoredProcParam(1) = IIf(Len(Trim(txt(0))) = 0, "%", Trim(txt(0)))
        .StoredProcParam(2) = DTDesde
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

Private Sub cmdCancelar_Click(Index As Integer)
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
    MostrarFormVentas Me, "C2"
    Call adll.llenacombo(Combo1, "select puntovtacodigo,puntovtadescripcion from vt_puntoventa", VGCNx)
    DTDesde = Date
    DtHasta = Date
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub


