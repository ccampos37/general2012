VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmRepRankClientesNegocios 
   Caption         =   "Clientes x Tipo de Negocios Mensualizado"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6690
   LinkTopic       =   "Form2"
   ScaleHeight     =   5010
   ScaleWidth      =   6690
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option4 
      Caption         =   "Resumido"
      Height          =   495
      Left            =   4080
      TabIndex        =   14
      Top             =   2880
      Width           =   1215
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Por Producto"
      Height          =   495
      Left            =   2160
      TabIndex        =   13
      Top             =   2880
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Por linea de Articulo"
      Height          =   495
      Left            =   3960
      TabIndex        =   12
      Top             =   2280
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Por Cliente"
      Height          =   495
      Left            =   2160
      TabIndex        =   11
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   5535
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
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   2610
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
         Index           =   1
         Left            =   3555
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   360
         Visible         =   0   'False
         Width           =   1305
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
         Left            =   3360
         TabIndex        =   2
         Top             =   3120
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
         Left            =   1680
         TabIndex        =   1
         Top             =   3120
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTHasta 
         Height          =   375
         Left            =   2640
         TabIndex        =   5
         Top             =   1320
         Width           =   1845
         _ExtentX        =   3254
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
         Format          =   62521345
         CurrentDate     =   37518
      End
      Begin MSComCtl2.DTPicker DTDesde 
         Height          =   405
         Left            =   2640
         TabIndex        =   6
         Top             =   720
         Width           =   1845
         _ExtentX        =   3254
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
         Format          =   62521345
         CurrentDate     =   37518
      End
      Begin Crystal.CrystalReport oCrystalReport 
         Left            =   480
         Top             =   1800
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
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
         Height          =   255
         Index           =   5
         Left            =   720
         TabIndex        =   10
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Top             =   2520
         Width           =   2475
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
         Left            =   1560
         TabIndex        =   8
         Top             =   720
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
         Left            =   1560
         TabIndex        =   7
         Top             =   1320
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmRepRankClientesNegocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TotalNeto As Double
Dim TotalBruto As Double
Dim d_porcentaje As Double
Dim d_monto As Double
Dim index_combo As Integer
''''''''''''''''''''''''''''''''''''''''''''''''''''
'Agregar:
Dim busca As New dll_apisgen.dll_apis
Dim adll As New dllgeneral.dll_general


Private Sub cmdAceptar_Click(Index As Integer)
On Error GoTo Errores
 If DTDesde > DTHasta Then
     MsgBox "Fecha Desde debe ser mayor a Fecha Hasta", vbInformation, "AVISO"
     Exit Sub
 End If
 
 Screen.MousePointer = 11
 
 Call Consulta_Reporte
  
 With oCrystalReport
        .Reset
        If Option1 = True Then
          .ReportFileName = VGParamSistem.Rutareport & "RepvtRankingNegocioClientexMes.rpt"
         ElseIf Option2 = True Then
             .ReportFileName = VGParamSistem.Rutareport & "RepvtRankingNegocioLineaxMes.rpt"
           ElseIf Option3 = True Then
                 .ReportFileName = VGParamSistem.Rutareport & "RepvtRankingNegocioProductoxMes.rpt"
              Else
                 .ReportFileName = VGParamSistem.Rutareport & "RepvtRankingnegocioxMes.rpt"
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
   '     .LogOnServer "pdssql.dll", _
   '      busca.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "SERVIDOR", ""), _
   '      busca.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "BDDATOS", ""), _
   '      busca.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "USUARIO", ""), _
   '      busca.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "PASSW", "")
        .Connect = _
        "DSN=" & busca.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "SERVIDOR", "") & ";" & _
        "DSQ=" & busca.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "BDDATOS", "") & ";" & _
        "UID=" & busca.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "USUARIO", "") & ";" & _
        "PWD=" & busca.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "PASSW", "")
        '''''''''''''''''''''''''''''''''''''''''''''''''''
        .DiscardSavedData = True
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintSetupBtn = True
        .WindowShowExportBtn = True
        .WindowShowZoomCtl = True
        .WindowShowNavigationCtls = True
        .WindowShowPrintBtn = True
        .WindowTitle = "Ranking de Clientes"
        .Formulas(0) = "Empresa='" & VGparametros.nomempresa & "'"
        .Formulas(3) = "Desde='" & DTDesde & "'"
        .Formulas(4) = "Hasta='" & DTHasta & "'"
        .StoredProcParam(0) = VGcnx.DefaultDatabase
        .StoredProcParam(1) = IIf(Trim(txt(1)) = "", "%", Trim(txt(1)))
        If Combo1.ListIndex <> -1 Then
            .Formulas(5) = "PuntoVta='" & Combo1.Text & "'"
         Else
            .Formulas(5) = "PuntoVta='TODOS'"
        End If
        .StoredProcParam(2) = DTDesde
        .StoredProcParam(3) = DTHasta
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

Private Sub Form_Load()
    MostrarForm Me, "C2"
    Call adll.llenacombo(Combo1, "select puntovtacodigo,puntovtadescripcion from vt_puntoventa", VGcnx)
    DTDesde = Date
    DTHasta = Date
End Sub

Private Sub Combo1_Click()
  If Combo1.ListCount > 0 Then
     txt(1) = adll.ComboDato(Combo1.Text)
  Else
     txt(1) = ""
  End If
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub



Private Sub txt_LostFocus(Index As Integer)
 txt(Index).Text = Format(txt(Index).Text, "###,##0.00")
End Sub

Private Function Consulta_Reporte()

Dim SQL_TOTAL_NETO As String
Dim SQL_TOTAL_BRUTO_Sol As String
Dim SQL_TOTAL_BRUTO_Dol As String
Dim rs As New ADODB.Recordset
Dim codpuntoventa As String

 If Trim(txt(1)) = "" Then
    codpuntoventa = "%"
 Else
    codpuntoventa = Trim(txt(1))
 End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Function


