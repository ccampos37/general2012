VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRRHH001 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Citaciones a trabajadores"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   Icon            =   "frmRRHH001.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   7170
   Begin Crystal.CrystalReport Reporte 
      Left            =   2835
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmListaAsist 
      Caption         =   "Lista de &Asistencia"
      Height          =   360
      Left            =   210
      TabIndex        =   9
      Top             =   5655
      Width           =   1830
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   5610
      TabIndex        =   11
      Top             =   5640
      Width           =   1380
   End
   Begin VB.CommandButton cmCitaciones 
      Caption         =   "&Imprimir Citaciones"
      Height          =   360
      Left            =   2175
      TabIndex        =   10
      Top             =   5655
      Width           =   1830
   End
   Begin VB.Frame Frame1 
      Caption         =   "Configuración del Reporte"
      Height          =   5400
      Left            =   210
      TabIndex        =   0
      Top             =   150
      Width           =   6780
      Begin VB.TextBox xHoraFin 
         Height          =   285
         Left            =   4305
         TabIndex        =   22
         Top             =   4125
         Width           =   1530
      End
      Begin VB.TextBox xHoraIni 
         Height          =   285
         Left            =   2130
         TabIndex        =   21
         Top             =   4125
         Width           =   1530
      End
      Begin VB.CommandButton cmSelecTrab 
         Caption         =   "Seleccion (F5)"
         Height          =   990
         Left            =   5625
         Picture         =   "frmRRHH001.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Seleccione aquí los trabajadores quienes particparan de la citación"
         Top             =   630
         Width           =   945
      End
      Begin VB.TextBox xSegundo 
         Height          =   285
         Left            =   1455
         MaxLength       =   250
         TabIndex        =   7
         Top             =   4545
         Width           =   5115
      End
      Begin MSComCtl2.DTPicker xFechaFin 
         Height          =   285
         Left            =   4305
         TabIndex        =   6
         Top             =   3780
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   62193665
         CurrentDate     =   36869
      End
      Begin MSComCtl2.DTPicker xFechaIni 
         Height          =   285
         Left            =   2130
         TabIndex        =   5
         Top             =   3780
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   503
         _Version        =   393216
         Format          =   62193665
         CurrentDate     =   36869
      End
      Begin VB.TextBox xTitulo 
         Height          =   285
         Left            =   1845
         MaxLength       =   80
         TabIndex        =   2
         Top             =   2160
         Width           =   4725
      End
      Begin VB.TextBox xFirma 
         Height          =   285
         Left            =   4230
         TabIndex        =   8
         Text            =   "La Administración"
         Top             =   4995
         Width           =   2340
      End
      Begin VB.TextBox xCuerpo 
         Height          =   840
         Left            =   150
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   2790
         Width           =   6420
      End
      Begin MSComctlLib.ListView Lista 
         Height          =   1395
         Left            =   150
         TabIndex        =   1
         Top             =   615
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   2461
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Descripción de los Datos"
            Object.Width           =   11289
         EndProperty
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Segundo parrafo"
         Height          =   195
         Left            =   150
         TabIndex        =   19
         Top             =   4620
         Width           =   1185
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "hasta"
         Height          =   195
         Left            =   3810
         TabIndex        =   18
         Top             =   4170
         Width           =   390
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Hora de Inicio"
         Height          =   195
         Left            =   150
         TabIndex        =   17
         Top             =   4155
         Width           =   990
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "hasta"
         Height          =   195
         Left            =   3810
         TabIndex        =   16
         Top             =   3810
         Width           =   390
      End
      Begin VB.Label xNombreFecha 
         AutoSize        =   -1  'True
         Caption         =   "Fecha a relizar"
         Height          =   195
         Left            =   150
         TabIndex        =   4
         ToolTipText     =   "Puede hacer doble click para cambiar"
         Top             =   3825
         Width           =   1035
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Título de la Citación"
         Height          =   195
         Left            =   150
         TabIndex        =   15
         Top             =   2205
         Width           =   1425
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Firma:"
         Height          =   195
         Left            =   3645
         TabIndex        =   14
         Top             =   5040
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Texto que aparecerá en el cuerpo del mensaje"
         Height          =   195
         Left            =   150
         TabIndex        =   13
         Top             =   2565
         Width           =   3300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Seleccione datos del trabajador a mostrar en la Citación"
         Height          =   195
         Left            =   165
         TabIndex        =   12
         Top             =   345
         Width           =   3930
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   -45
      Top             =   -180
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
            Picture         =   "frmRRHH001.frx":1194
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmRRHH001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ACAMPOS(5) As String

Private Sub CMCITACIONES_Click()
    If xTitulo.Text = "" Then
        MsgBox "DEBE TENER UN TITULO", vbInformation
        Exit Sub
    End If
    With Reporte
        .Reset
        .ReportFileName = REGSISTEMA.REPORTES & "\PLNAS005.RPT"
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "XCODTRAB='" & IIf(Lista.ListItems(1).Checked, "SI", "NO") & "'"
        .Formulas(2) = "XNOMBRES='" & IIf(Lista.ListItems(2).Checked, "SI", "NO") & "'"
        .Formulas(3) = "XAREA='" & IIf(Lista.ListItems(3).Checked, "SI", "NO") & "'"
        .Formulas(4) = "XFECHAING='" & IIf(Lista.ListItems(4).Checked, "SI", "NO") & "'"
        .Formulas(5) = "XAREA='" & IIf(Lista.ListItems(5).Checked, "SI", "NO") & "'"
        .Formulas(6) = "XTITULO='" & xTitulo.Text & "'"
        .Formulas(7) = "XCUERPO='" & xCuerpo.Text & "'"
        .Formulas(8) = "XNOMBREFECHA='" & xNombreFecha.Caption & "'"
        .Formulas(9) = "XFECHAINI='" & xFechaIni.Value & "'"
        .Formulas(10) = "XFECHAFIN='" & IIf(IsNull(xFechaFin.Value), "NO", xFechaFin.Value) & "'"
        .Formulas(11) = "XHORAINI='" & IIf(xHoraIni = "", "NO", xHoraIni.Text) & "'"
        .Formulas(12) = "XHORAFIN='" & IIf(xHoraFin = "", "NO", " - HASTA " & xHoraFin.Text) & "'"
        .Formulas(13) = "XSEGUNDO='" & xSegundo.Text & "'"
        .Formulas(14) = "XFIRMA=' " & xFirma.Text & "'"
        .WindowTitle = "PLNAS005- EMISION DE CITACIONES"
        .DataFiles(0) = App.PATH & "\BDAUXCOM.MDB"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        If .Status <> 2 Then .Action = 1
    End With
End Sub

Private Sub CMLISTAASIST_Click()
    If xTitulo.Text = "" Then
        MsgBox "DEBE TENER UN TITULO", vbInformation
        Exit Sub
    End If
    With Reporte
        .Reset
        .ReportFileName = REGSISTEMA.REPORTES & "PLNAS002.RPT"
        .WindowTitle = "PLNAS002- LISTA PARA FIRMAR"
        .Connect = "DSN=" & VGL_SERVER & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL & ""
        .DataFiles(0) = App.PATH & "\BDAUXCOM.MDB"
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "XTITULO='" & xTitulo.Text & "'"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        If .Status <> 2 Then .Action = 1
    End With
End Sub

Private Sub CMSELECTRAB_CLICK()
    REGSELECT.USARFECHACESE = False
    frSelect.Show 1
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub FORM_KEYDOWN(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = vbKeyF5 Then CMSELECTRAB_CLICK
End Sub

Private Sub Form_Load()
    ACAMPOS(0) = "CÓDIGO DEL TRABAJADOR"
    ACAMPOS(1) = "NOMBRES DEL TRABAJADOR"
    ACAMPOS(2) = "NÚMERO DE DOCUMENTO"
    ACAMPOS(3) = "FECHA DE INGRESO"
    ACAMPOS(4) = "AREA DE TRABAJO"
    CARGALISTA
End Sub

Public Sub CARGALISTA()
    Dim X As Byte
    Lista.ListItems.Clear
    For X = 0 To 4
        Lista.ListItems.Add , "C" & X, ACAMPOS(X), , 1
    Next
End Sub

Private Sub XNOMBREFECHA_Click()
    Dim S As String
    S = InputBox("ESCRIBA EL NOMBRE DEL TITULO PARA LA FECHA 1", "PERSONALIZACIÓN DEL REPORTE", xNombreFecha.Caption)
    If S <> "" Then xNombreFecha.Caption = S
End Sub

