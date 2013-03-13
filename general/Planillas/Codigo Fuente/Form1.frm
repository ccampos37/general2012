VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Begin VB.Form CalcPlan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cálculo de Planilla de Remuneraciones"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7560
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   7560
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Cargar Modo Edición"
      Height          =   1035
      Left            =   90
      TabIndex        =   25
      Top             =   3690
      Width           =   2295
      Begin VB.CheckBox Check7 
         Caption         =   "Neto de Vacaciones"
         Height          =   210
         Left            =   225
         TabIndex        =   29
         ToolTipText     =   "Carga los Netos de Vacaciones si los trabajadores estan de vacaciones"
         Top             =   750
         Value           =   1  'Checked
         Width           =   1800
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Nuevos Mov.Cta.Cte"
         Height          =   210
         Left            =   225
         TabIndex        =   27
         Top             =   510
         Width           =   1860
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Nuevos Adelantos"
         Height          =   240
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Formatos de Planilla"
      Height          =   1710
      Left            =   105
      TabIndex        =   20
      Top             =   4875
      Width           =   7290
      Begin VB.CheckBox xRedondeo 
         Caption         =   "Aplicar redondeo de pago neto"
         Height          =   225
         Left            =   120
         TabIndex        =   28
         Top             =   1365
         Width           =   2550
      End
      Begin VB.CommandButton xAbrir 
         Height          =   420
         Left            =   3240
         Picture         =   "Form1.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Abrir formato de planilla"
         Top             =   780
         Width           =   420
      End
      Begin MSDataGridLib.DataGrid DGRubs 
         Height          =   1350
         Left            =   3810
         TabIndex        =   24
         Top             =   255
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   2381
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483633
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Rubros a Usar"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            AllowRowSizing  =   0   'False
            RecordSelectors =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmFormatos 
         Caption         =   "&Editor de Formatos"
         Height          =   360
         Left            =   120
         TabIndex        =   23
         Top             =   855
         Width           =   1605
      End
      Begin AplisetControlText.Aplitext xFormato 
         Height          =   300
         Left            =   120
         TabIndex        =   22
         Top             =   435
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   529
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin VB.Label Label3 
         Caption         =   "Seleccionar Formato"
         Height          =   225
         Left            =   150
         TabIndex        =   21
         Top             =   225
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos a Cargar"
      Height          =   1305
      Left            =   90
      TabIndex        =   16
      Top             =   2310
      Width           =   2295
      Begin VB.CheckBox Check4 
         Caption         =   "Cuentas Corrientes"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   990
         Value           =   1  'Checked
         Width           =   1860
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Movimientos"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   765
         Value           =   1  'Checked
         Width           =   1860
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Asistencia"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   525
         Value           =   1  'Checked
         Width           =   1860
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Adelantos de Remun."
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   300
         Value           =   1  'Checked
         Width           =   1860
      End
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6105
      TabIndex        =   5
      Top             =   840
      Width           =   1275
   End
   Begin VB.CommandButton cmContinuar 
      Caption         =   "Continuar >>"
      Height          =   375
      Left            =   6105
      TabIndex        =   4
      Top             =   390
      Width           =   1275
   End
   Begin MSDataGridLib.DataGrid DGLista 
      Height          =   2325
      Left            =   2505
      TabIndex        =   14
      Top             =   2400
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   4101
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Trabajadores Seleccionados"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmSelecTrab 
      Caption         =   "Seleccion (F5)"
      Height          =   990
      Left            =   6510
      Picture         =   "Form1.frx":064C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1335
      Width           =   870
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Por Centros de Costo"
      Height          =   210
      Left            =   150
      TabIndex        =   7
      Top             =   1125
      Width           =   1830
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Por Areas de Trabajo"
      Height          =   210
      Left            =   150
      TabIndex        =   15
      Top             =   825
      Value           =   -1  'True
      Width           =   1830
   End
   Begin AplisetControlText.Aplitext xMes 
      Height          =   285
      Left            =   180
      TabIndex        =   0
      Top             =   390
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   503
      Locked          =   -1  'True
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker xFechaFin 
      Height          =   285
      Left            =   1065
      TabIndex        =   8
      Top             =   1770
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   16842753
      CurrentDate     =   36699
   End
   Begin MSComCtl2.DTPicker xFechaIni 
      Height          =   285
      Left            =   1065
      TabIndex        =   9
      Top             =   1440
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   16842753
      CurrentDate     =   36699
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   1965
      Left            =   2490
      TabIndex        =   1
      Top             =   375
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   3466
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Periodos en Cronogramas"
         Object.Width           =   5733
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "FechaIni"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "FechaFin"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label l2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Final"
      Height          =   195
      Left            =   105
      TabIndex        =   13
      Top             =   1830
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label l1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Inicial"
      Height          =   195
      Left            =   105
      TabIndex        =   12
      Top             =   1485
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Mes de Trabajo"
      Height          =   195
      Left            =   195
      TabIndex        =   11
      Top             =   150
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Periodos en Cronograma"
      Height          =   195
      Left            =   2520
      TabIndex        =   10
      Top             =   165
      Width           =   1740
   End
End
Attribute VB_Name = "CalcPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim XITEM As ListItem, CADIN As String
Dim RSTRAB As New ADODB.Recordset
Dim RSRUBS As New ADODB.Recordset

Private Sub CMCONTINUAR_CLICK()
    If xMes.Text = "" Then
        MsgBox "No ha seleccionado un mes de pago, por favor seleccione un mes activo y si no existe solicite al administrador del sistema la creación de un nuevo mes activo", vbCritical
        Exit Sub
    End If
    If Not xFechaIni.Visible Then
        MsgBox "NO HA SELECCIONADO UN PERIODO DE PAGO, POR FAVOR CARGUE EL MES Y SELECCIONE UN PERIODO DE PAGO QUE HALLA FIJADO EN SU CRONOGRAMA DE PAGOS", vbCritical
        Exit Sub
    End If
    If RSTRAB.RecordCount = 0 Then
        MsgBox "LA SELECCIÓN DE TRABAJADORES NO PUEDE ESTAR VACIA, DEBERÁ REALIZAR UNA NUEVA SELECCIÓN DE TRABAJADORES", vbCritical
        Exit Sub
    End If
    If xFormato.Text = "" Then
        MsgBox "NO HA SELECCIONADO UN FORMATO DE PLANILLA, POR FAVOR SELECCIONELO, VERIFIQUE LOS CONCEPTOS DE REMUNERACIONES QUE SE VAN A PROCESAR", vbCritical
        Exit Sub
    End If
    If RSRUBS.RecordCount = 0 Then
        MsgBox "EL FORMATO DE BOLETA QUE HA SELECCIONADO NO CONTIENE CONCEPTOS DE REMUNERACIONES, POR FAVOR AGREGELÉ NUEVOS CONCEPTOS DESDE EL EDITOR DE FORMATOS DE PLANILLA", vbCritical
        Exit Sub
    End If
    Dim RSDELS As New ADODB.Recordset
    CADIN = ""
    If DevuelveValor("SELECT USARCRONOGRAMA FROM EMPRESA", DBSYSTEM) = 1 Then
        If Option1.Value Then
            RSDELS.Open "SELECT DISTINCT CODREF FROM FECHAPAGO, NOMBOL WHERE TIPOAC=0 AND CERRADO=0 AND CODNOMBOL=" & Lista.SelectedItem.Tag, DBSYSTEM, adOpenStatic
        Else
            RSDELS.Open "SELECT DISTINCT CODREF FROM FECHAPAGO, NOMBOL WHERE TIPOAC=1 AND CERRADO=0 AND CODNOMBOL=" & Lista.SelectedItem.Tag, DBSYSTEM, adOpenStatic
        End If
        If RSDELS.RecordCount = 0 Then
            MsgBox "NO SE HAN ENCONTRADO AREAS O CENTROS DE COSTOS QUE ESTEN PROGRAMADOS PARA PAGOS DE ADELANTOS DE REMUNERACIONES EN EL PERIODO SELECCIONADO", vbCritical
            Set RSDELS = Nothing
            Exit Sub
        End If
        Do While Not RSDELS.EOF
            If CADIN = "" Then CADIN = "'" & RSDELS!CODREF & "'" Else CADIN = CADIN & ",'" & RSDELS!CODREF & "'"
            RSDELS.MoveNext
        Loop
        CADIN = "AND AREA IN (" & CADIN & ")"
    End If
    'CARGA DEL REGINPUT
    '------------------
    VPTRASPRM = xFormato.Tag
    REGINPUT.CADENA = CADIN
    REGINPUT.FECHAINI = xFechaIni.Value
    REGINPUT.FECHAFIN = xFechaFin.Value
    REGINPUT.Codigo = Lista.SelectedItem.Tag
    REGINPUT.MESACTIVO = xMes.Tag
    REGINPUT.BOL_TABLE = "BOL" & Format(Month(xMes.Tag), "00") & Year(xMes.Tag)
    REGINPUT.MOV_TABLE = "MOV" & Format(Month(xMes.Tag), "00") & Year(xMes.Tag)
    REGINPUT.NOMBRE = Lista.SelectedItem.Text
    REGINPUT.REDONDEO = IIf(xRedondeo.Value = 1, True, False)
    Set RSDELS = Nothing
    If Not ExisteTabla(REGINPUT.BOL_TABLE) Then
        MsgBox "NO EXISTE LA TABLA DE BOLETAS " & REGINPUT.BOL_TABLE, vbCritical
        Exit Sub
    End If
    If Not ExisteTabla(REGINPUT.MOV_TABLE) Then
        MsgBox "NO EXISTE LA TABLA DE MOVIMIENTOS " & REGINPUT.MOV_TABLE, vbCritical
        Exit Sub
    End If
    FrmDetAdel.FMES = xFechaIni.Value
    InputPl.Show 1
End Sub

Private Sub CMFORMATOS_Click()
    frFormatos.Show 1
    CARGARUBROS
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub DGLISTA_HEADCLICK(ByVal COLINDEX As Integer)
    RSTRAB.Sort = DGLista.Columns(COLINDEX).DataField
End Sub

Private Sub DGLISTA_KeyDown(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 13 Then SendKeys "{TAB}"
End Sub

Private Sub FORM_KEYDOWN(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = vbKeyF5 Then CMSELECTRAB_CLICK
End Sub

Private Sub Form_Load()
    Me.TOP = 0
    Me.Left = 0
    If ExisteTablaAux(" [##TMPSELECT" & VGL_COMPUTER & "] ") Then DBSTARPLAN.Execute "DROP TABLE [##TMPSELECT" & VGL_COMPUTER & "]"
    DBSTARPLAN.Execute "CREATE TABLE  [##TMPSELECT" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8), NOMBRES VARCHAR(50), FECHAING DATETIME, AREA VARCHAR(10), CENTROCOSTO VARCHAR(10), TIPOTRAB VARCHAR(2), BASICO  Numeric(20,2),basico1 numeric(20,2) )"
    RSTRAB.Open " [##TMPSELECT" & VGL_COMPUTER & "] ", DBSTARPLAN, adOpenStatic
    Set DGLista.DataSource = RSTRAB
    DGLista.Caption = "Trabajadores Seleccionados Nº:" & RSTRAB.RecordCount
    Dim XFEC As String
    XFEC = "01/" & MDIPrincipal.BarraEstado.Panels("Periodo").Text
    If IsDate(XFEC) Then
        xMes.Text = DevuelveValor("SELECT NOMBRE FROM MESESACT WHERE MESACTIVO=" & DateSQL(XFEC), DBSYSTEM)
        xMes.Tag = CDate(XFEC)
        CARGAMESES
    End If
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSTRAB = Nothing
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
    xFechaIni.Visible = True
    xFechaFin.Visible = True
    l1.Visible = True
    l2.Visible = True
    xFechaIni.Value = CDate(Item.SubItems(1))
    xFechaFin.Value = CDate(Item.SubItems(2))
End Sub

Private Sub CMSELECTRAB_CLICK()
    REGSELECT.FECHACESEMAX = xFechaFin.Value
    REGSELECT.FECHAINIMAX = xFechaFin.Value
    REGSELECT.FECHAINI = xFechaIni.Value
    REGSELECT.SITUACIONES = "'0','1'"
    REGSELECT.USARFECHACESE = True
    frSelect.Show 1
    REGSELECT.USARFECHACESE = False
    If Not xFechaIni.Visible Then
        MsgBox "DEBERÁ SELECCIONAR UN PERIODO DE PAGO", vbCritical
        Exit Sub
    End If
    RSTRAB.Requery
    Set DGLista.DataSource = RSTRAB
    DGLista.Caption = "Trabajdores Seleccionados Nº:" & RSTRAB.RecordCount
End Sub

Private Sub LISTA_KeyDown(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 13 Then SendKeys "{TAB}"
End Sub

Private Sub XABRIR_Click()
XFORMATO_DblClick
End Sub

Private Sub XFORMATO_DblClick()
    Dim RSFORMA As New ADODB.Recordset
    RSFORMA.Open "SELECT * FROM FORMATOS ORDER BY NOMBRE", DBSYSTEM, adOpenStatic
    If RSFORMA.EOF Or RSFORMA.RecordCount = 0 Then
        MsgBox "NO SE HAN ENCONTRADO FORMATOS DE PLANILLA ALMACENADOS EN LA BASE DE DATOS DEL SISTEMA. POR FAVOR CREE UNO NUEVO, EN EL EDITOR DE FORMATOS DE PLANILLA", vbCritical
        Set RSFORMA = Nothing
        Exit Sub
    End If
    frmComun.CONECTAR RSFORMA
    frmComun.Show 1
    If VGUTIL(1) = "" Then
        Set RSFORMA = Nothing
        Exit Sub
    End If
    xFormato.Text = RSFORMA!NOMBRE
    xFormato.Tag = RSFORMA!ID_FORMATO
    Set RSFORMA = Nothing
    CARGARUBROS
End Sub

Private Sub XMES_DBLCLICK()
    Dim RSMESES As New ADODB.Recordset
    RSMESES.Open "SELECT MESACTIVO, NOMBRE FROM MESESACT ORDER BY MESACTIVO", DBSYSTEM, adOpenStatic
    If RSMESES.RecordCount = 0 Then
        MsgBox "NO SE HAN ENCONTRADO MESES EN ACTIVIDAD", vbCritical
        Set RSMESES = Nothing
        Exit Sub
    End If
    frmComun.CONECTAR RSMESES
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xMes.Text = RSMESES!NOMBRE
        xMes.Tag = RSMESES!MESACTIVO
    Else
        Set RSMESES = Nothing
        Exit Sub
    End If
    Set RSMESES = Nothing
    'RECICLAJE DE RSMESES
    CARGAMESES
End Sub

Public Sub CARGARUBROS()
    If xFormato.Tag = "" Or xFormato.Text = "" Then Exit Sub
    Set RSRUBS = Nothing
    RSRUBS.Open "SELECT CODIGO,NOMBRE,CONCEPTOS.TIPO,FORMULA,COLPLANILLA FROM CONCEPTOS, FORMARUBS WHERE FORMARUBS.CONCEPTO=CONCEPTOS.CODIGO AND FORMARUBS.ID_FORMATO=" & xFormato.Tag & " ORDER BY CONCEPTOS.TIPO, NOMBRE", DBSYSTEM, adOpenStatic
    Set DGRubs.DataSource = RSRUBS
End Sub

Public Sub CARGAMESES()
    Dim RSMESES As New ADODB.Recordset
    Lista.ListItems.Clear
    RSMESES.Open "SELECT CODIGO, NOMBRE, FECHAINI, FECHAFIN FROM NOMBOL WHERE CERRADO<>1 AND MES=" & DateSQL(CDate(xMes.Tag)) & " ORDER BY FECHAINI", DBSYSTEM, adOpenStatic
    Do While Not RSMESES.EOF
        Set XITEM = Lista.ListItems.Add(, , RSMESES!NOMBRE, , 1)
        XITEM.SubItems(1) = RSMESES!FECHAINI
        XITEM.SubItems(2) = RSMESES!FECHAFIN
        XITEM.Tag = RSMESES!Codigo
        RSMESES.MoveNext
    Loop
    l1.Visible = False
    l2.Visible = False
    xFechaIni.Visible = False
    xFechaFin.Visible = False
    Set RSMESES = Nothing
End Sub

