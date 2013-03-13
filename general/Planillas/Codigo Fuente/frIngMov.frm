VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Begin VB.Form frIngMov 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Movimientos"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6420
   Icon            =   "frIngMov.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   6420
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmCancelar 
      Cancel          =   -1  'True
      Caption         =   "C&ancelar"
      Height          =   375
      Left            =   3217
      TabIndex        =   11
      Top             =   5235
      Width           =   1425
   End
   Begin VB.CommandButton cmContinuar 
      Caption         =   "&Continuar >>"
      Height          =   375
      Left            =   1552
      TabIndex        =   10
      Top             =   5235
      Width           =   1425
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   135
      Top             =   1800
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
            Picture         =   "frIngMov.frx":030A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Especificaciones de Ingreso"
      Height          =   2385
      Left            =   150
      TabIndex        =   2
      Top             =   2685
      Width           =   6195
      Begin VB.OptionButton Option2 
         Caption         =   "Por Centros de Costo"
         Height          =   210
         Left            =   195
         TabIndex        =   15
         Top             =   1200
         Width           =   1830
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Areas de Trabajo"
         Height          =   210
         Left            =   195
         TabIndex        =   14
         Top             =   945
         Value           =   -1  'True
         Width           =   1830
      End
      Begin MSComCtl2.DTPicker xFechaFin 
         Height          =   285
         Left            =   1125
         TabIndex        =   9
         Top             =   1920
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   53477377
         CurrentDate     =   36699
      End
      Begin MSComCtl2.DTPicker xFechaIni 
         Height          =   285
         Left            =   1125
         TabIndex        =   7
         Top             =   1575
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   53477377
         CurrentDate     =   36699
      End
      Begin MSComctlLib.ListView Lista 
         Height          =   1725
         Left            =   2535
         TabIndex        =   5
         Top             =   495
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   3043
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
      Begin AplisetControlText.Aplitext xMes 
         Height          =   285
         Left            =   165
         TabIndex        =   4
         Top             =   510
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   503
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Periodos en Cronograma"
         Height          =   195
         Left            =   2580
         TabIndex        =   13
         Top             =   270
         Width           =   1740
      End
      Begin VB.Label l2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final"
         Height          =   195
         Left            =   165
         TabIndex        =   8
         Top             =   1980
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial"
         Height          =   195
         Left            =   165
         TabIndex        =   6
         Top             =   1620
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Mes de Trabajo"
         Height          =   195
         Left            =   165
         TabIndex        =   3
         Top             =   270
         Width           =   1110
      End
   End
   Begin VB.CommandButton cmSelecTrab 
      Caption         =   "Seleccionar (F5)"
      Height          =   1080
      Left            =   165
      Picture         =   "frIngMov.frx":065E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   180
      Width           =   1005
   End
   Begin MSDataGridLib.DataGrid DGFiltro 
      Height          =   2325
      Left            =   1305
      TabIndex        =   1
      Top             =   195
      Visible         =   0   'False
      Width           =   5010
      _ExtentX        =   8837
      _ExtentY        =   4101
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483633
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   3
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
         AllowRowSizing  =   0   'False
         RecordSelectors =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Todos los trabajadores de la empresa seleccionada"
      Height          =   510
      Left            =   1305
      TabIndex        =   12
      Top             =   600
      Width           =   4755
   End
End
Attribute VB_Name = "frIngMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSAUX As New ADODB.Recordset
Dim XITEM As ListItem

Private Sub CMCANCELAR_CLICK()
    Unload Me
End Sub

Private Sub CMCONTINUAR_CLICK()
    If Not ExisteTablaAux(" [##TMPSELECT" & VGL_COMPUTER & "] ") Then MsgBox "SELECCIONE LOS TRABAJADORES": Exit Sub
    If xFechaIni.Visible Then
        With REGINGMOV
            .CODNOMBOL = Lista.SelectedItem.Tag
            .FECHAFIN = xFechaFin.Value
            .FECHAINI = xFechaIni.Value
            .NOMBRE = Lista.SelectedItem.Text
            .AREA = Option1.Value
        End With
        If DevuelveValor("SELECT USARCRONOGRAMA FROM EMPRESA", DBSYSTEM) = 1 Then
            Dim RSDELS As New ADODB.Recordset, CADIN As String
            If Option1.Value Then
                RSDELS.Open "SELECT CODREF FROM FECHAPAGO WHERE TIPOAC=0 AND CODNOMBOL=" & REGINGMOV.CODNOMBOL, DBSYSTEM, adOpenStatic
            Else
                RSDELS.Open "SELECT CODREF FROM FECHAPAGO WHERE TIPOAC=1 AND CODNOMBOL=" & REGINGMOV.CODNOMBOL, DBSYSTEM, adOpenStatic
            End If
            If RSDELS.RecordCount = 0 Then
                MsgBox "No se han encontrado Areas ó Centros de Costos que esten programados para pagos en el Periodo seleccionado", vbCritical
                Set RSDELS = Nothing
                Exit Sub
            End If
            CADIN = ""
            Do While Not RSDELS.EOF
                If CADIN = "" Then CADIN = "'" & RSDELS!CODREF & "'" Else CADIN = CADIN & ",'" & RSDELS!CODREF & "'"
                RSDELS.MoveNext
            Loop
            CADIN = "WHERE AREA IN (" & CADIN & ")"
            REGINGMOV.CADCONDI = CADIN
            Set RSDELS = Nothing
        End If
        frInputMov.Show 1
    Else
        MsgBox "Faltan completar datos", vbCritical
    End If
End Sub

Private Sub CMSELECTRAB_CLICK()
    If Not xFechaIni.Visible Then
        MsgBox "Debera seleccionar primero el Periodo de Pago", vbInformation
        Exit Sub
    End If
    REGSELECT.FECHACESEMAX = xFechaFin.Value
    REGSELECT.FECHAINIMAX = xFechaFin.Value
    REGSELECT.FECHAINI = xFechaIni.Value
    REGSELECT.USARFECHACESE = True
    frSelect.Show 1
    Set RSAUX = Nothing
    RSAUX.Open " [##TMPSELECT" & VGL_COMPUTER & "] ", DBAUXCOM, adOpenStatic
    Set DGFiltro.DataSource = RSAUX
    If RSAUX.RecordCount = 0 Then
        DGFiltro.Visible = False
    Else
        DGFiltro.Visible = True
    End If
End Sub

Private Sub FORM_KEYDOWN(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = vbKeyF5 Then CMSELECTRAB_CLICK
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSAUX = Nothing
End Sub
Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
    xFechaIni.Visible = True
    xFechaFin.Visible = True
    l1.Visible = True
    l2.Visible = True
    xFechaIni.Value = CDate(Item.SubItems(1))
    xFechaFin.Value = CDate(Item.SubItems(2))
End Sub
Private Sub XMES_DBLCLICK()
    Lista.ListItems.Clear
    Dim RSMESES As New ADODB.Recordset
    RSMESES.Open "SELECT MESACTIVO, NOMBRE FROM MESESACT ORDER BY MESACTIVO", DBSYSTEM, adOpenStatic
    If RSMESES.RecordCount = 0 Then
        MsgBox "No se han encontrado meses en actividad", vbCritical
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
    RSMESES.Open "SELECT CODIGO, NOMBRE, FECHAINI, FECHAFIN FROM NOMBOL WHERE CERRADO=0 AND MES=" & DateSQL(CDate(xMes.Tag)) & " ORDER BY FECHAINI", DBSYSTEM, adOpenStatic
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

