VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frBolEmit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Panel de Administración de Resultados de Planilla"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8460
   Icon            =   "frBolEmit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   8460
   Tag             =   "Panel de Boletas Emitidas"
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   6720
      MultiLine       =   -1  'True
      TabIndex        =   36
      Text            =   "frBolEmit.frx":0442
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton xVerRango 
      Caption         =   ">>"
      Height          =   300
      Left            =   4830
      TabIndex        =   35
      Top             =   150
      Width           =   390
   End
   Begin MSComCtl2.DTPicker xFechaFin 
      Height          =   315
      Left            =   3555
      TabIndex        =   34
      Top             =   150
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "MM - yyyy"
      Format          =   61931523
      CurrentDate     =   36867
   End
   Begin MSComCtl2.DTPicker xFechaIni 
      Height          =   315
      Left            =   2100
      TabIndex        =   32
      Top             =   150
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "MM - yyyy"
      Format          =   61931523
      CurrentDate     =   36867
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   390
      Left            =   135
      TabIndex        =   30
      Top             =   6090
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   688
      ButtonWidth     =   3254
      ButtonHeight    =   582
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Otros Procesos        "
            Object.ToolTipText     =   "Click aquí para más reportes de boletas"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   7
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Grafica10"
                  Text            =   "Grafica Estadistica de los 10 mayores"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "firmar"
                  Text            =   "Listado para firmar"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "EliminaTodasBol"
                  Text            =   "Eliminar Todas las Boletas"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "TipoF5"
                  Text            =   "Filtrar Trabajadores (Tipo F5)"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "QuitarTrab"
                  Text            =   "Quitar Registros Seleccionados"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RenamePeriodo"
                  Text            =   "Cambia nombre del periodo"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Key             =   "REPORTEMENSUAL"
                  Text            =   "Reporte Mensual"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdPlan 
      Caption         =   "P&lanilla Continua"
      Height          =   360
      Left            =   6885
      TabIndex        =   21
      Top             =   5700
      Width           =   1440
   End
   Begin VB.CommandButton cmFormatoPlanilla 
      Caption         =   "&Formato Planilla"
      Height          =   360
      Left            =   6885
      TabIndex        =   20
      Top             =   5210
      Width           =   1440
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Resumen"
      Height          =   360
      Left            =   6885
      TabIndex        =   19
      Top             =   4720
      Width           =   1440
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Debitos Cta. Cte"
      Height          =   360
      Left            =   6885
      TabIndex        =   17
      Top             =   3740
      Width           =   1440
   End
   Begin VB.CommandButton CmdDeta 
      Caption         =   "&Detalles"
      Height          =   360
      Left            =   6885
      TabIndex        =   18
      Top             =   4230
      Width           =   1440
   End
   Begin VB.CommandButton cmPlanilla 
      Caption         =   "&Planilla"
      Height          =   360
      Left            =   6885
      TabIndex        =   16
      Top             =   3250
      Visible         =   0   'False
      Width           =   1440
   End
   Begin AplisetControlText.Aplitext xArea 
      Height          =   285
      Left            =   3705
      TabIndex        =   25
      Top             =   2310
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   503
      Locked          =   -1  'True
      Text            =   ""
   End
   Begin VB.OptionButton Sel1 
      BackColor       =   &H00808080&
      Caption         =   "&Centros de Costo"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   1935
      TabIndex        =   24
      Top             =   2370
      Width           =   1530
   End
   Begin VB.OptionButton Sel1 
      BackColor       =   &H00808080&
      Caption         =   "&Areas de Trabajo"
      ForeColor       =   &H8000000E&
      Height          =   225
      Index           =   0
      Left            =   150
      TabIndex        =   23
      Top             =   2370
      Value           =   -1  'True
      Width           =   1560
   End
   Begin VB.CommandButton cmCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   6885
      TabIndex        =   22
      Top             =   6195
      Width           =   1440
   End
   Begin VB.CommandButton cmBilletes 
      Caption         =   "&Billetaje"
      Height          =   360
      Left            =   6885
      TabIndex        =   15
      Top             =   2760
      Width           =   1440
   End
   Begin Crystal.CrystalReport RptBoletas 
      Left            =   6195
      Top             =   5355
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmPagosBanco 
      Caption         =   "Pagos por Banco"
      Height          =   360
      Left            =   6885
      TabIndex        =   14
      Top             =   2270
      Width           =   1440
   End
   Begin VB.CommandButton cmPrtTodos 
      Caption         =   "&Imprimir Todos"
      Height          =   360
      Left            =   6885
      TabIndex        =   13
      Top             =   1780
      Width           =   1440
   End
   Begin VB.CommandButton cmPrtUno 
      Caption         =   "Imprimir &Uno"
      Height          =   360
      Left            =   6885
      TabIndex        =   12
      Top             =   1290
      Width           =   1440
   End
   Begin VB.ComboBox xMeses 
      Height          =   315
      Left            =   2115
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   150
      Width           =   2790
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   1830
      Left            =   135
      TabIndex        =   2
      Top             =   480
      Width           =   6510
      _ExtentX        =   11483
      _ExtentY        =   3228
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nombre"
         Object.Width           =   6774
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Fecha Inic."
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fecha Term."
         Object.Width           =   2011
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   135
      Top             =   4620
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frBolEmit.frx":16D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frBolEmit.frx":295C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid dgBoletas 
      Height          =   3165
      Left            =   150
      TabIndex        =   0
      Top             =   2625
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   5583
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16777215
      HeadLines       =   2
      RowHeight       =   17
      RowDividerStyle =   0
      FormatLocked    =   -1  'True
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Boletas de Remuneraciones"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "CODTRAB"
         Caption         =   "Código"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "NOMBRES"
         Caption         =   "Trabajador"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "INGRESOS"
         Caption         =   "Total Ingresos"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00 "
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "EGRESOS"
         Caption         =   "Total Egresos"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00 "
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "NETO"
         Caption         =   "Neto a Pagar"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00 "
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
         EndProperty
      EndProperty
   End
   Begin VB.TextBox SqlCad 
      Height          =   360
      Left            =   4650
      TabIndex        =   29
      Top             =   2910
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.ComboBox xVistaMes 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frBolEmit.frx":2DB0
      Left            =   150
      List            =   "frBolEmit.frx":2DBA
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   150
      Width           =   1905
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "al"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3345
      TabIndex        =   33
      Top             =   210
      Width           =   120
   End
   Begin VB.Image xError 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   7785
      Picture         =   "frBolEmit.frx":2DE5
      Top             =   75
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   7575
      Picture         =   "frBolEmit.frx":36AF
      Top             =   105
      Width           =   240
   End
   Begin VB.Line Line2 
      X1              =   3270
      X2              =   3360
      Y1              =   6030
      Y2              =   6030
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   15
      Left            =   3255
      TabIndex        =   28
      Top             =   6030
      Width           =   120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   6180
      X2              =   6300
      Y1              =   420
      Y2              =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Resultados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   7080
      TabIndex        =   27
      Top             =   570
      Width           =   1200
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "de Planillas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   7035
      TabIndex        =   26
      Top             =   810
      Width           =   1245
   End
   Begin VB.Label xCont 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0 Registros"
      Height          =   270
      Left            =   135
      TabIndex        =   11
      Top             =   5805
      Width           =   2655
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Egresos"
      Height          =   270
      Left            =   4035
      TabIndex        =   10
      Top             =   6075
      Width           =   1320
   End
   Begin VB.Label xSumEgr 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      Height          =   270
      Left            =   5355
      TabIndex        =   9
      Top             =   6075
      Width           =   1080
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Neto"
      Height          =   270
      Left            =   4035
      TabIndex        =   8
      Top             =   6345
      Width           =   1320
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Ingresos"
      Height          =   270
      Left            =   4035
      TabIndex        =   7
      Top             =   5805
      Width           =   1320
   End
   Begin VB.Label xSumNet 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      Height          =   270
      Left            =   5355
      TabIndex        =   6
      Top             =   6345
      Width           =   1080
   End
   Begin VB.Label xSumIng 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      Height          =   270
      Left            =   5355
      TabIndex        =   5
      Top             =   5805
      Width           =   1080
   End
   Begin VB.Image xVerDetalle 
      Height          =   240
      Left            =   2910
      Picture         =   "frBolEmit.frx":3AF1
      Top             =   5835
      Width           =   240
   End
   Begin VB.Image xMarcaTodos 
      Height          =   240
      Left            =   5325
      Picture         =   "frBolEmit.frx":3E33
      Top             =   165
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lVerDetalle 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ver detalle"
      Height          =   270
      Left            =   2805
      TabIndex        =   4
      Top             =   5805
      Width           =   1230
   End
   Begin VB.Label lMarcaTodos 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Marcar Todos"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5610
      TabIndex        =   3
      Top             =   195
      Width           =   990
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000003&
      BackStyle       =   1  'Opaque
      Height          =   6585
      Left            =   75
      Top             =   75
      Width           =   6645
   End
End
Attribute VB_Name = "frBolEmit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSBOLE As New ADODB.Recordset
Dim REGACT As REGWIN
Dim ITSOPEN As Boolean
Dim WithEvents RSLISTA As ADODB.Recordset
Attribute RSLISTA.VB_VarHelpID = -1
Dim XCONTINUA As Boolean
Dim XFLAG As Boolean
Dim xFechaPago As String
Dim xMes
Dim SNOMBOL As String
Private Sub CMBILLETES_Click()
    VPTAREA = "BOLETAS"
    frmBilletes.Show 1
End Sub
Private Sub cmCerrar_Click()
    Unload Me
End Sub

Private Sub CMDDETA_Click()
    If RSLISTA.RecordCount = 0 Then
        MsgBox "No existen registros para imprimir", vbExclamation
        Exit Sub
    End If
    ARMARCONSULTA
End Sub
Private Sub ARMARCONSULTA()
'---------------PARA MAÑANA--------------------
    CambiaPanelBD True
    Dim RUTPAGOSCTA As String, RUTMOVICTA As String
    Dim RUTCTAGRUPO As String, RUTNOMBOL As String
    Dim ANNO As String, RUTADEL As String
    Dim RUTBOLETA As String, RUTMOVI As String
    Dim INTO As String, RUTCONC As String
    
    ANNO = Left(xMeses.Text, 2) & Right(xMeses.Text, 4)

    Screen.MousePointer = 11
    DBSTARPLAN.Execute "EXECUTE SP_ARMARCONSULTA '" & REGSISTEMA.BASESQL & "', '" & ANNO & "','" & VGL_COMPUTER & "'"
   
   Dim RSCAMP As New ADODB.Recordset
   
   RSCAMP.Open "SELECT DISTINCT LTRIM( [##TMPDETAGROUP" & VGL_COMPUTER & "] .CODCONCEP) AS CODCONCEP,  [##TMPDETAGROUP" & VGL_COMPUTER & "] .NOMCONCEP  " & _
                " FROM  [##TMPDETAGROUP" & VGL_COMPUTER & "]  ", DBSTARPLAN, adOpenKeyset, adLockOptimistic

   Set FrmDetalle.DCmcampo.RowSource = RSCAMP
   FrmDetalle.DCmcampo.ListField = "NOMCONCEP"
   FrmDetalle.DCmcampo.BoundColumn = "CODCONCEP"
   FrmDetalle.DCmcampo.BoundText = "REMENS"

   CambiaPanelBD False
   FrmDetalle.Show 1
   Set RSCAMP = Nothing
End Sub

Private Sub CMDPLAN_Click()
    Dim RSPLAN As New ADODB.Recordset
    Dim RSAUX As New ADODB.Recordset
    Dim TIPO As Byte, I As Integer
    Screen.MousePointer = 11
    XCONTINUA = False
    CambiaPanelBD True
    Call CARGAPLAN
    CambiaPanelBD False
    If Not XCONTINUA Then
        Screen.MousePointer = 1
        Exit Sub
    End If
On Error GoTo ERRPLAN
    RSPLAN.Open " [##PLAN2000" & VGL_COMPUTER & "] ", DBSTARPLAN, adOpenKeyset, adLockReadOnly
    If RSPLAN.RecordCount = 0 Then Exit Sub
    CambiaPanelBD True
    'CREAR TEMPORAL
    RSPLAN.MoveFirst
    If ExisteTablaAux(" [##TMPLAN" & VGL_COMPUTER & "] ") Then DBSTARPLAN.Execute "DROP TABLE  [##TMPLAN" & VGL_COMPUTER & "] "
    DBSTARPLAN.Execute _
    "CREATE TABLE  [##TMPLAN" & VGL_COMPUTER & "]  (CLAVE INT IDENTITY (1,1) , MES DATETIME, TIPOPLANILLA INT, INUMBOL INT, CODTRAB VARCHAR(8), " & _
    "NOMBRES VARCHAR(100),CARGO VARCHAR(40),BASICO  Numeric(20,2) ,ING  Numeric(20,2) ,CONING VARCHAR(50)," & _
    "EGR  Numeric(20,2) ,CONEGR VARCHAR(50),APO  Numeric(20,2) ,CONAPO VARCHAR(50),FECHING DATETIME,CCOSTO VARCHAR(15),AFP VARCHAR(30)) "
    DBSTARPLAN.Execute " DELETE FROM  [##TMPLAN" & VGL_COMPUTER & "] "
    RSAUX.Open " [##TMPLAN" & VGL_COMPUTER & "] ", DBSTARPLAN, adOpenKeyset, adLockOptimistic
    Dim CONC As String
    With RSPLAN
        Do While Not .EOF
            For I = 22 To .Fields.Count - 1
                TIPO = DevuelveValor("SELECT TIPO FROM COLUMPL WHERE CODIGO='" & Trim(.Fields(I).Name) & "'", DBSYSTEM)
                CONC = DevuelveValor("SELECT NOMBRE FROM COLUMPL WHERE CODIGO='" & Trim(.Fields(I).Name) & "'", DBSYSTEM)
                If TIPO = 2 And UCase(.Fields(I).Name) <> "TOTEGR" And UCase(.Fields(I).Name) <> "TOTING" Then
                    If .Fields(I).Value > 0 Then
                        DBSTARPLAN.Execute "INSERT INTO  [##TMPLAN" & VGL_COMPUTER & "]  (MES, TIPOPLANILLA, INUMBOL, CODTRAB, NOMBRES, CARGO, BASICO, FECHING, CCOSTO, AFP, ING, CONING) VALUES " & _
                        "(" & DateSQL(RSPLAN!MES) & "," & IIf(RSPLAN!TIPOPLANILLA = True, 1, 0) & ", " & RSPLAN!INUMBOL & ", '" & RSPLAN!CODTRAB & "', '" & RSPLAN!NOMBRES & "', '" & RSPLAN!CARGO & "', " & RSPLAN!BASICO & ", " & DateSQL(RSPLAN!FECHAING) & ", '" & _
                        RSPLAN!CCosto & "', '" & RSPLAN!FONDOPENS & "', " & RSPLAN.Fields(I).Value & ", '" & CONC & "')"
                        'RSAUX.AddNew
                        'Call LLENARRS(RSAUX, RSPLAN)
                        'RSAUX!Ing = RSPLAN.Fields(I).Value
                        'RSAUX!CONING = CONC
                        'RSAUX.Update
                    End If
                End If
            Next I
            Set RSAUX = Nothing
            RSAUX.Open " [##TMPLAN" & VGL_COMPUTER & "] ", DBSTARPLAN, adOpenKeyset, adLockOptimistic
            RSAUX.Filter = "CODTRAB='" & RSPLAN!CODTRAB & "'"
            RSAUX.MoveFirst
            Do While Not RSAUX.EOF
                For I = 22 To .Fields.Count - 1
                    TIPO = DevuelveValor("SELECT TIPO FROM COLUMPL WHERE CODIGO='" & Trim(.Fields(I).Name) & "'", DBSYSTEM)
                    CONC = DevuelveValor("SELECT NOMBRE FROM COLUMPL WHERE CODIGO='" & Trim(.Fields(I).Name) & "'", DBSYSTEM)
                    If TIPO = 3 And UCase(.Fields(I).Name) <> "TOTEGR" And UCase(.Fields(I).Name) <> "TOTING" Then
                        If .Fields(I).Value > 0 Then
                            If RSAUX.EOF And Not EXISTECAM(RSPLAN!CODTRAB, "CONEGR", CONC) Then
                                DBSTARPLAN.Execute "INSERT INTO  [##TMPLAN" & VGL_COMPUTER & "]  (MES, TIPOPLANILLA, INUMBOL, CODTRAB, NOMBRES, CARGO, BASICO, FECHING, CCOSTO, AFP, EGR , CONEGR) VALUES " & _
                                "(" & DateSQL(RSPLAN!MES) & "," & IIf(RSPLAN!TIPOPLANILLA = True, 1, 0) & ", " & RSPLAN!INUMBOL & ", '" & RSPLAN!CODTRAB & "', '" & RSPLAN!NOMBRES & "', '" & RSPLAN!CARGO & "', " & RSPLAN!BASICO & ", " & DateSQL(RSPLAN!FECHAING) & ", '" & _
                                RSPLAN!CCosto & "', '" & RSPLAN!FONDOPENS & "', " & RSPLAN.Fields(I).Value & ", '" & CONC & "')"
                                Set RSAUX = Nothing
                                RSAUX.Open " [##TMPLAN" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic
                                RSAUX.MoveLast: RSAUX.MoveNext
                               Else:
                               If Not EXISTECAM(RSPLAN!CODTRAB, "CONEGR", CONC) Then
                                    RSAUX!EGR = RSPLAN.Fields(I).Value
                                    RSAUX!CONEGR = CONC
                                    RSAUX.Update
                               End If
                               If Not RSAUX.EOF Then RSAUX.MoveNext
                            End If
                         End If
                    End If
                Next
              If Not RSAUX.EOF Then RSAUX.MoveNext
            Loop
            Set RSAUX = Nothing
            RSAUX.Open " [##TMPLAN" & VGL_COMPUTER & "] ", DBSTARPLAN, adOpenKeyset, adLockOptimistic
            RSAUX.Filter = "CODTRAB='" & RSPLAN!CODTRAB & "'"
            RSAUX.MoveFirst
            Do While Not RSAUX.EOF
                For I = 22 To .Fields.Count - 1
                    TIPO = DevuelveValor("SELECT TIPO FROM COLUMPL WHERE CODIGO='" & Trim(.Fields(I).Name) & "'", DBSYSTEM)
                    CONC = DevuelveValor("SELECT NOMBRE FROM COLUMPL WHERE CODIGO='" & Trim(.Fields(I).Name) & "'", DBSYSTEM)
                    If TIPO = 4 And UCase(.Fields(I).Name) <> "TOTEGR" And UCase(.Fields(I).Name) <> "TOTING" Then
                        If .Fields(I).Value > 0 Then
                            If RSAUX.EOF And Not EXISTECAM(RSPLAN!CODTRAB, "CONAPO", CONC) Then
                                DBSTARPLAN.Execute "INSERT INTO  [##TMPLAN" & VGL_COMPUTER & "]  (MES, TIPOPLANILLA, INUMBOL, CODTRAB, NOMBRES, CARGO, BASICO, FECHING, CCOSTO, AFP, APO, CONAPO) VALUES " & _
                                "(" & DateSQL(RSPLAN!MES) & "," & IIf(RSPLAN!TIPOPLANILLA = True, 1, 0) & ", " & RSPLAN!INUMBOL & ", '" & RSPLAN!CODTRAB & "', '" & RSPLAN!NOMBRES & "', '" & RSPLAN!CARGO & "', " & RSPLAN!BASICO & ", " & DateSQL(RSPLAN!FECHAING) & ", '" & _
                                RSPLAN!CCosto & "', '" & RSPLAN!FONDOPENS & "', " & RSPLAN.Fields(I).Value & ", '" & CONC & "')"
                                Set RSAUX = Nothing
                                RSAUX.Open " [##TMPLAN" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic
                                RSAUX.MoveLast: RSAUX.MoveNext
                               Else:
                               If Not EXISTECAM(RSPLAN!CODTRAB, "CONAPO", CONC) Then
                                    RSAUX!APO = RSPLAN.Fields(I).Value
                                    RSAUX!CONAPO = CONC
                                    RSAUX.Update
                               End If
                               If Not RSAUX.EOF Then RSAUX.MoveNext
                            End If
                        End If
                    End If
                Next
               If Not RSAUX.EOF Then RSAUX.MoveNext
            Loop
            .MoveNext
        Loop
    End With
    DBSTARPLAN.Execute "UPDATE  [##TMPLAN" & VGL_COMPUTER & "]  SET ING=0 WHERE ING IS NULL "
    DBSTARPLAN.Execute "UPDATE  [##TMPLAN" & VGL_COMPUTER & "]  SET EGR=0 WHERE EGR IS NULL "
    DBSTARPLAN.Execute "UPDATE  [##TMPLAN" & VGL_COMPUTER & "]  SET APO=0 WHERE APO IS NULL "
     With rptBoletas
        .Reset
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=" & VGL_BASE & ""
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0051.RPT"
        .StoredProcParam(0) = " [##TMPLAN" & VGL_COMPUTER & "] "
        .Destination = crptToWindow
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowState = crptMaximized
        .WindowTitle = .ReportFileName
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "XPERIODO='" & UCase(Lista.SelectedItem.Text) & "'"
        CambiaPanelBD False
        If .Status <> 2 Then .PrintReport
    End With
    Screen.MousePointer = 1
    Exit Sub
ERRPLAN:
    MsgBox "NO SE ENCUENTRAN INGRESOS DE ALGUNOS TRABAJADORES POR FAVOR REVISE SU DATA " & Chr(13) & "O " & _
           "LOS TIPOS DE COLUMNAS EN SU PLANILLA DE SEGURO NO HAY COLUMNAS DE TIPO INGRESO", vbInformation
           Screen.MousePointer = 1
    Resume Next
End Sub
Private Sub LLENARRS(RS As ADODB.Recordset, RS2 As ADODB.Recordset)
Dim F1 As Date, F2 As Date
F1 = RS2!MES
F2 = RS2!FECHAING
    With RS
        !MES = DateSQL(F1)
        !TIPOPLANILLA = RS2!TIPOPLANILLA
        !INUMBOL = RS2!INUMBOL
        !CODTRAB = RS2!CODTRAB
        !NOMBRES = RS2!NOMBRES
        !CARGO = RS2!CARGO
        !BASICO = RS2!BASICO
        !feching = DateSQL(F2)
        !CCosto = RS2!CCosto
        !AFP = RS2!FONDOPENS
    End With
End Sub
Private Function EXISTECAM(CODTRAB As String, CAMPO As String, VALOR As String)
    Dim RS As New ADODB.Recordset
    EXISTECAM = False
    RS.Open "SELECT * FROM  [##TMPLAN" & VGL_COMPUTER & "]  WHERE CODTRAB='" & Trim(CODTRAB) & "' AND " & Trim(CAMPO) & "='" & _
    Trim(VALOR) & "'", DBSYSTEM, adOpenKeyset, adLockReadOnly
    If RS.RecordCount > 0 Then
        EXISTECAM = True
        Exit Function
    End If
End Function
Public Sub CMFORMATOPLANILLA_Click()
'On Error Resume Next
    
    SNOMBOL = Right(Lista.SelectedItem.KEY, Len(Lista.SelectedItem.KEY) - 1)
    If Lista.ListItems.Count = 0 Then
        MsgBox "NO EXISTEN REGISTROS", vbInformation
        Exit Sub
    End If
    Dim FMES As Date
    FMES = DevuelveValor("SELECT MES FROM NOMBOL WHERE CODIGO=" & SNOMBOL, DBSYSTEM)
    If Not ExisteTabla("BOL" & Format(Month(FMES), "00") & Year(FMES)) Then
        MsgBox "NO EXISTE LA TABLA CORRESPONDIENTE AL MES ESPECIFICADO, SI LO HA REGISTRADO ANTERIORMENTE ENTONCES DEBERÁ CARGAR LOS DATOS DESDE EL ALMACEN DE DATOS DE PLANILLAS. CONSULTE AL ADMINISTRADOR", vbInformation
        Exit Sub
    End If
    CambiaPanelBD True
    Load FrmOpPlanG
    FrmOpPlanG.Caption = "PLANILLA " & Lista.SelectedItem.Text
    CambiaPanelBD False
    FrmOpPlanG.Show 1
    If Not FrmOpPlanG.Aceptar Then Exit Sub
    
    'IF MSGBOX("DESEA VISUALIZAR LA PLANILLA DE: " & LISTA.SELECTEDITEM.TEXT, VBYESNO + VBQUESTION) = VBNO THEN EXIT SUB
    Screen.MousePointer = 11
    Dim VMES As Date
    Dim RSMESES As New ADODB.Recordset
    RSMESES.Open "EMPRESA", DBSYSTEM, adOpenStatic, adLockReadOnly
    If RSMESES.RecordCount = 0 Then
        MsgBox "SE HA ENCONTRADO UN PROBLEMA EN LA DEFINICIÓN DE LA TABLA EMPRESA", vbCritical
        Set RSMESES = Nothing
        Exit Sub
    Else
        REGSISTEMA.COLPLANADEL = RSMESES!ADELPLAN
    End If
    Dim xFilePlanilla As String
    Dim XMAL As Boolean
    XCONTINUA = False
    CambiaPanelBD True
    Call CARGAPLAN(XMAL)
    CambiaPanelBD False
    If Not XCONTINUA Then
        Set RSMESES = Nothing
        Exit Sub
    End If
    If XMAL Then Exit Sub
    With rptBoletas
        .Reset
        xFilePlanilla = DevNomRep(Trim(REGSISTEMA.RUC), Trim(REGSISTEMA.USER), FILEPLANILLA)
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=" & VGL_BASE & ""
        If UCase(Dir$(REGSISTEMA.REPORTES & xFilePlanilla)) <> UCase(xFilePlanilla) Then xFilePlanilla = "PLAN0032.RPT"
        .ReportFileName = REGSISTEMA.REPORTES & "\" & xFilePlanilla
        .StoredProcParam(0) = " [##PLAN2000" & VGL_COMPUTER & "] "
        .Destination = crptToWindow
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowState = crptMaximized
        .WindowTitle = .ReportFileName
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "XRUC='RUC N° " & REGSISTEMA.RUC & "'"
        .Formulas(2) = "XMES='" & IIf(xMes = "", Lista.SelectedItem.Text, xMes) & "'"
        .Formulas(3) = "XDIRECCION='" & REGSISTEMA.DIRECCION & "'"
        If .Status <> 2 Then .Action = 1
    End With
    Screen.MousePointer = 1
End Sub
Private Sub CARGAPLAN(Optional ByRef VALOR As Boolean)
    If Not COMPRUEBAPLAN Then Exit Sub
    Dim REGPLAN As TYPEREGPLAN
    Dim VARCODE As Long
    Dim SNOMBOL As String
    
    SNOMBOL = Right(Lista.SelectedItem.KEY, Len(Lista.SelectedItem.KEY) - 1)
    Dim FMES As Date
    FMES = DevuelveValor("SELECT MES FROM NOMBOL WHERE CODIGO=" & SNOMBOL, DBSYSTEM)
    VALOR = True
    With REGPLAN
        .AUTOR = REGSISTEMA.USER
        .DATABASE = "MASTER" 'SIGNIFICA QUE ES LA QUE ESTÁ ACTIVA
        .FECHA = Date
        .MES = CDate("01/" & Format(Month(FMES), "00") & "/" & Year(FMES))
        .TABLABOL = "BOL" & Format(Month(.MES), "00") & Year(.MES)
        .TABLAMOV = "MOV" & Format(Month(.MES), "00") & Year(.MES)
    End With
    Dim STRCAD As String
    STRCAD = "CREATE TABLE  [##PLAN2000" & VGL_COMPUTER & "]  (MES DATETIME, TIPOPLANILLA INT, INUMBOL INT, CODTRAB VARCHAR(8), NOMBRES VARCHAR(100), TIPOTRAB VARCHAR(2), FECHAING DATETIME, SITUACION VARCHAR(2), CCOSTO VARCHAR(10), CENTROCOSTO VARCHAR(25), DEPARTAMENTO VARCHAR(25), CARGO VARCHAR(25), BASICO  Numeric(20,2) , FONDOPENS VARCHAR(2), FECHACESE DATETIME, CODSCTR VARCHAR(6), EPS VARCHAR(8), CARNETSEG VARCHAR(15), VACINI DATETIME, VACFIN DATETIME, CUSPP VARCHAR(12), REDONDEO  Numeric(20,2),CODAREA VARCHAR(10),NOMAREA VARCHAR(50)"
    Dim RSAUX As New ADODB.Recordset
    RSAUX.Open "SELECT * FROM COLUMPL ORDER BY INDICE", DBSYSTEM, adOpenStatic
    Do While Not RSAUX.EOF
        STRCAD = STRCAD & ", " & RSAUX!Codigo & "  Numeric(20,2) "
        RSAUX.MoveNext
    Loop
    STRCAD = STRCAD & ")"
    If ExisteTablaAux(" [##PLAN2000" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##PLAN2000" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute STRCAD
    
    Dim RSBOLS As New ADODB.Recordset
    Dim RSMOVS As New ADODB.Recordset
    Dim RSPLAN2 As New ADODB.Recordset
    Dim VCAMP As String
    
    VCAMP = ""
    If FrmOpPlanG.Aceptar Then
        If FrmOpPlanG.OP = 0 Then VCAMP = " AND A.NOPDT=0"
        If FrmOpPlanG.OP = 1 Then VCAMP = " AND A.NOPDT=1"
    End If
    '/*
    Dim xLista As String
    Dim XITEM As ListItem
    Dim xCont As Integer
    xCont = 0
    For Each XITEM In Lista.ListItems
        If XITEM.Checked Then
            xLista = xLista & "" & Right(XITEM.KEY, Len(XITEM.KEY) - 1) & ","
            xCont = xCont + 1
        End If
    Next
    If Trim(xLista) = "" Then
        xLista = Right(Lista.SelectedItem.KEY, Len(Lista.SelectedItem.KEY) - 1)
      Else
        xLista = Left(xLista, Len(xLista) - 1)
    End If
    xLista = "(" & xLista & ")"
    xMes = ""
    If xCont > 1 Then
        xMes = "DEL MES DE " & DESMES(Month(FMES)) & " DEL " & Year(FMES)
    End If
    RSBOLS.Open "SELECT CODNOMBOL, A.CODTRAB, A.NOMBRES, INUMBOL, TIPOPLAN, TOTING, TOTEGR, A.TIPOTRAB, A.FECHAING, SITUACIÓN, BOL.CCOSTO, A.CENTRO, A.DEPARTAMENTO, A.CARGO, BOL.BASICO, BOL.CODAFP, A.FECHACESE, A.CODSCTR, A.RUCEPS, BOL.XREDONDEO,A.CODAREA,A.NOMBREAREA FROM VWTRABAJ A, " & REGPLAN.TABLABOL & " BOL WHERE BOL.CODTRAB=A.CODTRAB AND BOL.CODNOMBOL IN (SELECT CODIGO FROM NOMBOL WHERE MES=" & DateSQL(REGPLAN.MES) & " ) AND CODNOMBOL IN " & xLista & VCAMP & "  ORDER BY NOMBRES", DBSYSTEM, adOpenStatic
        
    RSPLAN2.Open " [##PLAN2000" & VGL_COMPUTER & "] ", DBSTARPLAN, adOpenDynamic, adLockOptimistic
    
    Do While Not RSBOLS.EOF
        RSPLAN2.AddNew
        RSPLAN2!MES = REGPLAN.MES
        RSPLAN2!TIPOPLANILLA = RSBOLS!TIPOPLAN
        RSPLAN2!CODTRAB = RSBOLS!CODTRAB
        RSPLAN2!NOMBRES = Trim(Left(RSBOLS!NOMBRES & String(35, " "), 35))
        RSPLAN2!TIPOTRAB = RSBOLS!TIPOTRAB
        RSPLAN2!FECHAING = CDate(RSBOLS!FECHAING)
        RSPLAN2!SITUACION = RSBOLS!SITUACIÓN
        RSPLAN2!CCosto = Trim(RSBOLS!CCosto)
        RSPLAN2!CENTROCOSTO = RSBOLS!CENTRO
        RSPLAN2!DEPARTAMENTO = RSBOLS!DEPARTAMENTO
        RSPLAN2!CARGO = RSBOLS!CARGO
        RSPLAN2!CODAREA = Left(RSBOLS!CODAREA, 2)
        RSPLAN2!NOMAREA = DevuelveValor("Select NOMBRE From dbo.AREASTRAB Where codccosto='" & Left(RSBOLS!CODAREA, 2) & "'", DBSYSTEM)
        RSPLAN2!BASICO = RSBOLS!BASICO
        RSPLAN2!FONDOPENS = RSBOLS!CODAFP
        If Not IsNull(RSBOLS!FECHACESE) Then RSPLAN2!FECHACESE = CDate(RSBOLS!FECHACESE)
        RSPLAN2!CODSCTR = RSBOLS!CODSCTR
        RSPLAN2!EPS = RSBOLS!RUCEPS
        RSPLAN2!INUMBOL = RSBOLS!INUMBOL
        RSPLAN2!CARNETSEG = "" & DevuelveValor("SELECT CARNETSEG FROM TRABAJADORES WHERE CODTRAB='" & RSBOLS!CODTRAB & "'", DBSYSTEM)
        RSPLAN2!CUSPP = "" & DevuelveValor("SELECT CUSPP FROM TRABAJADORES WHERE CODTRAB='" & RSBOLS!CODTRAB & "'", DBSYSTEM)
        RSPLAN2!INUMBOL = RSBOLS!INUMBOL
        If Not IsNull(DevuelveValor("SELECT CODIGO FROM HISTOVAC WHERE CERRADO=1 AND NOMBOL=" & RSBOLS!CODNOMBOL & " AND CODTRAB='" & RSBOLS!CODTRAB & "'", DBSYSTEM)) Then
            VARCODE = DevuelveValor("SELECT CODIGO FROM HISTOVAC WHERE CERRADO=1 AND NOMBOL=" & RSBOLS!CODNOMBOL & " AND CODTRAB='" & RSBOLS!CODTRAB & "'", DBSYSTEM)
            If VARCODE <> 0 Then
                RSPLAN2!VACINI = DevuelveValor("SELECT FECHAINI FROM HISTOVAC WHERE CODIGO=" & VARCODE, DBSYSTEM)
                RSPLAN2!VACFIN = DevuelveValor("SELECT FECHAFIN FROM HISTOVAC WHERE CODIGO=" & VARCODE, DBSYSTEM)
            End If
        End If
        RSPLAN2.Update
        
        DBSTARPLAN.Execute "UPDATE  [##PLAN2000" & VGL_COMPUTER & "]  SET TOTING=ISNULL(TOTING,0)+" & IIf(IsNull(RSBOLS!TOTING), 0, RSBOLS!TOTING) & ",TOTEGR=ISNULL(TOTEGR,0)+" & IIf(IsNull(RSBOLS!TOTEGR), 0, RSBOLS!TOTEGR) & ",REDONDEO=ISNULL(REDONDEO,0)+" & IIf(IsNull(RSBOLS!XREDONDEO), 0, RSBOLS!XREDONDEO) & " WHERE INUMBOL=" & RSBOLS!INUMBOL & " AND MES=" & DateSQL(REGPLAN.MES)
        
        '/*
        RSMOVS.Open "SELECT COLPLANILLA, MONTO FROM " & REGPLAN.TABLAMOV & " MOV, CONCEPTOS WHERE MOV.CONCEPTO=CONCEPTOS.CODIGO AND INUMBOL=" & RSBOLS!INUMBOL, DBSYSTEM, adOpenStatic
        
        Do While Not RSMOVS.EOF
            If Trim(RSMOVS!COLPLANILLA) <> "" Then DBSTARPLAN.Execute "UPDATE  [##PLAN2000" & VGL_COMPUTER & "]  SET " & Trim$(RSMOVS!COLPLANILLA) & "=ISNULL(" & Trim$(RSMOVS!COLPLANILLA) & ",0)+" & RSMOVS!MONTO & " WHERE INUMBOL=" & RSBOLS!INUMBOL & " AND MES=" & DateSQL(REGPLAN.MES)
            RSMOVS.MoveNext
        Loop
        RSMOVS.Close
        RSBOLS.MoveNext
    Loop
    Set RSMOVS = Nothing
    'VALIDAR CAMPOS
    If Not (ExisteCampo("NETO", " [##PLAN2000" & VGL_COMPUTER & "] ", DBSYSTEM) Or ExisteCampo("NETOPAGO", "PLAN2000", DBSYSTEM)) Then
        MsgBox "TIENE CREAR EL CAMPO NETO EN COLUMNAS DE PLANILLA", vbExclamation
        Exit Sub
    End If
    If Not ExisteCampo("TOTING", " [##PLAN2000" & VGL_COMPUTER & "] ", DBSYSTEM) Then
        MsgBox "TIENE CREAR EL CAMPO TOTING EN COLUMNAS DE PLANILLA", vbExclamation
        Exit Sub
    End If
    If Not ExisteCampo("TOTEGR", " [##PLAN2000" & VGL_COMPUTER & "] ", DBSYSTEM) Then
        MsgBox "TIENE CREAR EL CAMPO TOTEGR EN COLUMNAS DE PLANILLA", vbExclamation
        Exit Sub
    End If
    If ExisteCampo("NETO", " [##PLAN2000" & VGL_COMPUTER & "] ", DBSYSTEM) Then
        DBSYSTEM.Execute "UPDATE  [##PLAN2000" & VGL_COMPUTER & "]  SET NETO=TOTING-TOTEGR WHERE MES=" & DateSQL(REGPLAN.MES)
    Else
        DBSYSTEM.Execute "UPDATE  [##PLAN2000" & VGL_COMPUTER & "]  SET NETOPAGO=TOTING-TOTEGR WHERE MES=" & DateSQL(REGPLAN.MES)
    End If
    
    '---------------------------------------------------------------
    'COLOCANDO CERO A LOS VALORES NULOS
    RSBOLS.Close
    RSBOLS.Open "COLUMPL", DBSYSTEM, adOpenStatic
    Do While Not RSBOLS.EOF
        DBSTARPLAN.Execute "UPDATE [##" & REGSISTEMA.TABLAPLAN & VGL_COMPUTER & "]   SET " & RSBOLS!Codigo & " =0 WHERE " & RSBOLS!Codigo & " IS NULL"
        RSBOLS.MoveNext
    Loop
    '---------------------------------------------------------------
    'ASIGNACIÓN DE LOS ADELANTOS DE PAGO
    RSBOLS.Close
    RSBOLS.Open "SELECT BOL.INUMBOL, MONTO FROM " & REGPLAN.TABLABOL & " BOL, " & REGSISTEMA.TABLAADEL & " ADEL WHERE BOL.CODTRAB=ADEL.CODTRAB AND ADEL.ORIGEN IN (SELECT CODIGO FROM NOMBOL WHERE MES=" & DateSQL(REGPLAN.MES) & " AND CODIGO IN " & xLista & ")", DBSYSTEM, adOpenStatic
    Do While Not RSBOLS.EOF
        DBSTARPLAN.Execute "UPDATE [##" & REGSISTEMA.TABLAPLAN & VGL_COMPUTER & "] SET " & REGSISTEMA.COLPLANADEL & "= " & REGSISTEMA.COLPLANADEL & "+" & RSBOLS!MONTO & " WHERE INUMBOL=" & RSBOLS!INUMBOL & " AND MES=" & DateSQL(REGPLAN.MES)
        RSBOLS.MoveNext
    Loop
    '---------------------------------------------------------------
    'ASIGNANDO LOS VALORES DE CUENTAS CORRIENTES
    RSBOLS.Close
    RSBOLS.Open "SELECT PAGOSCTA.CODTRAB, TIPOPLAN, PLANILLA, BOL.INUMBOL, PAGOSCTA.MONTO FROM " & REGPLAN.TABLABOL & " BOL, PAGOSCTA, MOVICTA, CTAGRUPO WHERE BOL.CODTRAB=PAGOSCTA.CODTRAB AND PAGOSCTA.CODMOV=MOVICTA.CODMOV AND MOVICTA.CODGRUPO=CTAGRUPO.CODGRUPO AND PAGOSCTA.CODNOMBOL IN (SELECT CODIGO FROM NOMBOL WHERE MES=" & DateSQL(REGPLAN.MES) & " AND NOMBOL.CODIGO IN " & xLista & ") AND TIPOBOLETA='B' ", DBSYSTEM, adOpenStatic
    Do While Not RSBOLS.EOF
        If ExisteCampo(RSBOLS!PLANILLA, " [##PLAN2000" & VGL_COMPUTER & "] ", DBSYSTEM) Then
            DBSTARPLAN.Execute "UPDATE [##" & REGSISTEMA.TABLAPLAN & VGL_COMPUTER & "] SET " & RSBOLS!PLANILLA & "=" & RSBOLS!PLANILLA & "+" & RSBOLS!MONTO & " WHERE INUMBOL=" & RSBOLS!INUMBOL & " AND MES=" & DateSQL(REGPLAN.MES)
        Else
            MsgBox "NO EXISTE EL CAMPO DE PLANILLA " & RSBOLS!PLANILLA & ". ERROR EN LA CONFIGURACIÓN DE CUENTAS CORRIENTES. NO SE HAN CARGADO LOS DATOS", vbInformation
        End If
        RSBOLS.MoveNext
    Loop
    '---------------------------------------------------------------
    'TAREA CUMPLIDA, CERRANDO TABLAS
    Set RSBOLS = Nothing
    Set RSPLAN2 = Nothing
    Set RSMOVS = Nothing
    VALOR = False
    XCONTINUA = True
End Sub
Private Sub CMPAGOSBANCO_Click()
'MequedeAhijctr
    If RSLISTA.RecordCount = 0 Then
        MsgBox "NO EXISTEN REGISTROS A PROCESAR. FALTA SELECCIONAR UN PERIODO DE PAGO CONTENIENDO BOLETAS DE REMUNERACIONES", vbCritical
        Exit Sub
    End If
    CambiaPanelBD True
    If ExisteTablaAux(" [##TMPBANCOS" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPBANCOS" & VGL_COMPUTER & "] "
    If ExisteTablaAux(" [##PAGOSXBANCO" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##PAGOSXBANCO" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "SELECT CODTRAB, TIPDOC, DOCIDEN, CTABANCO, BANCO INTO  [##TMPBANCOS" & VGL_COMPUTER & "]  FROM " & REGSISTEMA.BASESQL & ".dbo.TRABAJADORES WHERE CODTRAB IN (SELECT CODTRAB FROM  [##_TMPLSTBOL" & VGL_COMPUTER & "] )"
    DBSYSTEM.Execute "SELECT A.CODTRAB, NOMBRES, NETO, TIPDOC, DOCIDEN, CTABANCO, BANCO INTO  [##PAGOSXBANCO" & VGL_COMPUTER & "]  FROM  [##_TMPLSTBOL" & VGL_COMPUTER & "]  A,  [##TMPBANCOS" & VGL_COMPUTER & "] ##TMPBANCOS  WHERE A.CODTRAB=##TMPBANCOS.CODTRAB"
    CambiaPanelBD False
    frPagoBco.Show 1
End Sub

Private Sub CMPLANILLA_Click()
    If RSLISTA.RecordCount = 0 Then
        MsgBox "NO EXISTEN REGISTROS SELECCIONADOS. LA LISTA SE ENCUENTRA VACIA, SELECCIONE UNO O MAS PERIODOS DE PAGOS Y QUE HAYAN SIDO PROCESADOS", vbCritical
        Exit Sub
    End If
    Dim XITEM As ListItem, CADCH As String, SNOMBOL As String
    CADCH = "("
    For Each XITEM In Lista.ListItems
        If XITEM.Checked Then
            SNOMBOL = Right(XITEM.KEY, Len(XITEM.KEY) - 1)
            CADCH = CADCH & IIf(CADCH = "(", "", ",") & SNOMBOL
        End If
    Next
    CADCH = CADCH & ")"
    VPTAREA = CADCH
    frPlanBol.Show 1
End Sub

Private Sub CMPRTTODOS_Click()
    Dim XFILEBOL As String, xDir As String
    XFILEBOL = DevNomRep(Trim(REGSISTEMA.RUC), Trim(REGSISTEMA.USER), FILEBOLETA)
    If UCase(Dir$(REGSISTEMA.REPORTES & XFILEBOL)) <> UCase(XFILEBOL) Then
        MsgBox "NO SE HA ENCONTRADO EL REPORTE. ASIGNE CORRECTAMENTE EL NOMBRE DEL REPORTE DE BOLETAS DE REMUNERACIONES", vbInformation, "FALTA: " & XFILEBOL
        Exit Sub
    End If
    DBSTARPLAN.Execute "DELETE FROM RPTBOLETAS"
    If RSLISTA.RecordCount = 0 Or RSLISTA.EOF Then
        MsgBox "NO SE HAN SELECCIONADO BOLETAS O NO EXISTEN EN ESTE NOMBRE DE PLANILLA", vbInformation
        Exit Sub
    End If
    RSLISTA.MoveFirst
    If InStr(XFILEBOL, "XX") > 0 Then
        rptBoletas.Reset
        If ExisteTablaAux(" [##RPTBOLCOLUMNAS" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##RPTBOLCOLUMNAS" & VGL_COMPUTER & "] "
        DBSYSTEM.Execute "SELECT CODTRAB INTO  [##RPTBOLCOLUMNAS" & VGL_COMPUTER & "]  FROM  [##_TMPLSTBOL" & VGL_COMPUTER & "] "
        Dim TMPNOMBRE2 As String
        If Dir$(App.PATH & "\" & XFILEBOL) = XFILEBOL Then Kill App.PATH & "\" & XFILEBOL
        FileCopy REGSISTEMA.REPORTES & XFILEBOL, App.PATH & "\" & XFILEBOL
        TMPNOMBRE2 = Replace(XFILEBOL, "XX", "0")
        If UCase(Dir$(REGSISTEMA.REPORTES & XFILEBOL)) <> UCase(XFILEBOL) Then
            MsgBox "NO EXISTE EL ARCHIVO AUXILIAR PARA LA IMPRESION DE ESTE REPORTE DE BOLETA DE REMUNERACIONES: FALTA " & TMPNOMBRE2, vbInformation
            Exit Sub
        End If
        If Dir$(App.PATH & "\" & TMPNOMBRE2) = TMPNOMBRE2 Then Kill App.PATH & "\" & TMPNOMBRE2
        FileCopy REGSISTEMA.REPORTES & TMPNOMBRE2, App.PATH & "\" & TMPNOMBRE2
    End If
    CambiaPanelBD True
    Do While Not RSLISTA.EOF
        CARGABOL
        RSLISTA.MoveNext
    Loop
    With rptBoletas
        .Reset
        'FRWAIT.SHOW 1
        If InStr(XFILEBOL, "XX") > 0 Then
            .ReportFileName = App.PATH & XFILEBOL
        Else
            .ReportFileName = REGSISTEMA.REPORTES & XFILEBOL
            .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
        End If
        .WindowTitle = "REPORTE DE BOLETAS DE REMUNERACIONES - TODAS LAS BOLETAS: " & .ReportFileName
        .StoredProcParam(0) = REGSISTEMA.BASESQL
'        .StoredProcParam(0) = "RPTBOLETAS"
'        .StoredProcParam(1) = REGSISTEMA.BASESQL & ".dbo.TRABAJADORES"
'        .StoredProcParam(2) = "CODTRAB"
'        .StoredProcParam(3) = "CODTRAB"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "XRUC='RUC N° " & REGSISTEMA.RUC & "'"
        .Formulas(2) = "XDIRECCION='" & xDir & "'"
        .Formulas(3) = "@FechPag='" & Format(xFechaPago, "dd/mm/yyyy") & "'"
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        If rptBoletas.Status <> 2 Then .Action = 1
    End With
    RSLISTA.MoveFirst
    XFLAG = False
    CambiaPanelBD False
    Screen.MousePointer = 1
End Sub

Private Sub CMPRTUNO_Click()
    Dim XFILEBOL As String, xDir As String
    XFILEBOL = DevNomRep(Trim(REGSISTEMA.RUC), Trim(REGSISTEMA.USER), FILEBOLETA)
    If UCase(Dir$(REGSISTEMA.REPORTES & XFILEBOL)) <> UCase(XFILEBOL) Then
        MsgBox "NO SE HA ENCONTRADO EL REPORTE. ASIGNE CORRECTAMENTE EL NOMBRE DEL REPORTE DE BOLETAS DE REMUNERACIONES", vbInformation, "FALTA: " & XFILEBOL
        Exit Sub
    End If
    DBSTARPLAN.Execute "DELETE FROM RPTBOLETAS"
    If RSLISTA.RecordCount = 0 Or RSLISTA.EOF Then
        MsgBox "NO SE HAN SELECCIONADO BOLETAS O NO EXISTEN EN ESTE NOMBRE DE PLANILLA", vbInformation
        Exit Sub
    End If
    Dim XBOOK As Variant
    If InStr(XFILEBOL, "XX") > 0 Then
        On Error Resume Next
        If ExisteTablaAux(" [##RPTBOLCOLUMNAS" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##RPTBOLCOLUMNAS" & VGL_COMPUTER & "] "
        DBSYSTEM.Execute "CREATE TABLE  [##RPTBOLCOLUMNAS" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8))"
        Dim TMPNOMBRE2 As String
        If Dir$(App.PATH & "\" & XFILEBOL) = XFILEBOL Then Kill App.PATH & "\" & XFILEBOL
        FileCopy REGSISTEMA.REPORTES & XFILEBOL, App.PATH & "\" & XFILEBOL
        TMPNOMBRE2 = Replace(XFILEBOL, "X", "0")
        If UCase(Dir$(REGSISTEMA.REPORTES & XFILEBOL)) <> UCase(XFILEBOL) Then
            MsgBox "NO EXISTE EL ARCHIVO AUXILIAR PARA LA IMPRESION DE ESTE REPORTE DE BOLETA DE REMUNERACIONES: FALTA " & TMPNOMBRE2, vbInformation
            Exit Sub
        End If
        If Dir$(App.PATH & "\" & TMPNOMBRE2) = TMPNOMBRE2 Then Kill App.PATH & "\" & TMPNOMBRE2
        FileCopy REGSISTEMA.REPORTES & TMPNOMBRE2, App.PATH & "\" & TMPNOMBRE2
    End If
    CambiaPanelBD True
    For Each XBOOK In dgBoletas.SelBookmarks
        RSLISTA.Bookmark = XBOOK
        CARGABOL
        If InStr(XFILEBOL, "XX") > 0 Then DBSYSTEM.Execute "INSERT INTO  [##RPTBOLCOLUMNAS" & VGL_COMPUTER & "]  VALUES ('" & RSLISTA!CODTRAB & "')"
    Next
    With rptBoletas
        .Reset
        xDir = DevuelveValor("SELECT DIRECCIÓN FROM EMPRESA", DBSYSTEM)
        .WindowTitle = "REPORTE DE BOLETAS DE REMUNERACIONES - SOLO LA BOLETA SELECCIONADA"
        If InStr(XFILEBOL, "XX") > 0 Then
            .ReportFileName = REGSISTEMA.REPORTES & "\" & XFILEBOL
        Else
            .ReportFileName = REGSISTEMA.REPORTES & XFILEBOL
            '.LogOnServer "pdssql.dll", VGL_SERVERREP, "MARFICE_PP", "SOPORTE", "SOPORTE"
            .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=" & VGL_BASE & ""
            
        End If
        .StoredProcParam(0) = REGSISTEMA.BASESQL
'        .StoredProcParam(0) = "RPTBOLETAS"
'        .StoredProcParam(1) = REGSISTEMA.BASESQL & ".dbo.TRABAJADORES"
'        .StoredProcParam(2) = "CODTRAB"
'        .StoredProcParam(3) = "CODTRAB"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "XRUC='RUC N° " & REGSISTEMA.RUC & "'"
        .Formulas(2) = "XDIRECCION='" & xDir & "'"
        .Formulas(3) = "@FechPag='" & Format(xFechaPago, "dd/mm/yyyy") & "'"
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        CambiaPanelBD False
        If rptBoletas.Status <> 2 Then .Action = 1
    End With
    XFLAG = False
End Sub

Private Sub Command1_Click()
    If RSLISTA.RecordCount = 0 Then
        MsgBox "NO HAY REGISTROS PARA IMPRIMIR", vbExclamation
        Exit Sub
    End If
    FrmDeb.Show 1
End Sub

Private Sub Command2_Click()
    If Not Lista.SelectedItem.Checked Then
       MsgBox "Debe de ubicarse y marcar el periodo de pago. Este reporte solo toma de un solo periodo de pago, sobre los registros seleccionados", vbInformation
       Exit Sub
    End If
    frmResumenBolEmit.Show 1
End Sub

Private Sub Command3_Click()

End Sub

Private Sub DGBOLETAS_DblClick()
    XVERDETALLE_Click
End Sub

Private Sub DGBOLETAS_HEADCLICK(ByVal COLINDEX As Integer)
    On Error Resume Next
    Dim XCOL As String
    XCOL = dgBoletas.Columns(COLINDEX).DataField
    If ITSOPEN Then
        RSLISTA.Close
        RSLISTA.Open "SELECT * FROM  [##_TMPLSTBOL" & VGL_COMPUTER & "]  ORDER BY " & XCOL
        Set dgBoletas.DataSource = RSLISTA
        FORMATEARDG
        dgBoletas.Tag = XCOL
    End If
End Sub

Private Sub Form_Activate()
    ActivarTools REGACT
End Sub

Private Sub Form_Load()
    Set RSLISTA = New ADODB.Recordset
    If ExisteTablaAux(" [##_TMPLSTBOL" & VGL_COMPUTER & "] ") Then
        DBSYSTEM.Execute "DELETE FROM  [##_TMPLSTBOL" & VGL_COMPUTER & "] "
    Else
        DBSYSTEM.Execute "CREATE TABLE  [##_TMPLSTBOL" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8) , NOMBRES VARCHAR(50), INGRESOS  Numeric(20,2) , EGRESOS  Numeric(20,2) , NETO  Numeric(20,2) , INUMBOL INT, NOMBOL INT, PERIODO VARCHAR(50),BASICO  Numeric(20,2) )"
    End If
    RSLISTA.Open " [##_TMPLSTBOL" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic
    ITSOPEN = False
    With REGACT
        .BUSCAR = True
        .EDITAR = False
        .ELIMINAR = True
        .FILTRAR = False
        .IMPRIMIR = True
        .NUEVO = False
        .PRELIMINAR = True
    End With
    CARGARMESES
    Me.TOP = 0
    Me.Left = 0
    xVistaMes.ListIndex = 0
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSBOLE = Nothing
    Set RSLISTA = Nothing
End Sub

Private Sub LISTA_ITEMCHECK(ByVal Item As MSComctlLib.ListItem)
    Dim SNOMBOL As String
    SNOMBOL = Right(Item.KEY, Len(Item.KEY) - 1)
    Dim FMES As Date
    FMES = DevuelveValor("SELECT MES FROM NOMBOL WHERE CODIGO=" & SNOMBOL, DBSYSTEM)
    xFechaPago = DevuelveValor("SELECT FECHAPAGO FROM NOMBOL WHERE CODIGO=" & SNOMBOL, DBSYSTEM)
    
    If Not ExisteTabla("BOL" & Format(Month(FMES), "00") & Year(FMES)) Then
        MsgBox "NO EXISTE LA TABLA CORRESPONDIENTE AL MES ESPECIFICADO, SI LO HA REGISTRADO ANTERIORMENTE ENTONCES DEBERÁ CARGAR LOS DATOS DESDE EL ALMACEN DE DATOS DE PLANILLAS. CONSULTE AL ADMINISTRADOR", vbInformation
        Exit Sub
    End If
    If Not ExisteCampo("XREDONDEO", "BOL" & Format(Month(FMES), "00") & Year(FMES), DBSYSTEM) Then
        DBSYSTEM.Execute "ALTER TABLE " & "BOL" & Format(Month(FMES), "00") & Year(FMES) & " ADD XREDONDEO  Numeric(20,2) "
        DBSYSTEM.Execute "UPDATE BOL" & Format(Month(FMES), "00") & Year(FMES) & " SET XREDONDEO=0"
    End If
    CambiaPanelBD True
    If Item.Checked Then
        Screen.MousePointer = 11
        DBSYSTEM.Execute "UPDATE BOL" & Format(Month(FMES), "00") & Year(FMES) & " SET XREDONDEO=0 WHERE (XREDONDEO)IS NULL"
        'If ExisteTablaSQL(" [##_TMPLSTBOL" & VGL_COMPUTER & "] ", DBAUXCOM) Then DBSYSTEM.Execute "DROP TABLE  [##_TMPLSTBOL" & VGL_COMPUTER & "] "
        DBSYSTEM.Execute "INSERT INTO  [##_TMPLSTBOL" & VGL_COMPUTER & "]  SELECT TR.CODTRAB, NOMBRES, TOTING AS INGRESOS, TOTEGR AS EGRESOS, TOTING-TOTEGR+XREDONDEO AS NETO, INUMBOL, CODNOMBOL AS NOMBOL,'" & Item.Text & "' AS PERIODO,BOLS.BASICO   FROM " & REGSISTEMA.BASESQL & ".dbo.VWTRABAJ TR, " & REGSISTEMA.BASESQL & ".dbo.BOL" & Format(Month(FMES), "00") & Year(FMES) & " BOLS  WHERE BOLS.CODTRAB=TR.CODTRAB AND CODNOMBOL=" & SNOMBOL & ""
        Screen.MousePointer = 1
    Else
        DBSYSTEM.Execute "DELETE FROM  [##_TMPLSTBOL" & VGL_COMPUTER & "]  WHERE NOMBOL=" & SNOMBOL
    End If
    Set RSLISTA = Nothing
    Set RSLISTA = New ADODB.Recordset
    RSLISTA.Open " [##_TMPLSTBOL" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic
    'RSLISTA.REQUERY
    Set dgBoletas.DataSource = RSLISTA
    FORMATEARDG
    xArea.Text = ""
    xArea.Tag = ""
    CambiaPanelBD False
End Sub


Public Sub CARGARMESES()
    Dim RSMESES As New ADODB.Recordset
    RSMESES.Open "SELECT DISTINCT MES FROM NOMBOL ORDER BY MES DESC", DBSYSTEM, adOpenKeyset, adLockOptimistic
    If RSMESES.RecordCount = 0 Then
        Set RSMESES = Nothing
        MsgBox "NO EXISTEN MESES PROCESADOS"
        Exit Sub
    End If
    xMeses.Clear
    Do While Not RSMESES.EOF
        xMeses.AddItem Format(Month(RSMESES!MES), "00") & "/" & Year(RSMESES!MES) & " : " & AMESES(Month(RSMESES!MES)) & " DE " & Year(RSMESES!MES)
        RSMESES.MoveNext
    Loop
    Set RSMESES = Nothing
    xMeses.ListIndex = 0
End Sub

Private Sub LMARCATODOS_Click()
    XMARCATODOS_CLICK
End Sub

Private Sub LVERDETALLE_Click()
    XVERDETALLE_Click
End Sub

Private Sub RSLISTA_MOVECOMPLETE(ByVal ADREASON As ADODB.EventReasonEnum, ByVal PERROR As ADODB.Error, ADSTATUS As ADODB.EventStatusEnum, ByVal PRECORDSET As ADODB.Recordset)
    On Error GoTo ERRMOVE
    If RSLISTA.EOF Then
        cmPagosBanco.Enabled = False
        cmBilletes.Enabled = False
        cmPrtUno.Enabled = False
        cmPrtTodos.Enabled = False
        cmPlanilla.Enabled = False
        Command2.Enabled = False
    Else
        cmPagosBanco.Enabled = True
        cmBilletes.Enabled = True
        cmPrtUno.Enabled = True
        cmPrtTodos.Enabled = True
        cmPlanilla.Enabled = True
        Command2.Enabled = True
    End If
    If ADREASON = adRsnMove Then
        If RSLISTA!Neto < 0 Then xError.Visible = True Else xError.Visible = False
    End If
    Exit Sub
ERRMOVE:
    Resume Next
End Sub

Private Sub SEL1_Click(INDEX As Integer)
    xArea.Text = ""
    xArea.Tag = ""
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    'MENU ESPECIAL DE OTROS PROCESOS
    Dim SNOMBOL As String
    SNOMBOL = Right(Lista.SelectedItem.KEY, Len(Lista.SelectedItem.KEY) - 1)
    Dim FMES As Date
    FMES = DevuelveValor("SELECT MES FROM NOMBOL WHERE CODIGO=" & SNOMBOL, DBSYSTEM)
    Select Case UCase(ButtonMenu.KEY)
        Case "GRAFICA10"
            CambiaPanelBD True
            If ExisteTablaAux(" [##TMPGRAF001" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPGRAF001" & VGL_COMPUTER & "] "
            DBSYSTEM.Execute "SELECT TOP 10 WITH TIES NETO, NOMBRES, CODTRAB INTO  [##TMPGRAF001" & VGL_COMPUTER & "]  FROM  [##_TMPLSTBOL" & VGL_COMPUTER & "]  ORDER BY NETO DESC"
            DBSTARPLAN.Execute "EXECUTE [ASISTMP] ' [##TMPGRAF001" & VGL_COMPUTER & "] '"
            With rptBoletas
                .Reset
                .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
                .ReportFileName = REGSISTEMA.REPORTES & "PLAN0062.RPT"
                .StoredProcParam(0) = " [##TMPGRAF001" & VGL_COMPUTER & "] "
                .Destination = crptToWindow
                .WindowShowPrintBtn = True
                .WindowShowRefreshBtn = True
                .WindowShowSearchBtn = True
                .WindowShowPrintSetupBtn = True
                .WindowState = crptMaximized
                .WindowTitle = .ReportFileName
                .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
                .Formulas(1) = "XMES='CORRESPONDIENTE A " & Lista.SelectedItem.Text & "'"
                If .Status <> 2 Then .Action = 1
                CambiaPanelBD False
            End With
        Case "FIRMAR"
            With rptBoletas
                CambiaPanelBD True
                .Reset
                .WindowTitle = "PLAN0063 - REPORTE DE BOLETAS DE REMUNERACIONES - NETOS A PAGAR"
                .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
                .ReportFileName = REGSISTEMA.REPORTES & "PLAN0063.RPT"
                .StoredProcParam(0) = " [##_TMPLSTBOL" & VGL_COMPUTER & "] "
                .SortFields(0) = "+{ASISTMP." & IIf(dgBoletas.Tag <> "", dgBoletas.Tag, "CODTRAB") & "}"
                .Destination = crptToWindow
                .WindowState = crptMaximized
                .WindowShowPrintBtn = True
                .WindowShowRefreshBtn = True
                .WindowShowSearchBtn = True
                .WindowShowPrintSetupBtn = True
                .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
                .Formulas(1) = "XRUC='RUC N° " & REGSISTEMA.RUC & "'"
                .Formulas(2) = "''"
                .Formulas(3) = "SEMA='" & Lista.SelectedItem.Text & "'"
                If .Status <> 2 Then .Action = 1
                CambiaPanelBD False
            End With
        Case "ELIMINATODASBOL"
            If RSLISTA.EOF Or RSLISTA.RecordCount = 0 Then
                MsgBox "NO SE HAN ENCONTRADO BOLETAS PARA ELIMINAR", vbCritical
                Exit Sub
            End If
            If MsgBox("REALMENTE DESEA ELIMINAR TODAS LAS BOLETAS DE REMUNERACIONES. LA ELIMINACIÓN ES PERMANENTE Y NO SE PODRÁ DESHACER LOS CAMBIOS", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
            CambiaPanelBD True
            RSLISTA.MoveFirst
            Screen.MousePointer = 11
            Dim RSAUX As New ADODB.Recordset, NUMV As Integer
            Do While Not RSLISTA.EOF
                FMES = DevuelveValor("SELECT MES FROM NOMBOL WHERE CODIGO=" & RSLISTA!NOMBOL, DBSYSTEM)
                DBSYSTEM.Execute "DELETE FROM MOV" & Format(Month(FMES), "00") & Year(FMES) & " WHERE INUMBOL=" & RSLISTA!INUMBOL
                DBSYSTEM.Execute "DELETE FROM BOL" & Format(Month(FMES), "00") & Year(FMES) & " WHERE INUMBOL=" & RSLISTA!INUMBOL
                DBSYSTEM.Execute "UPDATE " & REGSISTEMA.TABLAADEL & " SET NOMBOL=0, NUMBOL=0 WHERE NOMBOL=" & RSLISTA!NOMBOL & " AND CODTRAB='" & RSLISTA!CODTRAB & "'"
                'Limpieza de las vacaciones
                DBSYSTEM.Execute "UPDATE HISTOVAC SET CERRADO=0 WHERE CODTRAB='" & RSLISTA!CODTRAB & "' AND NOMBOL=" & RSLISTA!NOMBOL
                RSAUX.Open "SELECT CODMOV, MONTO FROM PAGOSCTA WHERE CODNOMBOL=" & RSLISTA!NOMBOL & " AND CODTRAB='" & RSLISTA!CODTRAB & "'", DBSYSTEM, adOpenStatic
                DBSYSTEM.Execute "DELETE FROM PAGOSCTA WHERE CODNOMBOL=" & RSLISTA!NOMBOL & " AND CODTRAB='" & RSLISTA!CODTRAB & "' AND TIPOBOLETA='B'"
                Do While Not RSAUX.EOF 'BORRAMOS LAS ASIGNACIONES ANTERIORES
'                    DBSYSTEM.Execute "UPDATE MOVICTA SET SALDO=SALDO+" & RSAUX!MONTO & " WHERE CODMOV=" & RSAUX!CODMOV, NUMV
'                    If NUMV = 0 Then MsgBox "SE PRODUJO UN ERROR AL INTENTAR CAMBIAR UN PAGO DE CUENTA CORRIENTE DEL TRABAJADOR - SOURCE: INPUTPL, CODE MOVICTA=" & RSAUX!CODMOV, vbCritical
                    Call ACTSALDO(RSAUX!CODMOV)
                    RSAUX.MoveNext
                Loop
                Set RSAUX = Nothing
                RSLISTA.MoveNext
            Loop
            Dim XITEM As ListItem
            For Each XITEM In Lista.ListItems
                If XITEM.Checked Then
                    XITEM.Checked = False
                    LISTA_ITEMCHECK XITEM
                End If
            Next
            Screen.MousePointer = 1
            CambiaPanelBD False
            MsgBox "TODAS LAS BOLETAS DE REMUNERACIONES SE HAN ELIMINADO SATISFACTORIAMENTE. SE HAN RESTABLECIDO LOS ADELANTOS Y DEBITOS DE CUENTAS CORRIENTES PENDIENTES SI LOS HUBIERA EXISTIDO", vbInformation
        Case "TIPOF5"
            If RSLISTA.EOF Or RSLISTA.RecordCount = 0 Then
                MsgBox "NO SE HAN ENCONTRADO BOLETAS PARA FILTRAR POR SELECIÓN DE TRABAJADORES", vbCritical
                Exit Sub
            End If
            CambiaPanelBD True
            Load frSelect
            frSelect.xCondFecha.ListIndex = 0
            CambiaPanelBD False
            frSelect.Show 1
            If VPTRASPRM = "CANCEL" Then Exit Sub
            CambiaPanelBD True
            DBSYSTEM.Execute "DELETE FROM  [##_TMPLSTBOL" & VGL_COMPUTER & "]  WHERE CODTRAB NOT IN (SELECT CODTRAB FROM  [##TMPSELECT" & VGL_COMPUTER & "] )"
            RSLISTA.Requery
            Set dgBoletas.DataSource = RSLISTA
            FORMATEARDG
            CambiaPanelBD False
        Case "QUITARTRAB"
            CambiaPanelBD True
            Dim XT As Variant
            For Each XT In dgBoletas.SelBookmarks
                RSLISTA.Bookmark = XT
                DBSYSTEM.Execute "DELETE FROM  [##_TMPLSTBOL" & VGL_COMPUTER & "]  WHERE INUMBOL=" & RSLISTA!INUMBOL
            Next
            RSLISTA.Requery
            Set dgBoletas.DataSource = RSLISTA
            FORMATEARDG
            CambiaPanelBD False
        Case "RENAMEPERIODO"
            Lista.LabelEdit = lvwAutomatic
            MsgBox "HABILITADO!" & Chr(13) & Chr(10) & "AHORA PUEDE CAMBIAR TEMPORALMENTE EL NOMBRE DE LOS PERIODOS DE PLANILLA", vbInformation
            Lista.SetFocus
            
        Case "REPORTEMENSUAL"
            CambiaPanelBD True
            Dim CAD As String
            Dim rsAuxiliar As ADODB.Recordset
            Dim TOTALREMU As Double, TOTALDESC As Double
            Dim rsAuxiliartmp As ADODB.Recordset
            Dim STRDATOX As String, STRCONCEPTOX As String, nCONTADOR As Integer
            Dim STRDATOX_NOMBRE As String
            Dim STRDATOX_BASICO As Double
            Dim STRDATOX_FECHAING As Date
            Dim STRDATOX_CARGO As String
            Dim STRDATOX_FONDOPENS As String
            'SNOMBOL
            Set rsAuxiliar = New ADODB.Recordset
            If ExisteTablaAux("[##_TMPBOLMEN" & VGL_COMPUTER & "]") Then DBSYSTEM.Execute "drop table [##_TMPBOLMEN" & VGL_COMPUTER & "]"
            
            DBSYSTEM.Execute "SELECT TMP.CODTRAB,TMP.NOMBRES,TMP.INUMBOL,TMP.NOMBOL,TMPLSTBOL.BASICO,TRA.CARGO,TRA.FECHAING,AFP.NOMBRE AS FONDOPENS,TMP.EGRESOS  INTO   [##_TMPBOLMEN" & VGL_COMPUTER & "]" & _
                             " FROM  [##_TMPLSTBOL" & VGL_COMPUTER & "] TMP LEFT JOIN TRABAJADORES TRA ON TMP.CODTRAB= TRA.CODTRAB,AFPS AFP,[##_TMPLSTBOL" & VGL_COMPUTER & "]  TMPLSTBOL WHERE TRA.FONDOPENS=AFP.CODAFP AND TMPLSTBOL.CODTRAB=TMP.CODTRAB"
            
            
            rsAuxiliar.Open "[##_TMPBOLMEN" & VGL_COMPUTER & "]", DBSYSTEM, adOpenDynamic, adLockOptimistic
            If rsAuxiliar.EOF Then Exit Sub
            
            If ExisteTablaAux("[##_TMPBOLMENFINALTOTAL" & VGL_COMPUTER & "]") Then DBSYSTEM.Execute "drop table [##_TMPBOLMENFINALTOTAL" & VGL_COMPUTER & "]"
            DBSYSTEM.Execute "CREATE TABLE [##_TMPBOLMENFINALTOTAL" & VGL_COMPUTER & "] (CODTRAB VARCHAR(8),NOMBRES VARCHAR(50),CONCEPTO  VARCHAR(50),FECHAING DATETIME ,BASICO  Numeric(20,2) ,CARGO VARCHAR(40),FONDOPENS VARCHAR(40),MONTO   Numeric(20,2)  DEFAULT 0 ,ORDEN  INT)"
             'DBSYSTEM.Execute "CREATE TABLE [##_TMPBOLMENFINALTOTAL" & VGL_COMPUTER & "] (CODTRAB VARCHAR(8),NOMBRE VARCHAR(40),MONTO   Numeric(20,2)  DEFAULT 0 ,ORDEN  INT)"
            Do While Not rsAuxiliar.EOF
                If ExisteTablaAux("[##_TMPBOLMENFINAL" & VGL_COMPUTER & "]") Then DBSYSTEM.Execute "drop table [##_TMPBOLMENFINAL" & VGL_COMPUTER & "]"
                CAD = " SELECT TMP.CODTRAB ,TMP.NOMBRES," & _
                " TMP.FECHAING," & _
                " " & _
                " TMP.BASICO,TMP.CARGO" & _
                " ,TMP.FONDOPENS , MOVX.MONTO, CON.NOMBRE, CON.TIPO, CON.CODIGO INTO [##_TMPBOLMENFINAL" & VGL_COMPUTER & "] FROM  MOV" & Format(Month(FMES), "00") & Format(Year(FMES), "00") & "  MOVX  LEFT JOIN CONCEPTOS CON ON MOVX.CONCEPTO=CON.CODIGO " & _
                " ,[##_TMPBOLMEN" & VGL_COMPUTER & "] TMP Where (MOVX.INUMBOL = " & rsAuxiliar!INUMBOL & ") And MOVX.CODNOMBOL = " & rsAuxiliar!NOMBOL & " And TMP.INUMBOL = MOVX.INUMBOL And (CON.TIPO = 0 Or CON.TIPO = 1 Or CON.TIPO = 2) ORDER BY CON.TIPO "
                    DBSYSTEM.Execute CAD
                    Set rsAuxiliartmp = New ADODB.Recordset
                    rsAuxiliartmp.Open "[##_TMPBOLMENFINAL" & VGL_COMPUTER & "]", DBSYSTEM, adOpenDynamic, adLockOptimistic
                    'If Not rsAuxiliartmp.EOF Then
                    '   DBSYSTEM.Execute "INSERT INTO [##_TMPBOLMENFINALTOTAL" & VGL_COMPUTER & "] VALUES('" & rsAuxiliartmp!CODTRAB & "' ,'" & rsAuxiliartmp!NOMBRES & "'," & rsAuxiliartmp!MONTO & "," & nCONTADOR & " )"
                    'End If
                If Not rsAuxiliartmp.EOF Then '/////////////////////////////////////////////////////////
                    STRDATOX = rsAuxiliartmp!CODTRAB
                    STRDATOX_NOMBRE = rsAuxiliartmp!NOMBRES
                    STRDATOX_BASICO = rsAuxiliartmp!BASICO
                    STRDATOX_FECHAING = rsAuxiliartmp!FECHAING
                    STRDATOX_CARGO = rsAuxiliartmp!CARGO
                    STRDATOX_FONDOPENS = rsAuxiliartmp!FONDOPENS
                    
                    STRCONCEPTOX = rsAuxiliartmp!NOMBRE
                    Do While Not rsAuxiliartmp.EOF                                                      'DATOS ,CONCEPTO ,MONTO ,ORDEN
                        nCONTADOR = INCREMENTAR(rsAuxiliartmp!Codigo)
                        If rsAuxiliartmp!TIPO = 1 Then
                                'DBSYSTEM.Execute "INSERT INTO [##_TMPBOLMENFINALTOTAL" & VGL_COMPUTER & "] VALUES('" & rsAuxiliartmp!DATOS & "' ,'" & rsAuxiliartmp!NOMBRE & "'," & rsAuxiliartmp!MONTO & "," & nCONTADOR & " )"
                                DBSYSTEM.Execute "INSERT INTO [##_TMPBOLMENFINALTOTAL" & VGL_COMPUTER & "] VALUES('" & rsAuxiliartmp!CODTRAB & "' ,'" & rsAuxiliartmp!NOMBRES & "','" & rsAuxiliartmp!NOMBRE & "','" & rsAuxiliartmp!FECHAING & "'," & rsAuxiliartmp!BASICO & ",'" & rsAuxiliartmp!CARGO & "','" & rsAuxiliartmp!FONDOPENS & "'," & rsAuxiliartmp!MONTO & "," & nCONTADOR & " )"
                        Else
                            If (rsAuxiliartmp!Codigo = "DIASTRAB" Or rsAuxiliartmp!Codigo = "HESIMPLE" Or rsAuxiliartmp!Codigo = "HEDOBLES") And rsAuxiliartmp!TIPO = 0 Then
                                                                                                                              'CODTRAB,                        NOMBRES,                   CONCEPTO,                        FECHAING,                      BASICO,                             CARGO,              FONDOPENS,                                    MONTO,             ORDEN
                                DBSYSTEM.Execute "INSERT INTO [##_TMPBOLMENFINALTOTAL" & VGL_COMPUTER & "] VALUES('" & rsAuxiliartmp!CODTRAB & "' ,'" & rsAuxiliartmp!NOMBRES & "','" & rsAuxiliartmp!NOMBRE & "','" & rsAuxiliartmp!FECHAING & "'," & rsAuxiliartmp!BASICO & ",'" & rsAuxiliartmp!CARGO & "','" & rsAuxiliartmp!FONDOPENS & "'," & rsAuxiliartmp!MONTO & "," & nCONTADOR & " )"
                            End If
                        End If
                        If rsAuxiliartmp!TIPO = 1 Then TOTALREMU = TOTALREMU + rsAuxiliartmp!MONTO
                        If rsAuxiliartmp!TIPO = 2 Then TOTALDESC = TOTALDESC + rsAuxiliartmp!MONTO
                        nCONTADOR = nCONTADOR + 1
                        rsAuxiliartmp.MoveNext
                    Loop
                    nCONTADOR = 100
                                                                                                                              'CODTRAB,            nOMBRES,                   CONCEPTO,                        FECHAING,                      BASICO,                     CARGO,              FONDOPENS,          MONTO,             ORDEN
                    DBSYSTEM.Execute "INSERT INTO [##_TMPBOLMENFINALTOTAL" & VGL_COMPUTER & "] VALUES('" & STRDATOX & "' ,'" & STRDATOX_NOMBRE & "','" & "TOTAL REMUNERACION" & "','" & STRDATOX_FECHAING & "'," & STRDATOX_BASICO & ",'" & STRDATOX_CARGO & "','" & STRDATOX_FONDOPENS & "'," & TOTALREMU & "," & nCONTADOR + 1 & " )"
                    DBSYSTEM.Execute "INSERT INTO [##_TMPBOLMENFINALTOTAL" & VGL_COMPUTER & "] VALUES('" & STRDATOX & "' ,'" & STRDATOX_NOMBRE & "','" & "DCTOS EN EFECTIVO" & "','" & STRDATOX_FECHAING & "'," & STRDATOX_BASICO & ",'" & STRDATOX_CARGO & "','" & STRDATOX_FONDOPENS & "'," & rsAuxiliar.Fields("EGRESOS") & "," & nCONTADOR + 2 & " )"
                    TOTALREMU = 0: TOTALDESC = 0: nCONTADOR = 0
                End If '///////////////////////////////////////////////////////////////////////////////
                rsAuxiliar.MoveNext
            Loop
            
            DBSTARPLAN.Execute "ASISTMP'[##_TMPBOLMENFINALTOTAL" & VGL_COMPUTER & "]'"
            With rptBoletas
                .Reset
                .WindowTitle = "REPORTE MENSUAL - COOPERATIVA"
                .ReportFileName = REGSISTEMA.REPORTES & "\REPPER\REPORT2.RPT"
                .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
                .StoredProcParam(0) = "[##_TMPBOLMENFINALTOTAL" & VGL_COMPUTER & "]"
                .Destination = crptToWindow
                .WindowState = crptMaximized
                .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
                .Formulas(1) = "XRUC='RUC N° " & REGSISTEMA.RUC & "'"
                .Formulas(2) = "XPERIODO='PERIODO :" & DESMES(CInt(Month(FMES))) & " " & Year(FMES) & "'"
                .WindowShowPrintBtn = True
                .WindowShowPrintSetupBtn = True
                .WindowShowRefreshBtn = True
                .WindowShowSearchBtn = True
                CambiaPanelBD False
              If rptBoletas.Status <> 2 Then .Action = 1
            End With
                
            
            
    End Select
End Sub

Private Sub XAREA_DblClick()
    Dim RSAUX As ADODB.Recordset
    Set RSAUX = New ADODB.Recordset
    If Sel1(0).Value Then 'SI ES POR AREAS
        RSAUX.Open "SELECT CODCCOSTO, NOMBRE FROM AREASTRAB ORDER BY CODCCOSTO", DBSYSTEM, adOpenStatic
    Else
        RSAUX.Open "SELECT CODCCOSTO, NOMBRE FROM CCOSTOS ORDER BY CODCCOSTO", DBSYSTEM, adOpenStatic
    End If
    If RSAUX.RecordCount = 0 Then
        MsgBox "NO SE HAN ENCONTRADO REGISTROS DE AREAS DE TRABAJO/CENTRO DE COSTO", vbCritical
        Set RSAUX = Nothing
        Exit Sub
    End If
    frmComun.CONECTAR RSAUX
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xArea.Text = RSAUX!CODCCOSTO & ": " & RSAUX!NOMBRE
        xArea.Tag = RSAUX!CODCCOSTO
        If ExisteTablaAux(" [##BOLSXAREA" & VGL_COMPUTER & "] ") Then
            DBSYSTEM.Execute "DROP TABLE  [##BOLSXAREA" & VGL_COMPUTER & "] "
        End If
        If Sel1(0).Value Then 'SI ES POR AREAS DE TRABAJO
            DBSYSTEM.Execute "SELECT CODTRAB INTO  [##BOLSXAREA" & VGL_COMPUTER & "]  FROM VWTRABAJ WHERE CODAREA LIKE '" & xArea.Tag & "%'"
        Else
            DBSYSTEM.Execute "SELECT CODTRAB INTO  [##BOLSXAREA" & VGL_COMPUTER & "]  FROM VWTRABAJ WHERE CODCCOSTO LIKE '" & xArea.Tag & "%'"
        End If
        frWait.Show 1
        DBSYSTEM.Execute "DELETE FROM  [##_TMPLSTBOL" & VGL_COMPUTER & "]  WHERE CODTRAB NOT IN (SELECT CODTRAB FROM  [##BOLSXAREA" & VGL_COMPUTER & "] )"
        RSLISTA.Requery
        FORMATEARDG
    End If
    Set RSAUX = Nothing
End Sub

Private Sub XAREA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        xArea.Text = ""
        xArea.Tag = ""
    End If
End Sub

Private Sub XFECHAFIN_KEYDOWN(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 13 Then SendKeys "{TAB}"
End Sub

Private Sub XFECHAINI_KEYDOWN(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 13 Then SendKeys "{TAB}"
End Sub

Private Sub XMARCATODOS_CLICK()
    Dim XITEM As ListItem
    For Each XITEM In Lista.ListItems
        If Not XITEM.Checked Then
            XITEM.Checked = True
            LISTA_ITEMCHECK XITEM
        End If
    Next
End Sub

Private Sub XMESES_Click()
    If xMeses.ListIndex = -1 Then Exit Sub
    Dim sMes As String
    If ITSOPEN Then
        RSBOLE.Close
    End If
    ITSOPEN = True
    sMes = "01/" & Left(xMeses.Text, 2) & "/" & Mid(xMeses.Text, 4, 4)
    RSBOLE.Open "SELECT * FROM NOMBOL WHERE MES=" & DateSQL(CDate(sMes)) & " ORDER BY FECHAINI, NOMBOL.NOMBRE", DBSYSTEM, adOpenKeyset
    If RSBOLE.RecordCount = 0 Then Exit Sub
    Dim xLista As ListItem
    RSBOLE.MoveFirst
    Lista.ListItems.Clear
    Do While Not RSBOLE.EOF
        Set xLista = Lista.ListItems.Add(, "C" & RSBOLE!Codigo, RSBOLE!NOMBRE, , 1)
        xLista.SubItems(1) = RSBOLE!FECHAINI
        xLista.SubItems(2) = RSBOLE!FECHAFIN
        RSBOLE.MoveNext
    Loop
    RSBOLE.MoveFirst
    If ExisteTablaAux(" [##_TMPLSTBOL" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DELETE FROM  [##_TMPLSTBOL" & VGL_COMPUTER & "] "
    RSLISTA.Requery
    Set dgBoletas.DataSource = RSLISTA
    FORMATEARDG
End Sub

Public Sub FORMATEARDG()
    Dim RSAUX As New ADODB.Recordset
    dgBoletas.Columns("Neto a Pagar").NumberFormat = "0.00 "
    RSAUX.Open "SELECT SUM(INGRESOS) AS TOTAL1,SUM(EGRESOS) AS TOTAL2,SUM(NETO) AS TOTAL3 FROM  [##_TMPLSTBOL" & VGL_COMPUTER & "] ", DBSYSTEM
    xSumIng.Caption = Format(IIf(IsNull(RSAUX!Total1), 0, RSAUX!Total1), "##,##0.00 ")
    xSumEgr.Caption = Format(IIf(IsNull(RSAUX!Total2), 0, RSAUX!Total2), "##,##0.00 ")
    xSumNet.Caption = Format(IIf(IsNull(RSAUX!TOTAL3), 0, RSAUX!TOTAL3), "##,##0.00 ")
    xCont.Caption = RSLISTA.RecordCount & " REGISTROS"
    Set RSAUX = Nothing
End Sub

Public Sub COMANDOTOOLBAR(COMANDO As String)
    Dim FMES As Date
    Select Case UCase(COMANDO)
        Case "BUSCAR"
            If RSLISTA.EOF Then Exit Sub
            CambiaPanelBD True
            Dim RSLISTA2 As New ADODB.Recordset
            Set RSLISTA2 = RSLISTA.Clone
            frmComun.CONECTAR RSLISTA2
            CambiaPanelBD False
            frmComun.Show 1
            If VGUTIL(1) <> "" Then
                RSLISTA.MoveFirst
                RSLISTA.FIND "CODTRAB='" & VGUTIL(1) & "'"
            End If
            Set RSLISTA2 = Nothing
        Case "IMPRIMIR", "PRELIMINAR"
            With rptBoletas
                CambiaPanelBD True
                .Reset
                .WindowTitle = "PLAN0025 - REPORTE DE BOLETAS DE REMUNERACIONES - NETOS A PAGAR"
                .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
                .ReportFileName = REGSISTEMA.REPORTES & "PLAN0025.RPT"
                .StoredProcParam(0) = " [##_TMPLSTBOL" & VGL_COMPUTER & "] "
                If Trim(dgBoletas.Tag) <> "" Then
                    .SortFields(0) = "+{ASISTMP." & dgBoletas.Tag & "}"
                End If
                .Destination = crptToWindow
                .WindowState = crptMaximized
                .WindowShowPrintBtn = True
                .WindowShowRefreshBtn = True
                .WindowShowSearchBtn = True
                .WindowShowPrintSetupBtn = True
                .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
                .Formulas(1) = "XRUC='RUC N° " & REGSISTEMA.RUC & "'"
                .Formulas(2) = "''"
                If .Status <> 2 Then .Action = 1
                CambiaPanelBD False
            End With
        Case "ELIMINAR"
            If RSLISTA.EOF Or RSLISTA.RecordCount = 0 Then
                MsgBox "NO SE HAN ENCONTRADO BOLETAS PARA ELIMINAR", vbCritical
                Exit Sub
            End If
            CambiaPanelBD True
            FMES = DevuelveValor("SELECT MES FROM NOMBOL WHERE CODIGO=" & RSLISTA!NOMBOL, DBSYSTEM)
            If MsgBox("REALMENTE DESEA ELIMINAR LA BOLETA DE " & RSLISTA!NOMBRES & ". LA ELIMINACIÓN ES PERMANENTE Y NO SE PODRÁ DESHACER LOS CAMBIOS", vbInformation + vbYesNo) = vbNo Then Exit Sub
            DBSYSTEM.Execute "DELETE FROM MOV" & Format(Month(FMES), "00") & Year(FMES) & " WHERE INUMBOL=" & RSLISTA!INUMBOL
            DBSYSTEM.Execute "DELETE FROM BOL" & Format(Month(FMES), "00") & Year(FMES) & " WHERE INUMBOL=" & RSLISTA!INUMBOL
            DBSYSTEM.Execute "UPDATE " & REGSISTEMA.TABLAADEL & " SET NOMBOL=0, NUMBOL=0 WHERE NOMBOL=" & RSLISTA!NOMBOL & " AND CODTRAB='" & RSLISTA!CODTRAB & "'"
            'Limpieza de las vacaciones
            DBSYSTEM.Execute "UPDATE HISTOVAC SET CERRADO=0 WHERE CODTRAB='" & RSLISTA!CODTRAB & "' AND NOMBOL=" & RSLISTA!NOMBOL
            Dim RSAUX As New ADODB.Recordset, NUMV As Integer
            RSAUX.Open "SELECT CODMOV, MONTO FROM PAGOSCTA WHERE CODNOMBOL=" & RSLISTA!NOMBOL & " AND CODTRAB='" & RSLISTA!CODTRAB & "'", DBSYSTEM, adOpenStatic
            DBSYSTEM.Execute "DELETE FROM PAGOSCTA WHERE CODNOMBOL=" & RSLISTA!NOMBOL & " AND CODTRAB='" & RSLISTA!CODTRAB & "' AND TIPOBOLETA='B'"
            Do While Not RSAUX.EOF 'BORRAMOS LAS ASIGNACIONES ANTERIORES
'                DBSYSTEM.Execute "UPDATE MOVICTA SET SALDO=SALDO+" & RSAUX!MONTO & " WHERE CODMOV=" & RSAUX!CODMOV, NUMV
'                If NUMV = 0 Then MsgBox "SE PRODUJO UN ERROR AL INTENTAR CAMBIAR UN PAGO DE CUENTA CORRIENTE DEL TRABAJADOR - SOURCE: INPUTPL, CODE MOVICTA=" & RSAUX!CODMOV, vbCritical
                Call ACTSALDO(RSAUX!CODMOV)
                RSAUX.MoveNext
            Loop
            Set RSAUX = Nothing
            RSLISTA.Delete
            If RSLISTA.RecordCount <> 0 Then
                RSLISTA.MovePrevious
                If RSLISTA.BOF Then RSLISTA.MoveFirst
            End If
            CambiaPanelBD False
            MsgBox "LA BOLETA DE REMUNERACIONES SE HA ELIMINADO SATISFACTORIAMENTE. SE HAN RESTABLECIDO LOS ADELANTOS Y DEBITOS DE CUENTAS CORRIENTES PENDIENTES SI LOS HUBIERA EXISTIDO", vbInformation
    End Select
End Sub

Private Sub XVERDETALLE_Click()
    CambiaPanelBD True
    If RSLISTA.EOF Then Exit Sub
    Dim FMES As Date
    FMES = DevuelveValor("SELECT MES FROM NOMBOL WHERE CODIGO=" & RSLISTA!NOMBOL, DBSYSTEM)
        
    If ExisteTablaAux(" [##TMPTRANS" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPTRANS" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "CREATE TABLE  [##TMPTRANS" & VGL_COMPUTER & "]  (CODIGO VARCHAR(15), DESCRIPCION VARCHAR(80), VALOR  Numeric(20,2) )"
    DBSYSTEM.Execute "INSERT INTO  [##TMPTRANS" & VGL_COMPUTER & "]  SELECT CONCEPTOS.CODIGO, CONCEPTOS.NOMBRE AS DESCRIPCION, MONTO AS VALOR FROM MOV" & Format(Month(FMES), "00") & Year(FMES) & " MOV, CONCEPTOS WHERE MOV.CONCEPTO=CONCEPTOS.CODIGO AND INUMBOL=" & RSLISTA!INUMBOL
    DBSYSTEM.Execute "INSERT INTO  [##TMPTRANS" & VGL_COMPUTER & "]  SELECT CODIGO, '<ADELANTO DE PAGO>' AS DESCRIPCION,MONTO AS VALOR FROM " & REGSISTEMA.TABLAADEL & " WHERE ORIGEN=" & RSLISTA!NOMBOL & " AND CODTRAB='" & RSLISTA!CODTRAB & "'"
    'FRWAIT.SHOW 1
    DBSYSTEM.Execute "INSERT INTO  [##TMPTRANS" & VGL_COMPUTER & "]  (CODIGO,DESCRIPCION,VALOR) SELECT MOVICTA.CODMOV, DESCRIPCION, MONTO FROM MOVICTA, PAGOSCTA WHERE MOVICTA.CODMOV=PAGOSCTA.CODMOV AND PAGOSCTA.CODTRAB='" & RSLISTA!CODTRAB & "' AND CODNOMBOL=" & RSLISTA!NOMBOL & " AND TIPOBOLETA='B'"
    DBSYSTEM.Execute "UPDATE  [##TMPTRANS" & VGL_COMPUTER & "]  SET VALOR=VALOR"
    'MSGBOX "SE CARGARON LOS DATOS SATISFACTORIAMENTE", VBINFORMATION
    If ExisteTablaAux(" [##_TMPDETBLT" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##_TMPDETBLT" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "SELECT * INTO  [##_TMPDETBLT" & VGL_COMPUTER & "]  FROM  [##TMPTRANS" & VGL_COMPUTER & "] "
    VPTAREA = "DETALLE DE BOLETA DE " & RSLISTA!NOMBRES
    CambiaPanelBD False
    ' INGRESOS              EGRESOS
    
        frDetBlt.BOLMEANO_NUMBOL = DevuelveValor("SELECT NUMBOL FROM BOL" & Format(Month(FMES), "00") & Year(FMES) & " WHERE INUMBOL=" & RSLISTA!INUMBOL & "", DBSYSTEM)
        frDetBlt.BOLMEANO_INUMBOL = DevuelveValor("SELECT INUMBOL FROM BOL" & Format(Month(FMES), "00") & Year(FMES) & " WHERE INUMBOL=" & RSLISTA!INUMBOL & "", DBSYSTEM)
        frDetBlt.BOLMEANO_CODNUMBOL = DevuelveValor("SELECT CODNOMBOL FROM BOL" & Format(Month(FMES), "00") & Year(FMES) & " WHERE INUMBOL=" & RSLISTA!INUMBOL & "", DBSYSTEM)
        frDetBlt.BOLMEANO_QYERYORIGINAL = "SELECT * FROM  [##TMPTRANS" & VGL_COMPUTER & "]"
        
        frDetBlt.FECHAPROCESO = FMES
        frDetBlt.BOLMEANO_CODTRAB = RSLISTA!CODTRAB
        frDetBlt.BOLMEANO_TOTIN = RSLISTA!INGRESOS
        frDetBlt.BOLMEANO_TOTEG = RSLISTA!EGRESOS
    frDetBlt.Show 1
    If ExisteTabla(" [##_TMPDETBLT" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##_TMPDETBLT" & VGL_COMPUTER & "] "
    If ExisteTabla(" [##TMPTRANS" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPTRANS" & VGL_COMPUTER & "] "
End Sub

Public Sub CARGABOL()
    
    Dim FMES As Date
    Dim ESVACACIONES As Boolean
    FMES = DevuelveValor("SELECT MES FROM NOMBOL WHERE CODIGO=" & RSLISTA!NOMBOL, DBSYSTEM)
    ESVACACIONES = False
    If Not ExisteCampo("XREDONDEO", "RPTBOLETAS", DBSTARPLAN) Then
        DBSTARPLAN.Execute "ALTER TABLE RPTBOLETAS ADD XREDONDEO  Numeric(20,2) "
    End If
    'SE CREAN LOS CAMPOS FIJOS
    If Not ExisteCampo("NOMBOL", "RPTBOLETAS", DBSTARPLAN) Then
        DBSTARPLAN.Execute "ALTER TABLE RPTBOLETAS ADD NOMBOL BIGINT"
    End If
    If Not ExisteCampo("FIJO1", "RPTBOLETAS", DBSTARPLAN) Then
        DBSTARPLAN.Execute "ALTER TABLE RPTBOLETAS ADD FIJO1  Numeric(20,2) "
        DBSTARPLAN.Execute "ALTER TABLE RPTBOLETAS ADD FIJO2  Numeric(20,2) "
        DBSTARPLAN.Execute "ALTER TABLE RPTBOLETAS ADD FIJO3  Numeric(20,2) "
        DBSTARPLAN.Execute "ALTER TABLE RPTBOLETAS ADD FIJO4  Numeric(20,2) "
        DBSTARPLAN.Execute "ALTER TABLE RPTBOLETAS ADD FIJO5  Numeric(20,2) "
        DBSTARPLAN.Execute "ALTER TABLE RPTBOLETAS ADD FIJO7  Numeric(20,2) "
        DBSTARPLAN.Execute "ALTER TABLE RPTBOLETAS ADD FIJO8  Numeric(20,2) "
        DBSTARPLAN.Execute "ALTER TABLE RPTBOLETAS ADD FIJO9  Numeric(20,2) "
    End If
    If Not ExisteCampo("CODCOSTO", "RPTBOLETAS", DBSTARPLAN) Then
        DBSTARPLAN.Execute "ALTER TABLE RPTBOLETAS ADD CODCOSTO INT"
        DBSTARPLAN.Execute "ALTER TABLE RPTBOLETAS ADD CODAREA INT"
    End If
    If Not ExisteCampo("CUSPP", "RPTBOLETAS", DBSTARPLAN) Then
        DBSTARPLAN.Execute "ALTER TABLE RPTBOLETAS ADD CUSPP VARCHAR(12)"
        DBSTARPLAN.Execute "ALTER TABLE RPTBOLETAS ADD AREA VARCHAR(25)"
    End If
    'On Error GoTo ERRPRTBOL
    Dim RSAUX As New ADODB.Recordset
    Dim RSBOL As New ADODB.Recordset
    If RSLISTA.EOF Then Exit Sub
    If ExisteTablaAux(" [##TMPTRANS" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPTRANS" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "CREATE TABLE  [##TMPTRANS" & VGL_COMPUTER & "]  (CODIGO VARCHAR(15), DESCRIPCION VARCHAR(30), VALOR  Numeric(20,2) , TIPO INT, ENLACE VARCHAR(8), FILA INT, IMPRESIONFIJA BIT)"
    'JALAR LOS OTROS INGRESOS (CUENTAS CORRIENTES)
    DBSYSTEM.Execute "INSERT INTO  [##TMPTRANS" & VGL_COMPUTER & "]  (CODIGO,DESCRIPCION,VALOR,TIPO,ENLACE,FILA, IMPRESIONFIJA) SELECT MOVICTA.CODMOV, DESCRIPCION, MONTO, 1 AS TIPO, ' ' AS ENLACE,11 AS FILA,0 AS IMPRESIONFIJA FROM MOVICTA, PAGOSCTA WHERE MOVICTA.CODMOV=PAGOSCTA.CODMOV AND MOVICTA.TIPOGRUPO=1 AND PAGOSCTA.CODTRAB='" & RSLISTA!CODTRAB & "' AND CODNOMBOL=" & RSLISTA!NOMBOL & " AND TIPOBOLETA='B'"
    'JALAR LOS CONCEPTOS
    DBSYSTEM.Execute "INSERT INTO  [##TMPTRANS" & VGL_COMPUTER & "]  SELECT CONCEPTOS.CODIGO, CONCEPTOS.NOMBRE AS DESCRIPCION, MONTO AS VALOR, CONCEPTOS.TIPO, CONCEPTOS.ENLACE, FILA, IMPRESIONFIJA FROM MOV" & Format(Month(FMES), "00") & Year(FMES) & " MOV, CONCEPTOS WHERE MOV.CONCEPTO=CONCEPTOS.CODIGO AND INUMBOL=" & RSLISTA!INUMBOL
    'JALAR LOS ADELANTOS DE PAGO
    DBSYSTEM.Execute "INSERT INTO  [##TMPTRANS" & VGL_COMPUTER & "]  SELECT CODIGO, '<ADELANTO DE PAGO>' AS DESCRIPCION,MONTO AS VALOR,2 AS TIPO,' ' AS ENLACE, 4 AS FILA,0 AS IMPRESIONFIJA FROM " & REGSISTEMA.TABLAADEL & " WHERE ORIGEN=" & RSLISTA!NOMBOL & " AND CODTRAB='" & RSLISTA!CODTRAB & "'"
    'JALAR LOS OTROS EGRESOS (CUENTAS CORRIENTES)
    DBSYSTEM.Execute "INSERT INTO  [##TMPTRANS" & VGL_COMPUTER & "]  (CODIGO,DESCRIPCION,VALOR,TIPO,ENLACE,FILA, IMPRESIONFIJA) SELECT MOVICTA.CODMOV, DESCRIPCION, MONTO, 2 AS TIPO, ' ' AS ENLACE,12 AS FILA,0 AS IMPRESIONFIJA FROM MOVICTA, PAGOSCTA WHERE MOVICTA.CODMOV=PAGOSCTA.CODMOV AND MOVICTA.TIPOGRUPO=2 AND PAGOSCTA.CODTRAB='" & RSLISTA!CODTRAB & "' AND CODNOMBOL=" & RSLISTA!NOMBOL & " AND TIPOBOLETA='B'"
    'COLOCANDO LA FECHA DE LA VACACIONES AL SISTEMA
    
    RSAUX.Open "SELECT * FROM VWTRABAJ WHERE CODTRAB='" & RSLISTA!CODTRAB & "'", DBSYSTEM, adOpenStatic
    RSBOL.Open "RPTBOLETAS", DBSTARPLAN, adOpenKeyset, adLockOptimistic
    'ADICION DE LA BOLETA
    RSBOL.AddNew
    RSBOL!NOMBOL = RSLISTA!NOMBOL
    RSBOL!CODTRAB = RSLISTA!CODTRAB
    RSBOL!NOMBRES = RSAUX!NOMBRES
    RSBOL!CENTROCOSTO = RSAUX!CENTRO
    RSBOL!FECHAING = FechS(RSAUX!FECHAING, Adof)
    RSBOL!PERIODO = RSLISTA!PERIODO
    RSBOL!TOTING = RSLISTA!INGRESOS
    RSBOL!TOTEGR = RSLISTA!EGRESOS
    RSBOL!XREDONDEO = DevuelveValor("SELECT XREDONDEO FROM BOL" & Format(Month(FMES), "00") & Year(FMES) & " WHERE INUMBOL=" & RSLISTA!INUMBOL, DBSYSTEM)
    RSBOL!BASICO = DevuelveValor("SELECT BASICO FROM BOL" & Format(Month(FMES), "00") & Year(FMES) & " WHERE INUMBOL=" & RSLISTA!INUMBOL, DBSYSTEM)
    RSBOL!AFP = "" & RSAUX!NOMBREAFP
    RSBOL!CARGO = RSAUX!CARGO
    RSBOL!DOCUMENTO = "" & RSAUX!DOCIDEN
    RSBOL!CARNETSEG = RSAUX!CARNETSEG
    RSBOL!CUENTABANCO = "" & RSAUX!CTABANCO
    RSBOL!CUSPP = "" & DevuelveValor("SELECT CUSPP FROM TRABAJADORES WHERE CODTRAB='" & RSLISTA!CODTRAB & "'", DBSYSTEM)
    RSBOL!AREA = "" & RSAUX!NOMBREAREA
    RSBOL!FECHACESE = FechS(RSAUX!FECHACESE, Adof)
    Dim IND1 As Byte, IND2 As Byte, IND3 As Byte, IND4 As Byte
    Dim CLASEPRT As Byte
    CLASEPRT = DevuelveValor("SELECT CLASEBOLETA FROM EMPRESA", DBSYSTEM)
    DevNomRep Trim(REGSISTEMA.RUC), Trim(REGSISTEMA.USER), FILEBOLETA, CLASEPRT
    Set RSAUX = Nothing
    
    Dim PERIODOVAC As String
    RSAUX.Open "SELECT * FROM HISTOVAC WHERE NOMBOL=" & RSLISTA!NOMBOL & " AND CODTRAB='" & RSLISTA!CODTRAB & "'", DBSYSTEM, adOpenKeyset, adLockReadOnly
    If Not RSAUX.EOF Or RSAUX.RecordCount <> 0 Then
        RSBOL("FECHAVAC1") = CDate(RSAUX!FECHAINI)
        RSBOL("FECHAVAC2") = CDate(RSAUX!FECHAFIN)
        PERIODOVAC = RSAUX!PERIODO
        ESVACACIONES = True
    End If
    Set RSAUX = Nothing
    RSAUX.Open "SELECT * FROM  [##TMPTRANS" & VGL_COMPUTER & "]  ORDER BY TIPO, FILA", DBSYSTEM, adOpenStatic
    IND1 = 0
    IND2 = 0
    IND3 = 0
    IND4 = 0
    Dim XCT As String, XCN As String, XCONTEO, XTIPCNPT As Byte, XTOTALAPORT As Single
    XCONTEO = 0
    XCT = "C"
    XCN = "I"
    XTIPCNPT = 0
    XTOTALAPORT = 0
    Do While Not RSAUX.EOF
        If RSAUX!IMPRESIONFIJA And Not XFLAG Then
            If RSAUX!FILA >= 10 Then
                MsgBox "La fila en el concepto: " & RSAUX!Codigo & "  de tipo impresión fija " & Chr(13) & _
                       "no debe ser mayor a 10,por lo tanto no va salir en la boleta", vbExclamation
                XFLAG = True
                Exit Sub
               Else
                RSBOL.Fields("FIJO" & RSAUX!FILA & "").Value = RSAUX!VALOR
            End If
        Else
            Select Case CLASEPRT
                Case 0
                    Select Case RSAUX!TIPO
                        Case 0: IND1 = IND1 + 1
                                RSBOL.Fields("C" & IND1).Value = RSAUX!DESCRIPCION
                                RSBOL.Fields("INF" & IND1).Value = RSAUX!VALOR
                        Case 1: IND1 = IND1 + 1
                                RSBOL.Fields("C" & IND1).Value = RSAUX!DESCRIPCION
                                If UCase(RSAUX!Codigo) = "REMUVAC" Then RSBOL.Fields("C" & IND1).Value = "REMU. VAC. " & PERIODOVAC
                                RSBOL.Fields("I" & IND1).Value = RSAUX!VALOR
                        Case 2: IND3 = IND3 + 1
                                RSBOL.Fields("R" & IND3).Value = RSAUX!DESCRIPCION
                                RSBOL.Fields("E" & IND3).Value = RSAUX!VALOR
                        Case 3: IND4 = IND4 + 1
                                RSBOL.Fields("G" & IND4).Value = RSAUX!DESCRIPCION
                                RSBOL.Fields("A" & IND4).Value = RSAUX!VALOR
                    End Select
                Case 1
                    Select Case RSAUX!TIPO
                        Case 0: IND1 = IND1 + 1
                                RSBOL.Fields("C" & IND1).Value = RSAUX!DESCRIPCION & "          (" & RSAUX!VALOR & ")"
                                RSBOL.Fields("INF" & IND1).Value = RSAUX!VALOR
                        Case 1: IND1 = IND1 + 1
                                RSBOL.Fields("C" & IND1).Value = RSAUX!DESCRIPCION
                                RSBOL.Fields("I" & IND1).Value = RSAUX!VALOR
                        Case 2: IND1 = IND1 + 1
                                RSBOL.Fields("C" & IND1).Value = RSAUX!DESCRIPCION
                                RSBOL.Fields("E" & IND1).Value = RSAUX!VALOR
                        Case 3: IND4 = IND4 + 1
                                RSBOL.Fields("G" & IND4).Value = RSAUX!DESCRIPCION
                                RSBOL.Fields("A" & IND4).Value = RSAUX!VALOR
                    End Select
                Case 2
                    If XCONTEO >= 21 Then
                        XCT = "R"
                        XCN = "E"
                        XCONTEO = XCONTEO - 20
                    End If
                    XCONTEO = XCONTEO + 1
                    If XTIPCNPT <> RSAUX!TIPO Then
                        XTIPCNPT = RSAUX!TIPO
                        Select Case XTIPCNPT
                            Case 2
                                If XCONTEO >= 21 Then
                                    XCT = "R"
                                    XCN = "E"
                                    XCONTEO = XCONTEO - 20
                                End If
                                RSBOL.Fields(XCT & XCONTEO).Value = "TOTAL INGRESOS"
                                RSBOL.Fields(XCN & XCONTEO).Value = RSLISTA!INGRESOS
                                XCONTEO = XCONTEO + 2
                                If XCONTEO >= 21 Then
                                    XCT = "R"
                                    XCN = "E"
                                    XCONTEO = XCONTEO - 20
                                End If
                                RSBOL.Fields(XCT & XCONTEO).Value = "RETENCIONES Y DESCUENTOS"
                                RSBOL.Fields(XCN & XCONTEO).Value = 0
                                XCONTEO = XCONTEO + 1
                            Case 3
                                If XCONTEO >= 21 Then
                                    XCT = "R"
                                    XCN = "E"
                                    XCONTEO = XCONTEO - 20
                                End If
                                RSBOL.Fields(XCT & XCONTEO).Value = "TOTAL EGRESOS"
                                RSBOL.Fields(XCN & XCONTEO).Value = RSLISTA!EGRESOS
                                XCONTEO = XCONTEO + 2
                                If XCONTEO >= 21 Then
                                    XCT = "R"
                                    XCN = "E"
                                    XCONTEO = XCONTEO - 20
                                End If
                                RSBOL.Fields(XCT & XCONTEO).Value = "NETO A PAGAR"
                                RSBOL.Fields(XCN & XCONTEO).Value = RSLISTA!INGRESOS - RSLISTA!EGRESOS
                                XCONTEO = XCONTEO + 2
                                If XCONTEO >= 21 Then
                                    XCT = "R"
                                    XCN = "E"
                                    XCONTEO = XCONTEO - 20
                                End If
                                If XCONTEO = 0 Then XCONTEO = XCONTEO + 1
                                RSBOL.Fields(XCT & XCONTEO).Value = "APORTACIONES DEL EMPLEADOR"
                                RSBOL.Fields(XCN & XCONTEO).Value = 0
                                XCONTEO = XCONTEO + 1
                        End Select
                    End If
                    If XCONTEO >= 21 Then
                        XCT = "R"
                        XCN = "E"
                        XCONTEO = XCONTEO - 20
                    End If
                    If RSAUX!TIPO = 3 Then XTOTALAPORT = XTOTALAPORT + RSAUX!VALOR
                    RSBOL.Fields(XCT & XCONTEO).Value = RSAUX!DESCRIPCION
                    RSBOL.Fields(XCN & XCONTEO).Value = RSAUX!VALOR
            End Select
        End If
        RSAUX.MoveNext
 '       RSBOL.Update
    Loop
    If CLASEPRT = 2 Then
        XCONTEO = XCONTEO + 1
        If XCONTEO >= 21 Then
            XCT = "R"
            XCN = "E"
            XCONTEO = XCONTEO - 20
        End If
        RSBOL.Fields(XCT & XCONTEO).Value = "TOTAL APORTACIONES"
        RSBOL.Fields(XCN & XCONTEO).Value = XTOTALAPORT
    End If
    RSBOL.Update
    If ESVACACIONES Then
        IND1 = IND1 + 1
        
    End If
    Set RSAUX = Nothing
    Set RSBOL = Nothing
'    Exit Sub
'ERRPRTBOL:
'    Resume Next
End Sub

Public Function COMPRUEBAPLAN() As Boolean
    Dim XVAL As Boolean
    XVAL = False
    COMPRUEBAPLAN = False
    REGSISTEMA.COLPLANADEL = DevuelveValor("SELECT ADELPLAN FROM EMPRESA", DBSYSTEM)
    If Not ExisteCampo(REGSISTEMA.COLPLANADEL, "PLAN2000", DBSYSTEM) Then
        MsgBox "NO SE ENCUENTRA EL CAMPO CORRESPONDIENTE A ADELANTOS DE REMUNERACIONES. LOS ADELANTOS DE REMUNERACIONES DEBEN DE SER ALMACENADOS EN UNA COLUMNA DE PLANILLA, LA CUAL NO EXISTE. DEFINA LA COLUMNA EN EL PANEL DE CONFIGURACIÓN DEL SISTEMA", vbInformation
        Exit Function
    End If
    If Not ExisteCampo("TOTING", "PLAN2000", DBSYSTEM) Then
        MsgBox "NO HA DEFINIDO O NO SE ENCUENTRA EL CAMPO TOTING, EL CUAL ES NECESARIO PARA EL FUNCIONAMIENTO DEL SISTEMA. PUEDE CREARLO DESDE EL MENU BASE DE DATOS-OTROS ARCHIVOS-COLUMNAS DE PLANILLA", vbInformation
        Exit Function
    End If
    If Not ExisteCampo("TOTEGR", "PLAN2000", DBSYSTEM) Then
        MsgBox "NO HA DEFINIDO O NO SE ENCUENTRA EL CAMPO TOTEGR, EL CUAL ES NECESARIO PARA EL FUNCIONAMIENTO DEL SISTEMA. PUEDE CREARLO DESDE EL MENU BASE DE DATOS-OTROS ARCHIVOS-COLUMNAS DE PLANILLA", vbInformation
        Exit Function
    End If
    Dim RSRUBROS As New ADODB.Recordset
    RSRUBROS.Open "CONCEPTOS", DBSYSTEM, adOpenKeyset, adLockOptimistic
    If RSRUBROS.EOF Then
        MsgBox "LOS CONCEPTOS DE REMUNERACIONES NO SE HAN DEFINIDO AÚN", vbInformation
        Set RSRUBROS = Nothing
        Exit Function
    End If
    Dim X As Long, Z As Byte
    X = 0
    Do While Not RSRUBROS.EOF
        If Trim(RSRUBROS!COLPLANILLA) <> "" Then
            DBSYSTEM.Execute "UPDATE COLUMPL SET TIPO=TIPO WHERE CODIGO='" & Trim(RSRUBROS!COLPLANILLA) & "'", X
            If X = 0 Then
                Z = MsgBox("EL CONCEPTO DE REMUNERACIÓN " & RSRUBROS!NOMBRE & " PRESENTA COMO COLUMNA DE PLANILLA EL CÓDIGO " & RSRUBROS!COLPLANILLA & " EL CUAL NO EXISTE DENTRO DE LA BASE DE DATOS. DESEA DEPURAR EL CONCEPTO DE REMUNERACIÓN", vbQuestion + vbYesNoCancel)
                If Z = vbCancel Or Z = vbNo Then Exit Function
                If Z = vbYes Then
                    VPTAREA = "EDITAR"
                    VPCODTMP = RSRUBROS!Codigo
                    Load frECnpt
                    frECnpt.cmCancela.Enabled = False
                    frECnpt.Show 1
                End If
            End If
        End If
        RSRUBROS.MoveNext
    Loop
    Set RSRUBROS = Nothing
    XVAL = True
    COMPRUEBAPLAN = True
End Function

Private Sub XVERRANGO_Click()
    xFechaIni.Day = 1
    xFechaFin.Day = 1
    If xFechaIni.Value > xFechaFin.Value Then
        MsgBox "EL RABGO DE FECHA ESTA ERRADO. EL MES DE INICIO DEBE SER MENOR O IGUAL A LA FECHA FINAL", vbInformation
        Exit Sub
    End If
    
    Dim sMes As String
    If ITSOPEN Then
        RSBOLE.Close
    End If
    ITSOPEN = True
    RSBOLE.Open "SELECT * FROM NOMBOL WHERE MES BETWEEN " & DateSQL(xFechaIni.Value) & " AND " & DateSQL(xFechaFin.Value) & " ORDER BY FECHAINI, NOMBOL.NOMBRE", DBSYSTEM, adOpenKeyset
    If RSBOLE.RecordCount = 0 Then Exit Sub
    Dim xLista As ListItem
    RSBOLE.MoveFirst
    Lista.ListItems.Clear
    Do While Not RSBOLE.EOF
        Set xLista = Lista.ListItems.Add(, "C" & RSBOLE!Codigo, RSBOLE!NOMBRE, , 1)
        xLista.SubItems(1) = RSBOLE!FECHAINI
        xLista.SubItems(2) = RSBOLE!FECHAFIN
        RSBOLE.MoveNext
    Loop
    RSBOLE.MoveFirst
    If ExisteTablaAux(" [##_TMPLSTBOL" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DELETE FROM  [##_TMPLSTBOL" & VGL_COMPUTER & "] "
    RSLISTA.Requery
    Set dgBoletas.DataSource = RSLISTA
    FORMATEARDG
End Sub

Private Sub XVISTAMES_Click()
    If xVistaMes.ListIndex = 0 Then
        xMeses.Visible = True
        xFechaIni.Visible = False
        xFechaFin.Visible = False
        xVerRango.Visible = False
        XMESES_Click
    Else
        xFechaIni.Value = Date
        xFechaFin.Value = Date
        xMeses.Visible = False
        xFechaIni.Visible = True
        xFechaFin.Visible = True
        xVerRango.Visible = True
    End If
End Sub

Private Sub XVISTAMES_KeyDown(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 13 Then SendKeys "{TAB}"
End Sub
Private Function INCREMENTAR(Optional CONC As String) As Integer
On Error GoTo handler
    Select Case CONC
        Case "DIASTRAB"
              INCREMENTAR = 20
        Case "REMBASIC"
              INCREMENTAR = 21
        Case "ASIGFAMI"
              INCREMENTAR = 22
        Case "HESIMPLE"
              INCREMENTAR = 23
        Case "MHE25"
              INCREMENTAR = 24
        Case "HEDOBLES"
              INCREMENTAR = 25
        Case "MHE100"
              INCREMENTAR = 26
        Case "REFRIGE"
              INCREMENTAR = 27
        Case "REINTEGR"
               INCREMENTAR = 28
        Case "INCPROD"
               INCREMENTAR = 29
        Case "LEY01"
               INCREMENTAR = 30
        Case "IAFPNVO"
              INCREMENTAR = 31
        Case Else
              INCREMENTAR = 32
    End Select
 Exit Function
handler:
   INCREMENTAR = 0
End Function

