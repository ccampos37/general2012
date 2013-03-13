VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frSelect 
   BackColor       =   &H8000000B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selector de Trabajadores"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7245
   Icon            =   "frSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   7245
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Todos los Trabajadores"
      Height          =   330
      Left            =   210
      TabIndex        =   1
      Top             =   4485
      Width           =   2130
   End
   Begin Crystal.CrystalReport RPT 
      Left            =   2040
      Top             =   2550
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1980
      Top             =   1860
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
            Picture         =   "frSelect.frx":08CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   5610
      TabIndex        =   3
      Top             =   4485
      Width           =   1305
   End
   Begin VB.CommandButton cmAtras 
      Caption         =   "<< &Atrás"
      Enabled         =   0   'False
      Height          =   330
      Left            =   2730
      TabIndex        =   2
      Top             =   4485
      Width           =   1305
   End
   Begin VB.CommandButton cmSiguiente 
      Caption         =   "&Siguiente >>"
      Default         =   -1  'True
      Height          =   330
      Left            =   4170
      TabIndex        =   0
      Top             =   4485
      Width           =   1305
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4140
      Left            =   2550
      TabIndex        =   4
      Top             =   165
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   7303
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Paso 1"
      TabPicture(0)   =   "frSelect.frx":0C1E
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmVacaciones"
      Tab(0).Control(1)=   "ctxBaja"
      Tab(0).Control(2)=   "xNocalculo"
      Tab(0).Control(3)=   "Frame1"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Paso 2"
      TabPicture(1)   =   "frSelect.frx":0C3A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(2)=   "Op2"
      Tab(1).Control(3)=   "Op1"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Paso 3"
      TabPicture(2)   =   "frSelect.frx":0C56
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "xNumTrab"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Image4"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "dgTrabs"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmQuitar"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmAdiciona"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Paso 4"
      TabPicture(3)   =   "frSelect.frx":0C72
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmImprimir"
      Tab(3).Control(1)=   "Image3"
      Tab(3).Control(2)=   "Image2"
      Tab(3).Control(3)=   "Label3"
      Tab(3).Control(4)=   "Label2"
      Tab(3).ControlCount=   5
      Begin VB.CommandButton cmVacaciones 
         Caption         =   "&Vacaciones"
         Height          =   330
         Left            =   -74715
         TabIndex        =   31
         Top             =   3720
         Width           =   1290
      End
      Begin VB.CheckBox ctxBaja 
         Caption         =   "No mostrar Tabajadores de Baja"
         Height          =   195
         Left            =   -74670
         TabIndex        =   5
         Top             =   720
         Value           =   1  'Checked
         Width           =   3945
      End
      Begin VB.CheckBox xNocalculo 
         Alignment       =   1  'Right Justify
         Caption         =   "No Calculados"
         Height          =   195
         Left            =   -72060
         TabIndex        =   13
         Top             =   3795
         Width           =   1485
      End
      Begin VB.Frame Frame3 
         Caption         =   "Fecha de Ingreso"
         Height          =   1200
         Left            =   -74850
         TabIndex        =   28
         Top             =   2745
         Width           =   4275
         Begin MSComCtl2.DTPicker xFecha2 
            Height          =   315
            Left            =   2640
            TabIndex        =   18
            Top             =   705
            Visible         =   0   'False
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            Format          =   62062593
            CurrentDate     =   36698
         End
         Begin MSComCtl2.DTPicker xFecha 
            Height          =   315
            Left            =   2640
            TabIndex        =   17
            Top             =   360
            Visible         =   0   'False
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            Format          =   62062593
            CurrentDate     =   36698
         End
         Begin VB.ComboBox xCondFecha 
            Height          =   315
            ItemData        =   "frSelect.frx":0C8E
            Left            =   1515
            List            =   "frSelect.frx":0CAA
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   360
            Width           =   1110
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Ingreso"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   405
            Width           =   1245
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Area de Trabajo"
         Height          =   1740
         Left            =   -74850
         TabIndex        =   27
         Top             =   930
         Width           =   4275
         Begin VB.CheckBox xDepende 
            Caption         =   "Incluir Trabajadores dependientes"
            Height          =   210
            Left            =   165
            TabIndex        =   15
            Top             =   1395
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   3660
         End
         Begin VB.CheckBox xTodaAC 
            Caption         =   "Todas las Areas de Trabajo de la Empresa"
            Height          =   240
            Left            =   150
            TabIndex        =   11
            Top             =   330
            Value           =   1  'Checked
            Width           =   3735
         End
         Begin VB.CheckBox NoAC 
            Caption         =   "No incluir"
            Height          =   195
            Left            =   150
            TabIndex        =   12
            Top             =   623
            Visible         =   0   'False
            Width           =   1005
         End
         Begin AplisetControlText.Aplitext xCodAC2 
            Height          =   300
            Left            =   1290
            TabIndex        =   29
            Top             =   960
            Visible         =   0   'False
            Width           =   2820
            _ExtentX        =   4974
            _ExtentY        =   529
            Locked          =   -1  'True
            Text            =   ""
         End
         Begin AplisetControlText.Aplitext xCodAC1 
            Height          =   300
            Left            =   1290
            TabIndex        =   30
            Top             =   615
            Visible         =   0   'False
            Width           =   2820
            _ExtentX        =   4974
            _ExtentY        =   529
            Locked          =   -1  'True
            Text            =   ""
         End
      End
      Begin VB.OptionButton Op2 
         Caption         =   "Centro de Costo"
         Height          =   255
         Left            =   -72585
         TabIndex        =   10
         Top             =   585
         Width           =   1500
      End
      Begin VB.OptionButton Op1 
         Caption         =   "Area de Trabajo"
         Height          =   255
         Left            =   -74835
         TabIndex        =   9
         Top             =   585
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.CommandButton cmAdiciona 
         Caption         =   "Adicionar"
         Height          =   315
         Left            =   2100
         TabIndex        =   20
         Top             =   3675
         Width           =   1125
      End
      Begin VB.CommandButton cmQuitar 
         Caption         =   "&Quitar"
         Height          =   315
         Left            =   915
         TabIndex        =   19
         Top             =   3675
         Width           =   1125
      End
      Begin VB.CommandButton cmImprimir 
         Caption         =   "&Imprimir"
         Height          =   345
         Left            =   -71910
         TabIndex        =   22
         Top             =   3585
         Width           =   1305
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tipo de Trabajador"
         Height          =   3255
         Left            =   -74850
         TabIndex        =   23
         Top             =   420
         Width           =   4275
         Begin MSComctlLib.ListView Lista 
            Height          =   2085
            Left            =   120
            TabIndex        =   8
            Top             =   1110
            Width           =   4020
            _ExtentX        =   7091
            _ExtentY        =   3678
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
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
               Text            =   "Tipo de Trabajador"
               Object.Width           =   6456
            EndProperty
         End
         Begin VB.CheckBox xNoTrab 
            Caption         =   "No incluir los marcados"
            Height          =   225
            Left            =   180
            TabIndex        =   7
            Top             =   825
            Width           =   2055
         End
         Begin VB.CheckBox xTodosTrab 
            Caption         =   "Todos los tipos de Trabajador"
            Height          =   195
            Left            =   180
            TabIndex        =   6
            Top             =   570
            Value           =   1  'Checked
            Width           =   2820
         End
      End
      Begin MSDataGridLib.DataGrid dgTrabs 
         Height          =   3015
         Left            =   60
         TabIndex        =   21
         Top             =   375
         Width           =   4440
         _ExtentX        =   7832
         _ExtentY        =   5318
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
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   240
         Picture         =   "frSelect.frx":0CCF
         Top             =   3495
         Width           =   480
      End
      Begin VB.Label xNumTrab 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "50000 Trabs"
         Height          =   270
         Left            =   3315
         TabIndex        =   26
         Top             =   3705
         Width           =   1110
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   -74730
         Picture         =   "frSelect.frx":1599
         Top             =   3090
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   465
         Left            =   -74715
         Picture         =   "frSelect.frx":19DB
         Stretch         =   -1  'True
         Top             =   720
         Width           =   450
      End
      Begin VB.Label Label3 
         Caption         =   "Si Ud. desea imprimir una hoja de trabajo sobre la selección de registros que acaba de hacer, sólo pulse el botón Imprimir"
         Height          =   795
         Left            =   -74040
         TabIndex        =   25
         Top             =   3075
         Width           =   3465
      End
      Begin VB.Label Label2 
         Caption         =   "La información seleccionada se agrupará y combinará respectivamente con el proceso que está ejecutando."
         Height          =   870
         Left            =   -74130
         TabIndex        =   24
         Top             =   690
         Width           =   3510
      End
   End
   Begin VB.Image Image1 
      Height          =   3975
      Left            =   75
      Picture         =   "frSelect.frx":1D1D
      Top             =   150
      Width           =   2355
   End
End
Attribute VB_Name = "frSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Temporales que se utilizan en este formulario
'##TMPSELECT2
'##TMPSELECT
Option Explicit
Dim RSTRAB As New ADODB.Recordset
Dim XITEM As ListItem

Private Sub CMADICIONA_Click()
    Dim RSAUX2 As New ADODB.Recordset
    RSAUX2.Open "SELECT CODTRAB, NOMBRES FROM VWTRABAJ WHERE CODTRAB NOT IN (SELECT CODTRAB FROM  [##TMPSELECT2" & VGL_COMPUTER & "] )", DBSYSTEM, adOpenStatic
    frmComun.CONECTAR RSAUX2
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        DBSYSTEM.Execute "INSERT INTO [##TMPSELECT2" & VGL_COMPUTER & "] (CODTRAB , NOMBRES , FECHAING , AREA ,CENTROCOSTO , TIPOTRAB, BASICO) SELECT CODTRAB, NOMBRES, FECHAING, CODAREA, CODCCOSTO, TIPOTRAB, BASICO FROM " & REGSISTEMA.BASESQL & ".dbo.VWTRABAJ WHERE CODTRAB='" & VGUTIL(1) & "'"
    End If
    If RSTRAB.State <> 0 Then
        RSTRAB.Requery
        Set dgTrabs.DataSource = RSTRAB
        xNumTrab.Caption = RSTRAB.RecordCount & " TRABS"
    End If
    Set RSAUX2 = Nothing
End Sub

Private Sub CMATRAS_Click()
    If SSTab1.Tab = 0 Then Exit Sub
    SSTab1.Tab = SSTab1.Tab - 1
    cmSiguiente.Caption = "&Siguiente >>"
    If SSTab1.Tab = 0 Then cmAtras.Enabled = False
End Sub

Private Sub CMCANCELAR_CLICK()
    If Not ExisteTablaAux(" [##TMPSELECT" & VGL_COMPUTER & "] ") Then DBSTARPLAN.Execute "CREATE TABLE  [##TMPSELECT" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8), NOMBRES VARCHAR(50), FECHAING DATETIME, AREA VARCHAR(10), CENTROCOSTO VARCHAR(10), TIPOTRAB VARCHAR(2), BASICO  Numeric(20,2),BASICO1 NUMERIC(20,2) )"
    VPTRASPRM = "CANCEL"
    Unload Me
End Sub
Private Sub CMIMPRIMIR_CLICK()
    If Not ExisteTablaAux(" [##TMPSELECT2" & VGL_COMPUTER & "] ") Then
        MsgBox "NO SE HA EJECUTADO EL PASO 3, VUELVA A LA FICHA NÚMERO TRES PARA PREVISUALIZAR Y CARGAR LOS REGISTROS A IMPRIMIR", vbCritical
        Exit Sub
    End If
        DBSTARPLAN.Execute "EXECUTE [SP_TRABAJADOR] ##TMPSELECT2" & VGL_COMPUTER
    With RPT
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0034.RPT"
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=" & VGL_BASE & ""
        .StoredProcParam(0) = "##TMPSELECT2" & VGL_COMPUTER & ""
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowTitle = "PLAN0034 "
        .WindowShowPrintSetupBtn = True
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Action = 1
    End With
End Sub

Private Sub CMQUITAR_CLICK()
    On Error GoTo ERRQUITAR
    If RSTRAB.EOF Then Exit Sub
    Dim XBOOK As Variant
    For Each XBOOK In dgTrabs.SelBookmarks
        'Aqui hay un horror con el sgte mensaje
        'El nombre de objeto '##TMPSELECT2PC02' no es válido.
        RSTRAB.Bookmark = XBOOK
        DBSTARPLAN.Execute "DELETE FROM  [##TMPSELECT2" & VGL_COMPUTER & "]  WHERE CODTRAB='" & RSTRAB!CODTRAB & "'"
    Next
    RSTRAB.Requery
    Set dgTrabs.DataSource = RSTRAB
    xNumTrab.Caption = RSTRAB.RecordCount & " TRABS"
    Exit Sub
ERRQUITAR:
    Resume Next
End Sub

Private Sub CMSIGUIENTE_Click()
    If SSTab1.Tab = 3 Then 'FINALIZAR
        If Not ExisteTablaAux("[##TMPSELECT" & VGL_COMPUTER & "]") Then
            DBSTARPLAN.Execute "CREATE TABLE  [##TMPSELECT" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8), NOMBRES VARCHAR(50), FECHAING DATETIME, AREA VARCHAR(10), CENTROCOSTO VARCHAR(10), TIPOTRAB VARCHAR(2), BASICO  Numeric(20,2), BASICO1  Numeric(20,2) )"
        Else
            DBSTARPLAN.Execute "DELETE FROM [##TMPSELECT" & VGL_COMPUTER & "] "
        End If
        SQL = "INSERT INTO  [##TMPSELECT" & VGL_COMPUTER & "]  SELECT * FROM  [##TMPSELECT2" & VGL_COMPUTER & "] "
        DBSTARPLAN.Execute SQL
        Unload Me
    Else
        SSTab1.Tab = SSTab1.Tab + 1
        cmSiguiente.Caption = IIf(SSTab1.Tab = 3, "&Finalizar", "&Siguiente >>")
    End If
    If SSTab1.Tab > 0 Then cmAtras.Enabled = True
End Sub

Private Sub cmVacaciones_Click()
    If Not ExisteTablaAux(" [##TMPSELECT" & VGL_COMPUTER & "] ") Then
        DBSTARPLAN.Execute "CREATE TABLE  [##TMPSELECT" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8), NOMBRES VARCHAR(50), FECHAING DATETIME, AREA VARCHAR(10), CENTROCOSTO VARCHAR(10), TIPOTRAB VARCHAR(2), BASICO  Numeric(20,2),BASICO1  Numeric(20,2) )"
    Else
        DBSYSTEM.Execute "DELETE FROM  [##TMPSELECT" & VGL_COMPUTER & "] "
    End If
    DBSYSTEM.Execute "INSERT INTO  [##TMPSELECT" & VGL_COMPUTER & "]  SELECT CODTRAB, NOMBRES, FECHAING, CODAREA AS AREA, CODCCOSTO AS CENTROCOSTO, TIPOTRAB, BASICO FROM VWTRABAJ " & IIf(xCondFecha.ListIndex = 0, "", " WHERE FECHAING<=" & DateSQL(xFecha.Value))
    DBSYSTEM.Execute "DELETE FROM  [##TMPSELECT" & VGL_COMPUTER & "]  WHERE CODTRAB NOT IN (SELECT CODTRAB FROM HISTOVAC WHERE CERRADO=0)"
    Unload Me
End Sub

Private Sub Command1_Click() 'TODOS LOS TRABAJADORES
'    If Not ExisteTablaAux(" [##TMPSELECT" & VGL_COMPUTER & "] ") Then DBSTARPLAN.Execute "CREATE TABLE  [##TMPSELECT" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8), NOMBRES VARCHAR(50), FECHAING DATETIME, AREA VARCHAR(10), CENTROCOSTO VARCHAR(10), TIPOTRAB VARCHAR(2), BASICO  Numeric(20,2), FECHACESE DATETIME )"
     If Not ExisteTablaAux(" [##TMPSELECT" & VGL_COMPUTER & "] ") Then DBSTARPLAN.Execute "CREATE TABLE  [##TMPSELECT" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8), NOMBRES VARCHAR(50), FECHAING DATETIME, AREA VARCHAR(10), CENTROCOSTO VARCHAR(10), TIPOTRAB VARCHAR(2), BASICO  Numeric(20,2),BASICO1 NUMERIC(20,2))"
    DBSYSTEM.Execute "DELETE FROM  [##TMPSELECT" & VGL_COMPUTER & "] "
'    DBSYSTEM.Execute "INSERT INTO  [##TMPSELECT" & VGL_COMPUTER & "]  SELECT CODTRAB, NOMBRES, FECHAING, CODAREA AS AREA, CODCCOSTO AS CENTROCOSTO, TIPOTRAB, BASICO,FECHACESE FROM " & REGSISTEMA.BASESQL & ".dbo.VWTRABAJ" & IIf(xCondFecha.ListIndex = 0, "", " WHERE FECHAING<=" & DateSQL(xFecha.Value))
    DBSYSTEM.Execute "INSERT INTO  [##TMPSELECT" & VGL_COMPUTER & "]  SELECT CODTRAB, NOMBRES, FECHAING, CODAREA AS AREA, CODCCOSTO AS CENTROCOSTO, TIPOTRAB, BASICO  FROM " & REGSISTEMA.BASESQL & ".dbo.VWTRABAJ" & IIf(xCondFecha.ListIndex = 0, "", " WHERE FECHAING<=" & DateSQL(xFecha.Value))
    If REGSELECT.USARFECHACESE Then DBSYSTEM.Execute "DELETE FROM  [##TMPSELECT" & VGL_COMPUTER & "]  WHERE CODTRAB IN (SELECT CODTRAB FROM " & REGSISTEMA.BASESQL & ".dbo.TRABAJADORES WHERE FECHACESE<" & DateSQL(REGSELECT.FECHAINI) & ")"
    Unload Me
End Sub

Private Sub dgTrabs_HeadClick(ByVal COLINDEX As Integer)
    Static COL As Integer
    If COL = COLINDEX Then
        If Trim(Right(RSTRAB.Sort, 4)) = "ASC" Then
            RSTRAB.Sort = dgTrabs.Columns(COLINDEX).DataField & " DESC"
          Else
            RSTRAB.Sort = dgTrabs.Columns(COLINDEX).DataField & " ASC "
        End If
        Exit Sub
    End If
    RSTRAB.Sort = dgTrabs.Columns(COLINDEX).DataField & " ASC "
    COL = COLINDEX
End Sub

Private Sub Form_Load()
    VPTRASPRM = "OK"
    CARGATIPOS
    XTODOSTRAB_Click
    xFecha.Value = Date
    xFecha2.Value = Date
    xCondFecha.ListIndex = 0
    'IMAGE1.PICTURE = LOADRESPICTURE(102, 0)
    SSTab1.Tab = 0
    If REGSELECT.USARFECHACESE Then
        xCondFecha.ListIndex = 6
        xFecha.Value = REGSELECT.FECHACESEMAX
    End If
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSTRAB = Nothing
End Sub

Private Sub OP1_Click()
    Frame2.Caption = "Area de Trabajo"
    xTodaAC.Caption = "Todas las areas de Trabajo de la Empresa"
    xCodAC1.Text = ""
    xCodAC1.Tag = ""
    xCodAC2.Text = ""
    xCodAC2.Tag = ""
    NoAC.Value = 0
End Sub

Private Sub OP2_Click()
    Frame2.Caption = "Centro de Costo de Trabajo"
    xTodaAC.Caption = "Todos los Centros de Costo de la Empresa"
    xCodAC1.Text = ""
    xCodAC1.Tag = ""
    xCodAC2.Text = ""
    xCodAC2.Tag = ""
    NoAC.Value = 0
End Sub

Public Sub CARGATIPOS()
    Dim RSAUX As New ADODB.Recordset
    RSAUX.Open "SELECT * FROM TIPOSTRAB ORDER BY TIPTRAB", DBSYSTEM, adOpenStatic
    Lista.ListItems.Clear
    Do While Not RSAUX.EOF
        Set XITEM = Lista.ListItems.Add(, "R" & RSAUX!TIPTRAB, RSAUX!TIPTRAB & ": " & RSAUX!DESCRIP, , 1)
        RSAUX.MoveNext
    Loop
    Set RSAUX = Nothing
End Sub

Private Sub SSTAB1_Click(PREVIOUSTAB As Integer)
    Select Case SSTab1.Tab
        Case 0: cmAtras.Enabled = False: cmSiguiente.Caption = "&Siguiente  >>"
        Case 1: cmAtras.Enabled = True: cmSiguiente.Caption = "&Siguiente >>"
        Case 2: cmAtras.Enabled = True
                cmSiguiente.Caption = "&Siguiente>>"
                If PREVIOUSTAB < 2 Then GENERASQL
        Case 3: cmAtras.Enabled = True: cmSiguiente.Caption = "&Finalizar"
    End Select
End Sub

Private Sub XCODAC1_DblClick()
    Dim RsAreas As New ADODB.Recordset
    If Op1.Value Then
        RsAreas.Open "SELECT CODCCOSTO AS CODIGO, NOMBRE FROM AREASTRAB ORDER BY CODCCOSTO", DBSYSTEM, adOpenStatic
    Else
        RsAreas.Open "SELECT CODCCOSTO AS CODIGO, NOMBRE FROM CCOSTOS ORDER BY CODCCOSTO", DBSYSTEM, adOpenStatic
    End If
    If RsAreas.RecordCount = 0 Then
        MsgBox "NO SE HAN ENCONTRADO REGISTROS PARA SELECCIONAR", vbCritical
        Set RsAreas = Nothing
    End If
    frmComun.CONECTAR RsAreas, , "CODCCOSTO"
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xCodAC1.Text = VGUTIL(1) & " :  " & VGUTIL(2)
        xCodAC1.Tag = VGUTIL(1)
    End If
    Set RsAreas = Nothing
End Sub

Private Sub XCODAC2_DblClick()
    Dim RsAreas As New ADODB.Recordset
    RsAreas.Open "SELECT CODCCOSTO AS CODIGO, NOMBRE FROM AREASTRAB ORDER BY CODCCOSTO", DBSYSTEM, adOpenStatic
    If RsAreas.RecordCount = 0 Then
        MsgBox "NO SE HAN ENCONTRADO AREAS DE TRABAJO", vbCritical
        Set RsAreas = Nothing
    End If
    frmComun.CONECTAR RsAreas, , "CODCCOSTO"
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xCodAC2.Text = VGUTIL(1) & " :  " & VGUTIL(2)
        xCodAC2.Tag = VGUTIL(1)
    End If
    Set RsAreas = Nothing
End Sub

Private Sub XCONDFECHA_Click()
    xFecha.Visible = False
    xFecha2.Visible = False
    If xCondFecha.ListIndex <> 0 Then
        xFecha.Visible = True
    End If
    If xCondFecha.ListIndex = 7 Then xFecha2.Visible = True
End Sub

Private Sub XTODAAC_Click()
    If xTodaAC.Value = 1 Then
        xCodAC1.Visible = False
        xCodAC2.Visible = False
        NoAC.Visible = False
        xDepende.Visible = False
    Else
        xCodAC1.Visible = True
        xCodAC2.Visible = True
        NoAC.Visible = True
        xDepende.Visible = True
    End If
End Sub

Private Sub XTODOSTRAB_Click()
    If xTodosTrab.Value = 1 Then
        Lista.Visible = False
    Else
        Lista.Visible = True
    End If
End Sub

Public Sub GENERASQL()
    Dim X As Integer, STRCAD As String, CAD1 As String, CAD2 As String
    STRCAD = ""
    If xTodosTrab.Value = 0 Then
        STRCAD = "("
        For Each XITEM In Lista.ListItems
            If XITEM.Checked Then STRCAD = STRCAD & IIf(STRCAD = "(", "", ",") & "'" & Left(XITEM.Text, 2) & "'"
        Next
        STRCAD = STRCAD & ")"
        If STRCAD <> "()" Then STRCAD = "(TIPOTRAB " & IIf(xNoTrab.Value = 1, "NOT", "") & " IN " & STRCAD & ")" Else STRCAD = ""
    End If
    If xTodaAC.Value = 0 Then
        If xCodAC1.Tag = "" Then
            MsgBox "DEBE SELECCIONAR POR LO MENOS UNA AREA O CENTRO DE COSTO, O EN TODO CASO ACTIVE LA OPCIÓN PARA TODOS LOS REGISTROS", vbCritical
            Exit Sub
        End If
        If Op1.Value Then 'SI ES POR AREAS
            If xDepende.Value = 0 Then
                CAD1 = "(CODAREA" & IIf(NoAC.Value = 1, "<>'", "='") & xCodAC1.Tag & "')"
            Else
                CAD1 = "(CODAREA" & IIf(NoAC.Value = 1, " NOT ", "") & " LIKE '" & xCodAC1.Tag & "%')"
            End If
            CAD2 = ""
            If xCodAC2.Tag <> "" And xCodAC2.Text <> "" Then 'SI EXISTE OTRA SELECCION DE AREA O CC
                If xDepende.Value = 0 Then
                    CAD2 = "(CODAREA" & IIf(NoAC.Value = 1, "<>'", "='") & xCodAC2.Tag & "')"
                Else
                    CAD2 = "(CODAREA" & IIf(NoAC.Value = 1, " NOT ", "") & " LIKE '" & xCodAC2.Tag & "%')"
                End If
            End If
            If CAD2 <> "" Then
                CAD1 = "(" & CAD1 & " OR " & CAD2 & ")"
            End If
        Else
            If xDepende.Value = 0 Then
                CAD1 = "(CODCCOSTO" & IIf(NoAC.Value = 1, "<>'", "='") & xCodAC1.Tag & "')"
            Else
                CAD1 = "(CODCCOSTO" & IIf(NoAC.Value = 1, " NOT ", "") & " LIKE '" & xCodAC1.Tag & "%')"
            End If
            CAD2 = ""
            If xCodAC2.Tag <> "" Then 'SI EXISTE OTRA SELECCION DE AREA O CC
                If xDepende.Value = 0 Then
                    CAD2 = "(CODCCOSTO" & IIf(NoAC.Value = 1, "<>'", "='") & xCodAC2.Tag & "')"
                Else
                    CAD2 = "(CODCCOSTO" & IIf(NoAC.Value = 1, " NOT ", "") & " LIKE '" & xCodAC2.Tag & "%')"
                End If
            End If
            If CAD2 <> "" Then
                CAD1 = "(" & CAD1 & " OR " & CAD2 & ")"
            End If
        End If
    End If
    If STRCAD = "" Then
        STRCAD = CAD1
    Else
        If CAD1 <> "" Then STRCAD = STRCAD & " AND " & CAD1
    End If
    If xCondFecha.ListIndex <> 0 Then
        If xCondFecha.ListIndex = 7 Then
            CAD1 = "(FECHAING BETWEEN " & DateSQL(xFecha.Value) & " AND " & DateSQL(xFecha2.Value) & ")"
        Else
            CAD1 = "floor(cast(FECHAING as real))" & xCondFecha.Text & FechS(xFecha.Value, Sqlf)
        End If
    End If
    If STRCAD = "" Then
        STRCAD = CAD1
    Else
        If CAD1 <> "" Then STRCAD = STRCAD & " AND " & CAD1
    End If
    Set RSTRAB = Nothing
    'CORRECION DEL ERROR PARA LA TABLA ##TMPSELECT
    If ExisteTablaSQL("[##TMPSELECT" & VGL_COMPUTER & "] ", DBSYSTEM) Then
        DBSTARPLAN.Execute "DROP TABLE  [##TMPSELECT" & VGL_COMPUTER & "] "
    End If
    'HASTA AQUI
    If ExisteTablaSQL("[##TMPSELECT2" & VGL_COMPUTER & "] ", DBSYSTEM) Then
        DBSTARPLAN.Execute "DROP TABLE  [##TMPSELECT2 " & VGL_COMPUTER & "] "
    End If
    Dim TXNOCALCULO As String
    If xNoCalculo.Value Then
        TXNOCALCULO = "NOCALCULO=1"
     Else
        TXNOCALCULO = "NOCALCULO=0"
    End If
    
    Dim Baja As String
    If ctxBaja.Value Then
        Baja = " AND SITUACIÓN < 2 "
     Else
        Baja = " AND SITUACIÓN >= 0 "
    End If

    If ExisteTablaAux("[##TMPSELECT2" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "drop table  [##TMPSELECT2" & VGL_COMPUTER & "] "
    DBSTARPLAN.Execute "CREATE TABLE  [##TMPSELECT2" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8), NOMBRES VARCHAR(50), FECHAING DATETIME, AREA VARCHAR(10), CENTROCOSTO VARCHAR(10), TIPOTRAB VARCHAR(2), BASICO  Numeric(20,2), BASICO1  Numeric(20,2) )"
    DBSTARPLAN.Execute "INSERT INTO  [##TMPSELECT2" & VGL_COMPUTER & "]  (CODTRAB , NOMBRES , FECHAING , AREA ,CENTROCOSTO , TIPOTRAB, BASICO,basico1) SELECT CODTRAB, NOMBRES, FECHAING, CODAREA, CODCCOSTO, TIPOTRAB, BASICO,opciona FROM " & VGL_SERVER & "." & REGSISTEMA.BASESQL & ".dbo.VWTRABAJ " & IIf(STRCAD = "", IIf(REGSELECT.USARFECHACESE, " WHERE (ISNULL(FECHACESE) OR NOT FECHACESE<" & DateSQL(REGSELECT.FECHACESEMAX) & ") AND " & TXNOCALCULO, ""), " WHERE " & TXNOCALCULO & Baja & " AND " & STRCAD)
    If REGSELECT.USARFECHACESE Then DBSTARPLAN.Execute "DELETE FROM  [##TMPSELECT2" & VGL_COMPUTER & "]  WHERE CODTRAB IN (SELECT CODTRAB FROM " & VGL_SERVER & "." & REGSISTEMA.BASESQL & ".dbo.TRABAJADORES WHERE floor(cast(FECHACESE as real))<" & FechS(REGSELECT.FECHAINI, Sqlf) & ")"
    RSTRAB.Open " [##TMPSELECT2" & VGL_COMPUTER & "] ", DBSTARPLAN, adOpenStatic
    Set dgTrabs.DataSource = RSTRAB
    xNumTrab.Caption = RSTRAB.RecordCount & " TRABS"
End Sub


