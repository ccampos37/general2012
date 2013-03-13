VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#8.0#0"; "ApliCTxt.ocx"
Begin VB.Form frNomBol 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nombres de Planilla"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   Icon            =   "frNomBol.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   7065
   Begin MSDataGridLib.DataGrid dgNombres 
      Height          =   2115
      Left            =   120
      TabIndex        =   16
      Top             =   960
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   3731
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      Caption         =   "Nombres de Planilla"
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
            LCID            =   2058
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
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Mes de Cargo de Planilla"
      Height          =   795
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   6855
      Begin VB.CommandButton cmAgregar 
         Caption         =   "&Agregar Nombre"
         Height          =   375
         Left            =   4920
         TabIndex        =   19
         Top             =   270
         Width           =   1815
      End
      Begin VB.ComboBox xMeses 
         Height          =   315
         ItemData        =   "frNomBol.frx":030A
         Left            =   2340
         List            =   "frNomBol.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   300
         Width           =   2055
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Seleccionar mes a procesar"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   360
         Width           =   1965
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Descripción"
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   3135
      Visible         =   0   'False
      Width           =   6855
      Begin VB.CheckBox chIncluye 
         Caption         =   "Incluir trabajadores dependientes del Centro de Costo"
         Height          =   390
         Left            =   4080
         TabIndex        =   21
         Top             =   1410
         Width           =   2655
      End
      Begin VB.CommandButton cmInput 
         Caption         =   "Input de Planillas"
         Height          =   1035
         Left            =   5640
         Picture         =   "frNomBol.frx":030E
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   300
         Width           =   1095
      End
      Begin VB.CommandButton cmDelete 
         Caption         =   "&Eliminar"
         Height          =   315
         Left            =   5400
         TabIndex        =   18
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton cmUpdate 
         Caption         =   "Actualizar"
         Height          =   315
         Left            =   3990
         TabIndex        =   17
         Top             =   2160
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker xFechaFin 
         Height          =   315
         Left            =   1740
         TabIndex        =   12
         Top             =   2100
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Format          =   24641537
         CurrentDate     =   36650
      End
      Begin MSComCtl2.DTPicker xFechaIni 
         Height          =   315
         Left            =   1740
         TabIndex        =   10
         Top             =   1740
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Format          =   24641537
         CurrentDate     =   36650
      End
      Begin MSComCtl2.DTPicker xMes 
         Height          =   315
         Left            =   1740
         TabIndex        =   8
         Top             =   1380
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MMM/yyyy"
         Format          =   24641539
         CurrentDate     =   36650
      End
      Begin AplisetControlText.Aplitext xCCosto 
         Height          =   315
         Left            =   1740
         TabIndex        =   6
         Top             =   1020
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xNombre 
         Height          =   315
         Left            =   1740
         TabIndex        =   4
         Top             =   660
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         MaxLength       =   50
         Text            =   ""
      End
      Begin VB.ComboBox xTipoPlanilla 
         Height          =   315
         ItemData        =   "frNomBol.frx":0618
         Left            =   1740
         List            =   "frNomBol.frx":062B
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   285
         Width           =   3855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Término"
         Height          =   195
         Left            =   225
         TabIndex        =   11
         Top             =   2160
         Width           =   1290
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Inicio"
         Height          =   195
         Left            =   225
         TabIndex        =   9
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Mes de Cargo"
         Height          =   195
         Left            =   225
         TabIndex        =   7
         Top             =   1440
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Centro de Costo"
         Height          =   195
         Left            =   225
         TabIndex        =   5
         Top             =   1080
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Left            =   225
         TabIndex        =   3
         Top             =   720
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Planilla"
         Height          =   195
         Left            =   225
         TabIndex        =   1
         Top             =   345
         Width           =   1080
      End
   End
End
Attribute VB_Name = "FRnOMbOL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsMeses As New ADODB.Recordset
Dim WithEvents RsNombres As ADODB.Recordset
Attribute RsNombres.VB_VarHelpID = -1
Dim CargadoN As Boolean
Dim Hacer As String

Private Sub cmAgregar_Click()
    If xMeses.ListCount = 0 Then
        MsgBox "No existen meses activos para trabajar con planillas", vbCritical
        Exit Sub
    End If
    RsMeses.MoveFirst
    RsMeses.Find "nombre='" & xMeses.Text & "'"
    vpFecha = RsMeses!MesActivo
    frAddPl.Show 1
    xMeses_Click
End Sub

Private Sub cmDelete_Click()
    If xMeses.ListCount = 0 Then
        MsgBox "No existen meses activos para trabajar con planillas", vbCritical
        Exit Sub
    End If
    If RsNombres.EOF Or RsNombres.BOF Then
        MsgBox "no existe un registro de Nombre de Planilla Activo", vbCritical
        Exit Sub
    End If
    If MsgBox("Realmente desea eliminar el registro de " & RsNombres!Nombre, vbYesNo) = vbNo Then Exit Sub
    RsNombres.Delete
    xMeses_Click
End Sub

Private Sub cmInput_Click()
    On Error GoTo cmInput
    Dim RsCnpt As New ADODB.Recordset
    RsCnpt.Open "select conceptos.* from conceptos, formatos where formatos.concepto=conceptos.codigo and ccosto='" & xCCosto.Tag & "' AND Formatos.Tipo=" & xTipoPlanilla.ListIndex & " order by conceptos.tipo,fila", DbSystem, adOpenKeyset, adLockOptimistic
    If RsCnpt.RecordCount = 0 Then
        MsgBox "No se ha definido un formato de boletas para el centro de costo seleccionado", vbCritical
        If MsgBox("Desea definir un formato en estos momentos", vbYesNo + vbQuestion) = vbYes Then frFormatos.Show
        Set RsCnpt = Nothing
        Exit Sub
    End If
    'Como existen conceptos se procede a crear la tabla temporal
    Dim strCrea As String
    strCrea = "CREATE TABLE InputBol (CodTrab Char(6), Nombres Char(35)"
    With RsCnpt
        Do While Not .EOF
            If !EsEscrito Then strCrea = strCrea & ", " & !Codigo & " Single"
            .MoveNext
        Loop
    End With
    strCrea = strCrea & ")"
    DbSystem.Execute strCrea
    With RegInput
        .CentroCosto = xCCosto.Tag
        .Codigo = RsNombres!Codigo
        .MesActivo = RsNombres!Mes
        .TipoPlanilla = RsNombres!TipoPlanilla
        .Nombre = RsNombres!Nombre
        .Bol_Table = "Bol" & Format(Month(RsNombres!Mes), "00") & Format(Year(RsNombres!Mes), "0000")
        .Mov_Table = "Mov" & Format(Month(RsNombres!Mes), "00") & Format(Year(RsNombres!Mes), "0000")
    End With
    Set RsCnpt = Nothing
    InputPl.Show 1
    DbSystem.Execute "DROP TABLE InputBol"
    Exit Sub
cmInput:
    MsgBox "El sistema corregirá los errores que se produjeron al cerrar indebidamente el sistema", vbInformation
    DbSystem.Execute "DROP TABLE InputBol"
    Resume
End Sub

Private Sub cmUpdate_Click()
    If Not ComprobarData Then Exit Sub
    Dim rsNombol As New ADODB.Recordset
    If Hacer = "Nuevo" Then
        'CAMBIAR ESTO POR INSERT -
        rsNombol.Open "nombol", DbSystem, adOpenKeyset, adLockOptimistic
        With rsNombol
            .AddNew
            !TipoPlanilla = xTipoPlanilla.ListIndex
            !Nombre = xNombre.Text
            !CCosto = xCCosto.Tag
            !Mes = xMes.Value
            !FechaIni = xFechaIni.Value
            !FechaFin = xFechaFin.Value
            .Update
            xMeses.Visible = True
            cmUpdate.Visible = True
            cmDelete.Visible = True
            dgNombres.Visible = True
            MsgBox "La información del nombre de la planilla se ha grabado satisfactoriamente", vbInformation
        End With
        'FIN DEL CAMBIO POR INSERT
    Else
           'CAMBIAR ESTO POR UPDATE -
        rsNombol.Open "nombol", DbSystem, adOpenKeyset, adLockOptimistic
        rsNombol.Find "Codigo=" & RsNombres!Codigo
        If rsNombol.EOF Then
            MsgBox "Se ha eleminado el nombre de boleta desde otro usuario", vbCritical
            Set rsNombol = Nothing
            Exit Sub
        End If
        'Si en caso lo encuentra
        With rsNombol
            !TipoPlanilla = xTipoPlanilla.ListIndex
            !Nombre = xNombre.Text
            !CCosto = xCCosto.Tag
            !Mes = xMes.Value
            !FechaIni = xFechaIni.Value
            !FechaFin = xFechaFin.Value
            .Update
            MsgBox "Las modificaciones se realizaron con éxito", vbInformation
        End With
        'FIN DEL CAMBIO POR INSERT
    End If
    Hacer = "Editar"
    Set rsNombol = Nothing
    xMeses_Click
End Sub


Private Sub Form_Load()
    Set RsNombres = New ADODB.Recordset
    CargadoN = False
    Hacer = "Editar"
    RsMeses.Open "mesesact", DbSystem, adOpenKeyset, adLockOptimistic
    CargaMes
    If RsMeses.RecordCount > 0 Then xMeses.ListIndex = 0
End Sub

Public Sub CargaMes()
    xMeses.Clear
    If RsMeses.RecordCount = 0 Then Exit Sub
    RsMeses.MoveFirst
    Do While Not RsMeses.EOF
        xMeses.AddItem RsMeses!Nombre, 0
        RsMeses.MoveNext
    Loop
    If xMeses.ListCount = 0 Then cmAgregar.Enabled = False
    RsMeses.MoveFirst
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RsMeses = Nothing
    Set RsNombres = Nothing
End Sub

Private Sub RsNombres_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If Not CargadoN Then Exit Sub
    If RsNombres.RecordCount = 0 Then Exit Sub
    With RsNombres
        xTipoPlanilla.ListIndex = !TipoPlanilla
        xNombre.Text = !Nombre
        xCCosto.Text = !CodCCosto & " : " & !CentroCosto
        xCCosto.Tag = !CodCCosto
        xMes.Value = !Mes
        xFechaIni.Value = !FechaIni
        xFechaFin.Value = !FechaFin
    End With
    Frame1.Visible = True
End Sub

Private Sub xCCosto_DblClick()
    Dim RsCCostos As New ADODB.Recordset
    RsCCostos.Open "Select CodCCosto,Nombre From CCostos Order By CodCCosto", DbSystem, adOpenKeyset, adLockOptimistic
    frmComun.Conectar RsCCostos
    frmComun.Show 1
    If vgUtil(1) <> "" Then
        xCCosto.Text = vgUtil(1) & " :  " & vgUtil(2)
        xCCosto.Tag = vgUtil(1)
    End If
    Set RsCCostos = Nothing
End Sub


Private Sub xMes_Validate(Cancel As Boolean)
    Dim a
    a = RsMeses.Bookmark
    xMes.Day = 1
    RsMeses.Find "MesActivo=#" & xMes.Value & "#"
    If RsMeses.EOF Then
        MsgBox "La fecha seleccionada no se encuentra dentro de los meses activos para la edición de planillas", vbCritical
        Cancel = True
    End If
    RsMeses.Bookmark = a
End Sub

Private Sub xMeses_Click()
    Dim strCAD As String
    If xMeses.ListCount = 0 Then
        MsgBox "No se han encontrado meses activos, por favor Active un nuevo mes para poder realizar esta tarea", vbCritical
        cmAgregar.Enabled = False
        Exit Sub
    End If
    RsMeses.MoveFirst
    RsMeses.Find "nombre='" & xMeses.Text & "'"
    If RsMeses.EOF Then
        MsgBox "Se ha producido un problema de conexion con los datos de meses. Por favor ingrese de nuevo"
        Unload Me
    End If
    If CargadoN Then RsNombres.Close
    strCAD = "SELECT CCostos.CodCCosto, CCostos.Nombre as CentroCosto, NomBol.* FROM CCostos INNER JOIN NomBol ON CCostos.CodCCosto = NomBol.CCosto WHERE Mes=#" & Format(RsMeses!MesActivo, "mm/dd/yyyy") & "# ORDER BY ccostos.codccosto"
    RsNombres.Open strCAD, DbSystem, adOpenKeyset, adLockOptimistic
    If Not RsNombres.EOF Then RsNombres.MoveFirst
    Set dgNombres.DataSource = RsNombres
    CargadoN = True
    formatearGrid
    Frame1.Visible = False
End Sub

Public Sub formatearGrid()
    With dgNombres
        .Columns("Codigo").Visible = False
        .Columns("TipoPlanilla").Visible = False
        .Columns("ccosto").Visible = False
    End With
End Sub

Public Function ComprobarData() As Boolean
    ComprobarData = False
    If xTipoPlanilla.ListIndex = -1 Then
        MsgBox "No ha seleccionado un tipo de planilla", vbCritical
        xTipoPlanilla.SetFocus
        Exit Function
    End If
    If Trim(xNombre.Text) = "" Then
        MsgBox "El nombre de la planilla no es valido de acuerdo a las especificaciones de planilla", vbCritical
        xNombre.SetFocus
        Exit Function
    End If
    If xCCosto.Tag = "" Then
        MsgBox "No ha seleccionado un centro de costo para el tipo de planilla", vbCritical
        xCCosto_DblClick
        Exit Function
    End If
    xMes.Day = 1
    ComprobarData = True
End Function
