VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#8.0#0"; "ApliCTxt.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frFchPgo2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Programación de Fechas de Pago"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   Icon            =   "frFchPg2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmImprimir 
      Caption         =   "&Imprimir"
      Height          =   330
      Left            =   2700
      TabIndex        =   21
      Top             =   6240
      Width           =   1095
   End
   Begin Crystal.CrystalReport RptPagos 
      Left            =   2385
      Top             =   2940
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.CommandButton cmBorrar 
      Caption         =   "Borrar Elemento"
      Height          =   330
      Left            =   1200
      TabIndex        =   20
      Top             =   6240
      Width           =   1365
   End
   Begin VB.CommandButton cmLimpiar 
      Caption         =   "&Limpiar"
      Height          =   315
      Left            =   120
      TabIndex        =   19
      Top             =   6240
      Width           =   945
   End
   Begin VB.CommandButton cmCerrar 
      Caption         =   "&Cerrar"
      Height          =   315
      Left            =   5085
      TabIndex        =   17
      Top             =   6255
      Width           =   1080
   End
   Begin VB.Frame Frame2 
      Caption         =   "Programación de Pagos"
      Height          =   4500
      Left            =   120
      TabIndex        =   3
      Top             =   1650
      Width           =   6045
      Begin AplisetControlText.Aplitext xDescripcion 
         Height          =   300
         Left            =   1140
         TabIndex        =   6
         Top             =   1530
         Width           =   4740
         _ExtentX        =   8361
         _ExtentY        =   529
         MaxLength       =   50
         Text            =   ""
      End
      Begin MSDataGridLib.DataGrid dgFechas 
         Height          =   2370
         Left            =   105
         TabIndex        =   16
         Top             =   1980
         Width           =   5760
         _ExtentX        =   10160
         _ExtentY        =   4180
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         AllowDelete     =   -1  'True
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
         Caption         =   "Cronograma de Pagos"
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
      Begin MSComCtl2.DTPicker xAdelanto 
         Height          =   300
         Left            =   2355
         TabIndex        =   15
         Top             =   1125
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "ddd dd MMM yyyy"
         Format          =   24576003
         CurrentDate     =   36679
      End
      Begin MSComCtl2.DTPicker xFechaPago 
         Height          =   300
         Left            =   105
         TabIndex        =   13
         Top             =   1125
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "ddd dd MMM yyyy"
         Format          =   24576003
         CurrentDate     =   36679
      End
      Begin MSComCtl2.DTPicker xFechaFin 
         Height          =   300
         Left            =   2370
         TabIndex        =   11
         Top             =   555
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "ddd dd MMM yyyy"
         Format          =   24576003
         CurrentDate     =   36679
      End
      Begin MSComCtl2.DTPicker xFechaIni 
         Height          =   300
         Left            =   120
         TabIndex        =   9
         Top             =   555
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "ddd dd MMM yyyy"
         Format          =   24576003
         CurrentDate     =   36679
      End
      Begin VB.CommandButton cmAgregar 
         Caption         =   "Agregar"
         Height          =   300
         Left            =   4605
         TabIndex        =   7
         Top             =   1125
         Width           =   1275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         Height          =   195
         Left            =   135
         TabIndex        =   18
         Top             =   1575
         Width           =   840
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Adelanto"
         Height          =   195
         Left            =   2415
         TabIndex        =   14
         Top             =   915
         Width           =   1350
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Pago"
         Height          =   195
         Left            =   150
         TabIndex        =   12
         Top             =   915
         Width           =   1095
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Termino"
         Height          =   195
         Left            =   2385
         TabIndex        =   10
         Top             =   315
         Width           =   1065
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicio"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   315
         Width           =   870
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Area de Trabajo"
      Height          =   1320
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6045
      Begin VB.TextBox xCCosto 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1455
         TabIndex        =   2
         Top             =   360
         Width           =   4290
      End
      Begin VB.Label xConfig 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label3"
         Height          =   285
         Left            =   1455
         TabIndex        =   5
         Top             =   795
         Width           =   4290
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Configuración"
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Area de Trabajo"
         Height          =   195
         Left            =   180
         TabIndex        =   1
         Top             =   405
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frFchPgo2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsCrono As New ADODB.Recordset
Dim xCfg(2) As Byte

Public Sub CargaCrono(ByVal vCosto As String)
    Dim RsAux As New ADODB.Recordset
    RsAux.Open "SELECT * FROM AreasTrab WHERE Cronograma=1 AND CodCCosto='" & vCosto & "'", DbSystem, adOpenStatic
    If RsAux.EOF Or RsAux.RecordCount = 0 Then
        MsgBox "El Area de Trabajo seleccionado no se encuentra configurado para aceptar Cronograma de Pagos de Remuneraciones", vbInformation
        Unload Me
        Exit Sub
    End If
    xCCosto.Text = RsAux!Nombre
    xCCosto.Tag = RsAux!CodCCosto
    xConfig.Caption = Choose(RsAux!tipo + 1, "Semanal", "Quincenal", "Cada # dias", "Mensual")
    If RsAux!tipo = 2 Then xConfig.Caption = "Cada " & RsAux!NumDias & " dias"
    If RsAux!tipo = 0 Then xConfig.Caption = xConfig.Caption & ". Comienza la semana el " & Choose(RsAux!Inidia + 1, "Domingo", "Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sábado")
    If RsAux!Adelantos = 0 Then xAdelanto.Enabled = False Else xAdelanto.Enabled = True
    xCfg(0) = RsAux!tipo
    xCfg(1) = RsAux!Inidia
    xCfg(2) = RsAux!NumDias
    Set RsAux = Nothing
    xFechaIni_CloseUp
    RsCrono.Open "SELECT * FROM FechaPago2 WHERE CodCCosto='" & vCosto & " ' ORDER BY FechaIni", DbSystem, adOpenKeyset
    Set dgFechas.DataSource = RsCrono
    FormatDG
End Sub

Private Sub cmAgregar_Click()
    If xDescripcion.Text = "" Then
        MsgBox "Debe ingresar una descripción para el presente registro", vbCritical
        xDescripcion.SetFocus
        Exit Sub
    End If
    DbSystem.Execute "INSERT INTO FechaPago2 (CodCCosto, FechaIni, FechaFin, Adelanto, FechaPago, Nombre) VALUES ('" & xCCosto.Tag & "', " & DateSQL(xFechaIni.Value) & "," & DateSQL(xFechaFin.Value) & ", " & DateSQL(xAdelanto.Value) & "," & DateSQL(xFechaPago.Value) & ",'" & xDescripcion.Text & "')"
    RsCrono.Requery
    If RsCrono.RecordCount > 0 Then RsCrono.MoveLast
    FormatDG
    xFechaIni.Value = DateAdd("d", 1, xFechaFin.Value)
    xFechaIni_CloseUp
End Sub

Private Sub cmBorrar_Click()
    If MsgBox("Realmente desea eliminar el registro seleccionado", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    DbSystem.Execute "DELETE FROM FechaPago2 WHERE ID_FechaPago=" & RsCrono!ID_FechaPago
    RsCrono.Requery
    FormatDG
End Sub

Private Sub cmCerrar_Click()
    Unload Me
End Sub

Private Sub cmImprimir_Click()
    With RptPagos
        .ReportFileName = RegSistema.Reportes & "Crono.rpt"
        .DataFiles(0) = RegSistema.PathEmpresa & "\planilla.mdb"
        .SelectionFormula = ""
        If MsgBox("Desea imprimir todos los registros (Si) o solo el selccionado (No)", vbYesNo + vbQuestion) = vbNo Then .SelectionFormula = "{CCostos.CodCCosto}='" & xCCosto.Tag & "'"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .PrintReport
    End With
End Sub

Private Sub cmLimpiar_Click()
    If MsgBox("Realmente desea eliminar todos los registros mostrados para este Centro de Costo. La eliminación no afectará a los valores ingresados en Boletas de Remuneraciones", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    DbSystem.Execute "DELETE FROM FechaPago2 WHERE CodCCosto='" & xCCosto.Tag & "'"
    RsCrono.Requery
    FormatDG
End Sub

Private Sub Form_Load()
    xFechaIni.Value = CDate("01/01/" & RegSistema.Anno)
    xFechaIni_CloseUp
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RsCrono = Nothing
End Sub

Private Sub xFechaFin_CloseUp()
    Select Case xCfg(0)
        Case 0: xDescripcion.Text = " Semana del " & xFechaIni.Value & " al " & xFechaFin.Value
        Case 1: xDescripcion.Text = "Quincena del " & xFechaIni.Value & " al " & xFechaFin.Value
        Case 2: xDescripcion.Text = "Remuneración del " & xFechaIni.Value & " al " & xFechaFin.Value
        Case 3: xDescripcion.Text = "Remuneración del mes de " & AMeses(xFechaIni.Month) & " del " & xFechaIni.Year
    End Select
    xFechaPago.Value = xFechaFin.Value
    If xFechaPago.DayOfWeek = 1 Then xFechaPago.Value = DateAdd("d", -1, xFechaPago.Value)
End Sub

Private Sub xFechaIni_CloseUp()
    Select Case xCfg(0)
        Case 0 'Semanal, comienza el
            If Weekday(xFechaIni.Value) - 1 <> xCfg(1) Then
                MsgBox "La fecha especificada no corresponde al inicio de semana configurado. La semana comienza el " & Choose(xCfg(1) + 1, "Domingo", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado")
            End If
            xFechaFin.Value = DateAdd("d", 6, xFechaIni)
            If xAdelanto.Enabled Then xAdelanto.Value = DateAdd("d", 3, xFechaIni.Value)
            xDescripcion.Text = " Semana del " & xFechaIni.Value & " al " & xFechaFin.Value
        Case 1
            If Not (xFechaIni.Day = 16 Or xFechaIni.Day = 1) Then
                MsgBox "La configuración especifica pago Quincenal. Entonces cada pago debe empeezar el 1 o el 15 de cada mes", vbInformation
            End If
            xFechaFin.Value = DateAdd("d", 14, xFechaIni.Value)
            If xFechaIni.Day > 14 Then xFechaFin.Value = DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & xFechaIni.Month & "/" & xFechaIni.Year)))
            If xAdelanto.Enabled Then xAdelanto.Value = DateAdd("d", 6, xFechaIni.Value)
            xDescripcion.Text = "Quincena del " & xFechaIni.Value & " al " & xFechaFin.Value
        Case 2
            xFechaFin.Value = DateAdd("d", xCfg(2) - 1, xFechaIni.Value)
            If xAdelanto.Enabled Then xAdelanto.Value = DateAdd("d", xCfg(2) \ 2, xFechaIni.Value)
            xDescripcion.Text = "Remuneración del " & xFechaIni.Value & " al " & xFechaFin.Value
        Case 3
            If xFechaIni.Day <> 1 Then
                MsgBox "La configuración especificada para el cronograma de pagos, especifica un pago mensual, el cual deberá empezar el primer dia de cada mes"
            End If
            xFechaFin.Value = DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & xFechaIni.Month & "/" & xFechaIni.Year)))
            If xAdelanto.Enabled Then xAdelanto.Value = DateAdd("d", 14, xFechaIni.Value)
            xDescripcion.Text = "Remuneración del mes de " & AMeses(xFechaIni.Month) & " del " & xFechaIni.Year
    End Select
    xFechaPago.Value = xFechaFin.Value
    If xFechaPago.DayOfWeek = 1 Then xFechaPago.Value = DateAdd("d", -1, xFechaPago.Value)
End Sub

Public Sub FormatDG()
    With dgFechas
        .Columns("ID_FechaPAgo").Visible = False
        .Columns("CodCCosto").Visible = False
        .Columns("Nombre").Width = .Columns("Nombre").Width * 2
        If Not xAdelanto.Enabled Then .Columns("Adelanto").Visible = False
        .Columns("Cerrado").Visible = False
        If RsCrono.EOF Or RsCrono.RecordCount = 0 Then
            cmBorrar.Enabled = False
            cmLimpiar.Enabled = False
        Else
            cmBorrar.Enabled = True
            cmLimpiar.Enabled = True
        End If
    End With
End Sub
