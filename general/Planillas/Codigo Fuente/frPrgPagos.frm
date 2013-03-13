VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form frPrgPagos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cronograma de Pagos de Remuneraciones"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   Icon            =   "frPrgPagos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   6795
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   5145
      TabIndex        =   20
      Top             =   4005
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.CommandButton cmGrabar 
      Caption         =   "&Grabar"
      Height          =   345
      Left            =   5145
      TabIndex        =   19
      Top             =   3555
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.CommandButton cmSalir 
      Caption         =   "&Salir"
      Height          =   345
      Left            =   5145
      TabIndex        =   18
      Top             =   5505
      Width           =   1545
   End
   Begin VB.CommandButton cmEditar 
      Caption         =   "&Editar"
      Height          =   345
      Left            =   5145
      TabIndex        =   17
      Top             =   5040
      Width           =   1545
   End
   Begin VB.CommandButton cmAgregar 
      Caption         =   "&Agregar"
      Height          =   345
      Left            =   5145
      TabIndex        =   16
      Top             =   4590
      Width           =   1545
   End
   Begin MSDataGridLib.DataGrid DGLista 
      Height          =   2295
      Left            =   135
      TabIndex        =   15
      Top             =   3555
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   4048
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
      Caption         =   "Pagos Existentes en la Base de Datos"
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
   Begin VB.Frame Frame1 
      Caption         =   "Especificaciones del Pago"
      Height          =   3375
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   6555
      Begin VB.CheckBox CHKAUX 
         Caption         =   "Corntrol  Auxiliar"
         Height          =   300
         Left            =   135
         TabIndex        =   21
         Top             =   2250
         Width           =   2085
      End
      Begin MSComCtl2.DTPicker xFechaAdelanto 
         Height          =   330
         Left            =   1395
         TabIndex        =   14
         Top             =   2850
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   582
         _Version        =   393216
         Format          =   24772609
         CurrentDate     =   36700
      End
      Begin VB.CheckBox xDarAdelanto 
         Caption         =   "Activar Adelanto de Pago"
         Height          =   195
         Left            =   135
         TabIndex        =   12
         Top             =   2550
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker xFechaPago 
         Height          =   330
         Left            =   1395
         TabIndex        =   11
         Top             =   1845
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   582
         _Version        =   393216
         Format          =   24772609
         CurrentDate     =   36700
      End
      Begin MSComCtl2.DTPicker xFechaFin 
         Height          =   330
         Left            =   1395
         TabIndex        =   9
         Top             =   1455
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   582
         _Version        =   393216
         Format          =   24772609
         CurrentDate     =   36700
      End
      Begin MSComCtl2.DTPicker xFechaIni 
         Height          =   330
         Left            =   1395
         TabIndex        =   7
         Top             =   1065
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   582
         _Version        =   393216
         Format          =   24772609
         CurrentDate     =   36700
      End
      Begin MSComCtl2.DTPicker xMes 
         Height          =   330
         Left            =   1395
         TabIndex        =   5
         Top             =   675
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "MMMM'del ' yyyy"
         Format          =   24772611
         CurrentDate     =   36700
      End
      Begin AplisetControlText.Aplitext xNombre 
         Height          =   315
         Left            =   1395
         TabIndex        =   2
         Top             =   330
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   556
         MaxLength       =   50
         Text            =   ""
      End
      Begin MSComCtl2.MonthView VistaMes 
         Height          =   2370
         Left            =   3420
         TabIndex        =   3
         Top             =   675
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483626
         Appearance      =   1
         MaxSelCount     =   31
         MonthBackColor  =   16777215
         MultiSelect     =   -1  'True
         ShowWeekNumbers =   -1  'True
         StartOfWeek     =   24772609
         TitleBackColor  =   -2147483635
         TitleForeColor  =   -2147483639
         TrailingForeColor=   32896
         CurrentDate     =   36526
         MinDate         =   2
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Adelanto"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   2925
         Width           =   1125
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Pago"
         Height          =   195
         Left            =   135
         TabIndex        =   10
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final"
         Height          =   195
         Left            =   135
         TabIndex        =   8
         Top             =   1530
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicio"
         Height          =   195
         Left            =   135
         TabIndex        =   6
         Top             =   1140
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Mes de Cargo"
         Height          =   195
         Left            =   135
         TabIndex        =   4
         Top             =   750
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Left            =   135
         TabIndex        =   1
         Top             =   420
         Width           =   555
      End
   End
End
Attribute VB_Name = "frPrgPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents RSCRONO As ADODB.Recordset
Attribute RSCRONO.VB_VarHelpID = -1

Private Sub CMAGREGAR_CLICK()
    cmSalir.Visible = False
    cmAgregar.Visible = False
    cmEditar.Visible = False
    cmGrabar.Visible = True
    cmCancelar.Visible = True
    DGLista.Visible = False
    xNombre.Text = ""
    xFechaIni.Value = Date
    xFechaFin.Value = Date
    xFechaPago.Value = Date
    xDarAdelanto.Value = 0
    xFechaAdelanto.Visible = False
    Frame1.Enabled = True
    xNombre.Tag = "AGREGAR"
    xNombre.SetFocus
End Sub

Private Sub CMCANCELAR_CLICK()
    cmSalir.Visible = True
    cmAgregar.Visible = True
    cmEditar.Visible = True
    cmGrabar.Visible = False
    cmCancelar.Visible = False
    If Not RSCRONO.EOF Then RSCRONO.MoveFirst
    Frame1.Enabled = False
    DGLista.Visible = True
End Sub

Private Sub CMEDITAR_CLICK()
    Frame1.Enabled = True
    cmSalir.Visible = False
    cmAgregar.Visible = False
    cmEditar.Visible = False
    cmGrabar.Visible = True
    cmCancelar.Visible = True
    xNombre.Tag = "EDITAR"
    DGLista.Visible = False
End Sub

Private Sub CMGRABAR_CLICK()
    If xNombre.Text = "" Then
        MsgBox "No es posible grabar los datos si el nombre del pago no tiene datos", vbCritical
        Exit Sub
    End If
    If xFechaIni.Value > xFechaFin.Value Then
        MsgBox "Error de usuario: La fecha de inicio no puede ser mayor o igual a la fecha final", vbCritical
        Exit Sub
    End If
    If xDarAdelanto.Value = 1 Then
        If Not (xFechaAdelanto.Value > xFechaIni.Value And xFechaAdelanto.Value < xFechaFin.Value) Then
            MsgBox "Error de usuario: La fecha del adelanto de remuneraciones no se encuentra dentro del rango de fechas de inicio y final", vbCritical
            Exit Sub
        End If
    End If
    If Not (xMes.Month = xFechaIni.Month And xMes.Year = xFechaIni.Year) Then
        MsgBox "La fecha de inicio del periodo de pago no  se encuentra dentro del mes de cargo. Debera corregir este problema para poder grabar los datos", vbCritical
    End If
    If xNombre.Tag = "EDITAR" Then
        DBSYSTEM.Execute "UPDATE NOMBOL SET NOMBRE='" & xNombre.Text & "', MES=" & DateSQL(xMes.Value) & ",FECHAINI=" & DateSQL(xFechaIni.Value) & ",FECHAFIN=" & DateSQL(xFechaFin.Value) & ",FECHAPAGO=" & DateSQL(xFechaPago.Value) & ",DARADELANTO=" & xDarAdelanto.Value & ",FECHAADELANTO=" & DateSQL(xFechaAdelanto.Value) & ",ULTMES=" & IIf(CHKAUX = 0, 0, 1) & " WHERE CODIGO=" & RSCRONO!Codigo
    Else
        DBSYSTEM.Execute "INSERT INTO NOMBOL (NOMBRE,MES,FECHAINI,FECHAFIN,FECHAPAGO,DARADELANTO,FECHAADELANTO,CERRADO,ULTMES) VALUES ('" & xNombre.Text & "'," & DateSQL(xMes.Value) & "," & DateSQL(xFechaIni.Value) & "," & DateSQL(xFechaFin.Value) & "," & DateSQL(xFechaPago.Value) & "," & xDarAdelanto.Value & "," & DateSQL(xFechaAdelanto.Value) & ",0," & IIf(CHKAUX = 0, 0, 1) & ")"
    End If
    RSCRONO.Requery
    Set DGLista.DataSource = RSCRONO
    cmSalir.Visible = True
    cmAgregar.Visible = True
    cmEditar.Visible = True
    cmGrabar.Visible = False
    cmCancelar.Visible = False
    Frame1.Enabled = False
    DGLista.Visible = True
End Sub

Private Sub CMSALIR_CLICK()
    Unload Me
End Sub
Private Sub Form_Load()
    Set RSCRONO = New ADODB.Recordset
    RSCRONO.Open "SELECT * FROM NOMBOL WHERE CERRADO=0 ORDER BY FECHAINI, FECHAFIN DESC", DBSYSTEM, adOpenStatic
    Set DGLista.DataSource = RSCRONO
    Frame1.Enabled = False
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSCRONO = Nothing
End Sub

Private Sub RSCRONO_MOVECOMPLETE(ByVal ADREASON As ADODB.EventReasonEnum, ByVal PERROR As ADODB.Error, ADSTATUS As ADODB.EventStatusEnum, ByVal PRECORDSET As ADODB.Recordset)
    On Error GoTo ERRVISTA
    If PRECORDSET.EOF Then Exit Sub
    With PRECORDSET
        xNombre.Text = "" & !NOMBRE
        xMes.Value = !MES
        xFechaIni.Value = !FECHAINI
        xFechaFin.Value = !FECHAFIN
        xFechaPago.Value = !FECHAPAGO
        xDarAdelanto.Value = !DARADELANTO
        xFechaAdelanto.Value = !FECHAADELANTO
        CHKAUX.Value = IIf(ESNULO(!ULTMES, 0) = 0, 0, 1)
        If xDarAdelanto.Value = 0 Then xFechaAdelanto.Visible = False Else xFechaAdelanto.Visible = True
        VistaMes.Value = xFechaIni.Value
        VistaMes.SelStart = xFechaIni.Value
        VistaMes.SelEnd = xFechaFin.Value
    End With
    Exit Sub
ERRVISTA:
    Resume Next
End Sub

Private Sub XDARADELANTO_CLICK()
    If xDarAdelanto.Value = 0 Then
        xFechaAdelanto.Visible = False
    Else
        xFechaAdelanto.Visible = True
    End If
End Sub

Private Sub XFECHAFIN_CHANGE()
    xFechaPago.Value = xFechaFin.Value
End Sub

Private Sub XFECHAINI_CHANGE()
    xFechaFin.Value = xFechaIni.Value
End Sub

Private Sub XMES_CHANGE()
    xMes.Day = 1
    xFechaIni.Value = xMes.Value
    xFechaFin.Value = DateAdd("D", -1, DateAdd("M", 1, xFechaIni.Value))
    xFechaPago.Value = xFechaFin.Value
    VistaMes.Value = xFechaIni.Value
End Sub

