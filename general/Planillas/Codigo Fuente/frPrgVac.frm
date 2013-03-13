VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form frPrgVac 
   Caption         =   "Programación de Vacaciones"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6615
   Icon            =   "frPrgVac.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Periodo de Cálculo"
      Height          =   1770
      Left            =   2355
      TabIndex        =   24
      Top             =   4230
      Width           =   2145
      Begin MSComCtl2.DTPicker xFechaIniCal 
         Height          =   300
         Left            =   225
         TabIndex        =   25
         Top             =   585
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         Format          =   24707073
         CurrentDate     =   36697
      End
      Begin MSComCtl2.DTPicker xFechaFinCal 
         Height          =   300
         Left            =   225
         TabIndex        =   26
         Top             =   1215
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         Format          =   24707073
         CurrentDate     =   36697
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Inicio"
         Height          =   195
         Left            =   225
         TabIndex        =   2
         Top             =   330
         Width           =   1095
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final"
         Height          =   195
         Left            =   225
         TabIndex        =   19
         Top             =   960
         Width           =   825
      End
   End
   Begin VB.CommandButton cmQuitar 
      Caption         =   "&Quitar Trabajador"
      Height          =   375
      Left            =   4785
      TabIndex        =   16
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton cmAgregar 
      Caption         =   "Agregar Trabajador"
      Height          =   375
      Left            =   4785
      TabIndex        =   15
      Top             =   4710
      Width           =   1695
   End
   Begin VB.CommandButton cmCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   4785
      TabIndex        =   14
      Top             =   5625
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid dgLista 
      Height          =   2400
      Left            =   75
      TabIndex        =   13
      Top             =   165
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   4233
      _Version        =   393216
      AllowUpdate     =   0   'False
      ColumnHeaders   =   -1  'True
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
      Caption         =   "Trabajadores Programados para Vacaciones"
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
   Begin VB.CheckBox xAlerta 
      Caption         =   "&Alerta antes de 7 dias de inicio del periodo"
      Height          =   240
      Left            =   3195
      TabIndex        =   12
      Top             =   6135
      Width           =   3315
   End
   Begin VB.CommandButton cmActualizar 
      Caption         =   "&Actualizar"
      Height          =   375
      Left            =   4785
      TabIndex        =   11
      Top             =   4260
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Periodo Vacacional"
      Height          =   2160
      Left            =   75
      TabIndex        =   4
      Top             =   4230
      Width           =   2145
      Begin MSComCtl2.DTPicker xFechaIni 
         Height          =   300
         Left            =   225
         TabIndex        =   9
         Top             =   585
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         Format          =   24707073
         CurrentDate     =   36697
      End
      Begin MSComCtl2.DTPicker xFechaFin 
         Height          =   300
         Left            =   225
         TabIndex        =   10
         Top             =   1215
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         Format          =   24707073
         CurrentDate     =   36697
      End
      Begin VB.Label xTotalDias 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Top             =   1650
         Width           =   585
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Total dias"
         Height          =   195
         Left            =   210
         TabIndex        =   7
         Top             =   1710
         Width           =   690
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final"
         Height          =   195
         Left            =   225
         TabIndex        =   6
         Top             =   960
         Width           =   825
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Inicio"
         Height          =   195
         Left            =   225
         TabIndex        =   5
         Top             =   330
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Definición"
      Height          =   1500
      Left            =   75
      TabIndex        =   0
      Top             =   2610
      Width           =   6435
      Begin MSComCtl2.DTPicker xFechaIng 
         Height          =   300
         Left            =   4785
         TabIndex        =   18
         Top             =   720
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   529
         _Version        =   393216
         Format          =   24707073
         CurrentDate     =   36709
      End
      Begin AplisetControlText.Aplitext xPeriodo 
         Height          =   285
         Left            =   1380
         TabIndex        =   27
         Top             =   675
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   503
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xTrab 
         Height          =   285
         Left            =   1380
         TabIndex        =   28
         Top             =   330
         Width           =   4905
         _ExtentX        =   8652
         _ExtentY        =   503
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin VB.Label xBasico 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   300
         Left            =   4785
         TabIndex        =   23
         Top             =   1065
         Width           =   1515
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Básico"
         Height          =   195
         Left            =   3870
         TabIndex        =   22
         Top             =   1065
         Width           =   480
      End
      Begin VB.Label xArea 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1395
         TabIndex        =   21
         Top             =   1050
         Width           =   2280
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Area de Trabajo"
         Height          =   195
         Left            =   195
         TabIndex        =   20
         Top             =   1095
         Width           =   1140
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Ing."
         Height          =   195
         Left            =   3855
         TabIndex        =   17
         Top             =   780
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Periodo"
         Height          =   195
         Left            =   195
         TabIndex        =   3
         Top             =   765
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Trabajador"
         Height          =   195
         Left            =   195
         TabIndex        =   1
         Top             =   435
         Width           =   765
      End
   End
End
Attribute VB_Name = "frPrgVac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents RSLISTA As ADODB.Recordset
Attribute RSLISTA.VB_VarHelpID = -1
Dim SWAGREGAR As Boolean
Dim RSAUX As New ADODB.Recordset

Private Sub CMACTUALIZAR_CLICK()
    If SWAGREGAR Then
        Set RSAUX = Nothing
        RSAUX.Open "SELECT * FROM HISTOVAC WHERE CODTRAB='" & xTrab.Tag & "' AND CERRADO=0", DbSystem, adOpenStatic
        If RSAUX.RecordCount > 0 Then
            MsgBox "EL TRABAJADOR YA TIENE UNA PROGRAMACIÓN DE VACACIONES PENDIENTE, POR FAVOR SELECCIONE OTRO TRABAJADOR O EDITE LA PROGRAMACIÓN ANTERIORMENTE GRABADA DEL TRABAJADOR SELECCIONADO", vbCritical
            Exit Sub
        End If
        If (DateDiff("D", xFechaIng.Value, xFechaIni.Value)) < 365 Then
            If MsgBox("DESDE LA FECHA DE INGRESO AL PERIODO INDICADO NO HA TRANSCURRIDO 365 DIAS. DESEA CONTINUAR", vbYesNo + vbQuestion) = vbNo Then Exit Sub
        End If
        DbSystem.Execute "INSERT INTO HISTOVAC (CODTRAB,PERIODO,FECHAING,AREA,BASICO,FECHAINI,FECHAFIN,DIAS,FECHAREG,CERRADO,FECHAINICAL,FECHAFINCAL) VALUES ('" & xTrab.Tag & "','" & xPeriodo.Text & "'," & DateSQL(xFechaIng.Value) & ",'" & xArea.Tag & "'," & xBasico.Caption & "," & DateSQL(xFechaIni.Value) & "," & DateSQL(xFechaFin.Value) & "," & xTotalDias.Caption & "," & DateSQL(Date) & ",0," & DateSQL(xFechaIniCal.Value) & "," & DateSQL(xFechaFinCal.Value) & ")"
        SWAGREGAR = False
        dgLista.Visible = True
        cmQuitar.Visible = True
        cmAgregar.Caption = "AGREGAR TRABAJADOR"
    Else
        DbSystem.Execute "DELETE FROM HISTOVAC WHERE CODIGO=" & RSLISTA!CODIGO
        DbSystem.Execute "INSERT INTO HISTOVAC (CODTRAB,PERIODO,FECHAING,AREA,BASICO,FECHAINI,FECHAFIN,DIAS,FECHAREG,CERRADO,FECHAINICAL,FECHAFINCAL) VALUES ('" & xTrab.Tag & "','" & xPeriodo.Text & "'," & DateSQL(xFechaIng.Value) & ",'" & xArea.Tag & "'," & xBasico.Caption & "," & DateSQL(xFechaIni.Value) & "," & DateSQL(xFechaFin.Value) & "," & xTotalDias.Caption & "," & DateSQL(Date) & ",0," & DateSQL(xFechaIniCal.Value) & "," & DateSQL(xFechaFinCal.Value) & ")"
    End If
    REFRESCARGRID
    dgLista.Visible = True
End Sub

Private Sub CMAGREGAR_Click()
    If cmAgregar.Caption = "AGREGAR TRABAJADOR" Then
        xTrab.Text = ""
        xTrab.Tag = ""
        XTRAB_DblClick
        dgLista.Visible = False
        cmQuitar.Visible = False
        cmAgregar.Caption = "CANCELAR"
        SWAGREGAR = True
    Else 'SI ES CANCELAR
        SWAGREGAR = False
        dgLista.Visible = True
        cmQuitar.Visible = True
        cmAgregar.Caption = "AGREGAR TRABAJADOR"
    End If
End Sub

Private Sub CMCERRAR_CLICK()
    Unload Me
End Sub

Private Sub CMQUITAR_Click()
    If RSLISTA.EOF Then
        MsgBox "ERROR DE USUARIO: LA LISTA SE ENCUENTRA VACIA", vbCritical
        Exit Sub
    End If
    If MsgBox("REALMENTE DESEA QUITAR DE LA PROGRAMACIÓN DE VACACIONES AL TRABAJADOR: " & RSLISTA!NOMBRES, vbYesNo + vbQuestion) = vbNo Then Exit Sub
    DbSystem.Execute "DELETE FROM HISTOVAC WHERE CODIGO=" & RSLISTA!CODIGO
    REFRESCARGRID
End Sub

Private Sub Form_Load()
    Set RSLISTA = New ADODB.Recordset
    xFechaIni.Value = Date
    xFechaIni.Day = 1
    xFechaFin.Value = DateAdd("M", 1, xFechaIni.Value)
    xFechaFin.Value = DateAdd("D", -1, xFechaFin.Value)
    XFECHAINICAL_CHANGE
    RSLISTA.Open "SELECT NOMBRES, NOMBREAREA, HISTOVAC.* FROM HISTOVAC, VWTRABAJ WHERE HISTOVAC.CODTRAB=VWTRABAJ.CODTRAB AND CERRADO=0 ORDER BY NOMBRES", DbSystem, adOpenStatic
    Set dgLista.DataSource = RSLISTA
    REFRESCARGRID
    If RSLISTA.RecordCount = 0 Then CUANDOESEOF False
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSLISTA = Nothing
    Set RSAUX = Nothing
End Sub

Private Sub RSLISTA_MOVECOMPLETE(ByVal ADREASON As ADODB.EventReasonEnum, ByVal PERROR As ADODB.Error, ADSTATUS As ADODB.EventStatusEnum, ByVal PRECORDSET As ADODB.Recordset)
    If RSLISTA.EOF Then
        CUANDOESEOF False
        Exit Sub
    Else
        CUANDOESEOF True
    End If
    xTrab.Text = RSLISTA!NOMBRES
    xTrab.Tag = RSLISTA!CODTRAB
    xPeriodo.Text = RSLISTA!PERIODO
    xFechaIng.Value = RSLISTA!FECHAING
    xArea.Caption = RSLISTA!NOMBREAREA
    xArea.Tag = RSLISTA!AREA
    xBasico.Caption = Format(RSLISTA!BASICO, "0.00 ")
    xFechaIni.Value = RSLISTA!FECHAINI
    xFechaFin.Value = RSLISTA!FECHAFIN
    xTotalDias.Caption = RSLISTA!Dias
End Sub

Private Sub XFECHAINICAL_CHANGE()
    xFechaFinCal.Value = DateAdd("D", 30, xFechaIniCal.Value)
    xTotalDias = DateDiff("D", xFechaIniCal.Value, xFechaFinCal.Value) + 1
End Sub

Private Sub XTRAB_DblClick()
    Dim RSTRAB As New ADODB.Recordset
    RSTRAB.Open "SELECT CODTRAB, NOMBRES,FECHAING,BASICO,NOMBREAREA, CODAREA FROM VWTRABAJ WHERE SITUACIÓN<'2' ORDER BY CODTRAB", DbSystem, adOpenKeyset, adLockOptimistic
    If RSTRAB.EOF Or RSTRAB.RecordCount = 0 Then
        MsgBox "NO SE HAN ENCONTRADO REGISTRO DE TRABAJADORES", vbCritical
        Set RSTRAB = Nothing
        Exit Sub
    End If
    frmComun.CONECTAR RSTRAB
    frmComun.Show 1
    If vgUtil(2) <> "" Then
        xTrab.Tag = RSTRAB!CODTRAB
        xTrab.Text = RSTRAB!CODTRAB & " : " & RSTRAB!NOMBRES
        xArea.Caption = RSTRAB!NOMBREAREA
        xArea.Tag = RSTRAB!CODAREA
        xBasico.Caption = Format(RSTRAB!BASICO, "0.00 ")
        xFechaIng.Value = RSTRAB!FECHAING
        xFechaIniCal.Value = xFechaIng.Value
        xFechaIniCal.Year = Year(Date)
        xFechaIniCal.Day = 1
        xPeriodo.Text = (Year(Date) - 1) & " - " & Year(Date)
    End If
    Set RSTRAB = Nothing
    CUANDOESEOF True
End Sub

Public Sub CUANDOESEOF(ByVal HABILITADO As Boolean)
    Frame1.Visible = HABILITADO
    Frame2.Visible = HABILITADO
    cmQuitar.Enabled = HABILITADO
    cmActualizar.Enabled = HABILITADO
End Sub

Public Sub REFRESCARGRID()
    RSLISTA.Requery
    Set dgLista.DataSource = RSLISTA
    If RSLISTA.RecordCount = 0 Then CUANDOESEOF False
    With dgLista
        .Columns("CODTRAB").Visible = False
        .Columns("PERIODO").Visible = False
        .Columns("FECHAING").Visible = False
        .Columns("AREA").Visible = False
        .Columns("BASICO").Visible = False
        .Columns("DIAS").Visible = False
        .Columns("FECHAREG").Visible = False
        .Columns("CERRADO").Visible = False
        .Columns("NOMBREAREA").Visible = False
        .Columns("CODIGO").Visible = False
        .Columns(0).Width = 2610.142
    End With
End Sub

