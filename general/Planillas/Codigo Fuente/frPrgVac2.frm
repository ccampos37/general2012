VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form frPrgVac2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Programación de Vacaciones"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   Icon            =   "frPrgVac2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   5745
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   3022
      TabIndex        =   23
      Top             =   5490
      Width           =   1245
   End
   Begin VB.CommandButton cmAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   1477
      TabIndex        =   22
      Top             =   5490
      Width           =   1245
   End
   Begin VB.Frame Frame3 
      Caption         =   "Periodo a Cancelar"
      Height          =   1545
      Left            =   105
      TabIndex        =   14
      Top             =   3825
      Width           =   5550
      Begin MSComCtl2.DTPicker xPerIni 
         Height          =   300
         Left            =   1395
         TabIndex        =   18
         Top             =   390
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   529
         _Version        =   393216
         Format          =   24641537
         CurrentDate     =   36844
      End
      Begin MSComCtl2.DTPicker xPerFin 
         Height          =   300
         Left            =   3765
         TabIndex        =   19
         Top             =   390
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   529
         _Version        =   393216
         Format          =   24641537
         CurrentDate     =   36844
      End
      Begin VB.Label Label11 
         Caption         =   $"frPrgVac2.frx":08CA
         Height          =   585
         Left            =   165
         TabIndex        =   20
         Top             =   825
         Width           =   5265
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   3135
         TabIndex        =   17
         Top             =   450
         Width           =   420
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   735
         TabIndex        =   16
         Top             =   450
         Width           =   465
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Salida de Vacaciones"
      Height          =   1770
      Left            =   90
      TabIndex        =   9
      Top             =   1995
      Width           =   5550
      Begin MSComCtl2.DTPicker xSalFin 
         Height          =   300
         Left            =   3780
         TabIndex        =   13
         Top             =   315
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   529
         _Version        =   393216
         Format          =   24641537
         CurrentDate     =   36844
      End
      Begin MSComCtl2.DTPicker xSalIni 
         Height          =   300
         Left            =   1410
         TabIndex        =   12
         Top             =   315
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   529
         _Version        =   393216
         Format          =   24641537
         CurrentDate     =   36844
      End
      Begin VB.Label xDias 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "30"
         Height          =   285
         Left            =   3780
         TabIndex        =   2
         Top             =   660
         Width           =   510
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Dias"
         Height          =   195
         Left            =   3150
         TabIndex        =   24
         Top             =   675
         Width           =   315
      End
      Begin VB.Label Label12 
         Caption         =   $"frPrgVac2.frx":0995
         Height          =   585
         Left            =   150
         TabIndex        =   21
         Top             =   1020
         Width           =   5250
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   3150
         TabIndex        =   11
         Top             =   375
         Width           =   420
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   750
         TabIndex        =   10
         Top             =   375
         Width           =   465
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Información del Trabajador"
      Height          =   1845
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   5550
      Begin AplisetControlText.Aplitext xTrab 
         Height          =   315
         Left            =   1650
         TabIndex        =   25
         Top             =   345
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker xUltFecha 
         Height          =   285
         Left            =   1665
         TabIndex        =   8
         ToolTipText     =   "Información otorgada por el sistema"
         Top             =   1380
         Width           =   3720
         _ExtentX        =   6562
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   24641536
         CurrentDate     =   36844
      End
      Begin MSComCtl2.DTPicker xFechaIng 
         Height          =   285
         Left            =   1665
         TabIndex        =   4
         Top             =   750
         Width           =   3720
         _ExtentX        =   6562
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   24641536
         CurrentDate     =   36844
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Ultimas Vacaciones"
         Height          =   195
         Left            =   150
         TabIndex        =   7
         Top             =   1425
         Width           =   1395
      End
      Begin VB.Label xArea 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1665
         TabIndex        =   6
         Top             =   1065
         Width           =   3720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Area de Trabajo"
         Height          =   195
         Left            =   150
         TabIndex        =   5
         Top             =   1110
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Ingreso"
         Height          =   195
         Left            =   150
         TabIndex        =   3
         Top             =   795
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Trabajador"
         Height          =   195
         Left            =   150
         TabIndex        =   1
         Top             =   480
         Width           =   765
      End
      Begin VB.Label Label8 
         Caption         =   "No registra"
         Height          =   240
         Left            =   1680
         TabIndex        =   15
         Top             =   1425
         Width           =   1380
      End
   End
End
Attribute VB_Name = "frPrgVac2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CMACEPTAR_CLICK()
    If xTrab.Tag = "" Then
        MsgBox "Deberá seleccionar un trabajador", vbInformation
        Exit Sub
    End If
    If Val(xDias.Caption) <= 0 Then
        MsgBox "El rango de fechas de inicio y salida de vacaciones es incorrecta", vbInformation
        Exit Sub
    End If
    If Val(xDias.Caption) > 30 Then
        MsgBox "El Numero de dias de Descanso Vacacional No puede execeder a 30 dias", vbInformation
        Exit Sub
    End If
    If VPTAREA <> "NUEVO" Then
        DBSYSTEM.Execute "DELETE FROM HISTOVAC WHERE CODIGO=" & VPTAREA
    End If
    'FALTA SEGURIDAD DE GRABACION
    DBSYSTEM.Execute "INSERT INTO HISTOVAC (CODTRAB, PERIODO, FECHAING, AREA, FECHAINI, FECHAFIN, DIAS, FECHAREG, CERRADO, FECHAINICAL, FECHAFINCAL, PROGRAMADO)" & _
        "VALUES ('" & xTrab.Tag & "','XXX'," & DateSQL(xFechaIng.Value) & ",'" & xArea.Tag & "'," & DateSQL(xSalIni.Value) & "," & DateSQL(xSalFin.Value) & "," & xDias.Caption & "," & DateSQL(Date) & ",0," & DateSQL(xPerIni.Value) & "," & DateSQL(xPerFin.Value) & ",1)"
    Unload Me
End Sub

Private Sub COMMAND2_CLICK()
    Unload Me
End Sub

Private Sub FORM_LOAD()
    xSalIni.Value = frVacaciones.xMes.Value
    xSalIni.Day = 1
    xSalFin.Value = DateAdd("D", 29, xSalIni.Value)
End Sub

Private Sub XSALFIN_CHANGE()
    xDias.Caption = DateDiff("D", xSalIni.Value, xSalFin.Value) + 1
End Sub

Private Sub XSALINI_CHANGE()
    xDias.Caption = DateDiff("D", xSalIni.Value, xSalFin.Value) + 1
End Sub

Private Sub XTRAB_DBLCLICK()
    Dim RSTRAB As New ADODB.Recordset
    RSTRAB.Open "SELECT CODTRAB, NOMBRES,FECHAING,NOMBREAREA, CODAREA FROM VWTRABAJ WHERE SITUACIÓN <'2' AND CODTRAB NOT IN (SELECT CODTRAB FROM HISTOVAC WHERE CERRADO=0)", DBSYSTEM, adOpenKeyset, adLockOptimistic
    If RSTRAB.RecordCount = 0 Then
        MsgBox "No se han encontrado registros de Trabajadores", vbCritical
        Set RSTRAB = Nothing
        Exit Sub
    End If
    frmComun.CONECTAR RSTRAB
    frmComun.Show 1
    If VGUTIL(2) <> "" Then
        xTrab.Tag = RSTRAB!CODTRAB
        xTrab.Text = RSTRAB!CODTRAB & " : " & RSTRAB!NOMBRES
        xArea.Caption = RSTRAB!NOMBREAREA
        xArea.Tag = RSTRAB!CODAREA
        xFechaIng.Value = RSTRAB!FECHAING
        xPerFin.Value = xFechaIng.Value
        xPerFin.Year = Year(Date)
        xPerFin.Day = 1
        xPerIni.Value = xPerFin.Value
        xPerIni.Year = xPerFin.Year - 1
        If IsNull(DevuelveValor("SELECT FECHAFIN FROM HISTOVAC, VWTRABAJ WHERE HISTOVAC.CODTRAB=VWTRABAJ.CODTRAB", DBSYSTEM)) Then
            xUltFecha.Visible = False
        Else
            xUltFecha.Visible = True
            If IsNull(DevuelveValor("SELECT MAX(FECHAFIN) AS MAX1 FROM HISTOVAC, VWTRABAJ WHERE HISTOVAC.CODTRAB=VWTRABAJ.CODTRAB AND HISTOVAC.CODTRAB='" & xTrab.Tag & "'", DBSYSTEM)) Then
                xUltFecha.Visible = False
            Else
                xUltFecha.Value = DevuelveValor("SELECT MAX(FECHAFIN) AS MAX1 FROM HISTOVAC, VWTRABAJ WHERE HISTOVAC.CODTRAB=VWTRABAJ.CODTRAB AND HISTOVAC.CODTRAB='" & xTrab.Tag & "'", DBSYSTEM)
            End If
        End If
    End If
    Set RSTRAB = Nothing
End Sub


