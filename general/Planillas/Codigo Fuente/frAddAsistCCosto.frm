VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form frAddAsistCCosto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Asistencia por Centros de Costos"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   Icon            =   "frAddAsistCCosto.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   6825
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   420
      Left            =   3637
      TabIndex        =   14
      Top             =   3270
      Width           =   1455
   End
   Begin VB.CommandButton cmAceptar 
      Caption         =   "&Aceptar"
      Height          =   420
      Left            =   1732
      TabIndex        =   13
      Top             =   3270
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Registro de Asistencia por Centro de Costos"
      Height          =   2985
      Left            =   82
      TabIndex        =   0
      Top             =   90
      Width           =   6660
      Begin AplisetControlText.Aplitext xValor 
         Height          =   315
         Left            =   1965
         TabIndex        =   12
         Top             =   2385
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         MaxLength       =   4
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xConcepto 
         Height          =   315
         Left            =   1965
         TabIndex        =   10
         Top             =   1980
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   556
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xCCosto 
         Height          =   315
         Left            =   1965
         TabIndex        =   6
         Top             =   1185
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   556
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xTrab 
         Height          =   315
         Left            =   1965
         TabIndex        =   4
         Top             =   810
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   556
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker xFecha 
         Height          =   330
         Left            =   1965
         TabIndex        =   2
         Top             =   420
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   582
         _Version        =   393216
         Format          =   24641537
         CurrentDate     =   36836
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Left            =   225
         TabIndex        =   11
         Top             =   2460
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Concepto Informativo"
         Height          =   195
         Left            =   225
         TabIndex        =   9
         Top             =   2025
         Width           =   1515
      End
      Begin VB.Label xBasico 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   315
         Left            =   1965
         TabIndex        =   8
         Top             =   1575
         Width           =   1125
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Rem. Básica"
         Height          =   195
         Left            =   225
         TabIndex        =   7
         Top             =   1650
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Centro de Costo"
         Height          =   195
         Left            =   225
         TabIndex        =   5
         Top             =   1260
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Trabajador"
         Height          =   195
         Left            =   225
         TabIndex        =   3
         Top             =   870
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Asistencia"
         Height          =   195
         Left            =   225
         TabIndex        =   1
         Top             =   480
         Width           =   1440
      End
   End
End
Attribute VB_Name = "frAddAsistCCosto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CMACEPTAR_CLICK()
    If xTrab.Tag = "" Then
        MsgBox "Falta seleccionar un Trabajdor", vbInformation
        xTrab.SetFocus
        Exit Sub
    End If
    If xCCosto.Tag = "" Then
        MsgBox "No se ha Seleccionado un Centro de Costo para el Trabajdor", vbInformation
        xCCosto.SetFocus
        Exit Sub
    End If
    If xConcepto.Text = "" Then
        MsgBox "No ha Seleccionado un Concepto de remuneración de Tipo Informativo", vbInformation
        xConcepto.SetFocus
        Exit Sub
    End If
    DBSYSTEM.Execute "DELETE FROM ASIS" & REGSISTEMA.ANNO & " WHERE CODTRAB='" & xTrab.Tag & "' AND CONCEPTO='" & xConcepto.Tag & "' AND DIA=" & DateSQL(xFecha.Value)
    DBSYSTEM.Execute "INSERT INTO ASIS" & REGSISTEMA.ANNO & " VALUES ('" & xTrab.Tag & "'," & DateSQL(xFecha.Value) & ",'" & xConcepto.Tag & "'," & xValor.Text & ",'" & xCCosto.Tag & "',1)"
End Sub
Private Sub CMCANCELAR_CLICK()
    Unload Me
End Sub
Private Sub XCCOSTO_DBLCLICK()
    If xTrab.Tag = "" Then Exit Sub
    Dim RSCOSTOS As New ADODB.Recordset
    RSCOSTOS.Open "SELECT A.CODCCOSTO, CCOSTOS.NOMBRE, A.BASICO FROM CCOSTOS, TRABXCOSTO A WHERE A.CODCCOSTO=CCOSTOS.CODCCOSTO AND A.CODTRAB='" & xTrab.Tag & "'", DBSYSTEM, adOpenStatic, adLockReadOnly
    frmComun.CONECTAR RSCOSTOS, , "A.CODCCOSTO"
    frmComun.Show 1
    If VGUTIL(2) <> "" Then
        xCCosto.Tag = RSCOSTOS!CODCCOSTO
        xCCosto.Text = RSCOSTOS!CODCCOSTO & " : " & RSCOSTOS!NOMBRE
        xBasico.Caption = Format(RSCOSTOS!BASICO, "0.00 ")
    End If
    Set RSCOSTOS = Nothing
End Sub

Private Sub XCONCEPTO_DBLCLICK()
    Dim RSCON As New ADODB.Recordset
    RSCON.Open "SELECT CODIGO, NOMBRE FROM CONCEPTOS WHERE TIPO=0 AND TIPOINFO<2", DBSYSTEM, adOpenStatic, adLockReadOnly
    If RSCON.EOF Then
        MsgBox "No se han encontrado registros de Conceptos de Remuneraciones de Tipo Informativos", vbInformation
        Set RSCON = Nothing
        Exit Sub
    End If
    frmComun.CONECTAR RSCON
    frmComun.Show 1
    If VGUTIL(2) <> "" Then
        xConcepto.Tag = RSCON!CODIGO
        xConcepto.Text = RSCON!CODIGO & " : " & RSCON!NOMBRE
    End If
    Set RSCON = Nothing
End Sub

Private Sub XFECHA_KEYDOWN(KEYCODE As Integer, SHIFT As Integer)
    If KEYCODE = 13 Then SendKeys "{TAB}"
End Sub

Private Sub FORM_LOAD()
    xFecha.Value = Date
End Sub

Private Sub XTRAB_DBLCLICK()
    Dim RSTRAB As New ADODB.Recordset
    RSTRAB.Open "SELECT CODTRAB, NOMBRES FROM VWTRABAJ WHERE CODTRAB IN (SELECT DISTINCT CODTRAB FROM TRABXCOSTO) ORDER BY CODTRAB", DBSYSTEM, adOpenKeyset, adLockOptimistic
    If RSTRAB.EOF Or RSTRAB.RecordCount = 0 Then
        MsgBox "No se ha encontrado registros de Trabajadores", vbCritical
        Set RSTRAB = Nothing
        Exit Sub
    End If
    frmComun.CONECTAR RSTRAB
    frmComun.Show 1
    If VGUTIL(2) <> "" Then
        xTrab.Tag = RSTRAB!CODTRAB
        xTrab.Text = RSTRAB!CODTRAB & " : " & RSTRAB!NOMBRES
    End If
    Set RSTRAB = Nothing
End Sub

