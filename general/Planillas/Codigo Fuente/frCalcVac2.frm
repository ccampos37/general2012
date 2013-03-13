VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Begin VB.Form frCalcVac2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cálculo de Vacaciones"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   Icon            =   "frCalcVac2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin AplisetControlText.Aplitext xCompensacion 
      Height          =   300
      Left            =   1515
      TabIndex        =   38
      Top             =   5835
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   529
      Locked          =   -1  'True
      Text            =   ""
      Iluminar        =   0   'False
   End
   Begin AplisetControlText.Aplitext xPeriodo 
      Height          =   300
      Left            =   1515
      TabIndex        =   35
      Top             =   5535
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   529
      Locked          =   -1  'True
      Text            =   ""
      Iluminar        =   0   'False
   End
   Begin AplisetControlText.Aplitext xDiasVac 
      Height          =   285
      Left            =   4185
      TabIndex        =   28
      Top             =   4980
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   503
      Locked          =   -1  'True
      Text            =   ""
   End
   Begin VB.CheckBox Check1 
      Height          =   240
      Left            =   5235
      TabIndex        =   25
      Top             =   3405
      Width           =   210
   End
   Begin VB.Frame Frame1 
      Caption         =   "Información del Trabajador"
      Height          =   1590
      Left            =   75
      TabIndex        =   16
      Top             =   60
      Width           =   5010
      Begin MSComCtl2.DTPicker xSalFin 
         Height          =   285
         Left            =   3570
         TabIndex        =   33
         Top             =   735
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   16842753
         CurrentDate     =   36845
      End
      Begin MSComCtl2.DTPicker xFechaFin 
         Height          =   285
         Left            =   3570
         TabIndex        =   24
         Top             =   1125
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   16842753
         CurrentDate     =   36844
      End
      Begin MSComCtl2.DTPicker xFechaIni 
         Height          =   285
         Left            =   1950
         TabIndex        =   22
         Top             =   1125
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   16842753
         CurrentDate     =   36844
      End
      Begin MSComCtl2.DTPicker xSalini 
         Height          =   285
         Left            =   1950
         TabIndex        =   20
         Top             =   735
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   16842753
         CurrentDate     =   36844
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "al"
         Height          =   195
         Index           =   1
         Left            =   3330
         TabIndex        =   36
         Top             =   810
         Width           =   120
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "al"
         Height          =   195
         Index           =   0
         Left            =   3330
         TabIndex        =   23
         Top             =   1185
         Width           =   120
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Periodo de Vacaciones"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   1170
         Width           =   1650
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salida de Vacaciones"
         Height          =   195
         Left            =   135
         TabIndex        =   19
         Top             =   810
         Width           =   1545
      End
      Begin VB.Label xTrab 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   960
         TabIndex        =   18
         Top             =   390
         Width           =   3900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Trabajador"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   420
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Quitar"
      Height          =   315
      Left            =   165
      TabIndex        =   15
      Top             =   5085
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Agregar"
      Height          =   315
      Left            =   165
      TabIndex        =   14
      Top             =   4710
      Visible         =   0   'False
      Width           =   960
   End
   Begin AplisetControlText.Aplitext xDias 
      Height          =   270
      Left            =   4515
      TabIndex        =   10
      Top             =   2010
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   476
      MaxLength       =   5
      Text            =   "0"
      Entero          =   -1  'True
      TipoDato        =   "N"
   End
   Begin AplisetControlText.Aplitext xMeses 
      Height          =   270
      Left            =   2910
      TabIndex        =   8
      Top             =   2010
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   476
      MaxLength       =   5
      Text            =   "0"
      Entero          =   -1  'True
      TipoDato        =   "N"
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   360
      Left            =   5220
      TabIndex        =   3
      Top             =   2730
      Width           =   1395
   End
   Begin VB.CommandButton cmGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   360
      Left            =   5220
      TabIndex        =   2
      Top             =   2167
      Width           =   1395
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Fórmulas de Vacaciones"
      Height          =   540
      Left            =   5220
      TabIndex        =   1
      Top             =   1425
      Width           =   1395
   End
   Begin VB.CommandButton cmCalcular 
      Caption         =   "&Calcular"
      Height          =   855
      Left            =   5745
      Picture         =   "frCalcVac2.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   180
      Width           =   870
   End
   Begin MSDataGridLib.DataGrid xDetalle 
      Height          =   2355
      Left            =   165
      TabIndex        =   13
      Top             =   2295
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   4154
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   15395562
      HeadLines       =   2
      RowHeight       =   17
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
      Caption         =   "Detalle del Cálculo"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Concepto"
         Caption         =   "Conceptos Computables"
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
         DataField       =   "Importe"
         Caption         =   "Importe"
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
         MarqueeStyle    =   2
         BeginProperty Column00 
            ColumnWidth     =   3030.236
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            DividerStyle    =   1
            ColumnWidth     =   1214.929
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar Prog 
      Height          =   180
      Left            =   10395
      TabIndex        =   4
      Top             =   5280
      Visible         =   0   'False
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   318
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Compensación a:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   195
      TabIndex        =   37
      Top             =   5888
      Width           =   1230
   End
   Begin VB.Label xProg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "**** Texto *****"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   10305
      TabIndex        =   5
      Top             =   5490
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo de Planilla"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   165
      TabIndex        =   34
      Top             =   5580
      Width           =   1305
   End
   Begin VB.Label xMontoBruto 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00 "
      Height          =   270
      Left            =   3660
      TabIndex        =   32
      Top             =   5265
      Width           =   1335
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Monto bruto de Vacaciones"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1530
      TabIndex        =   31
      Top             =   5310
      Width           =   1965
   End
   Begin VB.Label xTotalCalculo 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00 "
      Height          =   270
      Left            =   3660
      TabIndex        =   30
      Top             =   4710
      Width           =   1335
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total del Cálculo"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1530
      TabIndex        =   29
      Top             =   4740
      Width           =   1185
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Número de Días por Vacaciones"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1530
      TabIndex        =   27
      Top             =   5025
      Width           =   2325
   End
   Begin VB.Label Label2 
      Caption         =   "Grabar/mostrar los detalles en la planilla."
      Height          =   615
      Left            =   5475
      TabIndex        =   26
      Top             =   3405
      Width           =   1080
   End
   Begin VB.Label xFecha 
      Alignment       =   2  'Center
      Height          =   225
      Left            =   1005
      TabIndex        =   12
      Top             =   2040
      Width           =   1140
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " F. Ingreso"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   165
      TabIndex        =   11
      Top             =   2010
      Width           =   2040
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Dias"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   1
      Left            =   3570
      TabIndex        =   9
      Top             =   2010
      Width           =   960
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Meses"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   0
      Left            =   2175
      TabIndex        =   7
      Top             =   2010
      Width           =   780
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tiempo Computable"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   165
      TabIndex        =   6
      Top             =   1800
      Width           =   4860
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   4530
      Left            =   75
      Top             =   1725
      Width           =   5025
   End
End
Attribute VB_Name = "frCalcVac2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSCALC As New ADODB.Recordset
Dim ENPROCESO As Boolean
Private Sub CMACTUALIZAR_CLICK()
    CALCULOVAC
End Sub
Private Sub CMCALCULAR_CLICK()
 On Error GoTo CMCAL
 Dim GENERALTIPO As Boolean
    Screen.MousePointer = 11
    Dim XFEC As Date, NUMMESES As Integer, NUMDIAS As Integer, XFEC2 As Date
    ENPROCESO = True
    'CALCULO DEL TIEMPO COMPUTABLE
    Dim RSCNPT As New ADODB.Recordset
    Set RSCNPT = New ADODB.Recordset
    RSCNPT.Open "SELECT * FROM FORMULASVAC WHERE AFECTOPRO<>0", DBSYSTEM, adOpenStatic, adLockReadOnly
    If RSCNPT.RecordCount = 0 Then
        MsgBox "No ha definido Fórmulas de Vacaciones. Seleccione la opción de Fórmulas de Vacaciones", vbInformation
        Screen.MousePointer = 1
        Set RSCALC = Nothing
        Set RSCNPT = Nothing
        Exit Sub
    End If
    RSCNPT.MoveFirst
    Prog.Min = 0
    Prog.Max = Val(RSCNPT.RecordCount)
    Prog.Value = 0
    Prog.Visible = True
    xProg.Visible = True
    DBSYSTEM.Execute "DELETE FROM  [##TMPCTS2" & VGL_COMPUTER & "] "
    Dim VALOR As Single
    Do While Not RSCNPT.EOF
        GENERALTIPO = RSCNPT!GENE
        Prog.Value = Prog.Value + 1
        xProg.Caption = "CALCULANDO ... " & RSCNPT!NOMBRE
        If InStr(RSCNPT!FORMULA, "@") = 0 Then
            VALOR = DevuelveValor("SELECT " & RSCNPT!FORMULA & " AS VALOR_DEV FROM TRABAJADORES WHERE CODTRAB='" & xTrab.Tag & "'", DBSYSTEM)
            If IsNull(VALOR) Then VALOR = 0
        Else
            If Val(RSCNPT!CRITERIO) = 0 Then
                VALOR = DevuelveValor("SELECT " & CAMBIACADENA(RSCNPT!FORMULA, xTrab.Tag, 6, GENERALTIPO) & " AS VALOR_DEV FROM TRABAJADORES WHERE CODTRAB='" & xTrab.Tag & "'", DBSYSTEM)
            Else
                VALOR = DevuelveValor("SELECT " & CAMBIACADENA(RSCNPT!FORMULA, xTrab.Tag, Val(RSCNPT!CRITERIO), GENERALTIPO) & " AS VALOR_DEV FROM TRABAJADORES WHERE CODTRAB='" & xTrab.Tag & "'", DBSYSTEM)
            End If
        End If
        If VALOR <> 0 Then
            VALOR = Round(VALOR, 2)
            DBSYSTEM.Execute "INSERT INTO  [##TMPCTS2" & VGL_COMPUTER & "]  VALUES ('" & xTrab.Tag & "','" & RSCNPT!NOMBRE & "'," & VALOR & "," & IIf(RSCNPT!TIPO, 1, 0) & ")"
        End If
        RSCNPT.MoveNext
    Loop
    xProg.Visible = False
    Prog.Visible = False
    ENPROCESO = False
    cmGrabar.Enabled = True
    cmdAgregar.Visible = True
    cmdEliminar.Visible = True
    Screen.MousePointer = 1
    Set RSCALC = Nothing
    RSCALC.Open " [##TMPCTS2" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic
    Set xDetalle.DataSource = RSCALC
    CALCULOVAC
    Exit Sub
CMCAL:
    MsgBox "Error en Formula " & " detalle :" & Chr(13) & ERR.Description
    Resume Next
    Exit Sub
End Sub

Private Sub CMGRABAR_CLICK()
    If Val(xMontoBruto.Caption) = 0 Then
        MsgBox "No ha procesado o el monto de vacaciones no puede estar en valor cero", vbInformation
        cmCalcular.SetFocus
        Exit Sub
    End If
    If xPeriodo.Text = "" Then
        MsgBox "Las vacaciones deben de tener un periodo de pago", vbInformation
        xPeriodo.SetFocus
        Exit Sub
    End If
    If Val(xDiasVac.Text) <> 30 And xCompensacion.Text = "" Then
        MsgBox "No ha seleccionado un grupo donde almacenar el resto de vacaciones que forman parte de la compensacion (venta de vacaciones)", vbInformation
        xCompensacion.SetFocus
        Exit Sub
    End If
    If MsgBox("Desea grabar la información presentada", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    DBSYSTEM.Execute "UPdate HISTOVAC SET CERRADO=0, MONTO=" & xMontoBruto.Caption & ", FECHAREG=" & DateSQL(Date) & ",NOMBOL=" & xPeriodo.Tag & ",MODOCALCULO=0 WHERE CODIGO=" & Frame1.Tag
    If Check1.Value = 1 Then
        If RSCALC.RecordCount > 0 Then
            RSCALC.MoveFirst
            Do While Not RSCALC.EOF
                DBSYSTEM.Execute "INSERT INTO DETALLEVAC (CODIGO, DESCRIPCION,IMPORTE) VALUES (" & Frame1.Tag & ",'" & RSCALC!CONCEPTO & "'," & RSCALC!Importe & ")"
                RSCALC.MoveNext
            Loop
        End If
    End If
    Unload Me
End Sub
Private Sub Command1_Click()
    frFormulasVac.Show 1
End Sub
Private Sub Command2_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    If ExisteTablaAux(" [##TMPCTS2" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPCTS2" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "CREATE TABLE  [##TMPCTS2" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8), CONCEPTO varchar(35), IMPORTE  Numeric(20,2) , INDTIPO bit)"
    Select Case VPTAREA
        Case "NUEVO"
            Me.Caption = "Nuevo Calculo de Vacaciones"
        Case "MODIFICAR"
            Me.Caption = "Modificacion de Calculo de Vacaciones"
            Frame1.Enabled = False
            cmGrabar.Enabled = True
            cmdAgregar.Visible = True
            cmdEliminar.Visible = True
        Case "VISTA"
            Frame1.Enabled = False
            xDetalle.AllowUpdate = False
            cmCalcular.Visible = False
            Me.Caption = "Consulta de Calculo dd Vacaciones"
    End Select
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSCALC = Nothing
End Sub

Private Sub XCOMPENSACION_CLICK()
    If Val(xMontoBruto.Caption) = 0 Then
        MsgBox "No se puede continuar si no ha Calculado aun", vbInformation
        cmCalcular.SetFocus
        Exit Sub
    End If
    If Val(xDiasVac.Text) = 30 Then
        MsgBox "No existe saldo a pagar porque el trabajador recibirá por completo el pago de sus vacaciones", vbInformation
        Exit Sub
    End If
    Dim RSGRUPOS As New ADODB.Recordset
    RSGRUPOS.Open "SELECT CODGRUPO, NOMBRE FROM CTAGRUPO WHERE TIPO=1 ORDER BY NOMBRE", DBSYSTEM, adOpenKeyset, adLockOptimistic
    frmComun.CONECTAR RSGRUPOS
    frmComun.Show 1
    If VGUTIL(1) = "" Then
        Set RSGRUPOS = Nothing
        Exit Sub
    End If
    xCompensacion.Text = VGUTIL(1)
    xCompensacion.Tag = VGUTIL(0)
End Sub
Private Sub XDETALLE_AFTERCOLUPdatetime(ByVal COLINDEX As Integer)
    RSCALC.MOVE 0
    CMACTUALIZAR_CLICK
End Sub
Public Function CAMBIACADENA(ByVal CADENA As String, ByVal CODTRAB As String, ByVal Meses As Byte, ByVal GENERAL2 As Boolean) As String
    Dim POSARROBA As Integer, POS1 As Integer, PROCESO As String, CAMPO As String, POS2 As Integer
    Dim VALOR As Double
    Dim XFEC2 As Date
On Error GoTo FUNCT
    XFEC2 = DateAdd("M", -1 * Meses, xSalini.Value)
    If XFEC2 < xFechaIni.Value Then XFEC2 = xFechaIni.Value
    POSARROBA = 1
    POSARROBA = InStr(POSARROBA, CADENA, "@")
    Do While POSARROBA <> 0
        POS1 = InStr(POSARROBA, CADENA, "(")
        PROCESO = Mid(CADENA, POSARROBA + 1, POS1 - (POSARROBA + 1))
        POS2 = InStr(POSARROBA, CADENA, ")")
        CAMPO = Mid(CADENA, POS1 + 1, POS2 - (POS1 + 1))
        Select Case UCase(PROCESO)
            Case "PROMEDIO"
                VALOR = CALCULOCONCEPTOS(CODTRAB, XFEC2, DateAdd("M", -1, xSalini.Value), PROMEDIO, CAMPO, GENERAL2)
            Case "ULTIMOVALOR"
                VALOR = CALCULOCONCEPTOS(CODTRAB, XFEC2, DateAdd("M", -1, xSalini.Value), ULTIMOVALOR, CAMPO, GENERAL2)
            Case "PRIMERVALOR"
                VALOR = CALCULOCONCEPTOS(CODTRAB, XFEC2, DateAdd("M", -1, xSalini.Value), PRIMERVALOR, CAMPO, GENERAL2)
            Case "SUMA"
                VALOR = CALCULOCONCEPTOS(CODTRAB, XFEC2, DateAdd("M", -1, xSalini.Value), SUMA, CAMPO, GENERAL2)
            Case "MEDIA"
                VALOR = CALCULOCONCEPTOS(CODTRAB, XFEC2, DateAdd("M", -1, xSalini.Value), MEDIA, CAMPO, GENERAL2)
            Case "PROMEDIOVALOR"
                VALOR = CALCULOCONCEPTOS(CODTRAB, XFEC2, DateAdd("M", -1, xSalini.Value), PROMEDIOVALOR, CAMPO, GENERAL2)
            Case "PRIMERO"
                VALOR = CALCULOCONCEPTOS(CODTRAB, XFEC2, DateAdd("M", -1, xSalini.Value), PRIMERO, CAMPO, GENERAL2)
            Case "ULTIMO"
                VALOR = CALCULOCONCEPTOS(CODTRAB, XFEC2, DateAdd("M", -1, xSalini.Value), ULTIMO, CAMPO, GENERAL2)
            Case "MAYORVALOR"
                VALOR = CALCULOCONCEPTOS(CODTRAB, XFEC2, DateAdd("M", -1, xSalini.Value), MAYORVALOR, CAMPO, GENERAL2)
            Case "MENORVALOR"
                VALOR = CALCULOCONCEPTOS(CODTRAB, XFEC2, DateAdd("M", -1, xSalini.Value), MENORVALOR, CAMPO, GENERAL2)
            Case "NUMERO"
                VALOR = CALCULOCONCEPTOS(CODTRAB, XFEC2, DateAdd("M", -1, xSalini.Value), Numero, CAMPO, GENERAL2)
            Case "NSECUENCIA"
                VALOR = CALCULOCONCEPTOS(CODTRAB, XFEC2, DateAdd("M", -1, xSalini.Value), NSECUENCIA, CAMPO, GENERAL2)
            Case "CALCULOVALOR"
                   VALOR = CALCULOVALOR(GENERAL2, CAMPO, CODTRAB)
        End Select
        If IsNull(VALOR) Then VALOR = 0
        CADENA = Replace(CADENA, Mid(CADENA, POSARROBA, (POS2 - POSARROBA) + 1), "" & VALOR)
        POSARROBA = InStr(POSARROBA, CADENA, "@")
    Loop
    CAMBIACADENA = CADENA
    Exit Function
FUNCT:
    'EXIT FUNCTION
    MsgBox ERR.Description
End Function

Private Sub CMDAGREGAR_CLICK()
    FrmAgrCon.Show 1
    If FrmAgrCon.VarGrabar = False Then Exit Sub
    RSCALC.AddNew
    RSCALC!CONCEPTO = FrmAgrCon.CONCEPTO
    RSCALC!Importe = FrmAgrCon.Importe
    RSCALC!INDTIPO = FrmAgrCon.TIPO
    RSCALC.Update
    CALCULOVAC
End Sub

Private Sub CMDELIMINAR_CLICK()
    If RSCALC.RecordCount = 0 Then Exit Sub
    If MsgBox("Desea eliminar el registro seleccionado", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    If Not (RSCALC.EOF Or RSCALC.BOF) Then RSCALC.Delete
    CALCULOVAC
End Sub
Public Sub CALCULOVAC()
  On Error GoTo CAL
    xTotalCalculo.Caption = Format(DevuelveValor("SELECT SUM(IMPORTE) AS T1 FROM  [##TMPCTS2" & VGL_COMPUTER & "] ", DBSYSTEM), "0.00 ")
    xDiasVac.Text = DateDiff("D", xSalini.Value, xSalFin.Value) + 1
    xMontoBruto.Caption = Format(Val(xTotalCalculo.Caption) / 30 * Val(xDiasVac.Text), "0.00 ")
    Exit Sub
CAL:
    Exit Sub
End Sub

Private Sub XPERIODO_DBLCLICK()
    Dim RSMESES As New ADODB.Recordset
    RSMESES.Open "SELECT CODIGO, NOMBRE FROM NOMBOL WHERE MES IN (SELECT MESACTIVO FROM MESESACT)", DBSYSTEM, adOpenStatic, adLockReadOnly
    If RSMESES.RecordCount = 0 Then
        MsgBox "No se han encontrado meses en actividad", vbCritical
        Set RSMESES = Nothing
        Exit Sub
    End If
    frmComun.CONECTAR RSMESES
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xPeriodo.Text = RSMESES!NOMBRE
        xPeriodo.Tag = RSMESES!Codigo
    End If
    Set RSMESES = Nothing
End Sub

Function CALCULOVALOR(General3 As Boolean, CONCEPTO As String, ByRef CODTRAB As String) As Single
    Dim RS As ADODB.Recordset
    Dim RSCNPT As ADODB.Recordset
    Dim STRCALCSUM As String
    Dim XNUMMES As Integer, X As Integer, NUMOCURRE As Integer, SUMATOTAL As Double
    Dim FEC1 As Date, FEC2 As Date, STRMES As String, VALOR As Double, RESULTADO As Double
    NUMOCURRE = 0
    RESULTADO = 0
    SUMATOTAL = 0
    Dim STRCREA As String
    Dim ACUM As String
    ACUM = ""
        CONCEPTO = "'" + CONCEPTO + "'"
        For X = 1 To Len(CONCEPTO)
            ACUM = ACUM + Mid(CONCEPTO, X, 1)
            If Mid(CONCEPTO, X + 1, 1) = "," Then
                ACUM = ACUM + "'"
            End If
            If Mid(CONCEPTO, X, 1) = "," Then
                ACUM = ACUM + "'"
            End If
        Next
    CONCEPTO = ACUM
    If ExisteTablaAux("[##_TMPCALCULO" & VGL_COMPUTER & "]") Then DBSYSTEM.Execute "DROP TABLE _TMPCALCULO"
        'CREA LA TABLA PARA EL CALCULO
        STRCREA = "CREATE TABLE  [##_TMPCALCULO" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8), NOMBRES varchar(35), CODAREA varchar(10), CODCCOSTO varchar(10), BASICO  Numeric(20,2) , ASIGFAM  Numeric(20,2) , CODAFP varchar(2), TASASCTR  Numeric(20,2) , APOROBL  Numeric(20,2) , SEGURO  Numeric(20,2) , TOPESEGURO  Numeric(20,2) , COMISIONRA  Numeric(20,2) , SUMAAFP  Numeric(20,2) , SUMASALUD  Numeric(20,2) , TOTING  Numeric(20,2) , TOTEGR  Numeric(20,2) , _HORAST  Numeric(20,2) , _HOREXTRAS  Numeric(20,2) , _QUINTACAT  Numeric(20,2) "
        STRCREA = STRCREA + ", SUMAIES  Numeric(20,2) , SUMARENTA  Numeric(20,2) , SUMASCTR  Numeric(20,2) , SUMACTS  Numeric(20,2) , SUMAGRAT  Numeric(20,2) , SUMAVAC  Numeric(20,2) , T1  Numeric(20,2) , T2  Numeric(20,2) , T3  Numeric(20,2) , T4  Numeric(20,2) , T5  Numeric(20,2) , OTROSING  Numeric(20,2) , OTROSEGR  Numeric(20,2) , ADELANTO  Numeric(20,2) ,UBIGEO varchar(6),SEXO bit,TIPOTRAB varchar(2),FECHAING datetime, SITUACION varchar(2),CARGO varchar(25),BANCO varchar(4), ESSALUDVIDA BIT, RUCEPS varchar(11),NOPDT bit,OPCION01 bit, OPCION02 bit, OPCIONA varchar(15), OPCIONB varchar(15), XREDONDEO  Numeric(20,2) , AFECTOQUINTA bit)"
        DBSYSTEM.Execute STRCREA
        
    Set RS = New ADODB.Recordset
    Set RSCNPT = New ADODB.Recordset
    RSCNPT.Open "SELECT CONCEPTOS.* FROM CONCEPTOS,FORMARUBS WHERE CONCEPTO=CODIGO AND CODIGO IN (" & CONCEPTO & ") ORDER BY TIPO, FILA", DBSYSTEM, adOpenStatic
    With RSCNPT
        .MoveFirst
        Do While Not .EOF
            DBSYSTEM.Execute "ALTER TABLE  [##_TMPCALCULO" & VGL_COMPUTER & "]  ADD COLUMN " & RSCNPT!Codigo & "  Numeric(20,2) "
            DBSYSTEM.Execute "UPdate  [##_TMPCALCULO" & VGL_COMPUTER & "]  SET " & RSCNPT!Codigo & "=0"
            .MoveNext
        Loop
    End With
    DBSYSTEM.Execute "INSERT INTO  [##_TMPCALCULO" & VGL_COMPUTER & "]  (CODTRAB,NOMBRES,CODAREA,CODCCOSTO,BASICO,ASIGFAM,CODAFP,TASASCTR,APOROBL,SEGURO,TOPESEGURO,COMISIONRA,UBIGEO,SEXO,TIPOTRAB,FECHAING,SITUACION,CARGO,BANCO,ESSALUDVIDA,RUCEPS,NOPDT,OPCION01,OPCION02,OPCIONA,OPCIONB, AFECTOQUINTA) SELECT CODTRAB, NOMBRES,AREA,CCOSTO,BASICO,ASIGFAM,FONDOPENS,TASA,APOROBLI,SEGURO,TOPESEGURO,COMISIONRA,UBIGEO,SEXO,TIPOTRAB,FECHAING,SITUACIÓN,CARGO,BANCO,ESSALUDVIDA,RUCEPS,NOPDT,OPCION01,OPCION02,OPCIONA,OPCIONB,AFECTOQUINTA FROM TRABBOLETEAR IN '" & REGSISTEMA.PATHEMPRESA & "\PLANILLA.MDB' WHERE CODTRAB ='" & CODTRAB & "'"
    
    Dim RSAUX As New ADODB.Recordset
    DBSYSTEM.Execute "ALTER TABLE  [##_TMPCALCULO" & VGL_COMPUTER & "]  ADD COLUMN VALORDOLAR  Numeric(20,2) "
    DBSYSTEM.Execute "UPdate  [##_TMPCALCULO" & VGL_COMPUTER & "]  SET VALORDOLAR=" & MDIPrincipal.BarraEstado.Panels("DOLAR").Text
    If Not ExisteTabla("DATATRAB") Then
        DBSYSTEM.Execute "CREATE TABLE DATATRAB (CODDATA varchar(15),DESCDATA varchar(30),TIPODATA varchar(1))"
        MsgBox "El Sistema de Planillas ha actualizado su version", vbInformation
    End If
    RSAUX.Open "DATATRAB", DBSYSTEM, adOpenStatic, adLockReadOnly
    Do While Not RSAUX.EOF
        Select Case RSAUX!TIPODATA
            Case "N"
                DBSYSTEM.Execute "ALTER TABLE  [##_TMPCALCULO" & VGL_COMPUTER & "]  ADD COLUMN " & RSAUX!CODDATA & "  Numeric(20,2) "
                DBSYSTEM.Execute "UPdate  [##_TMPCALCULO" & VGL_COMPUTER & "]  INNER JOIN TRABAJADORES ON [##_TMPCALCULO" & VGL_COMPUTER & "].[CODTRAB]=TRABAJADORES.CODTRAB SET [##_TMPCALCULO" & VGL_COMPUTER & "].[" & RSAUX!CODDATA & "]=TRABAJADORES." & RSAUX!CODDATA
                DBSYSTEM.Execute "UPdate  [##_TMPCALCULO" & VGL_COMPUTER & "]  SET " & RSAUX!CODDATA & " =0 WHERE (" & RSAUX!CODDATA & ")IS NULL"
            Case "T"
                DBSYSTEM.Execute "ALTER TABLE  [##_TMPCALCULO" & VGL_COMPUTER & "]  ADD COLUMN " & RSAUX!CODDATA & " varchar(30)"
                DBSYSTEM.Execute "UPdate  [##_TMPCALCULO" & VGL_COMPUTER & "]  INNER JOIN TRABAJADORES ON [##_TMPCALCULO" & VGL_COMPUTER & "].[CODTRAB]=TRABAJADORES.CODTRAB SET [##_TMPCALCULO" & VGL_COMPUTER & "].[" & RSAUX!CODDATA & "]=TRABAJADORES." & RSAUX!CODDATA
                DBSYSTEM.Execute "UPdate  [##_TMPCALCULO" & VGL_COMPUTER & "]  SET " & RSAUX!CODDATA & " =' ' WHERE (" & RSAUX!CODDATA & ")IS NULL"
            Case "F"
                DBSYSTEM.Execute "ALTER TABLE  [##_TMPCALCULO" & VGL_COMPUTER & "]  ADD COLUMN " & RSAUX!CODDATA & " datetimeTIME"
                DBSYSTEM.Execute "UPdate  [##_TMPCALCULO" & VGL_COMPUTER & "]  INNER JOIN TRABAJADORES ON [##_TMPCALCULO" & VGL_COMPUTER & "].[CODTRAB]=TRABAJADORES.CODTRAB SET [##_TMPCALCULO" & VGL_COMPUTER & "].[" & RSAUX!CODDATA & "]=TRABAJADORES." & RSAUX!CODDATA
                DBSYSTEM.Execute "UPdate  [##_TMPCALCULO" & VGL_COMPUTER & "]  SET " & RSAUX!CODDATA & " =#01/01/1900# WHERE (" & RSAUX!CODDATA & ")IS NULL"
            Case "B"
                DBSYSTEM.Execute "ALTER TABLE  [##_TMPCALCULO" & VGL_COMPUTER & "]  ADD COLUMN " & RSAUX!CODDATA & " BIT"
                DBSYSTEM.Execute "UPdate  [##_TMPCALCULO" & VGL_COMPUTER & "]  INNER JOIN TRABAJADORES ON [##_TMPCALCULO" & VGL_COMPUTER & "].[CODTRAB]=TRABAJADORES.CODTRAB SET [##_TMPCALCULO" & VGL_COMPUTER & "].[" & RSAUX!CODDATA & "]=TRABAJADORES." & RSAUX!CODDATA
                DBSYSTEM.Execute "UPdate  [##_TMPCALCULO" & VGL_COMPUTER & "]  SET " & RSAUX!CODDATA & " =0 WHERE (" & RSAUX!CODDATA & ")IS NULL"
            Case Else
                Beep
        End Select
        RSAUX.MoveNext
    Loop
    Set RSAUX = Nothing
    DBSYSTEM.Execute "UPdate  [##_TMPCALCULO" & VGL_COMPUTER & "]  SET SUMAAFP=0,SUMASALUD=0,TOTING=0,TOTEGR=0,_HORAST=0,_HOREXTRAS=0,_QUINTACAT=0,SUMAIES=0,SUMARENTA=0,SUMASCTR=0,SUMACTS=0,SUMAGRAT=0,SUMAVAC=0,T1=0,T2=0,T3=0,T4=0,T5=0,OTROSING=0,OTROSEGR=0"
    'LIMPIA LOS CONCEPTOS DE LA TABLA TEMPORAL PARA EL CALCULO
        With RSCNPT
            .MoveFirst
            STRCREA = "ADELANTO=0"
            Do While Not .EOF
                STRCREA = STRCREA & ", " & RSCNPT!Codigo & "=0"
                .MoveNext
            Loop
    End With
    DBSYSTEM.Execute "UPdate  [##_TMPCALCULO" & VGL_COMPUTER & "]  SET " & STRCREA


    Dim CLA As Integer
    CLA = 0
    RSCNPT.MoveFirst
    'CARGA DE LOS DATOS TIPO FORMULA
    Do While Not RSCNPT.EOF
        If Not RSCNPT!ESESCRITO Then
            If RSCNPT!FORMULA = "" Then
                MsgBox "El ingreso de la formula no ha sido registrada, el sistema generará constantes avisos de errores sobre el fallo del siguiente rubro: " & RSCNPT!Codigo & ": " & RSCNPT!NOMBRE
                If MsgBox("Desea continuar con la carga del sistema", vbYesNo) = vbNo Then
                    Set RSCNPT = Nothing
                    Unload Me
                End If
            End If
        End If
        RSCNPT.MoveNext
    Loop
    
    'FORMULAS COMO SUMAAFP, TOTAL INGRESOS, ENTRE OTRAS
    Dim CADSUMAS(14) As String
    CADSUMAS(0) = "0+OTROSING"
    CADSUMAS(1) = "0+OTROSING"
    CADSUMAS(2) = "0+OTROSING"
    CADSUMAS(3) = "0+OTROSING"
    CADSUMAS(4) = "0+OTROSING"
    CADSUMAS(5) = "0+OTROSING"
    CADSUMAS(6) = "0+OTROSING"
    CADSUMAS(7) = "0+OTROSING"
    CADSUMAS(8) = "0+OTROSING"
    CADSUMAS(9) = "0+OTROSING"
    CADSUMAS(10) = "0+OTROSING"
    CADSUMAS(11) = "0+OTROSING"
    CADSUMAS(12) = "0+OTROSING"
    CADSUMAS(13) = "0+OTROSING"
    Dim CADHORAS As String, CADHOREXT As String, CAD5TA As String
    CADHORAS = "0"
    CADHOREXT = "0"
    CAD5TA = "0"
    'RECALCULAR LOS TOTALES DE INGRESOS AFECTOS A ...
    With RSCNPT
        .MoveFirst
        Do While Not .EOF
            If !TIPO = 0 Then
                'EN ESTA SECCIÓN SE CALCULAN LOS TIPOS DE INFORMACIÓN
                'COMO HORAS TRABAJADAS Y HORAS EXTRAS
                If !TIPOINFO < 3 Then
                    Select Case !TIPOINFO
                        Case 0: CADHORAS = CADHORAS & "+" & !Codigo & "* 8"
                        Case 1: CADHORAS = CADHORAS & "+" & !Codigo
                        Case 2: CADHOREXT = CADHOREXT & "+" & !Codigo
                    End Select
                End If
            End If
            If !TIPO = 1 Then
                    CADSUMAS(0) = CADSUMAS(0) & IIf(CADSUMAS(0) = "", "", "+") & !Codigo  'SOLO PARA EL CÁLCULO DE INGRESOS
                    If !SUMAAFP Then CADSUMAS(1) = CADSUMAS(1) & IIf(CADSUMAS(1) = "", "", "+") & !Codigo
                    If !SUMASALUD Then CADSUMAS(2) = CADSUMAS(2) & IIf(CADSUMAS(2) = "", "", "+") & !Codigo
                    If !SUMAIES Then CADSUMAS(3) = CADSUMAS(3) & IIf(CADSUMAS(3) = "", "", "+") & !Codigo
                    If !SUMARENTA Then
                        If Len(Trim(!COMENTARIO)) Then
                            CADSUMAS(4) = CADSUMAS(4) & IIf(CADSUMAS(4) = "", "", "+") & Trim(!COMENTARIO)
                        Else
                            CADSUMAS(4) = CADSUMAS(4) & IIf(CADSUMAS(4) = "", "", "+") & !Codigo
                        End If
                    End If
                    If !SUMASCTR Then CADSUMAS(5) = CADSUMAS(5) & IIf(CADSUMAS(5) = "", "", "+") & !Codigo
                    If !SUMACTS Then CADSUMAS(6) = CADSUMAS(6) & IIf(CADSUMAS(6) = "", "", "+") & !Codigo
                    If !SUMAGRAT Then CADSUMAS(7) = CADSUMAS(7) & IIf(CADSUMAS(7) = "", "", "+") & !Codigo
                    If !SUMAVAC Then CADSUMAS(8) = CADSUMAS(8) & IIf(CADSUMAS(8) = "", "", "+") & !Codigo
                    If !SUMAT1 Then CADSUMAS(9) = CADSUMAS(9) & IIf(CADSUMAS(9) = "", "", "+") & !Codigo
                    If !SUMAT2 Then CADSUMAS(10) = CADSUMAS(10) & IIf(CADSUMAS(10) = "", "", "+") & !Codigo
                    If !SUMAT3 Then CADSUMAS(11) = CADSUMAS(11) & IIf(CADSUMAS(11) = "", "", "+") & !Codigo
                    If !SUMAT4 Then
                        If Len(Trim(!COMENTARIO)) Then
                            CADSUMAS(12) = CADSUMAS(12) & IIf(CADSUMAS(12) = "", "", "+") & !COMENTARIO
                        Else
                            CADSUMAS(12) = CADSUMAS(12) & IIf(CADSUMAS(12) = "", "", "+") & !Codigo
                        End If
                    End If
                    If !SUMAT5 Then
                        If Len(Trim(!COMENTARIO)) Then
                            CADSUMAS(13) = CADSUMAS(13) & IIf(CADSUMAS(13) = "", "", "+") & !COMENTARIO
                        Else
                            CADSUMAS(13) = CADSUMAS(13) & IIf(CADSUMAS(13) = "", "", "+") & !Codigo
                        End If
                    End If
            End If
            .MoveNext
        Loop
    End With
    'ACTUALIZA LOS TOTALES
    STRCALCSUM = "UPdate  [##_TMPCALCULO" & VGL_COMPUTER & "]  SET  TOTING=" & CADSUMAS(0) & ", SUMAAFP=" & CADSUMAS(1) & ", SUMASALUD=" & CADSUMAS(2) & ", SUMAIES=" & CADSUMAS(3) & ", SUMARENTA=" & CADSUMAS(4) & ", SUMASCTR=" & CADSUMAS(5) & ", SUMACTS=" & CADSUMAS(6) & ", SUMAGRAT=" & CADSUMAS(7) & ", SUMAVAC=" & CADSUMAS(8) & ", T1=" & CADSUMAS(9) & ", T2=" & CADSUMAS(10) & ", T3=" & CADSUMAS(11) & ", T4=" & CADSUMAS(12) & ", T5=" & CADSUMAS(13) & ", _HORAST=" & CADHORAS & ", _HOREXTRAS=" & CADHOREXT
    
    'CALCULO DE FORMULAS TIPO INGRESO
    With RSCNPT
        .MoveFirst
        Do While Not .EOF
            If !TIPO <= 1 And Not !ESESCRITO And Not IsNull(!FORMULA) Then 'SI TIPO ES INGRESO
                DBSYSTEM.Execute "UPdate  [##_TMPCALCULO" & VGL_COMPUTER & "]  SET " & !Codigo & "=" & !FORMULA
            End If
            .MoveNext
        Loop
    End With
    '--------------------------------------
    'CALCULO DE LAS SUMAS DE AFECTOS A
    '--------------------------------------
    
    DBSYSTEM.Execute STRCALCSUM
    'VACIADO DE DATOS DE ASISTENCIA DE TRABAJADORES
        Set RSAUX = Nothing
        RSAUX.Open "SELECT CODTRAB, CONCEPTO, SUM(VALOR) AS CANTI FROM ASIS" & REGSISTEMA.ANNO & " WHERE (DIA BETWEEN " & DateSQL(REGINPUT.FECHAINI) & " AND " & DateSQL(REGINPUT.FECHAFIN) & ") AND CODTRAB='" & CODTRAB & "' GROUP BY CODTRAB, CONCEPTO", DBSYSTEM, adOpenStatic
        Do While Not RSAUX.EOF
            DBSYSTEM.Execute "UPdate  [##_TMPCALCULO" & VGL_COMPUTER & "]  SET " & RSAUX!CONCEPTO & "=" & RSAUX!CANTI & " WHERE CODTRAB='" & RSAUX!CODTRAB & "'"
            RSAUX.MoveNext
        Loop
    'VACIADO DE DATOS DE MOVIMIENTOS
        Set RSAUX = Nothing
        RSAUX.Open "SELECT CODTRAB, CONCEPTO, SUM(VALOR) AS CANTI FROM INGMOV2000 WHERE CODTRAB ='" & CODTRAB & "' AND CODNOMBOL=" & REGINPUT.Codigo & " GROUP BY CODTRAB, CONCEPTO", DBSYSTEM, adOpenStatic
        Do While Not RSAUX.EOF
            DBSYSTEM.Execute "UPdate  [##_TMPCALCULO" & VGL_COMPUTER & "]  SET " & RSAUX!CONCEPTO & "=" & RSAUX!CANTI & " WHERE CODTRAB='" & RSAUX!CODTRAB & "'"
            RSAUX.MoveNext
        Loop
End Function

