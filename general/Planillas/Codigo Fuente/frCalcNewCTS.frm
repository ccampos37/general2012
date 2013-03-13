VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frCalcCTS2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cálculo de C.T.S."
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9420
   Icon            =   "frCalcNewCTS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   9420
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Tasa"
      Height          =   660
      Left            =   5400
      TabIndex        =   23
      Top             =   165
      Width           =   2460
      Begin AplisetControlText.Aplitext xTasa 
         Height          =   300
         Left            =   1005
         TabIndex        =   24
         Top             =   255
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   529
         Text            =   "8.3333333"
         TipoDato        =   "N"
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   2190
         TabIndex        =   26
         Top             =   330
         Width           =   120
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tasa C.T.S."
         Height          =   195
         Left            =   90
         TabIndex        =   25
         Top             =   315
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   360
      Left            =   6315
      TabIndex        =   22
      Top             =   5625
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Agregar"
      Height          =   360
      Left            =   5265
      TabIndex        =   21
      Top             =   5625
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir Constancias"
      Height          =   360
      Left            =   7605
      TabIndex        =   20
      Top             =   5625
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.CommandButton cmActualizar 
      Caption         =   "&Actualizar"
      Height          =   360
      Left            =   7980
      TabIndex        =   19
      Top             =   5625
      Width           =   1275
   End
   Begin MSComctlLib.ProgressBar Prog 
      Height          =   135
      Left            =   120
      TabIndex        =   17
      Top             =   6135
      Visible         =   0   'False
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   360
      Left            =   8010
      TabIndex        =   12
      Top             =   390
      Width           =   1305
   End
   Begin VB.CommandButton cmGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   855
      Left            =   8220
      Picture         =   "frCalcNewCTS.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1020
      Width           =   870
   End
   Begin Crystal.CrystalReport Reporte 
      Left            =   225
      Top             =   2940
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmCalcular 
      Caption         =   "&Calcular"
      Height          =   855
      Left            =   6990
      Picture         =   "frCalcNewCTS.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1020
      Width           =   870
   End
   Begin MSDataGridLib.DataGrid xData 
      Height          =   3540
      Left            =   135
      TabIndex        =   8
      Top             =   2025
      Width           =   9150
      _ExtentX        =   16140
      _ExtentY        =   6244
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      Caption         =   "Trabajadores Seleccionados"
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "CodTrab"
         Caption         =   "Codigo"
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
         DataField       =   "Nombres"
         Caption         =   "Apellidos y Nombres"
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
      BeginProperty Column02 
         DataField       =   "ImporteCTS"
         Caption         =   "Importe del Cálc. C.T.S."
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
      BeginProperty Column03 
         DataField       =   "RemuAfec"
         Caption         =   "Remuneración Afecta"
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
      BeginProperty Column04 
         DataField       =   "Meses"
         Caption         =   "Meses"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0 "
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Dias"
         Caption         =   "Dias"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0 "
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "FechaIng"
         Caption         =   "Fecha de Ingreso"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yyyy"
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
         ScrollBars      =   3
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
         BeginProperty Column05 
            Alignment       =   1
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmSelecTrab 
      Caption         =   "(F5)"
      Height          =   855
      Left            =   5400
      Picture         =   "frCalcNewCTS.frx":15D6
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1020
      Width           =   870
   End
   Begin VB.Frame Frame1 
      Caption         =   "Planilla de CTS"
      Height          =   1725
      Left            =   135
      TabIndex        =   0
      Top             =   165
      Width           =   5145
      Begin VB.CheckBox Check1 
         Caption         =   "&Dias efectivos"
         Height          =   195
         Left            =   3720
         TabIndex        =   10
         Top             =   855
         Width           =   1365
      End
      Begin MSComCtl2.DTPicker xFechaFin 
         Height          =   330
         Left            =   1065
         TabIndex        =   6
         Top             =   1215
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd' de'MMMM'del ' yyyy"
         Format          =   23658499
         CurrentDate     =   36816
      End
      Begin MSComCtl2.DTPicker xFechaIni 
         Height          =   330
         Left            =   1065
         TabIndex        =   5
         Top             =   780
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd' de'MMMM'del ' yyyy"
         Format          =   23658499
         CurrentDate     =   36816
      End
      Begin AplisetControlText.Aplitext xPeriodo 
         Height          =   300
         Left            =   1065
         TabIndex        =   2
         Top             =   390
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   529
         Text            =   ""
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Left            =   300
         TabIndex        =   4
         Top             =   1283
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Desde:"
         Height          =   195
         Left            =   300
         TabIndex        =   3
         Top             =   848
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Periodo"
         Height          =   195
         Left            =   300
         TabIndex        =   1
         Top             =   450
         Width           =   540
      End
   End
   Begin VB.Label xProg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "**** Texto *****"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   135
      TabIndex        =   18
      Top             =   5925
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label xNumTrabs 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   2100
      TabIndex        =   16
      Top             =   5610
      Width           =   465
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Número de Trabajadores"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   135
      TabIndex        =   15
      Top             =   5640
      Width           =   1755
   End
   Begin VB.Label xTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00 "
      Height          =   255
      Left            =   3945
      TabIndex        =   14
      Top             =   5610
      Width           =   1035
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Planilla"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2910
      TabIndex        =   13
      Top             =   5640
      Width           =   900
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   4365
      Left            =   90
      Top             =   1980
      Width           =   9240
   End
End
Attribute VB_Name = "frCalcCTS2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSTRABS As ADODB.Recordset
Dim RSCALC As New ADODB.Recordset
Dim ENPROCESO As Boolean

Private Sub CMCALCULAR_CLICK()
 On Error GoTo ERRCALC
 Dim GENERAL As Boolean
    Screen.MousePointer = 11
    If RSTRABS.EOF Or RSTRABS.RecordCount = 0 Then
        MsgBox "MENSAJE DEL SISTEMA: No se puede Procesar la tarea requerida, si no ha Seleccionado uno o mas Trabajadores. Presione F5 para Seleccionar los Trabajadores", vbInformation
        Screen.MousePointer = 1
        Exit Sub
    End If
    DBSYSTEM.Execute "DELETE FROM  [##TMPCTS2" & VGL_COMPUTER & "] "
    Dim XFEC As Date, NUMMESES As Integer, NUMDIAS As Integer, XFEC2 As Date
    'PONER A 1 EL DIA DEL MES DE INICIO
    xFechaIni.Day = 1
    'PONER EL ULTIMO DIA DEL MES PARA LA FECHA FINAL
    xFechaFin.Day = 1
    xFechaFin.Value = DateAdd("M", 1, xFechaFin.Value)
    xFechaFin.Value = DateAdd("D", -1, xFechaFin.Value)
    If MsgBox("El Proceso de Calculo de C.T.S. puede tardar minutos, Desea Continuar : ", vbYesNo + vbQuestion) = vbNo Then
        Screen.MousePointer = 1
        Exit Sub
    End If
    'CREAPLANTILLA
    Prog.Min = 0
    Prog.Max = Val(xNumTrabs.Caption)
    Prog.Visible = True
    Prog.Value = 0
    xProg.Visible = True
    xProg.Caption = "Asignando Tiempo Computable"
    ENPROCESO = True
    'CALCULO DEL TIEMPO COMPUTABLE
    Dim RSCNPT As New ADODB.Recordset
    If Check1.Value = 1 Then
        RSCNPT.Open "SELECT * FROM CONCEPTOS WHERE TIPO=0 AND TIPOINFO<>5", DBSYSTEM, adOpenStatic, adLockReadOnly
        If RSCNPT.EOF Or RSCNPT.RecordCount = 0 Then
            MsgBox "No se han encontrado conceptos informativos de Dias u horas Trabajadas Computables a Beneficios Sociales", vbInformation
            Set RSCNPT = Nothing
            Screen.MousePointer = 1
            Exit Sub
        End If
        RSTRABS.MoveFirst
        Do While Not RSTRABS.EOF
            RSCNPT.MoveFirst
            Prog.Value = Prog.Value + 1
            NUMDIAS = 0
            Do While Not RSCNPT.EOF
                Select Case RSCNPT!TIPOINFO
                    Case 0
                        NUMDIAS = NUMDIAS + CALCULOCONCEPTOS(RSTRABS!CODTRAB, xFechaIni.Value, xFechaFin.Value, SUMA, RSCNPT!Codigo, False)
                    Case 1
                        NUMDIAS = NUMDIAS + CALCULOCONCEPTOS(RSTRABS!CODTRAB, xFechaIni.Value, xFechaFin.Value, SUMA, RSCNPT!Codigo, False) / 8
                    Case 3
                        NUMDIAS = NUMDIAS - CALCULOCONCEPTOS(RSTRABS!CODTRAB, xFechaIni.Value, xFechaFin.Value, SUMA, RSCNPT!Codigo, False)
                    Case 4
                        NUMDIAS = NUMDIAS - CALCULOCONCEPTOS(RSTRABS!CODTRAB, xFechaIni.Value, xFechaFin.Value, SUMA, RSCNPT!Codigo, False) / 8
                End Select
                RSCNPT.MoveNext
            Loop
            RSTRABS!Meses = NUMDIAS \ 30
            RSTRABS!Dias = NUMDIAS Mod 30
            RSTRABS.Update
            RSTRABS.MoveNext
        Loop
        Set RSCNPT = Nothing
    Else
        RSTRABS.MoveFirst
        Do While Not RSTRABS.EOF
            Prog.Value = Prog.Value + 1
            XFEC = DevuelveValor("SELECT FECHAING FROM TRABAJADORES WHERE CODTRAB='" & RSTRABS!CODTRAB & "'", DBSYSTEM)
            If IsNull(XFEC) Then
                MsgBox "El Trabajador " & RSTRABS!NOMBRES & " no presenta Fecha de Ingreso, El Sistema abortara el Proceso", vbCritical
                Screen.MousePointer = 1
                Exit Sub
            End If
            If XFEC > xFechaIni.Value Then
                If Day(XFEC) <> 1 Then
                    NUMMESES = DateDiff("M", XFEC, xFechaFin.Value)
                    XFEC2 = DateAdd("D", -1, DateAdd("M", 1, CDate("01/" & Month(XFEC) & "/" & Year(XFEC))))
                    NUMDIAS = XFEC2 - XFEC
                Else
                    NUMMESES = DateDiff("M", XFEC, xFechaFin.Value) + 1
                    NUMDIAS = 0
                End If
            Else
                NUMMESES = DateDiff("M", xFechaIni.Value, xFechaFin.Value) + 1
                NUMDIAS = 0
            End If
            RSTRABS!Meses = NUMMESES
            RSTRABS!Dias = NUMDIAS
            RSTRABS!FECHAING = XFEC
            RSTRABS.Update
            RSTRABS.MoveNext
        Loop
    End If
    Dim VALOR As Single
    RSCNPT.Open "SELECT * FROM FORMULASCTS WHERE AFECTOPRO<>0", DBSYSTEM, adOpenStatic, adLockReadOnly
    If RSCNPT.EOF Or RSCNPT.RecordCount = 0 Then
        MsgBox "MENSAJE DEL SISTEMA: El Sistema no ha Encontrado Formulas de CTS", vbInformation
        Set RSCNPT = Nothing
        Screen.MousePointer = 1
        Exit Sub
    End If
    Prog.Min = 0
    Prog.Max = Val(RSCNPT.RecordCount)
    Prog.Value = 0
    Prog.Visible = True
    xProg.Visible = True
    Do While Not RSCNPT.EOF
        GENERAL = RSCNPT!GENE
        Prog.Value = Prog.Value + 1
        xProg.Caption = "Calculando " & RSCNPT!NOMBRE
        RSTRABS.MoveFirst
        Do While Not RSTRABS.EOF
            If InStr(RSCNPT!FORMULA, "@") = 0 Then
                VALOR = DevuelveValor("SELECT " & RSCNPT!FORMULA & " AS VALOR_DEV FROM VWTRABAJ WHERE CODTRAB='" & RSTRABS!CODTRAB & "'", DBSYSTEM)
                If IsNull(VALOR) Then VALOR = 0
            Else
                If RSCNPT!CRITERIO = "" Then
                    VALOR = DevuelveValor("SELECT " & CAMBIACADENA(GENERAL, RSCNPT!FORMULA, RSTRABS!CODTRAB) & " AS VALOR_DEV FROM TRABAJADORES WHERE CODTRAB='" & RSTRABS!CODTRAB & "'", DBSYSTEM)
                Else
                    VALOR = DevuelveValor("SELECT " & CAMBIACADENA(GENERAL, RSCNPT!FORMULA, RSTRABS!CODTRAB, RSCNPT!CRITERIO) & " AS VALOR_DEV FROM TRABAJADORES WHERE CODTRAB='" & RSTRABS!CODTRAB & "'", DBSYSTEM)
                End If
            End If
            If VALOR <> 0 Then
                VALOR = Round(VALOR, 2)
                DBSYSTEM.Execute "INSERT INTO  [##TMPCTS2" & VGL_COMPUTER & "]  VALUES ('" & RSTRABS!CODTRAB & "','" & RSCNPT!NOMBRE & "'," & VALOR & "," & IIf(RSCNPT!TIPO = False, 0, 1) & ")"
            End If
            Me.Refresh
            RSTRABS.MoveNext
        Loop
        Me.Refresh
        RSCNPT.MoveNext
    Loop
    Prog.Min = 0
    Prog.Max = Val(xNumTrabs.Caption)
    Prog.Visible = True
    Prog.Value = 0
    xProg.Visible = True
    xProg.Caption = "Calculando C.T.S."
    ENPROCESO = True
    RSTRABS.MoveFirst
    Do While Not RSTRABS.EOF
        Prog.Value = Prog.Value + 1
        ACTUALIZACTS
        RSTRABS.MoveNext
    Loop
    xProg.Visible = False
    Prog.Visible = False
    ENPROCESO = False
    RSTRABS.MoveFirst
    TOTALPLANILLA
    cmGrabar.Enabled = True
    cmdAgregar.Visible = True
    cmdEliminar.Visible = True
    Screen.MousePointer = 1
    Exit Sub
ERRCALC:
    MsgBox ERR.Description
    Screen.MousePointer = 1
    Resume Next
    Exit Sub
End Sub
Private Sub CmdImprimir_Click()
'REPORTE IMPRIMIR
    
End Sub
Private Sub CREATRABSEL()
   'CREAR TABLA TEMPORAL DE LOS TRABAJADORES QUE SE HAN SELECCIONADO
    Dim RSTEMP As New ADODB.Recordset
    If ExisteTablaAux(" [##SELCODTRAB" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##SELCODTRAB" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "CREATE TABLE  [##SELCODTRAB" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8))"
    RSTEMP.Open " [##SELCODTRAB" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic
    Dim NUM As Variant
    For Each NUM In xData.SelBookmarks
        xData.Bookmark = NUM
        xData.COL = 0
        RSTEMP.AddNew
        RSTEMP!CODTRAB = Trim(xData.Text)
        RSTEMP.Update
    Next
End Sub
Private Sub CMGRABAR_CLICK()
    Dim xCodigo As Long, xSoles As Double
    If MsgBox("Seguro de Grabar los cambios de la Planilla de C.T.S.", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    If RSTRABS.RecordCount = 0 Then
        MsgBox "No existe registro para grabar", vbInformation
        Exit Sub
    End If
    
    If VPTAREA = "NUEVO" Then
        If Trim(xPeriodo.Text) = "" Then
            MsgBox "Falta ingresar el dato descriptivo de la formula para el calculo de  C.T.S.", vbInformation
            xPeriodo.SetFocus
            Exit Sub
        End If
    Else
        'SI ES MODIFICAR
        xCodigo = Val(VPTRASPRM)
        DBSYSTEM.Execute "DELETE FROM CTS WHERE CODIGO=" & xCodigo
        DBSYSTEM.Execute "DELETE FROM PLANCTS WHERE CODIGO=" & xCodigo
        DBSYSTEM.Execute "DELETE FROM DETALLECTS WHERE CODIGO=" & xCodigo
    End If
    xSoles = DevuelveValor("SELECT SUM(IMPORTECTS) AS T1 FROM  [##TMPCTS1" & VGL_COMPUTER & "] ", DBSYSTEM)
    DBSYSTEM.Execute "INSERT INTO CTS (NOMBRE, CERRADO, FECHAINI, FECHAFIN, SOLES) VALUES ('" & xPeriodo.Text & "',0," & DateSQL(xFechaIni.Value) & "," & DateSQL(xFechaFin.Value) & "," & xSoles & ")"
    xCodigo = DevuelveValor("SELECT MAX(CODIGO) AS COD1 FROM CTS", DBSYSTEM)
    DBSYSTEM.Execute "INSERT INTO PLANCTS (CODIGO, CODTRAB, NOMBRES, IMPORTECTS, MESES, DIAS, FECHAING) SELECT " & xCodigo & " AS CODIGO, CODTRAB,LTRIM(RTRIM(NOMBRES)),IMPORTECTS, MESES, DIAS, FECHAING FROM  [##TMPCTS1" & VGL_COMPUTER & "]  WHERE IMPORTECTS<>0"
    DBSYSTEM.Execute "INSERT INTO DETALLECTS SELECT " & xCodigo & " AS CODIGO, CODTRAB, CONCEPTO, IMPORTE, INDTIPO FROM  [##TMPCTS2" & VGL_COMPUTER & "]   WHERE IMPORTE<>0"
    MsgBox "La Información fue grabada satisfactorimente", vbInformation
End Sub

Private Sub CMSELECTRAB_CLICK()
    If xFechaFin.Value < xFechaIni.Value Then
        MsgBox "La Fecha dFinal no puede ser menor que la Fecha de Inicio", vbCritical
        Exit Sub
    End If
    REGSELECT.USARFECHACESE = True
    REGSELECT.FECHACESEMAX = xFechaFin.Value
    REGSELECT.FECHAINIMAX = xFechaFin.Value
    REGSELECT.FECHAINI = xFechaIni.Value
    frSelect.Show 1
    REGSELECT.USARFECHACESE = False
    If MsgBox("Desea asignar la CTS a los Trabajadores", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    DBSYSTEM.Execute "DELETE FROM  [##TMPCTS1" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "DELETE FROM  [##TMPCTS2" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "INSERT INTO  [##TMPCTS1" & VGL_COMPUTER & "]  (CODTRAB,NOMBRES) SELECT CODTRAB, LTRIM(RTRIM(NOMBRES)) FROM  [##TMPSELECT" & VGL_COMPUTER & "] "
    Set RSTRABS = Nothing
    Set RSTRABS = New ADODB.Recordset
    RSTRABS.Open " [##TMPCTS1" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic
    Set xData.DataSource = RSTRABS
    xNumTrabs.Caption = RSTRABS.RecordCount
End Sub
Private Sub Command2_Click()
    Unload Me
End Sub
Private Sub FORM_KEYDOWN(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = vbKeyF5 Then CMSELECTRAB_CLICK
End Sub
Private Sub Form_Load()
xFechaIni.Value = "01/" & frAdminCTS.xMes & "/" & frAdminCTS.xAnno
xFechaFin.Value = xFechaIni.Value
xFechaIni.Enabled = False
xFechaFin.Enabled = False
    xPeriodo.Text = AMESES(xFechaIni.Month) & "." & Year(xFechaIni.Value) & " A " & AMESES(xFechaFin.Month) & "." & Year(xFechaFin.Value)
    If ExisteTablaAux(" [##TMPCTS1" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPCTS1" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "CREATE TABLE  [##TMPCTS1" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8), NOMBRES varchar(100), IMPORTECTS  Numeric(20,2) , MESES int, DIAS int, FECHAING datetime)"
    If ExisteTablaAux(" [##TMPCTS2" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPCTS2" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "CREATE TABLE  [##TMPCTS2" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8), CONCEPTO varchar(100), IMPORTE  Numeric(20,2) , INDTIPO int)"
    DBSYSTEM.Execute "CREATE INDEX CODTRAB ON  [##TMPCTS2" & VGL_COMPUTER & "]  (CODTRAB) "
    Select Case VPTAREA
        Case "NUEVO"
            Me.Caption = "NUEVO CÁLCULO DE C.T.S."
            CmdImprimir.Visible = False
        Case "MODIFICAR"
            CmdImprimir.Visible = False
            Me.Caption = "MODIFICACIÓN DEL CÁLCULO DE C.T.S."
            Frame1.Enabled = False
            CARGADATOS
            cmGrabar.Enabled = True
            cmdAgregar.Visible = True
            cmdEliminar.Visible = True
        Case "VISTA"
            CmdImprimir.Visible = True
            Frame1.Enabled = False
            cmSelecTrab.Enabled = False
            cmActualizar.Visible = False
            cmCalcular.Visible = False
            Me.Caption = "CONSULTA DEL CÁLCULO DE C.T.S."
            CARGADATOS
        Case "PRUEBA"
            CmdImprimir.Visible = False
            cmGrabar.Visible = False
            cmdAgregar.Visible = True
            cmdEliminar.Visible = True
    End Select
    Set RSTRABS = New ADODB.Recordset
    RSTRABS.Open " [##TMPCTS1" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic
    Set xData.DataSource = RSTRABS
    TOTALPLANILLA
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSTRABS = Nothing
    Set RSCALC = Nothing
End Sub

Public Sub ACTUALIZACTS()
    On Error Resume Next
    Dim VALOR As Single
    VALOR = DevuelveValor("SELECT SUM(IMPORTE) AS T1 FROM  [##TMPCTS2" & VGL_COMPUTER & "]  WHERE CODTRAB='" & RSTRABS!CODTRAB & "' AND INDTIPO=0", DBSYSTEM)
    If VALOR <> 0 Then
        VALOR = IIf(RSTRABS!Meses = 0, 0, (VALOR * (Valc(xTasa.Text) / 100))) + IIf(RSTRABS!Dias = 0, 0, (((VALOR * (Valc(xTasa.Text) / 100)) / 30) * RSTRABS!Dias))
        'VALOR = (VALOR * (Val(xTasa.Text) / 100))
    Else
        VALOR = 0
    End If
    RSTRABS!IMPORTECTS = Round(VALOR, 2)
    VALOR = 0
    On Error Resume Next
    VALOR = DevuelveValor("SELECT SUM(IMPORTE) AS T1 FROM  [##TMPCTS2" & VGL_COMPUTER & "]  WHERE CODTRAB='" & RSTRABS!CODTRAB & "' AND INDTIPO=1", DBSYSTEM)
    If VALOR <> 0 Then
        RSTRABS!IMPORTECTS = RSTRABS!IMPORTECTS + Round(VALOR, 2)
    End If
    RSTRABS.Update
End Sub

Public Sub TOTALPLANILLA()
    DBSYSTEM.Execute "UPDATE  [##TMPCTS1" & VGL_COMPUTER & "]  SET IMPORTECTS=0 WHERE (IMPORTECTS)IS NULL"
    xTotal.Caption = Format(DevuelveValor("SELECT SUM(IMPORTECTS) AS T1 FROM  [##TMPCTS1" & VGL_COMPUTER & "] ", DBSYSTEM), "0.00 ")
    xNumTrabs.Caption = RSTRABS.RecordCount
End Sub

Private Sub XDATA_HEADCLICK(ByVal COLINDEX As Integer)
On Error Resume Next
    RSTRABS.Sort = xData.Columns(COLINDEX).DataField
End Sub

Public Sub CARGADATOS()
On Error Resume Next
    xPeriodo.Text = DevuelveValor("SELECT NOMBRE FROM CTS WHERE CODIGO=" & VPTRASPRM, DBSYSTEM)
    xFechaIni.Value = DevuelveValor("SELECT FECHAINI FROM CTS WHERE CODIGO=" & VPTRASPRM, DBSYSTEM)
    xFechaFin.Value = DevuelveValor("SELECT FECHAFIN FROM CTS WHERE CODIGO=" & VPTRASPRM, DBSYSTEM)
    DBSYSTEM.Execute "INSERT INTO  [##TMPCTS1" & VGL_COMPUTER & "]  SELECT CODTRAB, NOMBRES, IMPORTECTS, MESES, DIAS, FECHAING FROM PLANCTS WHERE CODIGO=" & VPTRASPRM
    DBSYSTEM.Execute "INSERT INTO  [##TMPCTS2" & VGL_COMPUTER & "]  SELECT CODTRAB, CONCEPTO, IMPORTE, INDTIPO FROM DETALLECTS WHERE CODIGO=" & VPTRASPRM
End Sub

Public Function CAMBIACADENA(GENERAL As Boolean, ByVal CADENA As String, ByVal CODTRAB As String, Optional MES As String = "NONE") As String
    Dim POSARROBA As Integer, POS1 As Integer, PROCESO As String, CAMPO As String, POS2 As Integer
    Dim VALOR As Double
    POSARROBA = 1
    POSARROBA = InStr(POSARROBA, CADENA, "@")
    Do While POSARROBA <> 0
        POS1 = InStr(POSARROBA, CADENA, "(")
        PROCESO = Mid(CADENA, POSARROBA + 1, POS1 - (POSARROBA + 1))
        POS2 = InStr(POSARROBA, CADENA, ")")
        CAMPO = Mid(CADENA, POS1 + 1, POS2 - (POS1 + 1))
        Select Case UCase(PROCESO)
            Case "PROMEDIO"
                VALOR = CALCULOCONCEPTOS(CODTRAB, xFechaIni.Value, xFechaFin.Value, PROMEDIO, CAMPO, GENERAL)
            Case "ULTIMOVALOR"
                VALOR = CALCULOCONCEPTOS(CODTRAB, xFechaIni.Value, xFechaFin.Value, ULTIMOVALOR, CAMPO, GENERAL)
            Case "PRIMERVALOR"
                VALOR = CALCULOCONCEPTOS(CODTRAB, xFechaIni.Value, xFechaFin.Value, PRIMERVALOR, CAMPO, GENERAL)
            Case "SUMA"
                VALOR = CALCULOCONCEPTOS(CODTRAB, xFechaIni.Value, xFechaFin.Value, SUMA, CAMPO, GENERAL)
            Case "MEDIA"
                VALOR = CALCULOCONCEPTOS(CODTRAB, xFechaIni.Value, xFechaFin.Value, MEDIA, CAMPO, GENERAL)
            Case "PROMEDIOVALOR"
                VALOR = CALCULOCONCEPTOS(CODTRAB, xFechaIni.Value, xFechaFin.Value, PROMEDIOVALOR, CAMPO, GENERAL)
            Case "PRIMERO"
                VALOR = CALCULOCONCEPTOS(CODTRAB, xFechaIni.Value, xFechaFin.Value, PRIMERO, CAMPO, GENERAL)
            Case "ULTIMO"
                VALOR = CALCULOCONCEPTOS(CODTRAB, xFechaIni.Value, xFechaFin.Value, ULTIMO, CAMPO, GENERAL)
            Case "MAYORVALOR"
                VALOR = CALCULOCONCEPTOS(CODTRAB, xFechaIni.Value, xFechaFin.Value, MAYORVALOR, CAMPO, GENERAL)
            Case "MENORVALOR"
                VALOR = CALCULOCONCEPTOS(CODTRAB, xFechaIni.Value, xFechaFin.Value, MENORVALOR, CAMPO, GENERAL)
            Case "NUMERO"
                VALOR = CALCULOCONCEPTOS(CODTRAB, xFechaIni.Value, xFechaFin.Value, Numero, CAMPO, GENERAL)
            Case "NSECUENCIA"
                VALOR = CALCULOCONCEPTOS(CODTRAB, xFechaIni.Value, xFechaFin.Value, NSECUENCIA, CAMPO, GENERAL)
            Case "INFOBOLETA"
                If MES = "NONE" Then
                    VALOR = CALCULOMES(CODTRAB, CAMPO, , xFechaIni.Value, xFechaFin.Value)
                Else
                    VALOR = CALCULOMES(CODTRAB, CAMPO, MES)
                End If
        End Select
        If IsNull(VALOR) Then VALOR = 0
        CADENA = Replace(CADENA, Mid(CADENA, POSARROBA, (POS2 - POSARROBA) + 1), "" & VALOR)
        POSARROBA = InStr(POSARROBA, CADENA, "@")
    Loop
    CAMBIACADENA = CADENA
End Function

Private Sub CMDAGREGAR_CLICK()
    On Error Resume Next
    FrmAgrCon.Show 1
    If FrmAgrCon.VarGrabar = False Then Exit Sub
    RSCALC.AddNew
    RSCALC!CODTRAB = RSTRABS!CODTRAB
    RSCALC!CONCEPTO = FrmAgrCon.CONCEPTO
    RSCALC!Importe = FrmAgrCon.Importe
    RSCALC!INDTIPO = FrmAgrCon.TIPO
    RSCALC.Update
End Sub

Private Sub CMDELIMINAR_CLICK()
    On Error Resume Next
    If RSCALC.RecordCount = 0 Then Exit Sub
    If MsgBox("Desea eliminar el Registro Seleccionado", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    If Not (RSCALC.EOF Or RSCALC.BOF) Then RSCALC.Delete
End Sub

Private Sub XFECHAFIN_CHANGE()
    xPeriodo.Text = AMESES(xFechaIni.Month) & "." & Year(xFechaIni.Value) & " A " & AMESES(xFechaFin.Month) & "." & Year(xFechaFin.Value)
End Sub

Private Sub XFECHAINI_CHANGE()
    xPeriodo.Text = AMESES(xFechaIni.Month) & "." & Year(xFechaIni.Value) & " A " & AMESES(xFechaFin.Month) & "." & Year(xFechaFin.Value)
End Sub



