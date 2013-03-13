VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frCalcGrati 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cálculo de Gratificaciones"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9420
   Icon            =   "frCalcGrati.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   9420
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmEliminar 
      Caption         =   "&Eliminar"
      Height          =   360
      Left            =   165
      TabIndex        =   32
      Top             =   5625
      Width           =   960
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Quitar"
      Height          =   360
      Left            =   6330
      TabIndex        =   30
      Top             =   5625
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Agregar"
      Height          =   360
      Left            =   5265
      TabIndex        =   29
      Top             =   5625
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.CommandButton cmActualizar 
      Caption         =   "&Actualizar"
      Height          =   360
      Left            =   7980
      TabIndex        =   28
      Top             =   5625
      Width           =   1275
   End
   Begin AplisetControlText.Aplitext xDias 
      Height          =   270
      Left            =   8775
      TabIndex        =   24
      Top             =   2235
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   476
      MaxLength       =   5
      Text            =   "0"
      Entero          =   -1  'True
      TipoDato        =   "N"
   End
   Begin AplisetControlText.Aplitext xMeses 
      Height          =   270
      Left            =   7845
      TabIndex        =   22
      Top             =   2235
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   476
      MaxLength       =   5
      Text            =   "0"
      Entero          =   -1  'True
      TipoDato        =   "N"
   End
   Begin MSComctlLib.ProgressBar Prog 
      Height          =   135
      Left            =   165
      TabIndex        =   18
      Top             =   6225
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
      Left            =   7950
      TabIndex        =   13
      Top             =   270
      Width           =   1305
   End
   Begin VB.CommandButton cmGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   360
      Left            =   7950
      TabIndex        =   12
      Top             =   1455
      Width           =   1305
   End
   Begin Crystal.CrystalReport Reporte 
      Left            =   225
      Top             =   2940
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Formulas de Gratificaciones"
      Height          =   510
      Left            =   5415
      TabIndex        =   11
      Top             =   1380
      Width           =   1830
   End
   Begin VB.CommandButton cmCalcular 
      Caption         =   "&Calcular"
      Height          =   855
      Left            =   6360
      Picture         =   "frCalcGrati.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   270
      Width           =   870
   End
   Begin MSDataGridLib.DataGrid xData 
      Height          =   3540
      Left            =   135
      TabIndex        =   8
      Top             =   2025
      Width           =   5145
      _ExtentX        =   9075
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
      ColumnCount     =   5
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
         Caption         =   "Importe de Gratificación"
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
      BeginProperty Column04 
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         ScrollBars      =   2
         BeginProperty Column00 
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2700.284
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   404.787
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   404.787
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmSelecTrab 
      Caption         =   "(F5)"
      Height          =   855
      Left            =   5415
      Picture         =   "frCalcGrati.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   270
      Width           =   870
   End
   Begin VB.Frame Frame1 
      Caption         =   "Planilla de Gratificación"
      Height          =   1725
      Left            =   135
      TabIndex        =   0
      Top             =   165
      Width           =   5145
      Begin VB.CheckBox xMesesCompletos 
         Caption         =   "&Meses completos"
         Height          =   195
         Left            =   3510
         TabIndex        =   31
         Top             =   1283
         Width           =   1545
      End
      Begin VB.CheckBox Check1 
         Caption         =   "&Dias efectivos"
         Height          =   195
         Left            =   3510
         TabIndex        =   10
         Top             =   848
         Visible         =   0   'False
         Width           =   1365
      End
      Begin MSComCtl2.DTPicker xFechaFin 
         Height          =   330
         Left            =   1065
         TabIndex        =   6
         Top             =   1215
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "MMMM'del ' yyyy"
         Format          =   23658499
         CurrentDate     =   36816
      End
      Begin MSComCtl2.DTPicker xFechaIni 
         Height          =   330
         Left            =   1065
         TabIndex        =   5
         Top             =   780
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "MMMM'del ' yyyy"
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
   Begin MSDataGridLib.DataGrid xDetalle 
      Height          =   3060
      Left            =   5265
      TabIndex        =   27
      Top             =   2505
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   5398
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            DividerStyle    =   1
            ColumnWidth     =   1005.165
         EndProperty
      EndProperty
   End
   Begin VB.Label xFecha 
      Alignment       =   2  'Center
      Height          =   225
      Left            =   6105
      TabIndex        =   26
      Top             =   2265
      Width           =   1140
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " F. Ingreso"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5265
      TabIndex        =   25
      Top             =   2235
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
      Left            =   8325
      TabIndex        =   23
      Top             =   2235
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
      Left            =   7275
      TabIndex        =   21
      Top             =   2235
      Width           =   600
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tiempo Computable"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   5265
      TabIndex        =   20
      Top             =   2025
      Width           =   4020
   End
   Begin VB.Label xProg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "**** Texto *****"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   180
      TabIndex        =   19
      Top             =   6015
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label xNumTrabs 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   2100
      TabIndex        =   17
      Top             =   5610
      Width           =   465
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Núm. Trabs"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1185
      TabIndex        =   16
      Top             =   5640
      Width           =   825
   End
   Begin VB.Label xTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00 "
      Height          =   255
      Left            =   3945
      TabIndex        =   15
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
      TabIndex        =   14
      Top             =   5640
      Width           =   900
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   4440
      Left            =   90
      Top             =   1980
      Width           =   9240
   End
End
Attribute VB_Name = "frCalcGrati"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSTRABS As ADODB.Recordset
Dim RSCALC As New ADODB.Recordset
Dim ENPROCESO As Boolean

Private Sub CMACTUALIZAR_CLICK()
    If Val(xDias.Text) > 30 Then
        MsgBox "No se puede actualizar. Los dias no pueden ser mayores a 30."
        Exit Sub
    End If
    If Val(xMeses.Text) > 12 Then
        MsgBox "No se puede actualizar. Los meses no pueden ser mayores a 12."
        Exit Sub
    End If
    RSTRABS!Meses = Val(xMeses.Text)
    RSTRABS!Dias = Val(xDias.Text)
    RSTRABS.Update
    ACTUALIZACTS
    TOTALPLANILLA
End Sub

Private Sub CMCALCULAR_CLICK()
 On Error GoTo ERRCALC
 Dim GENERAL As Boolean
    Screen.MousePointer = 11
    If RSTRABS.EOF Or RSTRABS.RecordCount = 0 Then
        MsgBox "Mensaje del Sistema: No se puede procesar la tarea requerida, si no ha seleccionado uno o mas trabajadores. Presione F5 para seleccionar trabajadores", vbInformation
        Exit Sub
    End If
    DBSYSTEM.Execute "DELETE FROM  [##TMPCTS2" & VGL_COMPUTER & "] "
    Dim XFEC As Date, NUMMESES As Integer, NUMDIAS As Integer, XFEC2 As Date
    'PONER A 1 EL DIA DEL MES DE INICIO
    'PONER EL ULTIMO DIA DEL MES PARA LA FECHA FINAL
    If MsgBox("El proceso de cálculo de Gratificacion puede tardar varios minutos, desea continuar: ", vbYesNo + vbQuestion) = vbNo Then
        Screen.MousePointer = 1
        Exit Sub
    End If

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
            MsgBox "No se han encontrado Conceptos Informativos de Dias u horas trabajadas computables a beneficios sociales", vbInformation
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
                GENERAL = RSCNPT!GENE
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
            If xMesesCompletos.Value = 1 Then
                RSTRABS!Meses = Int(NUMDIAS \ 30)
                RSTRABS!Dias = 0
            Else
                RSTRABS!Meses = NUMDIAS \ 30
                RSTRABS!Dias = NUMDIAS Mod 30
            End If
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
                MsgBox "El trabajador " & RSTRABS!NOMBRES & " no presenta fecha de ingreso, el sistema abortará el proceso", vbCritical
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
            If xMesesCompletos.Value = 0 Then RSTRABS!Dias = NUMDIAS Else RSTRABS!Dias = 0
            RSTRABS!FECHAING = XFEC
            RSTRABS.Update
            RSTRABS.MoveNext
        Loop
    End If
    Dim VALOR As Single

    RSCNPT.Open "SELECT * FROM FORMULASGRATI WHERE AFECTOPRO<>0", DBSYSTEM, adOpenStatic, adLockReadOnly
    If RSCNPT.EOF Or RSCNPT.RecordCount = 0 Then
        MsgBox "Mensaje del Sistema: El sistema no ha encontrado Fórmulas de Gratificación", vbInformation
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
                VALOR = DevuelveValor("SELECT " & RSCNPT!FORMULA & " AS VALOR_DEV FROM TRABAJADORES WHERE CODTRAB='" & RSTRABS!CODTRAB & "'", DBSYSTEM)
                If IsNull(VALOR) Then VALOR = 0
            Else
                VALOR = DevuelveValor("SELECT " & CAMBIACADENA(RSCNPT!FORMULA, RSTRABS!CODTRAB, GENERAL) & " AS VALOR_DEV FROM TRABAJADORES WHERE CODTRAB='" & RSTRABS!CODTRAB & "'", DBSYSTEM)
            End If
            If VALOR <> 0 Then
                VALOR = Round(VALOR, 2)
                DBSYSTEM.Execute "INSERT INTO  [##TMPCTS2" & VGL_COMPUTER & "]  VALUES ('" & RSTRABS!CODTRAB & "','" & RSCNPT!NOMBRE & "'," & VALOR & "," & IIf(RSCNPT!TIPO, 1, 0) & ")"
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
    xProg.Caption = "Calculando " & xPeriodo.Text
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
    Exit Sub
    Screen.MousePointer = 1
End Sub

Private Sub CMELIMINAR_CLICK()
    If RSTRABS.EOF Then Exit Sub
    DBSYSTEM.Execute "DELETE FROM  [##TMPCTS1" & VGL_COMPUTER & "]  WHERE CODTRAB='" & RSTRABS!CODTRAB & "'"
    DBSYSTEM.Execute "DELETE FROM  [##TMPCTS2" & VGL_COMPUTER & "]  WHERE CODTRAB='" & RSTRABS!CODTRAB & "'"
    RSTRABS.Requery
    Set xData.DataSource = RSTRABS
    xNumTrabs.Caption = RSTRABS.RecordCount
End Sub

Private Sub CMGRABAR_CLICK()
    Dim xCodigo As Long, xSoles As Double
    If MsgBox("Seguro de grabar los cambios en la planilla de Gratificación", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    If RSTRABS.RecordCount = 0 Then
        MsgBox "No existe nada por grabar", vbInformation
        Exit Sub
    End If
    If VPTAREA = "NUEVO" Then
        If Trim(xPeriodo.Text) = "" Then
            MsgBox "Falta especificar un nombre descriptivo de Cálculo para la Gratificación", vbInformation
            xPeriodo.SetFocus
            Exit Sub
        End If
        VPTAREA = "MODIFICAR"
        VPTRASPRM = "" & DevuelveValor("SELECT MAX(CODIGO) AS COD1 FROM GRATIFICACION", DBSYSTEM)
    Else
        'SI ES MODIFICAR
        xCodigo = Val(VPTRASPRM)
        DBSYSTEM.Execute "DELETE FROM GRATIFICACION WHERE CODIGO=" & xCodigo
        DBSYSTEM.Execute "DELETE FROM PLANGRATI WHERE CODIGO=" & xCodigo
        DBSYSTEM.Execute "DELETE FROM DETALLEGRATI WHERE CODIGO=" & xCodigo
    End If
    xSoles = DevuelveValor("SELECT SUM(IMPORTECTS) AS T1 FROM  [##TMPCTS1" & VGL_COMPUTER & "] ", DBSYSTEM)
    DBSYSTEM.Execute "INSERT INTO GRATIFICACION (NOMBRE, CERRADO, FECHAINI, FECHAFIN, SOLES) VALUES ('" & xPeriodo.Text & "',0," & DateSQL(xFechaIni.Value) & "," & DateSQL(xFechaFin.Value) & "," & xSoles & ")"
    xCodigo = DevuelveValor("SELECT MAX(CODIGO) AS COD1 FROM GRATIFICACION", DBSYSTEM)
    DBSYSTEM.Execute "INSERT INTO PLANGRATI (CODIGO, CODTRAB, NOMBRES, IMPORTEGRATI, MESES, DIAS, FECHAING) SELECT " & xCodigo & " AS CODIGO, CODTRAB,NOMBRES,IMPORTECTS, MESES, DIAS, FECHAING FROM  [##TMPCTS1" & VGL_COMPUTER & "]  WHERE IMPORTECTS<>0"
    DBSYSTEM.Execute "INSERT INTO DETALLEGRATI SELECT " & xCodigo & " AS CODIGO, CODTRAB, CONCEPTO, IMPORTE, INDTIPO FROM  [##TMPCTS2" & VGL_COMPUTER & "]  WHERE IMPORTE<>0"
    MsgBox "La Información se grabó Satisfactorimente", vbInformation
End Sub

Private Sub CMSELECTRAB_CLICK()
    If xFechaIni.Value = xFechaFin.Value Then
        MsgBox "Fechas no validas", vbInformation
        Exit Sub
    End If
    If xFechaFin.Value < xFechaIni.Value Then
        MsgBox "La Fecha Final no puede ser menor que la Fecha de Inicio", vbCritical
        Exit Sub
    End If
    REGSELECT.USARFECHACESE = True
    REGSELECT.FECHACESEMAX = xFechaIni.Value
    REGSELECT.FECHAINIMAX = xFechaFin.Value
    REGSELECT.FECHAINI = xFechaFin.Value
    Load frSelect
    frSelect.xFecha.Value = xFechaFin.Value
    frSelect.Show 1
    REGSELECT.USARFECHACESE = False
    If MsgBox("Desea asignar los trabajadores al cálculo de Gratificación", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    DBSYSTEM.Execute "DELETE FROM  [##TMPCTS1" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "DELETE FROM  [##TMPCTS2" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "INSERT INTO  [##TMPCTS1" & VGL_COMPUTER & "]  (CODTRAB,NOMBRES) SELECT CODTRAB, LTRIM(RTRIM(NOMBRES)) FROM  [##TMPSELECT" & VGL_COMPUTER & "] "
    Set RSTRABS = Nothing
    Set RSTRABS = New ADODB.Recordset
    RSTRABS.Open " [##TMPCTS1" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic
    Set xData.DataSource = RSTRABS
    xNumTrabs.Caption = RSTRABS.RecordCount
End Sub

Private Sub Command1_Click()
    frFormulasGrati.Show 1
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub FORM_KEYDOWN(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = vbKeyF5 Then CMSELECTRAB_CLICK
End Sub

Private Sub Form_Load()
On Error GoTo handler
    xFechaFin.Value = Date
    xFechaIni.Value = Date
    xFechaIni.Day = 1
    If Month(xFechaIni.Value) > 7 Then
        'ASUMIR QUE ES LA GRATIFICACIÓN DE NAVIDAD
        xFechaIni.Month = 7
        xFechaFin.Day = Ultmes(xFechaFin.Value)
        xFechaFin.Month = 12
        xPeriodo.Text = "Gratificacion por Navidad - " & Year(Date)
    Else
        xFechaIni.Month = 1
        xFechaFin.Day = Ultmes(xFechaFin.Value)
        xFechaFin.Month = 6
        xPeriodo.Text = "Gratificacion por Fiestas Patrias - " & Year(Date)
    End If
    If ExisteTablaAux(" [##TMPCTS1" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPCTS1" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "CREATE TABLE  [##TMPCTS1" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8), NOMBRES varchar(100), IMPORTECTS  Numeric(20,2) , MESES int, DIAS int, FECHAING datetime)"
    If ExisteTablaAux(" [##TMPCTS2" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPCTS2" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "CREATE TABLE  [##TMPCTS2" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8), CONCEPTO varchar(100), IMPORTE  Numeric(20,2) , INDTIPO bit)"
    Select Case VPTAREA
        Case "NUEVO"
            Me.Caption = "Nuevo Calculo de Gratificación"
        Case "MODIFICAR"
            Me.Caption = "Modificación del Calculo de Gratificación"
            Frame1.Enabled = True
            CARGADATOS
            cmGrabar.Enabled = True
            cmdAgregar.Visible = True
            cmdEliminar.Visible = True
        Case "VISTA"
            Frame1.Enabled = False
            xDetalle.AllowUpdate = False
            xMeses.Locked = True
            xDias.Locked = True
            cmSelecTrab.Enabled = False
            cmActualizar.Visible = False
            cmCalcular.Visible = False
            Me.Caption = "Consulta del Calculo de Gratificación"
            CARGADATOS
        Case "PRUEBA"
            cmGrabar.Visible = False
            cmdAgregar.Visible = True
            cmdEliminar.Visible = True
    End Select
    Set RSTRABS = New ADODB.Recordset
    RSTRABS.Open " [##TMPCTS1" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic
    Set xData.DataSource = RSTRABS
    TOTALPLANILLA
    XDATA_ROWCOLCHANGE 0, 0
Exit Sub
handler:
    MsgBox ERR.Description, vbCritical, "Revise la Fecha "
 Exit Sub
 Resume
 
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSTRABS = Nothing
    Set RSCALC = Nothing
End Sub

Public Sub ACTUALIZACTS()
    Dim VALOR
    VALOR = DevuelveValor("SELECT SUM(IMPORTE) AS T1 FROM  [##TMPCTS2" & VGL_COMPUTER & "]  WHERE CODTRAB='" & RSTRABS!CODTRAB & "' AND INDTIPO=0", DBSYSTEM)
    If IsNull(VALOR) Then VALOR = 0
    If VALOR <> 0 Then
        VALOR = IIf(RSTRABS!Meses = 0, 0, (VALOR / 6 * RSTRABS!Meses)) + IIf(RSTRABS!Dias = 0, 0, (VALOR / 6 / 30 * RSTRABS!Dias))
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
    RSTRABS.Sort = xData.Columns(COLINDEX).DataField
End Sub

Private Sub XDATA_ROWCOLCHANGE(LASTROW As Variant, ByVal LASTCOL As Integer)
    If RSTRABS.EOF Or RSTRABS.RecordCount = 0 Then
        cmCalcular.Enabled = False
        cmActualizar.Enabled = False
        xMeses.Text = 0
        xDias.Text = 0
        Exit Sub
    Else
        cmActualizar.Enabled = True
        cmCalcular.Enabled = True
    End If
    If ENPROCESO Then Exit Sub
    xMeses.Text = IIf(IsNull(RSTRABS!Meses), 0, RSTRABS!Meses)
    xDias.Text = IIf(IsNull(RSTRABS!Dias), 0, RSTRABS!Dias)
    xFecha.Caption = IIf(IsNull(RSTRABS!FECHAING), "", RSTRABS!FECHAING)
    Set RSCALC = Nothing
    RSCALC.Open "SELECT * FROM  [##TMPCTS2" & VGL_COMPUTER & "]  WHERE CODTRAB='" & RSTRABS!CODTRAB & "'", DBSYSTEM, adOpenStatic, adLockOptimistic
    Set xDetalle.DataSource = RSCALC
    If RSCALC.RecordCount = 0 Then cmActualizar.Enabled = False
End Sub

Private Sub XDETALLE_AFTERCOLUPDATE(ByVal COLINDEX As Integer)
    RSCALC.MOVE 0
    CMACTUALIZAR_CLICK
End Sub

Public Sub CARGADATOS()
    xPeriodo.Text = DevuelveValor("SELECT NOMBRE FROM GRATIFICACION WHERE CODIGO=" & VPTRASPRM, DBSYSTEM)
    xFechaIni.Value = DevuelveValor("SELECT FECHAINI FROM GRATIFICACION WHERE CODIGO=" & VPTRASPRM, DBSYSTEM)
    xFechaFin.Value = DevuelveValor("SELECT FECHAFIN FROM GRATIFICACION WHERE CODIGO=" & VPTRASPRM, DBSYSTEM)
    DBSYSTEM.Execute "INSERT INTO  [##TMPCTS1" & VGL_COMPUTER & "]  SELECT CODTRAB, NOMBRES, IMPORTEGRATI AS IMPORTECTS, MESES, DIAS, FECHAING FROM PLANGRATI WHERE CODIGO=" & VPTRASPRM
    DBSYSTEM.Execute "INSERT INTO  [##TMPCTS2" & VGL_COMPUTER & "]  SELECT CODTRAB, CONCEPTO, IMPORTE, INDTIPO FROM DETALLEGRATI WHERE CODIGO=" & VPTRASPRM
End Sub

Public Function CAMBIACADENA(ByVal CADENA As String, ByVal CODTRAB As String, GENERAL2 As Boolean) As String
On Error GoTo ERRCAM
    Dim POSARROBA As Integer, POS1 As Integer, PROCESO As String, CAMPO As String, POS2 As Integer
    Dim VALOR As Double
    POSARROBA = 1
    POSARROBA = InStr(POSARROBA, CADENA, "@")
    Do While POSARROBA <> 0
        POS1 = InStr(POSARROBA, CADENA, "(")
        PROCESO = Mid(CADENA, POSARROBA + 1, POS1 - (POSARROBA + 1))
        POS2 = InStr(POSARROBA, CADENA, ")")
        CAMPO = Mid(CADENA, POS1 + 1, POS2 - (POS1 + 1))
        xFechaIni.Value = DateAdd("m", -1, xFechaIni.Value)
        xFechaFin.Value = DateAdd("m", -1, xFechaFin.Value)
        Select Case UCase(PROCESO)
            Case "PROMEDIO"
                VALOR = CALCULOCONCEPTOS(CODTRAB, xFechaIni.Value, xFechaFin.Value, PROMEDIO, CAMPO, GENERAL2)
            Case "ULTIMOVALOR"
                VALOR = CALCULOCONCEPTOS(CODTRAB, xFechaIni.Value, xFechaFin.Value, ULTIMOVALOR, CAMPO, GENERAL2)
            Case "PRIMERVALOR"
                VALOR = CALCULOCONCEPTOS(CODTRAB, xFechaIni.Value, xFechaFin.Value, PRIMERVALOR, CAMPO, GENERAL2)
            Case "SUMA"
                VALOR = CALCULOCONCEPTOS(CODTRAB, xFechaIni.Value, xFechaFin.Value, SUMA, CAMPO, GENERAL2)
            Case "MEDIA"
                VALOR = CALCULOCONCEPTOS(CODTRAB, xFechaIni.Value, xFechaFin.Value, MEDIA, CAMPO, GENERAL2)
            Case "PROMEDIOVALOR"
                VALOR = CALCULOCONCEPTOS(CODTRAB, xFechaIni.Value, xFechaFin.Value, PROMEDIOVALOR, CAMPO, GENERAL2)
            Case "PRIMERO"
                VALOR = CALCULOCONCEPTOS(CODTRAB, xFechaIni.Value, xFechaFin.Value, PRIMERO, CAMPO, GENERAL2)
            Case "ULTIMO"
                VALOR = CALCULOCONCEPTOS(CODTRAB, xFechaIni.Value, xFechaFin.Value, ULTIMO, CAMPO, GENERAL2)
            Case "MAYORVALOR"
                VALOR = CALCULOCONCEPTOS(CODTRAB, xFechaIni.Value, xFechaFin.Value, MAYORVALOR, CAMPO, GENERAL2)
            Case "MENORVALOR"
                VALOR = CALCULOCONCEPTOS(CODTRAB, xFechaIni.Value, xFechaFin.Value, MENORVALOR, CAMPO, GENERAL2)
            Case "NUMERO"
                VALOR = CALCULOCONCEPTOS(CODTRAB, xFechaIni.Value, xFechaFin.Value, Numero, CAMPO, GENERAL2)
            Case "NSECUENCIA"
                VALOR = CALCULOCONCEPTOS(CODTRAB, xFechaIni.Value, xFechaFin.Value, NSECUENCIA, CAMPO, GENERAL2)
        End Select
        xFechaIni.Value = DateAdd("m", 1, xFechaIni.Value)
        xFechaFin.Value = DateAdd("m", 1, xFechaFin.Value)
        
        If IsNull(VALOR) Then VALOR = 0
        CADENA = Replace(CADENA, Mid(CADENA, POSARROBA, (POS2 - POSARROBA) + 1), "" & VALOR)
        POSARROBA = InStr(POSARROBA, CADENA, "@")
    Loop
    CAMBIACADENA = CADENA
    Exit Function
ERRCAM:
    Exit Function
End Function

Private Sub CMDAGREGAR_CLICK()
    FrmAgrCon.Show 1
    If FrmAgrCon.VarGrabar = False Then Exit Sub
    RSCALC.AddNew
    RSCALC!CODTRAB = RSTRABS!CODTRAB
    RSCALC!CONCEPTO = FrmAgrCon.CONCEPTO
    RSCALC!Importe = FrmAgrCon.Importe
    RSCALC!INDTIPO = FrmAgrCon.TIPO
    RSCALC.Update
    Call CMACTUALIZAR_CLICK
End Sub

Private Sub CMDELIMINAR_CLICK()
    If RSCALC.RecordCount = 0 Then Exit Sub
    If MsgBox("Desea eliminar el registro Seleccionado", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    If Not (RSCALC.EOF Or RSCALC.BOF) Then RSCALC.Delete
    Call CMACTUALIZAR_CLICK
End Sub

