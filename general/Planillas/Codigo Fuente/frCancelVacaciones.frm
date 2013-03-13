VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form frCancelVacaciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editor de Vacaciones"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   Icon            =   "frCancelVacaciones.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   7980
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Forma de goce de Vacaciones"
      Height          =   1680
      Left            =   105
      TabIndex        =   24
      Top             =   4860
      Width           =   6015
      Begin VB.CheckBox xProg 
         Caption         =   "Programar después"
         Height          =   225
         Left            =   3765
         TabIndex        =   17
         Top             =   240
         Value           =   1  'Checked
         Width           =   1680
      End
      Begin AplisetControlText.Aplitext xDiasCompensados 
         Height          =   285
         Left            =   1965
         TabIndex        =   16
         Top             =   1170
         Visible         =   0   'False
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   503
         MaxLength       =   2
         Text            =   "15"
         Entero          =   -1  'True
         TipoDato        =   "N"
      End
      Begin VB.OptionButton xForma 
         Caption         =   "Compensación de dias de vacaciones"
         Height          =   240
         Index           =   1
         Left            =   195
         TabIndex        =   15
         Top             =   870
         Width           =   3120
      End
      Begin VB.OptionButton xForma 
         Caption         =   "Descanso vacacional normal"
         Height          =   240
         Index           =   0
         Left            =   195
         TabIndex        =   14
         Top             =   525
         Value           =   -1  'True
         Width           =   2400
      End
      Begin MSComCtl2.DTPicker xFechaIni 
         Height          =   300
         Left            =   4425
         TabIndex        =   18
         Top             =   885
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   24838145
         CurrentDate     =   36697
      End
      Begin MSComCtl2.DTPicker xFechaFin 
         Height          =   300
         Left            =   4425
         TabIndex        =   19
         Top             =   1260
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   24838145
         CurrentDate     =   36697
      End
      Begin VB.Label xDiasFisicos 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4425
         TabIndex        =   33
         Top             =   540
         Width           =   585
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Días"
         Height          =   195
         Left            =   3765
         TabIndex        =   32
         Top             =   600
         Width           =   345
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "F. Final"
         Height          =   195
         Left            =   3765
         TabIndex        =   31
         Top             =   1320
         Width           =   510
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "F. Inicio"
         Height          =   195
         Left            =   3765
         TabIndex        =   30
         Top             =   975
         Width           =   555
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Dias Compensados"
         Height          =   195
         Left            =   465
         TabIndex        =   25
         Top             =   1215
         Width           =   1365
      End
   End
   Begin VB.CommandButton cmCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   6450
      TabIndex        =   21
      Top             =   5835
      Width           =   1380
   End
   Begin VB.CommandButton cmAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   6450
      TabIndex        =   20
      Top             =   5310
      Width           =   1380
   End
   Begin VB.Frame Frame1 
      Caption         =   "Descripción del Trabajador"
      Height          =   4575
      Left            =   105
      TabIndex        =   0
      Top             =   135
      Width           =   7770
      Begin VB.Frame Frame3 
         Caption         =   "Periodo de Cálculo"
         Height          =   2310
         Left            =   5490
         TabIndex        =   38
         Top             =   180
         Width           =   2145
         Begin AplisetControlText.Aplitext xDescPeriodo 
            Height          =   300
            Left            =   225
            TabIndex        =   13
            Top             =   1845
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            MaxLength       =   30
            Text            =   ""
         End
         Begin MSComCtl2.DTPicker xFechaIniCal 
            Height          =   300
            Left            =   225
            TabIndex        =   11
            Top             =   585
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            Format          =   24838145
            CurrentDate     =   36697
         End
         Begin MSComCtl2.DTPicker xFechaFinCal 
            Height          =   300
            Left            =   225
            TabIndex        =   12
            Top             =   1215
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            Format          =   24838145
            CurrentDate     =   36697
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Desc. Periodo"
            Height          =   195
            Left            =   225
            TabIndex        =   41
            Top             =   1605
            Width           =   1005
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Inicio"
            Height          =   195
            Left            =   225
            TabIndex        =   40
            Top             =   330
            Width           =   1095
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Final"
            Height          =   195
            Left            =   225
            TabIndex        =   39
            Top             =   960
            Width           =   825
         End
      End
      Begin VB.OptionButton xtipo 
         Caption         =   "Un solo &monto de vacaciones"
         Height          =   210
         Index           =   0
         Left            =   135
         TabIndex        =   4
         Top             =   1935
         Value           =   -1  'True
         Width           =   2580
      End
      Begin VB.OptionButton xtipo 
         Caption         =   "Cálculo &detallado"
         Height          =   210
         Index           =   1
         Left            =   135
         TabIndex        =   5
         Top             =   2295
         Width           =   1680
      End
      Begin VB.CommandButton cmAddDetalle 
         Caption         =   "Agregar"
         Height          =   360
         Left            =   5475
         TabIndex        =   9
         Top             =   2640
         Width           =   1200
      End
      Begin VB.CommandButton cmQuitar 
         Caption         =   "&Quitar"
         Height          =   360
         Left            =   5475
         TabIndex        =   10
         Top             =   3105
         Width           =   1200
      End
      Begin MSComCtl2.DTPicker xFechaIng 
         Height          =   315
         Left            =   1515
         TabIndex        =   23
         Top             =   1065
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   24838145
         CurrentDate     =   36801
      End
      Begin AplisetControlText.Aplitext xTrab 
         Height          =   300
         Left            =   135
         TabIndex        =   1
         Top             =   585
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   529
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xPeriodo 
         Height          =   300
         Left            =   1965
         TabIndex        =   2
         Top             =   1500
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   529
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xDet2 
         Height          =   285
         Left            =   4170
         TabIndex        =   8
         ToolTipText     =   "Monto del concepto de vacaciones"
         Top             =   2655
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         MaxLength       =   10
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xDet1 
         Height          =   285
         Left            =   945
         TabIndex        =   7
         ToolTipText     =   "Escriba aqui el concepto de remuneración para vacaciones"
         Top             =   2655
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   503
         MaxLength       =   30
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xMontoManual 
         Height          =   285
         Left            =   2910
         TabIndex        =   6
         Top             =   1905
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   503
         Text            =   ""
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1365
         Left            =   945
         TabIndex        =   34
         Top             =   2985
         Width           =   4440
         _ExtentX        =   7832
         _ExtentY        =   2408
         _Version        =   393216
         AllowUpdate     =   -1  'True
         Enabled         =   0   'False
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
         Caption         =   "Detalles del Cálculo"
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label xTotalDetalle 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   285
         Left            =   5475
         TabIndex        =   37
         Top             =   4080
         Width           =   1200
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Total Vacaciones "
         Height          =   195
         Left            =   5475
         TabIndex        =   36
         Top             =   3795
         Width           =   1290
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Periodo en Cronograma"
         Height          =   195
         Left            =   135
         TabIndex        =   35
         Top             =   1545
         Width           =   1665
      End
      Begin VB.Label xCCosto 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3795
         TabIndex        =   29
         Top             =   1050
         Width           =   1575
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "C.Costo:"
         Height          =   195
         Left            =   3150
         TabIndex        =   28
         Top             =   1095
         Width           =   600
      End
      Begin VB.Label xCodigo 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-"
         Height          =   255
         Left            =   4500
         TabIndex        =   27
         Top             =   270
         Width           =   870
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Código Interno"
         Height          =   195
         Left            =   3330
         TabIndex        =   26
         Top             =   285
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Ingreso"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   1110
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Trabajador"
         Height          =   195
         Left            =   165
         TabIndex        =   3
         Top             =   330
         Width           =   765
      End
   End
End
Attribute VB_Name = "frCancelVacaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSCALC As ADODB.Recordset

Private Sub CMACEPTAR_CLICK()
    If xTrab.Tag = "" Then
        MsgBox "Falta especificar Trabajador", vbInformation
        xTrab.SetFocus
        Exit Sub
    End If
    If xPeriodo.Text = "" Then
        MsgBox "Falta especificar Periodo de Pago.", vbInformation
        xPeriodo.SetFocus
        Exit Sub
    End If
    If DateDiff("D", xFechaIniCal.Value, xFechaFinCal.Value) < 260 Then
        MsgBox "En términos de ley, para tener derecho al goce de vacaciones, los trabajadores deberán tener un record de 260 días de trabajo efectivo durante el año de servicios, en el caso de obreros haber percibido 40 dominicales", vbInformation
        If MsgBox("Desea continuar", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    End If
    If xForma(1).Value Then
        If Val(xDiasCompensados.Text) + Val(xDiasFisicos.Caption) <> 30 Then
            MsgBox "El total de dias de vacaciones entre los dias físicos gozados más los dias compensados deben de ser 30", vbInformation
            Exit Sub
        End If
    End If
    If Trim(xDescPeriodo.Text) = "" Then
        MsgBox "Deberá especificar una descripción del periodo de vacaciones", vbInformation
        xDescPeriodo.SetFocus
        Exit Sub
    End If
    Dim SGMONTO As Single
    If xTipo(0).Value Then
        If Val(xMontoManual.Text) <= 0 Then
            MsgBox "Falta ingresar un monto de remuneración vacacional. El valor no puede ser inferior o igual a cero", vbInformation
            Exit Sub
        End If
        SGMONTO = Val(xMontoManual.Text)
    Else
        If Val(xTotalDetalle.Caption) <= 0 Then
            MsgBox "Faltan ingresar detalles para la remuneración vacacional. El valor total de la remuneración no puede ser cero", vbInformation
            Exit Sub
        End If
        SGMONTO = Val(xTotalDetalle.Caption)
    End If
    If MsgBox("Desea aceptar los cambios en el registro de vacaciones", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    'UN POCO TONTO, PERO ES LA UNICA SALIDA QUE SE ME OCURRIO
    'LO IMPORTANTE ES QUE FUNCIONA
    If VPTAREA <> "NUEVO" Then
        DBSYSTEM.Execute "DELETE FROM HISTOVAC WHERE CODIGO=" & VPTAREA
        DBSYSTEM.Execute "DELETE FROM DETALLEVAC WHERE CODIGO=" & VPTAREA
        VPTAREA = "NUEVO"
    End If
    If VPTAREA = "NUEVO" Then
        DBSYSTEM.Execute "INSERT INTO HISTOVAC (CODTRAB, PERIODO, FECHAING, AREA, FECHAINI, FECHAFIN, DIAS, FECHAREG,CERRADO,MONTO,FECHAINICAL,FECHAFINCAL,NOMBOL,FORMADESCANSO,DIASCOMPENSADOS,MODOCALCULO,PROGRAMADO) VALUES ('" & xTrab.Tag & "','" & xDescPeriodo.Text & "'," & DateSQL(xFechaIng.Value) & ",''," & DateSQL(xFechaIni.Value) & "," & DateSQL(xFechaFin.Value) & ",30," & DateSQL(Date) & ",0," & SGMONTO & "," & DateSQL(xFechaIniCal.Value) & "," & DateSQL(xFechaFinCal.Value) & "," & xPeriodo.Tag & "," & IIf(xForma(0).Value, 0, 1) & "," & xDiasCompensados.Text & ",1," & IIf(xProg.Value = 1, 1, 0) & ")"
        'SI ES POR DETALLE HAY QUE GUARDARLOS PARA DESPUES IMPRIMIRLOS
        'PERSONALIZACIÓN TECSUR
        'OCTUBRE DEL 2000
        If xTipo(1).Value And RSCALC.RecordCount <> 0 Then
            SGMONTO = DevuelveValor("SELECT CODIGO FROM HISTOVAC WHERE AREA='NONE'", DBSYSTEM)
            RSCALC.MoveFirst
            Do While Not RSCALC.EOF
                DBSYSTEM.Execute "INSERT INTO DETALLEVAC (CODIGO,DESCRIPCION,IMPORTE) VALUES (" & SGMONTO & ",'" & RSCALC!DESCRIPCION & "'," & RSCALC!Importe & ")"
                RSCALC.MoveNext
            Loop
        End If
        DBSYSTEM.Execute "UPDATE HISTOVAC SET AREA=' ' WHERE AREA='NONE'"
        MsgBox "Informació Grabada Satisfactorimente", vbInformation
    End If
    Unload Me
End Sub

Private Sub CMADDDETALLE_CLICK()
    If Trim(xDet1.Text) = "" Then
        MsgBox "Falta registrar la Descripción del Pago", vbInformation
        xDet1.SetFocus
        Exit Sub
    End If
    If Val(xDet2.Text) = 0 Then
        MsgBox "El monto no puede ser igual o inferior a cero", vbInformation
        xDet2.SetFocus
        Exit Sub
    End If
    DBSYSTEM.Execute "INSERT INTO  [##TMPCANCELVAC" & VGL_COMPUTER & "]  (DESCRIPCION,IMPORTE) VALUES ('" & xDet1.Text & "'," & xDet2.Text & ")"
    SUMATMP
    REFRESCAR
    xDet1.Text = ""
    xDet2.Text = 0
    xDet1.SetFocus
End Sub

Private Sub CMCANCELAR_CLICK()
    Unload Me
End Sub

Private Sub CMQUITAR_CLICK()
    If RSCALC.EOF Or RSCALC.RecordCount = 0 Then
        MsgBox "No existe nada por eliminar", vbInformation
        Exit Sub
    End If
        If MsgBox("Confirma que desea quitar el registro seleccionado:" & Chr(13) & Chr(10) & RSCALC!DESCRIPCION, vbYesNo + vbQuestion) = vbNo Then Exit Sub
        DBSYSTEM.Execute "DELETE FROM  [##TMPCANCELVAC" & VGL_COMPUTER & "]  WHERE DESCRIPCION='" & RSCALC!DESCRIPCION & "'"
    REFRESCAR
    SUMATMP
End Sub

Private Sub Form_Load()
    If ExisteTablaAux(" [##TMPCANCELVAC" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPCANCELVAC" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "CREATE TABLE  [##TMPCANCELVAC" & VGL_COMPUTER & "]  (DESCRIPCION VARCHAR(30), IMPORTE  Numeric(20,2) )"
    Set RSCALC = New ADODB.Recordset
    RSCALC.Open " [##TMPCANCELVAC" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockBatchOptimistic
    REFRESCAR
    xFechaFinCal.Value = Date
    xFechaFinCal.Day = 1
    xFechaIniCal.Value = DateAdd("D", -360, xFechaFinCal.Value)
    xDescPeriodo.Text = Year(xFechaIniCal.Value) & " - " & Year(xFechaFinCal.Value)
    XFORMA_CLICK (0)
    If VPTAREA <> "NUEVO" Then
        CARGADATOS
    End If
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSCALC = Nothing
End Sub

Private Sub XDIASCOMPENSADOS_CHANGE()
    XFECHAINI_CHANGE
End Sub

Private Sub XDIASCOMPENSADOS_LOSTFOCUS()
    If Val(xDiasCompensados.Text) > 30 Then
        MsgBox "Los dias compensados no pueden exceder a 30, pues el descanzo vacacional no puede exceder a 30 dias", vbInformation
        xDiasCompensados.Text = "30"
    Else
        If Val(xDiasCompensados.Text) > 15 Then
            MsgBox "Las dias de vacaciones compensados, en términos de ley," & Chr(13) & Chr(10) & "no pueden ser mayores a 15 dias", vbInformation
        End If
    End If
End Sub

Private Sub XFECHAFIN_CHANGE()
    xDiasFisicos.Caption = DateDiff("D", xFechaIni.Value, xFechaFin.Value)
End Sub

Private Sub XFECHAFIN_KEYDOWN(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 13 Then SendKeys "{TAB}"
End Sub

Private Sub XFECHAFINCAL_KEYDOWN(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 13 Then SendKeys "{TAB}"
End Sub

Private Sub XFECHAINI_CHANGE()
    If xForma(1).Value Then
        xFechaFin.Value = DateAdd("D", 30 - Val(xDiasCompensados.Text), xFechaIni.Value)
    Else
        xFechaFin.Value = DateAdd("D", 30, xFechaIni.Value)
    End If
    xDiasFisicos.Caption = DateDiff("D", xFechaIni.Value, xFechaFin.Value)
End Sub

Private Sub XFECHAINI_KEYDOWN(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 13 Then SendKeys "{TAB}"
End Sub

Private Sub XFECHAINICAL_CHANGE()
    xDescPeriodo.Text = Year(xFechaIniCal.Value) & " - " & Year(xFechaFinCal.Value)
End Sub

Private Sub XFECHAINICAL_KEYDOWN(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 13 Then SendKeys "{TAB}"
End Sub

Private Sub XFORMA_CLICK(INDEX As Integer)
    If INDEX = 1 Then
        xDiasCompensados.Visible = True
    Else
        xDiasCompensados.Visible = False
    End If
    XFECHAINI_CHANGE
End Sub

Private Sub XFORMA_KEYDOWN(INDEX As Integer, KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 13 Then SendKeys "{TAB}"
End Sub

Private Sub XMONTOMANUAL_KEYDOWN(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 13 Then SendKeys "{TAB}"
End Sub

Private Sub XPERIODO_DBLCLICK()
    Dim RSMESES As New ADODB.Recordset
    RSMESES.Open "SELECT CODIGO, NOMBRE,FECHAINI,FECHAFIN FROM NOMBOL WHERE MES IN (SELECT MESACTIVO FROM MESESACT)", DBSYSTEM, adOpenStatic, adLockReadOnly
    If RSMESES.RecordCount = 0 Then
        MsgBox "No se han encontrado meses en actividad", vbCritical
        cmAceptar.Enabled = False
        Set RSMESES = Nothing
        Exit Sub
    End If
    frmComun.CONECTAR RSMESES
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xPeriodo.Text = RSMESES!NOMBRE
        xPeriodo.Tag = RSMESES!Codigo
    End If
    xFechaIni.Value = RSMESES!FECHAINI
    xFechaIni.MinDate = RSMESES!FECHAINI
    XPROG_CLICK
    Set RSMESES = Nothing
End Sub

Private Sub XPROG_CLICK()
    If xProg.Value = 1 Then
        xFechaIni.Enabled = False
    Else
        xFechaIni.Enabled = True
    End If
End Sub

Private Sub XTIPO_CLICK(INDEX As Integer)
    On Error Resume Next
    If INDEX = 0 Then
        xMontoManual.Visible = True
        xMontoManual.SetFocus
        DataGrid1.Enabled = False
    Else
        xMontoManual.Visible = False
        DataGrid1.Enabled = True
        xDet1.SetFocus
    End If
End Sub

Private Sub XTIPO_KEYDOWN(INDEX As Integer, KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 13 Then SendKeys "{TAB}"
End Sub

Private Sub XTRAB_DBLCLICK()
    Dim RSAUX As New ADODB.Recordset
    RSAUX.Open "SELECT CODTRAB, NOMBRES, FECHAING, CODCCOSTO,CENTRO FROM VWTRABAJ WHERE SITUACIÓN<'2' AND CODTRAB NOT IN (SELECT CODTRAB FROM HISTOVAC WHERE CERRADO=0)", DBSYSTEM, adOpenStatic, adLockReadOnly
    If RSAUX.RecordCount = 0 Or RSAUX.EOF Then
        MsgBox "No se han encontrado Trabajadores", vbInformation
        Set RSAUX = Nothing
        cmAceptar.Enabled = False
        Exit Sub
    End If
    frmComun.CONECTAR RSAUX
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xTrab.Text = RSAUX!NOMBRES
        xTrab.Tag = RSAUX!CODTRAB
        xFechaIng.Value = RSAUX!FECHAING
        xCCosto.Caption = RSAUX!CODCCOSTO
        xCCosto.ToolTipText = Trim$(RSAUX!CENTRO)
    End If
    Set RSAUX = Nothing
End Sub

Public Sub REFRESCAR()
    RSCALC.Requery
    Set DataGrid1.DataSource = RSCALC
    With DataGrid1
        .Columns("DESCRIPCION").Width = 2800
        .Columns("IMPORTE").NumberFormat = "0.00 "
        .Columns("IMPORTE").Alignment = dbgRight
    End With
End Sub

Public Sub SUMATMP()
    Dim RSAUX As New ADODB.Recordset
    RSAUX.Open "SELECT SUM(IMPORTE) AS TOTAL FROM  [##TMPCANCELVAC" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenStatic, adLockReadOnly
    If RSAUX.RecordCount = 0 Or RSAUX.EOF Then
        xTotalDetalle.Caption = "0.00 "
    Else
        xTotalDetalle.Caption = Format(IIf(IsNull(RSAUX!TOTAL), 0, RSAUX!TOTAL), "0.00 ")
    End If
    Set RSAUX = Nothing
End Sub

Public Sub CARGADATOS()
    Dim RSEDIT As New ADODB.Recordset, X As Integer
    RSEDIT.Open "SELECT * FROM HISTOVAC WHERE CODIGO=" & VPTAREA, DBSYSTEM, adOpenStatic, adLockReadOnly
    If RSEDIT.RecordCount = 0 Or RSEDIT.EOF Then
        MsgBox "No se han podido recuperar los datos, posiblemente han sido bloqueados por otro usuario", vbCritical
        Set RSEDIT = Nothing
        Exit Sub
    End If
    xCodigo.Caption = VPTAREA
    With RSEDIT
        xCodigo.Caption = VPTAREA
        xTrab.Text = DevuelveValor("SELECT NOMBRES FROM VWTRABAJ WHERE CODTRAB='" & !CODTRAB & "'", DBSYSTEM)
        xTrab.Tag = !CODTRAB
        If Not IsNull(!NOMBOL) Then
            xPeriodo.Text = "" & DevuelveValor("SELECT NOMBRE FROM NOMBOL WHERE CODIGO=" & !NOMBOL, DBSYSTEM)
            xPeriodo.Tag = !NOMBOL
        End If
        xFechaIng.Value = !FECHAING
        xCCosto.Caption = DevuelveValor("SELECT CCOSTO FROM TRABAJADORES WHERE CODTRAB='" & !CODTRAB & "'", DBSYSTEM)
        xFechaIni.Value = !FECHAINI
        xFechaFin.Value = !FECHAFIN
        xFechaIniCal.Value = !FECHAINICAL
        xFechaFinCal.Value = !FECHAFINCAL
        If !FORMADESCANSO = 0 Then
            xForma(0).Value = True
        Else
            xForma(1).Value = True
            xDiasCompensados.Text = IIf(IsNull(!DIASCOMPENSADOS), 0, !DIASCOMPENSADOS)
            XFORMA_CLICK (1)
        End If
        If !PROGRAMADO = 0 Then
            xProg.Value = 0
            XPROG_CLICK
        End If
        If IsNull(!NOMBOL) Then
            xForma(0).Value = True
            XFORMA_CLICK (0)
        End If
        xMontoManual.Text = !MONTO
        xDescPeriodo.Text = !PERIODO
        DBSYSTEM.Execute "INSERT INTO  [##TMPCANCELVAC" & VGL_COMPUTER & "]  SELECT DESCRIPCION,IMPORTE FROM DETALLEVAC IN '" & REGSISTEMA.PATHEMPRESA & "\PLANILLA.MDB' WHERE CODIGO=" & VPTAREA, X
        If X > 0 Then
            xTipo(1).Value = True
            XTIPO_CLICK (1)
            REFRESCAR
            SUMATMP
        End If
    End With
    Set RSEDIT = Nothing
End Sub

