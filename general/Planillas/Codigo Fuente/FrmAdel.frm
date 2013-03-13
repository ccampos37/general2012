VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FrmAdel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Adelantos Detallado"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8370
   Icon            =   "FrmAdel.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   8370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport Reporte 
      Left            =   1110
      Top             =   4695
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox xTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5580
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "FrmAdel.frx":08CA
      Top             =   6195
      Width           =   1350
   End
   Begin VB.CommandButton cmGrabar 
      Caption         =   "&Grabar"
      Height          =   405
      Left            =   7065
      TabIndex        =   6
      Top             =   6135
      Width           =   1230
   End
   Begin VB.CommandButton cmSelecTrab 
      Caption         =   "Seleccion (F5)"
      Height          =   990
      Left            =   7335
      Picture         =   "FrmAdel.frx":08D1
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1515
      Width           =   945
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Por Centros de Costo"
      Height          =   210
      Left            =   60
      TabIndex        =   3
      Top             =   945
      Width           =   1830
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Por Areas de Trabajo"
      Height          =   210
      Left            =   60
      TabIndex        =   2
      Top             =   690
      Value           =   -1  'True
      Width           =   1830
   End
   Begin VB.CommandButton cmdCuentasCtes 
      Caption         =   "Aplicar &Cuentas Corrientes"
      Height          =   345
      Left            =   90
      TabIndex        =   1
      Top             =   2190
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   405
      Left            =   2400
      TabIndex        =   0
      Top             =   6135
      Width           =   1230
   End
   Begin MSDataGridLib.DataGrid xData 
      Height          =   3480
      Left            =   60
      TabIndex        =   5
      Top             =   2565
      Width           =   8250
      _ExtentX        =   14552
      _ExtentY        =   6138
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
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin AplisetControlText.Aplitext xMes 
      Height          =   285
      Left            =   90
      TabIndex        =   8
      Top             =   255
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   503
      Locked          =   -1  'True
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker xFechaFin 
      Height          =   285
      Left            =   1335
      TabIndex        =   9
      Top             =   1665
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   24444929
      CurrentDate     =   36699
   End
   Begin MSComCtl2.DTPicker xFechaIni 
      Height          =   285
      Left            =   1335
      TabIndex        =   10
      Top             =   1320
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   24444929
      CurrentDate     =   36699
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   1725
      Left            =   2910
      TabIndex        =   11
      Top             =   255
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   3043
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Periodos en Cronogramas"
         Object.Width           =   5733
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "FechaIni"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "FechaFin"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Image xAuto 
      Height          =   240
      Left            =   5025
      Picture         =   "FrmAdel.frx":119B
      Top             =   2265
      Width           =   240
   End
   Begin VB.Label xlAuto 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Auto-rellenado"
      Height          =   270
      Left            =   4890
      TabIndex        =   18
      Top             =   2265
      Width           =   1485
   End
   Begin VB.Image Image4 
      Height          =   240
      Left            =   5175
      Picture         =   "FrmAdel.frx":14DD
      ToolTipText     =   "Suma total de los adelantos por aceptar"
      Top             =   6210
      Width           =   240
   End
   Begin VB.Label l2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Final"
      Height          =   195
      Left            =   90
      TabIndex        =   17
      Top             =   1725
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label l1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Inicial"
      Height          =   195
      Left            =   90
      TabIndex        =   16
      Top             =   1365
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Mes de Trabajo"
      Height          =   195
      Left            =   105
      TabIndex        =   15
      Top             =   15
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Periodos en Cronograma"
      Height          =   195
      Left            =   2910
      TabIndex        =   14
      Top             =   15
      Width           =   1740
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   7695
      Picture         =   "FrmAdel.frx":181F
      Top             =   60
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Adelantos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   7170
      TabIndex        =   13
      Top             =   555
      Width           =   1065
   End
   Begin VB.Label xNumTrab 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " 0 Trabajadores"
      Height          =   285
      Left            =   30
      TabIndex        =   12
      Top             =   6210
      Width           =   2205
   End
   Begin VB.Image i1 
      Height          =   240
      Left            =   2415
      Picture         =   "FrmAdel.frx":1B29
      ToolTipText     =   "Indica que se han cargo Cuentas Corrientes a descontar en este adelanto"
      Top             =   2235
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image i2 
      Height          =   240
      Left            =   2655
      Picture         =   "FrmAdel.frx":1E6B
      ToolTipText     =   "Indica que se han cargo Cuentas Corrientes a descontar en este adelanto"
      Top             =   2235
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Line Line1 
      X1              =   5550
      X2              =   5625
      Y1              =   2340
      Y2              =   2340
   End
End
Attribute VB_Name = "FrmAdel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSTMPADEL As New ADODB.Recordset
Dim RSMESES As New ADODB.Recordset
Dim RSADELANTO As New ADODB.Recordset
Dim REGACT As REGWIN, CADIN As String
Dim SQLSTR As String
Public Columna As Integer
Dim FLAG As Boolean
Dim FLAGHEAT As Boolean
Private Sub CMDCUENTASCTES_CLICK()
    frAdelMoviCta.Show 1
    If UCase(Dir$(App.PATH & "\ADELCC.DYB")) = "ADELCC.DYB" Then
        i1.Visible = True
        i2.Visible = True
    Else
        i1.Visible = False
        i2.Visible = False
    End If
End Sub

Private Sub CMGRABAR_CLICK()
    If RSTMPADEL.RecordCount = 0 Or Val(xTotal.Text) = 0 Then
        MsgBox "NO EXISTE NADA POR GRABAR", vbCritical
        Exit Sub
    End If
    If xMes.Tag = "" Then
        MsgBox "No ha seleccionado un mes, cambie por favor el mes", vbCritical
        xMes.SetFocus
        Exit Sub
    End If
    Dim RSTBADEL As New ADODB.Recordset
    DBSTARPLAN.Execute "DELETE FROM  [##_TMPADELANTO" & VGL_COMPUTER & "]  WHERE MONTO=0"
    Set RSTMPADEL = New ADODB.Recordset
    RSTMPADEL.Open "[##_TMPADELANTO" & VGL_COMPUTER & "]", DBAUXCOM, adOpenKeyset, adLockOptimistic
    FORMATEARDBG
    Dim CARGACC As Boolean
    Dim RSAUX As New ADODB.Recordset
    CARGACC = False
    If UCase(Dir$(App.PATH & "\ADELCC.DYB")) = "ADELCC.DYB" Then
        If MsgBox("Esta a punto de cargar los Debitos por Cuentas Corrientes de Trabajadores.. desea hacer efectivo los debitos especificados", vbYesNo + vbQuestion) = vbNo Then CARGACC = False Else CARGACC = True
    End If
    If MsgBox("Desea continuar", vbYesNo) = vbYes Then
        If CARGACC Then
            RSAUX.Open App.PATH & "\ADELCC.DYB", , adOpenStatic, adLockReadOnly, adCmdFile
            Do While Not RSAUX.EOF
                DBSYSTEM.Execute "INSERT INTO PAGOSCTA (CODMOV,NUMBOL,CODNOMBOL,TIPOBOLETA,MONTO,DOLAR,CODTRAB,TIPO,SECUENCIA) VALUES (" & RSAUX!CODMOV & ",0," & Lista.SelectedItem.Tag & ",'A'," & Round(RSAUX!DEBITO, 2) & ",0,'" & RSAUX!CODTRAB & "'," & IIf(RSAUX!Tip = "E", 2, 1) & "," & RSAUX!SECUENCIA & ")"
                DBSYSTEM.Execute "UPDATE MOVICTA SET SALDO=SALDO-" & RSAUX!DEBITO & " WHERE CODMOV=" & RSAUX!CODMOV
                RSAUX.MoveNext
            Loop
        End If
Dim RsAuxTotal As ADODB.Recordset
Dim TOTAL As Single
TOTAL = 0
    SQLSTR = "SELECT SUM(MONTO"
    Set Rs_aux_1 = New ADODB.Recordset
    Rs_aux_1.Open "SELECT * FROM CONFIADEL WHERE TIPO=1", DBSYSTEM, adOpenKeyset, adLockOptimistic
    If Rs_aux_1.RecordCount > 0 Then
        While Not Rs_aux_1.EOF
            SQLSTR = SQLSTR + "+[" & Rs_aux_1.Fields(1) & "]"
            Rs_aux_1.MoveNext
        Wend
    End If
    RESTA = "SUM(0"
    Set Rs_aux_1 = New ADODB.Recordset
    Rs_aux_1.Open "SELECT * FROM CONFIADEL WHERE TIPO=2", DBSYSTEM, adOpenKeyset, adLockOptimistic
    If Rs_aux_1.RecordCount > 0 Then
        While Not Rs_aux_1.EOF
            RESTA = RESTA + "+[" & Rs_aux_1.Fields(1) & "]"
            Rs_aux_1.MoveNext
        Wend
    End If
    SQLSTR = SQLSTR & ")-" & RESTA & ") AS TOTAL FROM  [##_TMPADELANTO" & VGL_COMPUTER & "] "
        Do While Not RSTMPADEL.EOF
            Set RsAuxTotal = New ADODB.Recordset
            RsAuxTotal.Open SQLSTR & " WHERE CODTRAB='" & RSTMPADEL!CODTRAB & "'", DBSYSTEM
            If RsAuxTotal.RecordCount > 0 Then
                TOTAL = RsAuxTotal.Fields(0)
            End If
            DBSYSTEM.Execute "INSERT INTO " & REGSISTEMA.TABLAADEL & " (CODTRAB,MES,FECHAING,MONTO,NUMBOL,NOMBOL, ORIGEN) VALUES ('" & RSTMPADEL!CODTRAB & "'," & DateSQL(xMes.Tag) & "," & DateSQL(Date) & "," & TOTAL & ",0,0," & Lista.SelectedItem.Tag & ")"
            RSTMPADEL.MoveNext
        Loop
Dim Codigo As String
Dim Empleado  As String
        'DBSYSTEM.Execute "DELETE FROM DETADEL"
        Dim RSCON As New ADODB.Recordset
        Set RSAUX = New ADODB.Recordset
        RSAUX.Open "SELECT * FROM DETADEL", DBSYSTEM, adOpenDynamic, adLockOptimistic
        'GRABA EN LA TABLA DETALLE DE ADELANTOS
                RSTMPADEL.MoveFirst
                Codigo = RSTMPADEL!CODTRAB
                Empleado = RSTMPADEL!NOMBRES
                While Not RSTMPADEL.EOF
                    For X = 5 To xData.Columns.Count - 1
                        If X Mod 2 <> 0 Then
                            If Codigo <> RSTMPADEL!CODTRAB Then
                                    RSAUX.AddNew
                                    RSAUX!NOMBOL = Lista.SelectedItem.Tag
                                    RSAUX!CODTRAB = Codigo
                                    RSAUX!NOMBRE = Empleado
                                    RSAUX!MES = xFechaIni.Value
                                    RSAUX!CODCONCEP = "XXX"
                                    RSAUX!CONCEPTO = "PRESTAMO"
                                    RSAUX!MONTO = DevuelveValor("SELECT MONTO FROM PAGOSCTA WHERE CODTRAB='" & Codigo & "' AND CODNOMBOL=" & Lista.SelectedItem.Tag, DBSYSTEM)
                                    RSAUX!IE = "2"
                                    SQLSTR = "SELECT SUM(MONTO"
                                    Set Rs_aux_1 = New ADODB.Recordset
                                    Rs_aux_1.Open "SELECT * FROM CONFIADEL WHERE TIPO=1", DBSYSTEM, adOpenKeyset, adLockOptimistic
                                    If Rs_aux_1.RecordCount > 0 Then
                                        While Not Rs_aux_1.EOF
                                            SQLSTR = SQLSTR + "+[" & Rs_aux_1.Fields(1) & "]"
                                            Rs_aux_1.MoveNext
                                        Wend
                                    End If
                                    RESTA = "SUM(0"
                                    Set Rs_aux_1 = New ADODB.Recordset
                                    Rs_aux_1.Open "SELECT * FROM CONFIADEL WHERE TIPO=2", DBSYSTEM, adOpenKeyset, adLockOptimistic
                                    If Rs_aux_1.RecordCount > 0 Then
                                        While Not Rs_aux_1.EOF
                                            RESTA = RESTA + "+[" & Rs_aux_1.Fields(1) & "]"
                                            Rs_aux_1.MoveNext
                                        Wend
                                    End If
                                    SQLSTR = SQLSTR & ")-" & RESTA & "+" & RSAUX!MONTO & ") AS TOTAL FROM  [##_TMPADELANTO" & VGL_COMPUTER & "] "
                                    Set RsAuxTotal = New ADODB.Recordset
                                    RsAuxTotal.Open SQLSTR & " WHERE CODTRAB='" & RSTMPADEL!CODTRAB & "'", DBSYSTEM
                                    TOTALGENERAL = DevuelveValor(SQLSTR, DBSYSTEM)
                                    If RsAuxTotal.RecordCount > 0 Then
                                        TOTAL = RsAuxTotal.Fields(0)
                                    End If
                                    RSAUX!TOTAL = TOTAL
                                    RSAUX.Update
                                    DBSYSTEM.Execute "UPDATE DETADEL SET TOTAL=" & TOTAL & " WHERE CODTRAB='" & Codigo & "'"
                            End If
                            RSAUX.AddNew
                            RSAUX!NOMBOL = Lista.SelectedItem.Tag
                            RSAUX!CODCONCEP = DevuelveValor("SELECT CODIGO FROM CONFIADEL WHERE NOMBRE='" & RSTMPADEL.Fields(X).Name & "'", DBSYSTEM)
                            RSAUX!CODTRAB = Trim(RSTMPADEL!CODTRAB)
                            RSAUX!NOMBRE = Trim(RSTMPADEL!NOMBRES)
                            RSAUX!MES = xFechaIni.Value
                            RSAUX!CONCEPTO = RSTMPADEL.Fields(X).Name
                            RSAUX!MONTO = RSTMPADEL.Fields(X)
                            
                            'EN ESTA SECCION SE GRABA LOS CONCEPTOS DE ADELANTOS DETALLADOS
                            'EN LA TABLA INGRESO DE MOVIMIENTOS
                            If RSAUX!CODCONCEP <> "" Then
                                DBSYSTEM.Execute "INSERT INTO INGMOV2000(CODTRAB,CONCEPTO,VALOR,CODNOMBOL) VALUES " & _
                                "('" & RSTMPADEL!CODTRAB & "','" & RSAUX!CODCONCEP & "'," & RSTMPADEL.Fields(X - 1) & "," & Lista.SelectedItem.Tag & ")"
                            End If
                            
                            Set RSCON = New ADODB.Recordset
                            RSCON.Open "Select TIPO FROM CONFIADEL WHERE NOMBRE='" & RSTMPADEL.Fields(X).Name & "'", DBSYSTEM, adOpenKeyset, adLockOptimistic
                            If RSCON.RecordCount > 0 Then
                                RSAUX!IE = Trim(RSCON.Fields(0))
                            Else
                                RSAUX!IE = "1"
                            End If
                            SQLSTR = "SELECT SUM(MONTO"
                            Set Rs_aux_1 = New ADODB.Recordset
                            Rs_aux_1.Open "SELECT * FROM CONFIADEL WHERE TIPO=1", DBSYSTEM, adOpenKeyset, adLockOptimistic
                            If Rs_aux_1.RecordCount > 0 Then
                                While Not Rs_aux_1.EOF
                                    SQLSTR = SQLSTR + "+[" & Rs_aux_1.Fields(1) & "]"
                                    Rs_aux_1.MoveNext
                                Wend
                            End If
                            RESTA = "SUM(0"
                            Set Rs_aux_1 = New ADODB.Recordset
                            Rs_aux_1.Open "SELECT * FROM CONFIADEL WHERE TIPO=2", DBSYSTEM, adOpenKeyset, adLockOptimistic
                            If Rs_aux_1.RecordCount > 0 Then
                                While Not Rs_aux_1.EOF
                                    RESTA = RESTA + "+[" & Rs_aux_1.Fields(1) & "]"
                                    Rs_aux_1.MoveNext
                                Wend
                            End If
                            SQLSTR = SQLSTR & ")-" & RESTA & ") AS TOTAL FROM  [##_TMPADELANTO" & VGL_COMPUTER & "] "
                            Set RsAuxTotal = New ADODB.Recordset
                            RsAuxTotal.Open SQLSTR & " WHERE CODTRAB='" & RSTMPADEL!CODTRAB & "'", DBSYSTEM
                            If RsAuxTotal.RecordCount > 0 Then
                                TOTAL = RsAuxTotal.Fields(0)
                            End If
                            RSAUX!TOTAL = TOTAL
                            RSAUX.Update
                            Codigo = RSAUX!CODTRAB 'FC
                        End If
                    Next
                    RSTMPADEL.MoveNext
                Wend
                    'AGREGO LA CTA. CTE.
                RSAUX.AddNew
                RSAUX!NOMBOL = Lista.SelectedItem.Tag
                RSAUX!CODTRAB = Codigo
                RSAUX!NOMBRE = Empleado
                RSAUX!MES = xFechaIni.Value
                RSAUX!CODCONCEP = "XXX"
                RSAUX!CONCEPTO = "PRESTAMO"
                RSAUX!MONTO = DevuelveValor("SELECT MONTO FROM PAGOSCTA WHERE CODTRAB='" & Codigo & "' AND CODNOMBOL=" & Lista.SelectedItem.Tag, DBSYSTEM)
                RSAUX!IE = "2"
                SQLSTR = "SELECT SUM(MONTO"
                Set Rs_aux_1 = New ADODB.Recordset
                Rs_aux_1.Open "SELECT * FROM CONFIADEL WHERE TIPO=1", DBSYSTEM, adOpenKeyset, adLockOptimistic
                If Rs_aux_1.RecordCount > 0 Then
                    While Not Rs_aux_1.EOF
                        SQLSTR = SQLSTR + "+[" & Rs_aux_1.Fields(1) & "]"
                        Rs_aux_1.MoveNext
                    Wend
                End If
                RESTA = "SUM(0"
                Set Rs_aux_1 = New ADODB.Recordset
                Rs_aux_1.Open "SELECT * FROM CONFIADEL WHERE TIPO=2", DBSYSTEM, adOpenKeyset, adLockOptimistic
                If Rs_aux_1.RecordCount > 0 Then
                    While Not Rs_aux_1.EOF
                        RESTA = RESTA + "+[" & Rs_aux_1.Fields(1) & "]"
                        Rs_aux_1.MoveNext
                    Wend
                End If
                SQLSTR = SQLSTR & ")-" & RESTA & "+" & RSAUX!MONTO & ") AS TOTAL FROM  [##_TMPADELANTO" & VGL_COMPUTER & "] "
                Set RsAuxTotal = New ADODB.Recordset
                RsAuxTotal.Open SQLSTR & " WHERE CODTRAB='" & Codigo & "'", DBSYSTEM
                TOTALGENERAL = DevuelveValor(SQLSTR, DBSYSTEM)
                If RsAuxTotal.RecordCount > 0 Then
                    TOTAL = RsAuxTotal.Fields(0)
                End If
                RSAUX!TOTAL = TOTAL
                RSAUX.Update
                DBSYSTEM.Execute "UPDATE DETADEL SET TOTAL=" & TOTALGENERAL & " WHERE CODTRAB='" & Codigo & "'"
                                               
    End If
    If MsgBox("Se han grabado los datos satisfactoriamente. Desea Imprimir las Boletas?", vbYesNo, "Confirmar") = vbYes Then
        With Reporte
            .Reset
            .ReportFileName = REGSISTEMA.REPORTES & "REPORT1.RPT"
            .Connect = "DSN=" & VGL_SERVER & ";USER=SOPORTE;PWD=SOPORTE;DSQ=" & REGSISTEMA.BASESQL & ""
            .Destination = crptToWindow
            .WindowShowPrintBtn = True
            .WindowShowSearchBtn = True
            .WindowShowPrintSetupBtn = True
            .WindowTitle = "REPORT1 - RECIBO DE ADELANTO DE QUINCENA"
            .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
            If .Status <> 2 Then .Action = 1
        End With
        
    End If
    Set RSTBADEL = Nothing
    Set RSAUX = Nothing
    cmGrabar.Enabled = False
End Sub

Private Sub CMSELECTRAB_CLICK()
    If Not xFechaIni.Visible Then
        MsgBox "Deberá seleccionar un periodo de pago", vbCritical
        Exit Sub
    End If
    CADIN = ""
    If DevuelveValor("SELECT USARCRONOGRAMA FROM EMPRESA", DBSYSTEM) = 1 Then
        Dim RSDELS As New ADODB.Recordset
        If Option1.Value Then
            RSDELS.Open "SELECT DISTINCT CODREF FROM FECHAPAGO, NOMBOL WHERE TIPOAC=0 AND DARADELANTO=1 AND CODNOMBOL=" & Lista.SelectedItem.Tag, DBSYSTEM, adOpenStatic
        Else
            RSDELS.Open "SELECT DISTINCT CODREF FROM FECHAPAGO, NOMBOL WHERE TIPOAC=1 AND DARADELANTO=1 AND CODNOMBOL=" & Lista.SelectedItem.Tag, DBSYSTEM, adOpenStatic
        End If
        If RSDELS.RecordCount = 0 Then
            MsgBox "No se han encontrado Areas o Centros de Costos que esten programados para pagos de Adelantos de Remuneraciones en el PERIODO seleccionado", vbCritical
            Set RSDELS = Nothing
            Exit Sub
        End If
        CADIN = ""
        Do While Not RSDELS.EOF
            If CADIN = "" Then CADIN = "'" & RSDELS!CODREF & "'" Else CADIN = CADIN & ",'" & RSDELS!CODREF & "'"
            RSDELS.MoveNext
        Loop
        CADIN = "(" & CADIN & ")"
        Set RSDELS = Nothing
    End If
    REGSELECT.USARFECHACESE = True
    REGSELECT.FECHACESEMAX = xFechaFin.Value
    REGSELECT.FECHAINIMAX = xFechaFin.Value
    REGSELECT.FECHAINI = xFechaIni.Value
    frSelect.Show 1
    REGSELECT.USARFECHACESE = False
    FLAG = False
    XADDTRAB_CLICK
    FLAG = True
End Sub

Private Sub COMMAND1_CLICK()
    Unload Me
End Sub

Private Sub Command2_Click()
     Screen.MousePointer = 11
     If FLAG Then frAutoAd2.SumaAdel
     Screen.MousePointer = 1
End Sub

Private Sub FORM_KEYDOWN(KEYCODE As Integer, SHIFT As Integer)
If KEYCODE = vbKeyF5 Then CMSELECTRAB_CLICK
If KEYCODE = 84 And SHIFT = 4 Then XAUTO_CLICK
End Sub

Private Sub Form_Load()
FLAGHEAT = True
Dim Rs_aux_1 As ADODB.Recordset
    'BORRAMOS EL TEMPORAL DE CUENTAS CORRIENTES
    If UCase(Dir$(App.PATH & "\ADELCC.DYB")) = "ADELCC.DYB" Then Kill App.PATH & "\ADELCC.DYB"
    Me.Tag = "PANEL DE ADELANTOS DE PAGO"
    RSADELANTO.Open "SELECT CODTRAB,SUM(MONTO) AS SALDO FROM " & REGSISTEMA.TABLAADEL & " WHERE NOMBOL=0 GROUP BY CODTRAB ORDER BY CODTRAB", DBSYSTEM, adOpenKeyset, adLockOptimistic
    If ExisteTablaAux(" [##_TMPADELANTO" & VGL_COMPUTER & "] ") Then
        DBSTARPLAN.Execute "DROP TABLE  [##_TMPADELANTO" & VGL_COMPUTER & "] "
    End If
    SQLSTR = "CREATE TABLE  [##_TMPADELANTO" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(6),NOMBRES VARCHAR(60),FECHAING DATETIME,PENDIENTE MONEY NULL DEFAULT 0, BASICO MONEY,MONTO MONEY NULL DEFAULT 0"
    
    Set Rs_aux_1 = New ADODB.Recordset
    Rs_aux_1.Open "CONFIADEL", DBSYSTEM, adOpenKeyset, adLockOptimistic
    If Rs_aux_1.RecordCount > 0 Then
        While Not Rs_aux_1.EOF
            SQLSTR = SQLSTR + ", [M " & Rs_aux_1.Fields(0) & "] MONEY NULL DEFAULT 0, [" & Rs_aux_1.Fields(1) & "] MONEY NULL DEFAULT 0"
            Rs_aux_1.MoveNext
        Wend
    End If
    SQLSTR = SQLSTR & ",[TOTAL ADELANTO] MONEY NULL DEFAULT 0)"
    DBSTARPLAN.Execute SQLSTR
    RSTMPADEL.Open " [##_TMPADELANTO" & VGL_COMPUTER & "] ", DBSTARPLAN, adOpenKeyset, adLockOptimistic
    Set xData.DataSource = RSTMPADEL
    FORMATEARDBG
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSTMPADEL = Nothing
    Set RSMESES = Nothing
    Set RSADELANTO = Nothing
    If UCase(Dir$(App.PATH & "\ADELCC.DYB")) = "ADELCC.DYB" Then Kill App.PATH & "\ADELCC.DYB"
End Sub

Private Sub I1_CLICK()
    MsgBox "Indica que se han cargado las Cuentas Corrientes a descontar en este adelanto", vbInformation
End Sub

Private Sub I2_CLICK()
    MsgBox "Indica que se han cargado las Cuentas Corrientes a descontar en este adelanto", vbInformation
End Sub

Private Sub IMAGE1_CLICK()
                                                                                                                                                                            MsgBox "danielyafac@hotmail.com", vbInformation
End Sub

Private Sub Image4_Click()
    Dim RS As New ADODB.Recordset
    
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
    xFechaIni.Visible = True
    xFechaFin.Visible = True
    l1.Visible = True
    l2.Visible = True
    xFechaIni.Value = CDate(Item.SubItems(1))
    xFechaFin.Value = CDate(Item.SubItems(2))
    REGINPUT.FECHAFIN = xFechaFin.Value
End Sub

Private Sub XADDTRAB_CLICK()
    On Error GoTo ERRSALIR
    DBSTARPLAN.Execute "DELETE FROM  [##_TMPADELANTO" & VGL_COMPUTER & "] "
    If UCase(Dir$(App.PATH & "\ADELCC.DYB")) = "ADELCC.DYB" Then Kill App.PATH & "\ADELCC.DYB"
    Dim RSTRAB As New ADODB.Recordset
    Dim STRCAD As String
    'ELIMINAMOS AQUELLOS QUE YA HAYAN TENIDO ADELANTOS EN ESTE PERIODO
    DBSTARPLAN.Execute "DELETE FROM  [##_TMPSELECT" & VGL_COMPUTER & "]  WHERE CODTRAB IN (SELECT CODTRAB FROM " & REGSISTEMA.BASESQL & ".dbo.ADEL" & REGSISTEMA.ANNO & " WHERE ORIGEN=" & Lista.SelectedItem.Tag & ")"
    If DevuelveValor("SELECT USARCRONOGRAMA FROM EMPRESA", DBSYSTEM) = 1 Then
        If Option1.Value Then
            RSTRAB.Open "SELECT * FROM  [##_TMPSELECT" & VGL_COMPUTER & "]  WHERE AREA IN " & CADIN & " ORDER BY NOMBRES", DBSTARPLAN, adOpenStatic
        Else
            RSTRAB.Open "SELECT * FROM  [##_TMPSELECT" & VGL_COMPUTER & "]  WHERE CENTROCOSTO IN " & CADIN & " ORDER BY NOMBRES", DBSTARPLAN, adOpenStatic
        End If
    Else
        RSTRAB.Open "SELECT * FROM  [##_TMPSELECT" & VGL_COMPUTER & "]  ORDER BY NOMBRES", DBSTARPLAN, adOpenStatic
    End If
    If RSTRAB.RecordCount = 0 Then
        MsgBox "Los Seleccionados ya tienen provisionados adelantos para este PERIODO " & Chr(13) & "Si desea eliminarlos ir al panel administración de adelantos", vbExclamation
        Set RSTRAB = Nothing
        cmdCuentasCtes.Enabled = False
        Exit Sub
    Else
        cmdCuentasCtes.Enabled = True
    End If
    Do While Not RSTRAB.EOF
        With RSTMPADEL
            .AddNew
            !CODTRAB = RSTRAB!CODTRAB
            !NOMBRES = RSTRAB!NOMBRES
            !BASICO = RSTRAB!BASICO
            !MONTO = 0
            !FECHAING = RSTRAB!FECHAING
            If RSADELANTO.RecordCount > 0 Then
                RSADELANTO.MoveFirst
                RSADELANTO.FIND "CODTRAB='" & RSTRAB!CODTRAB & "'"
                If Not RSADELANTO.EOF Then
                    RSTMPADEL!PENDIENTE = RSADELANTO.Fields("SALDO")
                End If
            End If
            .Update
        End With
        RSTRAB.MoveNext
    Loop
    RSTMPADEL.Requery
    FORMATEARDBG
    Set RSTRAB = Nothing
ERRSALIR: Exit Sub
End Sub

Private Sub XAUTO_CLICK()
    FLAGHEAT = False
    If RSTMPADEL.RecordCount = 0 Then
        MsgBox "No existen registros de trabajadores", vbCritical
        Exit Sub
    End If
    frAutoAd2.MES_X = Left(Me.xFechaIni.Value, 2)
    frAutoAd2.ANNO_X = Right(Me.xFechaIni.Value, 4)
    VAR = 2
    frAutoAd2.Show 1
    FLAGHEAT = True
End Sub


Private Sub XDATA_AFTERCOLUPDATE(ByVal ColIndex As Integer)
    RSTMPADEL.MOVE 0
End Sub

Private Sub XDATA_AFTERUPDATE()
    Dim RSSUMA As New ADODB.Recordset
    Dim SQLSTR As String
    Dim RESTA As String
    Call SumaAdel2
    SQLSTR = "SELECT SUM(MONTO"
    Set Rs_aux_1 = New ADODB.Recordset
    Rs_aux_1.Open "SELECT * FROM CONFIADEL WHERE TIPO=1", DBSYSTEM, adOpenKeyset, adLockOptimistic
    If Rs_aux_1.RecordCount > 0 Then
        While Not Rs_aux_1.EOF
            SQLSTR = SQLSTR + "+[" & Rs_aux_1.Fields(1) & "]"
            Rs_aux_1.MoveNext
        Wend
    End If
    RESTA = "SUM(0"
    Set Rs_aux_1 = New ADODB.Recordset
    Rs_aux_1.Open "SELECT * FROM CONFIADEL WHERE TIPO=2", DBSYSTEM, adOpenKeyset, adLockOptimistic
    If Rs_aux_1.RecordCount > 0 Then
        While Not Rs_aux_1.EOF
            RESTA = RESTA + "+[" & Rs_aux_1.Fields(1) & "]"
            Rs_aux_1.MoveNext
        Wend
    End If
    SQLSTR = SQLSTR & ")-" & RESTA & ") AS TOTAL FROM  [##_TMPADELANTO" & VGL_COMPUTER & "] "
    RSSUMA.Open SQLSTR, DBSYSTEM, adOpenStatic
    xTotal.Text = IIf(IsNull(RSSUMA!TOTAL), "0.00", Format(RSSUMA!TOTAL, "0.00"))
    Set RSSUMA = Nothing
    
End Sub

Private Sub XDATA_HEADCLICK(ByVal ColIndex As Integer)
On Error Resume Next
    If ColIndex < 5 Then Exit Sub
    If xData.Columns(ColIndex).Caption = "TOTAL ADELANTO" Then Exit Sub
    If xData.Columns(ColIndex).Caption = "MONTO" Or Left(xData.Columns(ColIndex).Caption, 1) <> "M" Then
        Columna = ColIndex
        XAUTO_CLICK
    Else
        RSTMPADEL.Sort = xData.Columns(ColIndex).DataField
    End If
    RSTMPADEL.Requery
End Sub

Private Sub XLAUTO_CLICK()
    XAUTO_CLICK
End Sub

Public Sub FORMATEARDBG()
    xData.Columns("FECHAING").Visible = False
    xData.Columns("MONTO").NumberFormat = "0.00 "
    xData.Columns("BASICO").NumberFormat = "0.00 "
    xData.Columns("PENDIENTE").NumberFormat = "0.00 "
    xData.Columns("CODTRAB").Locked = True
    xData.Columns("CODTRAB").AllowSizing = True
    xData.Columns("NOMBRES").Locked = True
    xData.Columns("BASICO").Locked = True
    xData.Columns("FECHAING").Locked = True
    xData.Columns("MONTO").Alignment = dbgRight
    xData.Columns("BASICO").Alignment = dbgRight
    xData.Columns("MONTO").AllowSizing = False
    xData.Columns("PENDIENTE").Locked = True
    xData.Columns("PENDIENTE").Alignment = dbgRight
    xData.Columns("NOMBRES").Width = 2399.811
    xData.Columns("MONTO").Width = 1000
    xData.Columns("BASICO").Width = 1000
    xData.Columns("PENDIENTE").Width = 1000
    For X = 6 To xData.Columns.Count - 1
            xData.Columns(X).NumberFormat = "0.00 "
            xData.Columns(X).Alignment = dbgRight
            xData.Columns(X).Width = 1000
            xData.Columns(X).AllowSizing = False
    Next
    xData.Columns("TOTAL ADELANTO").Locked = True
    xNumTrab.Caption = " " & RSTMPADEL.RecordCount & " TRABAJADORES"
End Sub

Public Function GETDATA() As ADODB.Recordset
    Set GETDATA = RSTMPADEL.Clone(adLockOptimistic)
End Function

Private Sub XMES_DBLCLICK()
    Lista.ListItems.Clear
    Dim RSMESES As New ADODB.Recordset
    RSMESES.Open "SELECT MESACTIVO, NOMBRE FROM MESESACT ORDER BY MESACTIVO", DBSYSTEM, adOpenStatic
    If RSMESES.RecordCount = 0 Then
        MsgBox "No se han encontrado meses en actividad", vbCritical
        Set RSMESES = Nothing
        Exit Sub
    End If
    frmComun.CONECTAR RSMESES
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xMes.Text = RSMESES!NOMBRE
        xMes.Tag = RSMESES!MESACTIVO
    Else
        Set RSMESES = Nothing
        Exit Sub
    End If
    Set RSMESES = Nothing
    'RECICLAJE DE RSMESES
    RSMESES.Open "SELECT CODIGO, NOMBRE, FECHAINI, FECHAFIN FROM NOMBOL WHERE CERRADO<>1 AND MES=" & DateSQL(CDate(xMes.Tag)) & " ORDER BY FECHAINI", DBSYSTEM, adOpenStatic
    Do While Not RSMESES.EOF
        Set XITEM = Lista.ListItems.Add(, , RSMESES!NOMBRE, , 1)
        XITEM.SubItems(1) = RSMESES!FECHAINI
        XITEM.SubItems(2) = RSMESES!FECHAFIN
        XITEM.Tag = RSMESES!Codigo
        RSMESES.MoveNext
    Loop
    l1.Visible = False
    l2.Visible = False
    xFechaIni.Visible = False
    xFechaFin.Visible = False
    Set RSMESES = Nothing
End Sub
Public Sub SumaAdel2()
    'SUMAR LOS TOTALES DE ADELANTO
On Error GoTo ERRSALIR
Dim RSADEL2 As ADODB.Recordset
Dim I As Integer
Dim ACUM As Double
Dim ULTPOS As Variant
Dim CONCEP As Double
Dim TIPO As String
    If Not FLAG Then Exit Sub
    If Not FLAGHEAT Then Exit Sub
    Set RSADEL2 = New ADODB.Recordset
    RSADEL2.Open "SELECT * FROM [##_TMPADELANTO" & VGL_COMPUTER & "] WHERE CODTRAB='" & RSTMPADEL!CODTRAB & "'", DBAUXCOM, adOpenKeyset, adLockOptimistic
    If RSADEL2.RecordCount = 0 Then Exit Sub
        With RSADEL2
            ACUM = 0
            For I = 0 To RSADEL2.Fields.Count - 1
                If I < 5 Then GoTo FINAL
                If Not (.Fields(I).Name = "TOTAL ADELANTO" Or Left(.Fields(I).Name, 2) = "M ") Then
                    CONCEP = ESNULO(.Fields(I).Value, 0)
                    If .Fields(I).Name <> "MONTO" Then
                        TIPO = DevuelveValor("SELECT TIPO FROM CONFIADEL WHERE NOMBRE='" & .Fields(I).Name & "'", DBSYSTEM)
                        Select Case TIPO
                            Case 2: CONCEP = CONCEP * -1
                        End Select
                    End If
                    ACUM = ACUM + CONCEP
                End If
FINAL:
            Next
            xData.Columns("TOTAL ADELANTO") = ACUM
        End With
    Set RSADEL2 = Nothing
ERRSALIR: Exit Sub
End Sub

