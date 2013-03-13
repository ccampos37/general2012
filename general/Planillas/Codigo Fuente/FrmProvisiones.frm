VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Begin VB.Form FrmProvisiones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proceso de Provisiones"
   ClientHeight    =   5616
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7548
   Icon            =   "FrmProvisiones.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5616
   ScaleWidth      =   7548
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   390
      Left            =   6015
      TabIndex        =   8
      Top             =   5190
      Width           =   1440
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Procesar"
      Height          =   390
      Left            =   6030
      TabIndex        =   7
      Top             =   4740
      Width           =   1440
   End
   Begin VB.Frame Frame1 
      Height          =   4635
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   7410
      Begin VB.CommandButton cmSelecTrab 
         Caption         =   "Seleccion (F5)"
         Height          =   990
         Left            =   6450
         Picture         =   "FrmProvisiones.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Seleccion de Trabajadores"
         Top             =   210
         Width           =   870
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   345
         Left            =   2430
         TabIndex        =   11
         Top             =   3030
         Width           =   2025
         _ExtentX        =   3577
         _ExtentY        =   614
         _Version        =   393216
         CustomFormat    =   "MMMM -    yyyy"
         Format          =   61603843
         CurrentDate     =   36983
      End
      Begin VB.Frame Frame2 
         Height          =   525
         Left            =   2445
         TabIndex        =   2
         Top             =   3975
         Width           =   4275
         Begin VB.OptionButton Option1 
            Caption         =   "CTS"
            Height          =   300
            Index           =   2
            Left            =   3165
            TabIndex        =   5
            Top             =   150
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Gratificaciones"
            Height          =   300
            Index           =   1
            Left            =   1485
            TabIndex        =   4
            Top             =   150
            Width           =   1485
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Vacaciones"
            Height          =   270
            Index           =   0
            Left            =   135
            TabIndex        =   3
            Top             =   165
            Width           =   1260
         End
      End
      Begin AplisetControlText.Aplitext xDivisor 
         Height          =   285
         Left            =   2415
         TabIndex        =   14
         Top             =   3600
         Width           =   795
         _ExtentX        =   1397
         _ExtentY        =   508
         MaxLength       =   2
         Text            =   ""
         Redondear       =   -1  'True
         TipoDato        =   "N"
      End
      Begin MSDataGridLib.DataGrid DGLista 
         Height          =   2745
         Left            =   120
         TabIndex        =   16
         ToolTipText     =   "Trabajadores seleccionados para el proceso de planillas"
         Top             =   150
         Width           =   6225
         _ExtentX        =   10986
         _ExtentY        =   4847
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Trabajadores Seleccionados"
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Image Image2 
         Height          =   555
         Left            =   6120
         Picture         =   "FrmProvisiones.frx":0D0C
         Stretch         =   -1  'True
         Top             =   3210
         Width           =   660
      End
      Begin VB.Image Image1 
         Height          =   420
         Left            =   6435
         Picture         =   "FrmProvisiones.frx":114E
         Stretch         =   -1  'True
         Top             =   2970
         Width           =   435
      End
      Begin VB.Label Label3 
         Caption         =   "Divisor"
         Height          =   345
         Index           =   1
         Left            =   450
         TabIndex        =   15
         Top             =   3615
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Mes"
         Height          =   345
         Index           =   0
         Left            =   435
         TabIndex        =   6
         Top             =   3045
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo de Provisones"
         Height          =   315
         Left            =   420
         TabIndex        =   1
         Top             =   4125
         Width           =   1665
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   180
      Left            =   105
      TabIndex        =   9
      Top             =   4935
      Width           =   5790
      _ExtentX        =   10224
      _ExtentY        =   318
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   180
      Left            =   90
      TabIndex        =   12
      Top             =   5370
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   318
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label5 
      Caption         =   "Calculando"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   105
      TabIndex        =   13
      Top             =   5130
      Width           =   4065
   End
   Begin VB.Label Label4 
      Caption         =   "Procesando :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   10
      Top             =   4695
      Width           =   4065
   End
End
Attribute VB_Name = "FrmProvisiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As ADODB.Recordset
Dim AREA As String
Dim AREACOD As String
Dim RSFORMULAS As ADODB.Recordset
Dim OP As Integer
Dim RSTRAB As New ADODB.Recordset
Private Sub CMSELECTRAB_CLICK()
    REGSELECT.FECHACESEMAX = DTPicker1.Value
    REGSELECT.FECHAINIMAX = DTPicker1.Value
    REGSELECT.FECHAINI = DTPicker1.Value
    REGSELECT.SITUACIONES = "'0','1'"
    REGSELECT.USARFECHACESE = True
    frSelect.Show 1
    REGSELECT.USARFECHACESE = False
    RSTRAB.Requery
    Set DGLista.DataSource = RSTRAB
End Sub
Private Sub Command1_Click()
Dim FECHA As String
Dim TABLA As String
Dim V_GENERAL As Boolean
Dim X, Y As Integer
    'VERIFICA QUE OPCION SE ENCUENTRA
    Select Case OP
        Case 0
            TABLA = "SUMAVAC"
        Case 1
            TABLA = "SUMAGRAT"
        Case 2
            TABLA = "SUMACTS"
    End Select
    Y = 0
    FECHA = Format(DTPicker1.Value, "MMYYYY")
    Set RSFORMULAS = New ADODB.Recordset
    DBSYSTEM.Execute "UPDATE CONCEPTOS SET CODIGO=CODIGO WHERE " & TABLA & " = 1", Y
    Set RSFORMULAS = New ADODB.Recordset
    RSFORMULAS.Open "SELECT * FROM CONCEPTOS WHERE " & TABLA & " = 1", DBSYSTEM, adOpenKeyset, adLockPessimistic
    If RSFORMULAS.RecordCount = 0 Then
        MsgBox "NO EXISTEN FORMULAS PARA ESTE PROCESO", vbInformation, "MENSAJE"
        Exit Sub
    End If
    ProgressBar1.Value = 0
    ProgressBar2.Value = 0
    X = 0
    Dim SUMAVALOR As Single
    If ExisteTablaAux(" [##_TMPPROVISIONES" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##_TMPPROVISIONES" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "CREATE TABLE  [##_TMPPROVISIONES" & VGL_COMPUTER & "]  (CODCCOSTO varchar(20), CCOSTO varchar(80), CODTRAB VARCHAR(8), NOMBRES varchar(100), CONCEPTO varchar(50), MONTO  Numeric(20,2) )"
    If ExisteTablaAux(" [##_TMPPROVCC" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##_TMPPROVCC" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "CREATE TABLE  [##_TMPPROVCC" & VGL_COMPUTER & "]  (CODCCOSTO varchar(20), CCOSTO varchar(80), CODTRAB VARCHAR(8), NOMBRES varchar(100), MONTO  Numeric(20,2) )"
    DBSYSTEM.Execute "UPDATE  [##TMPSELECT" & VGL_COMPUTER & "]  SET CODTRAB=CODTRAB", X
    If RSTRAB.RecordCount = 0 Then
        MsgBox "NO EXISTEN TRABAJADORES EN ESTA AREA DE TRABAJO VERIFICAR", vbCritical, "INFORMACION"
        Exit Sub
    End If
    Screen.MousePointer = 11
    Label4.Visible = True
    ProgressBar1.Visible = True
    ProgressBar1.Max = X
    Label5.Visible = True
    ProgressBar2.Visible = True
    ProgressBar2.Max = Y
    Dim DIASTRAB As Integer
    Dim SWTCH As Boolean
    While Not RSTRAB.EOF
        Label4.Caption = "PROCESANDO : " & RSTRAB!NOMBRES
        ProgressBar1.Value = ProgressBar1.Value + 1
        Label4.Refresh
        RSFORMULAS.MoveFirst
        While Not RSFORMULAS.EOF
            Label5.Caption = "CALCULANDO : " & RSFORMULAS!NOMBRE
            ProgressBar2.Value = ProgressBar2.Value + 1
            Label5.Refresh
            If ExisteTabla("BOL" & FECHA) Then
                VALOR = Round(DevuelveValor("SELECT SUM(MONTO) AS SUMADEMONTO FROM BOL" & FECHA & " BOL INNER JOIN MOV" & FECHA & " MOV ON BOL.INUMBOL = MOV.INUMBOL WHERE (((MOV.CONCEPTO) IN('" & RSFORMULAS!Codigo & "')) AND ((BOL.CODTRAB)='" & RSTRAB!CODTRAB & "'))", DBSYSTEM), 2)
                    SUMAVALOR = SUMAVALOR + VALOR
                    VALOR = Round(VALOR, 2)
                    'OBTUVO LOS VALORES --------------INSERCION
                    DBSYSTEM.Execute "INSERT INTO  [##_TMPPROVISIONES" & VGL_COMPUTER & "]  (CODCCOSTO, CCOSTO, CODTRAB, NOMBRES, CONCEPTO, MONTO) VALUES ('" & RSTRAB!CENTROCOSTO & "', '" & RSTRAB!NOMBRE & "', '" & RSTRAB!CODTRAB & "', '" & RSTRAB!NOMBRES & "', '" & RSFORMULAS!NOMBRE & "', " & VALOR & ")"
            End If
            Me.Refresh
            RSFORMULAS.MoveNext
        Wend
        DBSYSTEM.Execute "INSERT INTO  [##_TMPPROVISIONES" & VGL_COMPUTER & "]  (CODCCOSTO, CCOSTO, CODTRAB, NOMBRES, CONCEPTO, MONTO) VALUES ('" & RSTRAB!CENTROCOSTO & "', '" & RSTRAB!NOMBRE & "', '" & RSTRAB!CODTRAB & "', '" & RSTRAB!NOMBRES & "', 'ZZ TOTAL BASE', " & SUMAVALOR & ")"
        'VERIFICA LA FECHA DE INGRESO Y LA DE CESE DEL TRABAJADOR
        DIASTRAB = 0
        SWTCH = False
        If Year(DTPicker1.Value) = Year(RSTRAB!FECHAING) Then
            If Month(DTPicker1.Value) = Month(RSTRAB!FECHAING) Then
                If Not IsNull(RSTRAB!FECHACESE) Or Not Len(Trim(RSTRAB!FECHACESE)) = 0 Then
                    If Year(DTPicker1.Value) = Year(RSTRAB!FECHACESE) Then
                        If Month(DTPicker1.Value) = Month(RSTRAB!FECHACESE) Then
                            DIASTRAB = DateDiff("D", RSTRAB!FECHAING, RSTRAB!FECHACESE)
                            SWTCH = True
                        Else
                            DIASTRAB = 30 - Day(RSTRAB!FECHAING)
                            SWTCH = True
                        End If
                    Else
                        DIASTRAB = 30 - Day(RSTRAB!FECHAING)
                        SWTCH = True
                    End If
                Else
                    DIASTRAB = 30 - Day(RSTRAB!FECHAING)
                    SWTCH = True
                End If
            End If
        End If
        If SWTCH Then
            DBSYSTEM.Execute "INSERT INTO  [##_TMPPROVISIONES" & VGL_COMPUTER & "]  (CODCCOSTO, CCOSTO, CODTRAB, NOMBRES, CONCEPTO, MONTO) VALUES ('" & RSTRAB!CENTROCOSTO & "', '" & RSTRAB!NOMBRE & "', '" & RSTRAB!CODTRAB & "', '" & RSTRAB!NOMBRES & "', 'ZZ TOTAL PROVISIONES', " & ((SUMAVALOR / Val(xDivisor.Text)) / 30) * DIASTRAB & ")"
            DBSYSTEM.Execute "INSERT INTO  [##_TMPPROVCC" & VGL_COMPUTER & "]  (CODCCOSTO, CCOSTO, CODTRAB, NOMBRES, MONTO) VALUES ('" & RSTRAB!CENTROCOSTO & "', '" & RSTRAB!NOMBRE & "', '" & RSTRAB!CODTRAB & "', '" & RSTRAB!NOMBRES & "', " & ((SUMAVALOR / Val(xDivisor.Text)) / 30) * DIASTRAB & ")"
        Else
            DBSYSTEM.Execute "INSERT INTO  [##_TMPPROVISIONES" & VGL_COMPUTER & "]  (CODCCOSTO, CCOSTO, CODTRAB, NOMBRES, CONCEPTO, MONTO) VALUES ('" & RSTRAB!CENTROCOSTO & "', '" & RSTRAB!NOMBRE & "', '" & RSTRAB!CODTRAB & "', '" & RSTRAB!NOMBRES & "', 'ZZ TOTAL PROVISIONES', " & SUMAVALOR / Val(xDivisor.Text) & ")"
            DBSYSTEM.Execute "INSERT INTO  [##_TMPPROVCC" & VGL_COMPUTER & "]  (CODCCOSTO, CCOSTO, CODTRAB, NOMBRES, MONTO) VALUES ('" & RSTRAB!CENTROCOSTO & "', '" & RSTRAB!NOMBRE & "', '" & RSTRAB!CODTRAB & "', '" & RSTRAB!NOMBRES & "', " & SUMAVALOR / Val(xDivisor.Text) & ")"
        End If
        SUMAVALOR = 0
        ProgressBar2.Value = 0
        RSTRAB.MoveNext
    Wend
    Label4.Visible = False
    ProgressBar1.Visible = False
    Label5.Visible = False
    ProgressBar2.Visible = False
    Screen.MousePointer = 1
    Set RSTRAB = Nothing
    Set RSFORMULAS = Nothing
    
    Select Case OP
    Case 0
        XtipRep.Reporte = "FORMULAS VACACIONES"
        XtipRep.TITLE = " VACACIONES "
    Case 1
        XtipRep.Reporte = "FORGRAPRO"
        XtipRep.TITLE = " GRATIFICACIONES "
    Case 2
        XtipRep.Reporte = "FORCTSPRO"
        XtipRep.TITLE = " CTS "
    End Select
    XtipRep.MES = Format(DTPicker1.Value, "MMMM")
    If MsgBox("EL PROCESO DE PROVISIONES DEL MES DE " & XtipRep.MES & " CONCLUYÓ SATISFACTORIAMENTE, DESEA IMPRIMIR EL FORMATO ", vbYesNo, "CONFIRMAR") = vbYes Then
        XtipRep.Show 1
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub
Private Sub FORM_KEYUP(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = vbKeyF5 Then CMSELECTRAB_CLICK
End Sub

Private Sub Form_Load()
    xDivisor.Text = 12
    If ExisteTablaAux(" [##TMPCOSTOS" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPCOSTOS" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "SELECT * INTO  [##TMPCOSTOS" & VGL_COMPUTER & "]  FROM CCOSTOS "
    If ExisteTablaAux("##TMPTRABAJADORES") Then DBSYSTEM.Execute "DROP TABLE ##TMPTRABAJADORES"
    DBSYSTEM.Execute "SELECT * INTO ##TMPTRABAJADORES FROM TRABAJADORES "
    Set RSTRAB = New ADODB.Recordset
    If Not ExisteTablaAux(" [##TMPSELECT" & VGL_COMPUTER & "] ") Then DBSTARPLAN.Execute "CREATE TABLE  [##TMPSELECT" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8), NOMBRES VARCHAR(50), FECHAING DATETIME, AREA VARCHAR(10), CENTROCOSTO VARCHAR(10), TIPOTRAB VARCHAR(2), BASICO  Numeric(20,2),BASICO1 NUMERIC(20,2) )"
    RSTRAB.Open "SELECT [##TMPSELECT" & VGL_COMPUTER & "].CODTRAB, [##TMPSELECT" & VGL_COMPUTER & "].NOMBRES, [##TMPSELECT" & VGL_COMPUTER & "].AREA, [##TMPSELECT" & VGL_COMPUTER & "].CENTROCOSTO,  [##TMPCOSTOS" & VGL_COMPUTER & "] .NOMBRE, [##TMPSELECT" & VGL_COMPUTER & "].TIPOTRAB, [##TMPSELECT" & VGL_COMPUTER & "].BASICO, ##TMPTRABAJADORES.FECHAING, ##TMPTRABAJADORES.FECHACESE" & _
                " FROM  [##TMPCOSTOS" & VGL_COMPUTER & "]  INNER JOIN (##TMPTRABAJADORES INNER JOIN  [##TMPSELECT" & VGL_COMPUTER & "]  ON ##TMPTRABAJADORES.CODTRAB = [##TMPSELECT" & VGL_COMPUTER & "].CODTRAB) ON  [##TMPCOSTOS" & VGL_COMPUTER & "] .CODCCOSTO = [##TMPSELECT" & VGL_COMPUTER & "].CENTROCOSTO", DBSYSTEM, adOpenStatic
    Set DGLista.DataSource = RSTRAB
    DTPicker1.Value = Date
    ProgressBar1.Value = 0
    ProgressBar1.Visible = False
    Me.Label4.Visible = False
    Option1(0).Value = True
    ProgressBar2.Value = 0
    ProgressBar2.Visible = False
    Me.Label5.Visible = False
End Sub
Private Sub OPTION1_CLICK(INDEX As Integer)
    OP = INDEX
End Sub
Public Function CAMBIACADENA(GENERAL_2 As Boolean, ByVal BOLETA As String, ByVal CADENA As String, ByVal CODTRAB As String, Optional MES As String = "NONE") As String
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
                VALOR = CALCULOCONCEPTOS(CODTRAB, BOLETA, PROMEDIO, CAMPO, GENERAL_2)
            Case "ULTIMOVALOR"
                VALOR = CALCULOCONCEPTOS(CODTRAB, BOLETA, ULTIMOVALOR, CAMPO, GENERAL_2)
            Case "PRIMERVALOR"
                VALOR = CALCULOCONCEPTOS(CODTRAB, BOLETA, PRIMERVALOR, CAMPO, GENERAL_2)
            Case "SUMA"
                VALOR = CALCULOCONCEPTOS(CODTRAB, BOLETA, SUMA, CAMPO, GENERAL_2)
            Case "MEDIA"
                VALOR = CALCULOCONCEPTOS(CODTRAB, BOLETA, MEDIA, CAMPO, GENERAL_2)
            Case "PROMEDIOVALOR"
                VALOR = CALCULOCONCEPTOS(CODTRAB, BOLETA, PROMEDIOVALOR, CAMPO, GENERAL_2)
            Case "PRIMERO"
                VALOR = CALCULOCONCEPTOS(CODTRAB, BOLETA, PRIMERO, CAMPO, GENERAL_2)
            Case "ULTIMO"
                VALOR = CALCULOCONCEPTOS(CODTRAB, BOLETA, ULTIMO, CAMPO, GENERAL_2)
            Case "MAYORVALOR"
                VALOR = CALCULOCONCEPTOS(CODTRAB, BOLETA, MAYORVALOR, CAMPO, GENERAL_2)
            Case "MENORVALOR"
                VALOR = CALCULOCONCEPTOS(CODTRAB, BOLETA, MENORVALOR, CAMPO, GENERAL_2)
            Case "NUMERO"
                VALOR = CALCULOCONCEPTOS(CODTRAB, BOLETA, Numero, CAMPO, GENERAL_2)
            Case "NSECUENCIA"
                VALOR = CALCULOCONCEPTOS(CODTRAB, BOLETA, NSECUENCIA, CAMPO, GENERAL_2)
            Case "INFOBOLETA"
                If MES = "NONE" Then
                    VALOR = CALCULOMES(CODTRAB, CAMPO, , xFechaIni.Value, xFechaFin.Value)
                Else
                    VALOR = CALCULOMES(CODTRAB, CAMPO, MES)
                End If
            Case "PLAN"
                If MES = "NONE" Then
                    VALOR = CALCULOMES2(CODTRAB, CAMPO, MES)
                Else
                    VALOR = CALCULOMES2(CODTRAB, CAMPO, MES)
                End If
        End Select
        If IsNull(VALOR) Then VALOR = 0
        CADENA = Replace(CADENA, Mid(CADENA, POSARROBA, (POS2 - POSARROBA) + 1), "" & VALOR)
        POSARROBA = InStr(POSARROBA, CADENA, "@")
    Loop
    CAMBIACADENA = CADENA
End Function

Public Function CALCULOCONCEPTOS(CODTRAB As String, MES As String, TIPO As TipoCalculo, CONCEPTO As String, GENERAL_3 As Boolean) As Double
    Dim XNUMMES As Integer, X As Integer, NUMOCURRE As Integer, SUMATOTAL As Double
    Dim FEC1 As Date, FEC2 As Date, STRMES As String, VALOR As Double, RESULTADO As Double
    NUMOCURRE = 0
    RESULTADO = 0
    SUMATOTAL = 0
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
    'PARA CALCULAR LA SECUENCIA DE UN VALOR
    Dim IX As Integer, SX As Integer
    Dim CADX As String
    Dim RSSECX As New ADODB.Recordset
    SX = 0
    RSSECX.Fields.Append "NUMERO", adInteger
    RSSECX.Open
        If ExisteTabla("BOL" & MES) Then
            If Not GENERAL_3 Then
                VALOR = Round(DevuelveValor("SELECT SUM(MONTO) AS SUMADEMONTO FROM BOL" & MES & " BOL INNER JOIN MOV" & MES & " MOV ON BOL.INUMBOL = MOV.INUMBOL WHERE (((MOV.CONCEPTO) IN(" & CONCEPTO & ")) AND ((BOL.CODTRAB)='" & CODTRAB & "'))", DBSYSTEM), 2)
            Else
                VALOR = Round(DevuelveValor("SELECT SUM(" & CONCEPTO & ") AS SUMADEMONTO FROM BOL" & MES & " BOL WHERE BOL.CODTRAB='" & CODTRAB & "'", DBSYSTEM), 2)
            End If
            Select Case TIPO
                Case PRIMERVALOR
                    If RESULTADO = 0 And VALOR <> 0 Then
                        RESULTADO = VALOR
                    End If
                Case ULTIMOVALOR
                    If VALOR <> 0 Then RESULTADO = VALOR
                Case MAYORVALOR
                    If X = 1 Then RESULTADO = VALOR Else If VALOR > RESULTADO Then RESULTADO = VALOR
                Case MENORVALOR
                    If RESULTADO = 0 Then RESULTADO = VALOR Else If VALOR < RESULTADO Then RESULTADO = VALOR
                Case NSECUENCIA
                    If VALOR <> 0 Then
                        SX = SX + 1
                       Else: SX = 0
                    End If
                    If SX <> 0 Or XNUMMES = X Then
                        RSSECX.AddNew
                        RSSECX!Numero = SX: RSSECX.Update
                    End If
            End Select
            If VALOR <> 0 Then NUMOCURRE = NUMOCURRE + 1
            SUMATOTAL = SUMATOTAL + VALOR
        End If
    Select Case TIPO
        Case MEDIA
            RESULTADO = SUMATOTAL / 2
        Case PROMEDIO
            RESULTADO = SUMATOTAL / XNUMMES
        Case PROMEDIOVALOR
            If SUMATOTAL = 0 Then RESULTADO = 0 Else RESULTADO = SUMATOTAL / NUMOCURRE
        Case SUMA
            RESULTADO = SUMATOTAL
        Case Numero
            RESULTADO = NUMOCURRE
        Case NSECUENCIA
            RSSECX.Sort = "NUMERO DESC"
            If RSSECX.RecordCount > 0 Then
               RSSECX.MoveFirst
               RESULTADO = RSSECX!Numero
            End If
    End Select
    CALCULOCONCEPTOS = Round(RESULTADO, 2)
End Function

Private Sub XDIVISOR_LOSTFOCUS()
    If Len(Trim(xDivisor.Text)) = 0 Or Val(xDivisor.Text) = 0 Then xDivisor.Text = 1
End Sub


