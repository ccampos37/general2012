VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Begin VB.Form frAdelantos 
   Caption         =   "Adelanto de Remuneraciones"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8580
   Icon            =   "frAdelantos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   8580
   Begin VB.CommandButton cmdCuentasCtes 
      Caption         =   "Aplicar &Cuentas Corrientes"
      Height          =   345
      Left            =   225
      TabIndex        =   5
      Top             =   2318
      Width           =   2295
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Por Areas de Trabajo"
      Height          =   210
      Left            =   195
      TabIndex        =   1
      Top             =   855
      Value           =   -1  'True
      Width           =   1830
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Por Centros de Costo"
      Height          =   210
      Left            =   195
      TabIndex        =   2
      Top             =   1110
      Width           =   1830
   End
   Begin VB.CommandButton cmSelecTrab 
      Caption         =   "Seleccion (F5)"
      Height          =   990
      Left            =   7470
      Picture         =   "frAdelantos.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1695
      Width           =   945
   End
   Begin MSDataGridLib.DataGrid xData 
      Height          =   3450
      Left            =   165
      TabIndex        =   7
      Top             =   2775
      Width           =   8250
      _ExtentX        =   14552
      _ExtentY        =   6085
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
      Left            =   225
      TabIndex        =   0
      Top             =   420
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   503
      Locked          =   -1  'True
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker xFechaFin 
      Height          =   285
      Left            =   1470
      TabIndex        =   8
      Top             =   1830
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   60358657
      CurrentDate     =   36699
   End
   Begin MSComCtl2.DTPicker xFechaIni 
      Height          =   285
      Left            =   1470
      TabIndex        =   9
      Top             =   1485
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   60358657
      CurrentDate     =   36699
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   1725
      Left            =   3045
      TabIndex        =   3
      Top             =   420
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
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   585
      Left            =   120
      TabIndex        =   15
      Top             =   6240
      Width           =   8430
      Begin VB.TextBox xTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5595
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   18
         Text            =   "frAdelantos.frx":0BD4
         Top             =   165
         Width           =   1350
      End
      Begin VB.CommandButton cmGrabar 
         Caption         =   "&Grabar"
         Height          =   405
         Left            =   7065
         TabIndex        =   17
         Top             =   105
         Width           =   1230
      End
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   405
         Left            =   2415
         TabIndex        =   16
         Top             =   105
         Width           =   1230
      End
      Begin VB.Image Image4 
         Height          =   240
         Left            =   5190
         Picture         =   "frAdelantos.frx":0BDB
         ToolTipText     =   "Suma total de los adelantos por aceptar"
         Top             =   180
         Width           =   240
      End
      Begin VB.Label xNumTrab 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 0 Trabajadores"
         Height          =   285
         Left            =   45
         TabIndex        =   19
         Top             =   165
         Width           =   2205
      End
   End
   Begin VB.Line Line1 
      X1              =   5685
      X2              =   5760
      Y1              =   2505
      Y2              =   2505
   End
   Begin VB.Image i2 
      Height          =   240
      Left            =   2790
      Picture         =   "frAdelantos.frx":0F1D
      ToolTipText     =   "Indica que se han cargo Cuentas Corrientes a descontar en este adelanto"
      Top             =   2370
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image i1 
      Height          =   240
      Left            =   2550
      Picture         =   "frAdelantos.frx":125F
      ToolTipText     =   "Indica que se han cargo Cuentas Corrientes a descontar en este adelanto"
      Top             =   2370
      Visible         =   0   'False
      Width           =   240
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
      Left            =   7305
      TabIndex        =   14
      Top             =   720
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   7830
      Picture         =   "frAdelantos.frx":15A1
      Top             =   225
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Periodos en Cronograma"
      Height          =   195
      Left            =   3045
      TabIndex        =   13
      Top             =   180
      Width           =   1740
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Mes de Trabajo"
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   180
      Width           =   1110
   End
   Begin VB.Label l1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Inicial"
      Height          =   195
      Left            =   225
      TabIndex        =   11
      Top             =   1530
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label l2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Final"
      Height          =   195
      Left            =   225
      TabIndex        =   10
      Top             =   1890
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label xlAuto 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Auto-rellenado"
      Height          =   270
      Left            =   5025
      TabIndex        =   6
      Top             =   2310
      Width           =   1485
   End
   Begin VB.Image xAuto 
      Height          =   240
      Left            =   5160
      Picture         =   "frAdelantos.frx":18AB
      Top             =   2310
      Width           =   240
   End
End
Attribute VB_Name = "frAdelantos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSTMPADEL As New ADODB.Recordset
Dim RSMESES As New ADODB.Recordset
Dim RSADELANTO As New ADODB.Recordset
Dim REGACT As REGWIN, CADIN As String

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
        MsgBox "No existe nada por Grabar", vbCritical
        Exit Sub
    End If
    If xMes.Tag = "" Then
        MsgBox "No ha Seleccionado un mes, Cambie por favor el mes", vbCritical
        xMes.SetFocus
        Exit Sub
    End If
    Dim RSTBADEL As New ADODB.Recordset
    
    DBSYSTEM.Execute "DELETE FROM  [##_TMPADELANTO" & VGL_COMPUTER & "]  WHERE MONTO=0"
    RSTMPADEL.Requery
    FORMATEARDBG
    Dim CARGACC As Boolean
    Dim RSAUX As New ADODB.Recordset
    CARGACC = False
    If UCase(Dir$(App.PATH & "\ADELCC.DYB")) = "ADELCC.DYB" Then
        If MsgBox("Esta punto de grabar los Debitos de Cuentas Corrientes de Trabajadores.. Desea hacer efectivo los Debitos especificos", vbYesNo + vbQuestion) = vbNo Then CARGACC = False Else CARGACC = True
    End If
    If MsgBox("Desea Continuar", vbYesNo) = vbYes Then
        If CARGACC Then
            RSAUX.Open App.PATH & "\ADELCC.DYB", , adOpenStatic, adLockReadOnly, adCmdFile
            Dim NUMPAGO As Integer
            Do While Not RSAUX.EOF
                'NUMPAGO = DevuelveValor("SELECT MAX(NUMPAGO) AS MAXI FROM PAGOSCTA", DBSYSTEM) + 1
                DBSYSTEM.Execute "INSERT INTO PAGOSCTA (CODMOV,NUMBOL,CODNOMBOL,TIPOBOLETA,MONTO,DOLAR,CODTRAB,TIPO,SECUENCIA) VALUES (" & RSAUX!CODMOV & ",0," & Lista.SelectedItem.Tag & ",'A'," & Round(RSAUX!DEBITO, 2) & ",0,'" & RSAUX!CODTRAB & "'," & IIf(RSAUX!Tip = "E", 2, 1) & "," & RSAUX!SECUENCIA & ")"
                DBSYSTEM.Execute "UPDATE MOVICTA SET SALDO=SALDO-" & RSAUX!DEBITO & " WHERE CODMOV=" & RSAUX!CODMOV
                RSAUX.MoveNext
            Loop
        End If
        Do While Not RSTMPADEL.EOF
            Dim Codigo As Long
            'CODIGO = DevuelveValor("SELECT MAX(CODIGO) AS MAXI FROM " & REGSISTEMA.TABLAADEL, DBSYSTEM) + 1
            DBSYSTEM.Execute "INSERT INTO " & REGSISTEMA.TABLAADEL & " (CODTRAB,MES,FECHAING,MONTO,NUMBOL,NOMBOL, ORIGEN) VALUES ('" & RSTMPADEL!CODTRAB & "'," & DateSQL(xMes.Tag) & "," & DateSQL(Date) & "," & RSTMPADEL!MONTO & ",0,0," & Lista.SelectedItem.Tag & ")"
            RSTMPADEL.MoveNext
        Loop
    End If
    MsgBox "Los datos se han grabado satisfactoriamente", vbInformation
    Set RSTBADEL = Nothing
    Set RSAUX = Nothing
    Unload Me
End Sub

Private Sub CMSELECTRAB_CLICK()
    If Not xFechaIni.Visible Then
        MsgBox "Debera seleccionar un periodo de Pago", vbCritical
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
            MsgBox "No se han encontrado Areas o Centros de Costos que esten programados para pagos de Adelantos de Remuneraciones en el periodo seleccionado", vbCritical
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
    XADDTRAB_CLICK
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub FORM_KEYDOWN(KEYCODE As Integer, Shift As Integer)
If KEYCODE = vbKeyF5 Then CMSELECTRAB_CLICK
If KEYCODE = 84 And Shift = 4 Then XAUTO_CLICK
End Sub

Private Sub Form_Load()
    'BORRAMOS EL TEMPORAL DE CUENTAS CORRIENTES
    If UCase(Dir$(App.PATH & "\ADELCC.DYB")) = "ADELCC.DYB" Then Kill App.PATH & "\ADELCC.DYB"
    Me.Tag = "PANEL DE ADELANTOS DE PAGO"
    RSADELANTO.Open "SELECT CODTRAB,SUM(MONTO) AS SALDO FROM " & REGSISTEMA.TABLAADEL & " WHERE NOMBOL=0 GROUP BY CODTRAB ORDER BY CODTRAB", DBSYSTEM, adOpenKeyset, adLockOptimistic
    If ExisteTablaAux(" [##_TMPADELANTO" & VGL_COMPUTER & "] ") Then
        DBSYSTEM.Execute "DROP TABLE  [##_TMPADELANTO" & VGL_COMPUTER & "] "
    End If
    DBSYSTEM.Execute "CREATE TABLE  [##_TMPADELANTO" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8),NOMBRES varchar(60),BASICO  Numeric(20,2) , FECHAING DATETIME, MONTO  Numeric(20,2) , PENDIENTE  Numeric(20,2) )"
    RSTMPADEL.Open " [##_TMPADELANTO" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic
    Set xData.DataSource = RSTMPADEL
    FORMATEARDBG
End Sub

Private Sub Form_Resize()
    If Me.Width < 8460 Then Exit Sub
    If Me.Height < 7005 Then Exit Sub
    
    '********************************
    Frame1.TOP = Me.ScaleHeight - 555
    '********************************
    xData.Height = Me.ScaleHeight - 3525
    xData.Width = Me.ScaleWidth - 210

End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSTMPADEL = Nothing
    Set RSMESES = Nothing
    Set RSADELANTO = Nothing
    If UCase(Dir$(App.PATH & "\ADELCC.DYB")) = "ADELCC.DYB" Then Kill App.PATH & "\ADELCC.DYB"
End Sub

Private Sub I1_CLICK()
    MsgBox "Indica que se han cargo Cuentas Corrientes a descontar en este adelanto", vbInformation
End Sub

Private Sub I2_CLICK()
    MsgBox "Indica que se han cargo Cuentas Corrientes a descontar en este adelanto", vbInformation
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
    DBSYSTEM.Execute "DELETE FROM  [##_TMPADELANTO" & VGL_COMPUTER & "] "
    If UCase(Dir$(App.PATH & "\ADELCC.DYB")) = "ADELCC.DYB" Then Kill App.PATH & "\ADELCC.DYB"
    Dim RSTRAB As New ADODB.Recordset
    Dim STRCAD As String
    'ELIMINAMOS AQUELLOS QUE YA HAYAN TENIDO ADELANTOS EN ESTE PERIODO
    DBSYSTEM.Execute "DELETE FROM  [##TMPSELECT" & VGL_COMPUTER & "]  WHERE CODTRAB IN (SELECT CODTRAB FROM ADEL2000 WHERE ORIGEN=" & Me.Lista.SelectedItem.Tag & ")"
    If DevuelveValor("SELECT USARCRONOGRAMA FROM EMPRESA", DBSYSTEM) = 1 Then
        If Option1.Value Then
            RSTRAB.Open "SELECT * FROM  [##TMPSELECT" & VGL_COMPUTER & "]  WHERE AREA IN " & CADIN & " ORDER BY NOMBRES", DBSYSTEM, adOpenStatic
        Else
            RSTRAB.Open "SELECT * FROM  [##TMPSELECT" & VGL_COMPUTER & "]  WHERE CENTROCOSTO IN " & CADIN & " ORDER BY NOMBRES", DBSYSTEM, adOpenStatic
        End If
    Else
        RSTRAB.Open "SELECT * FROM  [##TMPSELECT" & VGL_COMPUTER & "]  ORDER BY NOMBRES", DBSYSTEM, adOpenStatic
    End If
    If RSTRAB.RecordCount = 0 Then
        MsgBox "Los Seleccionados ya tienen provisionados adelantos para este periodo " & Chr(13) & "Si desea eliminarlos ir al panel administración de adelantos", vbExclamation
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
End Sub

Private Sub XAUTO_CLICK()
    If RSTMPADEL.RecordCount = 0 Then
        MsgBox "No existen registros de Trabajadores", vbCritical
        Exit Sub
    End If
    VAR = 1
    frAutoAd.Show 1
End Sub

Private Sub XDATA_AFTERCOLUPDATE(ByVal COLINDEX As Integer)
    RSTMPADEL.MOVE 0
End Sub

Private Sub XDATA_AFTERUPDATE()
    Dim RSSUMA As New ADODB.Recordset
    RSSUMA.Open "SELECT SUM(MONTO) AS TOTAL FROM  [##_TMPADELANTO" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenStatic
    xTotal.Text = Format(RSSUMA!TOTAL, "0.00")
    Set RSSUMA = Nothing
End Sub

Private Sub XDATA_HEADCLICK(ByVal COLINDEX As Integer)
    If xData.Columns(COLINDEX).Caption = "MONTO" Then XAUTO_CLICK Else RSTMPADEL.Sort = xData.Columns(COLINDEX).Caption
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
    xNumTrab.Caption = " " & RSTMPADEL.RecordCount & " Trabajadores"
End Sub

Public Function GETDATA() As ADODB.Recordset
    Set GETDATA = RSTMPADEL
End Function

Private Sub XMES_DBLCLICK()
    Lista.ListItems.Clear
    Dim RSMESES As New ADODB.Recordset
    RSMESES.Open "SELECT MESACTIVO, NOMBRE FROM MESESACT ORDER BY MESACTIVO", DBSYSTEM, adOpenStatic
    If RSMESES.RecordCount = 0 Then
        MsgBox "No se han encontrado Meses en Actividad", vbCritical
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
    RSMESES.Open "SELECT CODIGO, NOMBRE, FECHAINI, FECHAFIN FROM NOMBOL WHERE CERRADO=0 AND MES=" & DateSQL(CDate(xMes.Tag)) & " ORDER BY FECHAINI", DBSYSTEM, adOpenStatic
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

