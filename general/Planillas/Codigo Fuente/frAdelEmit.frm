VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frAdelEmit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Panel de Administración de Adelantos de Remuneraciones"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   Icon            =   "frAdelEmit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   8250
   Tag             =   "Panel de Boletas Emitidas"
   Begin VB.CheckBox chkDetadel 
      Caption         =   "Tomar Adelanto Detallado"
      Height          =   285
      Left            =   2370
      TabIndex        =   24
      Top             =   6750
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reporte Especial"
      Height          =   360
      Left            =   6390
      TabIndex        =   23
      Top             =   6705
      Visible         =   0   'False
      Width           =   1785
   End
   Begin AplisetControlText.Aplitext xArea 
      Height          =   285
      Left            =   3510
      TabIndex        =   18
      Top             =   2805
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   503
      Locked          =   -1  'True
      Text            =   ""
   End
   Begin VB.OptionButton Sel1 
      Caption         =   "Centros de Costo"
      Height          =   225
      Index           =   1
      Left            =   1935
      TabIndex        =   17
      Top             =   2865
      Width           =   1530
   End
   Begin VB.OptionButton Sel1 
      Caption         =   "Areas de Trabajo"
      Height          =   225
      Index           =   0
      Left            =   150
      TabIndex        =   16
      Top             =   2865
      Value           =   -1  'True
      Width           =   1560
   End
   Begin VB.CommandButton cmCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   6570
      TabIndex        =   15
      Top             =   2730
      Width           =   1560
   End
   Begin VB.CommandButton cmBilletes 
      Caption         =   "&Billetaje"
      Height          =   360
      Left            =   6570
      TabIndex        =   14
      Top             =   2256
      Width           =   1560
   End
   Begin Crystal.CrystalReport RptBoletas 
      Left            =   5655
      Top             =   1785
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmPagosBanco 
      Caption         =   "Pagos por Banco"
      Height          =   360
      Left            =   6570
      TabIndex        =   13
      Top             =   1784
      Width           =   1560
   End
   Begin VB.CommandButton cmPrtTodos 
      Caption         =   "&Imprimir Todos"
      Height          =   360
      Left            =   6570
      TabIndex        =   12
      Top             =   1312
      Width           =   1560
   End
   Begin VB.CommandButton cmPrtUno 
      Caption         =   "Imprimir Uno"
      Height          =   360
      Left            =   6570
      TabIndex        =   11
      Top             =   840
      Width           =   1560
   End
   Begin VB.ComboBox xMeses 
      Height          =   315
      Left            =   165
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   420
      Width           =   2925
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   1935
      Left            =   135
      TabIndex        =   1
      Top             =   840
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   3413
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FlatScrollBar   =   -1  'True
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
         Text            =   "Nombre"
         Object.Width           =   6774
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Fecha Inic."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fecha Term."
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   135
      Top             =   4620
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frAdelEmit.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frAdelEmit.frx":158E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid dgBoletas 
      Height          =   3165
      Left            =   150
      TabIndex        =   0
      Top             =   3165
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   5583
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   17
      RowDividerStyle =   0
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
      Caption         =   "Boletas de Remuneraciones"
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "CODTRAB"
         Caption         =   "Código"
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
         DataField       =   "NOMBRES"
         Caption         =   "Trabajador"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "ADELANTO"
         Caption         =   "Adelanto"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00 "
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "OTROSING"
         Caption         =   "Otros Ingresos"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00 "
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "OTROSEGR"
         Caption         =   "Otros Egresos"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00 "
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "NETO"
         Caption         =   "Adelanto Neto"
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
         MarqueeStyle    =   3
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   689.953
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2594.835
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1049.953
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1035.213
         EndProperty
      EndProperty
   End
   Begin VB.TextBox SqlCad 
      Height          =   285
      Left            =   1275
      TabIndex        =   20
      Text            =   "SqlCad"
      Top             =   4500
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.TextBox SqlText 
      Height          =   285
      Left            =   1260
      TabIndex        =   21
      Text            =   "SqlText"
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   390
      Left            =   150
      TabIndex        =   22
      Top             =   6690
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   688
      ButtonWidth     =   3254
      ButtonHeight    =   582
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Otros Procesos        "
            Object.ToolTipText     =   "Click aquí para más reportes de boletas"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "EliminaTodasBol"
                  Text            =   "Eliminar Todos los adelantos"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   7635
      Picture         =   "frAdelEmit.frx":19E2
      Top             =   150
      Width           =   480
   End
   Begin VB.Label Label2 
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
      Left            =   6645
      TabIndex        =   19
      Top             =   420
      Width           =   1065
   End
   Begin VB.Label xCont 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0 Registros"
      Height          =   270
      Left            =   150
      TabIndex        =   10
      Top             =   6360
      Width           =   2640
   End
   Begin VB.Label sum1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00 "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4680
      TabIndex        =   9
      Top             =   6360
      Width           =   1050
   End
   Begin VB.Label sum2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5745
      TabIndex        =   8
      Top             =   6360
      Width           =   1050
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total "
      Height          =   270
      Left            =   2805
      TabIndex        =   7
      Top             =   6360
      Width           =   795
   End
   Begin VB.Label xSumNet 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   6810
      TabIndex        =   6
      Top             =   6360
      Width           =   1065
   End
   Begin VB.Label xSumIng 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3615
      TabIndex        =   5
      Top             =   6360
      Width           =   1050
   End
   Begin VB.Image xMarcaTodos 
      Height          =   240
      Left            =   5160
      Picture         =   "frAdelEmit.frx":1CEC
      Top             =   450
      Width           =   240
   End
   Begin VB.Image xError 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Left            =   3195
      Picture         =   "frAdelEmit.frx":202E
      Top             =   210
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Mes de Adelantos"
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   180
      Width           =   1275
   End
   Begin VB.Label lMarcaTodos 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Marcar Todos"
      Height          =   270
      Left            =   5040
      TabIndex        =   4
      Top             =   465
      Width           =   1410
   End
End
Attribute VB_Name = "frAdelEmit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSBOLE As New ADODB.Recordset
Dim REGACT As REGWIN
Dim ITSOPEN As Boolean
Dim WithEvents RSLISTA As ADODB.Recordset
Attribute RSLISTA.VB_VarHelpID = -1

Private Sub CMBILLETES_Click()
    VPTAREA = "ADELANTO"
    frmBilletes.Show 1
End Sub

Private Sub cmCerrar_Click()
    Unload Me
End Sub
Private Sub CMPAGOSBANCO_Click()
    If RSLISTA.RecordCount = 0 Then
        MsgBox "No existen registros a procesar procesar. Falta selecionar un Periodo de Pago conteniendo Boletas de Remuneraciones", vbCritical
        Exit Sub
    End If
    If ExisteTablaAux(" [##TMPBANCOS" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPBANCOS" & VGL_COMPUTER & "] "
    If ExisteTablaAux(" [##PAGOSXBANCO" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##PAGOSXBANCO" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "SELECT CODTRAB, TIPDOC, DOCIDEN, CTABANCO, BANCO INTO  [##TMPBANCOS" & VGL_COMPUTER & "]  FROM " & REGSISTEMA.BASESQL & ".dbo.TRABAJADORES WHERE CODTRAB IN (SELECT CODTRAB FROM  [##_TMPLSTADEL" & VGL_COMPUTER & "] )"
    DBSYSTEM.Execute "SELECT A.CODTRAB, NOMBRES, NETO, TIPDOC, DOCIDEN, CTABANCO, BANCO INTO  [##PAGOSXBANCO" & VGL_COMPUTER & "]  FROM  [##_TMPLSTADEL" & VGL_COMPUTER & "]  A,  [##TMPBANCOS" & VGL_COMPUTER & "] ##TMPBANCOS WHERE A.CODTRAB=##TMPBANCOS.CODTRAB"
    frPagoBco.Show 1
End Sub

Private Sub CMPRTTODOS_Click()
    SqlCad = ""
    Select Case MsgBox("Desea Imprimir el detallado(SI) o resumido(NO)", vbQuestion + vbYesNoCancel + vbDefaultButton2)
        Case vbYes: ImpDeta
        Case vbNo: CONSULTAREP SqlCad
        Case Else: Exit Sub
    End Select
End Sub
Private Sub ImpDeta()
    Dim NUM As Variant
    Dim CONT As Integer
    If RSLISTA.RecordCount = 0 Then
        MsgBox "No existen registros para imprimir ", vbExclamation
    End If
    Screen.MousePointer = 11
    SqlCad.Text = ""
    SqlText.Text = ""
    If MsgBox("Desea Imprimir los registros seleccionado", vbQuestion + vbYesNo) = vbYes Then
        For Each NUM In dgBoletas.SelBookmarks
            dgBoletas.Bookmark = NUM
            dgBoletas.COL = 0
            SqlCad.Text = SqlCad.Text & "'" & Trim(dgBoletas.Text) & "',"
        Next
        If SqlCad.Text <> "" Then
            SqlCad.Text = Left(SqlCad.Text, Len(SqlCad.Text) - 1)
            SqlCad.Text = " AND  DETADEL.CODTRAB IN (" & SqlCad.Text & ")"
          Else
        End If
    End If
    If ExisteTablaAux("[##TMPDETADEL" & VGL_COMPUTER & "]") Then DBAUXCOM.Execute "DROP TABLE " & "[##TMPDETADEL" & VGL_COMPUTER & "]"
    DBSYSTEM.Execute "SELECT *,'[##TMPDETADEL" & VGL_COMPUTER & "]' AS TMP  INTO " & "[##TMPDETADEL" & VGL_COMPUTER & "] FROM  DETADEL WHERE NOMBOL=" & Lista.SelectedItem.Tag & SqlCad.Text
    
    With rptBoletas
        .Reset
        .ReportFileName = REGSISTEMA.REPORTES & "REPORT1.RPT"
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
        .StoredProcParam(0) = "[##TMPDETADEL" & VGL_COMPUTER & "]"
        '.StoredProcParam(1) = "[##TMPDETADEL" & VGL_COMPUTER & "]"
        
        .Destination = crptToWindow
        .WindowShowPrintBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowTitle = "REPORT1 - RECIBO DE ADELANTO DE QUINCENA"
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .SubreportToChange = "Report1AUX.rpt"
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
        .ParameterFields(0) = "@TABLATMP;[##TMPDETADEL" & VGL_COMPUTER & "];TRUE"
        .StoredProcParam(0) = "[##TMPDETADEL" & VGL_COMPUTER & "]"
        If .Status <> 2 Then .Action = 1
    End With
    Screen.MousePointer = 1
End Sub

Private Sub CMPRTUNO_Click()
    IMPRIMIR
End Sub
Private Sub IMPRIMIR()
Dim NUM As Variant
    Dim CONT As Integer
    If RSLISTA.RecordCount = 0 Then
        MsgBox "No existen registros para imprimir ", vbExclamation
    End If
    SqlCad.Text = ""
    SqlText.Text = ""
    For Each NUM In dgBoletas.SelBookmarks
        dgBoletas.Bookmark = NUM
        dgBoletas.COL = 0
        SqlCad.Text = SqlCad.Text & "'" & Trim(dgBoletas.Text) & "',"
    Next
    If SqlCad.Text <> "" Then
        SqlCad.Text = Left(SqlCad.Text, Len(SqlCad.Text) - 1)
        SqlCad.Text = " AND  [##_TMPLSTADEL" & VGL_COMPUTER & "].CODTRAB IN (" & SqlCad.Text & ")"
      Else
        MsgBox "Tiene que selecionar por lo menos un registro"
        Exit Sub
    End If
    CONSULTAREP SqlCad
End Sub

Private Sub CONSULTAREP(TEXTO As TextBox)
Dim INTO As String
Dim RUTA As String, RUTA2 As String
'ON ERROR GOTO ERRPRINT
    Screen.MousePointer = 11
    RUTA = REGSISTEMA.BASESQL & ".dbo."
    RUTA2 = " [##_TMPLSTADEL" & VGL_COMPUTER & "] "
    INTO = " INTO   [##TMPBOLADEL" & VGL_COMPUTER & "]  "
    SqlText = "" & _
    "SELECT VWTRABAJ.CODTRAB, VWTRABAJ.NOMBRES, VWTRABAJ.CODCCOSTO, " & _
    "VWTRABAJ.CENTRO , VWTRABAJ.CODAREA, VWTRABAJ.NOMBREAREA, " & _
    "[##_TMPLSTADEL" & VGL_COMPUTER & "].ADELANTO , [##_TMPLSTADEL" & VGL_COMPUTER & "].OTROSING, " & _
    "[##_TMPLSTADEL" & VGL_COMPUTER & "].OTROSEGR , NOMBOL.FECHAADELANTO " & INTO & _
    " FROM " & RUTA2 & ", NOMBOL, " & RUTA & "VWTRABAJ " & _
    " WHERE((([##_TMPLSTADEL" & VGL_COMPUTER & "].NOMBOL) = [NOMBOL].[CODIGO]) AND (([##_TMPLSTADEL" & VGL_COMPUTER & "].CODTRAB) = [VWTRABAJ].[CODTRAB])) " & _
    TEXTO.Text ' LA OTRA PARTE DE LA COSULTA DONDE INDICA QUE TRABAJADORES SE HAN SELECCIONADO
    If ExisteTablaAux(" [##TMPBOLADEL" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPBOLADEL" & VGL_COMPUTER & "] "
    'TENEMOS QUE PROGRAMAR LA EJECUCION DEL EXECUTE DESDE LA CONCEXION DE DESTINO
    DBSYSTEM.Execute SqlText.Text
    DBSTARPLAN.Execute "EXECUTE  [ASISTMP]  ' [##TMPBOLADEL" & VGL_COMPUTER & "] '"
    With rptBoletas
        .Reset
        .WindowTitle = "PLAN0039 - REPORTE DE BOLETAS DE ADELANTOS - NETOS A PAGAR"
        .ReportFileName = REGSISTEMA.REPORTES & "\PLAN0039.RPT"
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
        .StoredProcParam(0) = " [##TMPBOLADEL" & VGL_COMPUTER & "] "
        .Destination = crptToWindow
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowState = crptMaximized
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "XRUC='" & REGSISTEMA.RUC & "'"
        .Formulas(2) = "XMES='" & Lista.SelectedItem.Text & "'"
        If .Status <> 2 Then .Action = 1
    End With
    Screen.MousePointer = 1
    Exit Sub
    
ERRPRINT:
    MsgBox ERR.Description & ": " & ERR.Number
    Screen.MousePointer = 1
    
End Sub


Private Sub Command1_Click()
CambiaPanelBD True
    If ExisteTablaAux("[##DETADELPER" & VGL_COMPUTER & "]") Then DBSYSTEM.Execute "DROP TABLE  [##DETADELPER" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "SELECT * INTO  [##DETADELPER" & VGL_COMPUTER & "]  FROM DETADEL WHERE NOMBOL=" & Lista.SelectedItem.Tag
    
    If ExisteTablaAux("[##DETADELPER2" & VGL_COMPUTER & "]") Then DBSYSTEM.Execute "DROP TABLE  [##DETADELPER2" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "SELECT TRA.CODTRAB,TRA.APEPAT+' '+ TRA.APEMAT + ' '+ TRA.NOMBRE AS NOMBRES ,TRA.CARGO,TRA.FECHAING,TRA.BASICO,TRA.FONDOPENS INTO  [##DETADELPER2" & VGL_COMPUTER & "]  FROM TRABAJADORES TRA,DETADEL DET WHERE DET.NOMBOL=" & Lista.SelectedItem.Tag & " AND DET.CODTRAB=TRA.CODTRAB"
    DBSYSTEM.Execute "ALTER TABLE [##DETADELPER2" & VGL_COMPUTER & "] ADD  TOTAL  Numeric(20,2)  DEFAULT 0"
    DBSYSTEM.Execute "UPDATE  [##DETADELPER2" & VGL_COMPUTER & "] SET TOTAL=TMP.MONTO  FROM  [##DETADELPER2" & VGL_COMPUTER & "] TMP2 ,[##DETADELPER" & VGL_COMPUTER & "] TMP WHERE TMP2.CODTRAB=TMP.CODTRAB AND CONCEPTO='MONTO'"
    '---------------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------------
    If ExisteTablaAux("[##DETADELPER3" & VGL_COMPUTER & "]") Then DBSYSTEM.Execute "DROP TABLE  [##DETADELPER3" & VGL_COMPUTER & "] "
    If ExisteTablaAux("[##DETADELPER4" & VGL_COMPUTER & "]") Then DBSYSTEM.Execute "DROP TABLE  [##DETADELPER4" & VGL_COMPUTER & "]"
    
    DBSYSTEM.Execute "SELECT  CODTRAB,DCTO=CASE WHEN IE=1 THEN SUM(MONTO) WHEN IE=2 THEN SUM(MONTO*-1) End INTO [##DETADELPER4" & VGL_COMPUTER & "] From [##DETADELPER" & VGL_COMPUTER & "] Where CONCEPTO<>'MONTO' GROUP BY CODTRAB,IE ORDER BY CODTRAB"
    DBSYSTEM.Execute "SELECT CODTRAB,SUM(DCTO)AS DCTO INTO [##DETADELPER3" & VGL_COMPUTER & "] FROM [##DETADELPER4" & VGL_COMPUTER & "] GROUP BY CODTRAB ORDER BY CODTRAB"
    
    DBSYSTEM.Execute "ALTER TABLE [##DETADELPER2" & VGL_COMPUTER & "] ADD  DCTO  Numeric(20,2)  DEFAULT 0"
    DBSYSTEM.Execute "UPDATE  [##DETADELPER2" & VGL_COMPUTER & "] SET DCTO=TMP3.DCTO  FROM  [##DETADELPER2" & VGL_COMPUTER & "] TMP2 ,[##DETADELPER3" & VGL_COMPUTER & "] TMP3 WHERE TMP2.CODTRAB=TMP3.CODTRAB"
'-------------------------------------------------------------------------------------------------
    If ExisteTablaAux("[##DETADELPER5" & VGL_COMPUTER & "]") Then DBSYSTEM.Execute "DROP TABLE  [##DETADELPER5" & VGL_COMPUTER & "]"
    DBSYSTEM.Execute "SELECT TMP.*,AFPS.NOMBRE AS FONDOPENSX INTO [##DETADELPER5" & VGL_COMPUTER & "] FROM [##DETADELPER2" & VGL_COMPUTER & "] TMP,AFPS WHERE TMP.FONDOPENS=AFPS.CODAFP"
    
    DBSTARPLAN.Execute "ASISTMP '[##DETADELPER5" & VGL_COMPUTER & "]'"
    
    
With rptBoletas
        .Reset
        .WindowTitle = "PLAN0001PER - REPORTE DE COOPERATIVAS - QUINCENAL"
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
        .ReportFileName = REGSISTEMA.REPORTES & "REPPER\Coope_Quin.RPT"
        .Destination = crptToWindow
        .StoredProcParam(0) = "[##DETADELPER5" & VGL_COMPUTER & "]" '"[##DETADELPER2" & VGL_COMPUTER & "]"
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "XRUC='" & REGSISTEMA.RUC & "'"

        '.WindowShowPrintBtn = True
        '.WindowShowRefreshBtn = True
        '.WindowShowSearchBtn = True
        '.WindowShowPrintSetupBtn = True
        .WindowState = crptMaximized
        If .Status <> 2 Then .Action = 1
End With
    
CambiaPanelBD False
End Sub

Private Sub DGBOLETAS_HEADCLICK(ByVal COLINDEX As Integer)
    Dim XCOL As String
    XCOL = dgBoletas.Columns(COLINDEX).DataField
    If ITSOPEN Then
        RSLISTA.Close
        RSLISTA.Open "SELECT * FROM  [##_TMPLSTADEL" & VGL_COMPUTER & "]  ORDER BY " & XCOL
        Set dgBoletas.DataSource = RSLISTA
        FORMATEARDG
    End If
End Sub

Private Sub Form_Activate()
    ActivarTools REGACT
End Sub

Private Sub Form_Load()
    Set RSLISTA = New ADODB.Recordset
    If ExisteTablaAux(" [##_TMPLSTADEL" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##_TMPLSTADEL" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "CREATE TABLE  [##_TMPLSTADEL" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8), NOMBRES VARCHAR(100), ADELANTO  Numeric(20,2) , OTROSING  Numeric(20,2) , OTROSEGR  Numeric(20,2) , NETO  Numeric(20,2) , INUMBOL INT, NOMBOL INT, PERIODO VARCHAR(50))"
    RSLISTA.Open " [##_TMPLSTADEL" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic
    ITSOPEN = False
    With REGACT
        .BUSCAR = True
        .EDITAR = False
        .ELIMINAR = True
        .FILTRAR = False
        .IMPRIMIR = True
        .NUEVO = False
        .PRELIMINAR = True
    End With
    CARGARMESES

End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSBOLE = Nothing
    Set RSLISTA = Nothing
End Sub


Public Sub CARGARMESES()
    Dim RSMESES As New ADODB.Recordset
    RSMESES.Open "SELECT DISTINCT MES FROM NOMBOL ORDER BY MES DESC", DBSYSTEM, adOpenKeyset, adLockOptimistic
    If RSMESES.RecordCount = 0 Then
        Set RSMESES = Nothing
        MsgBox "NO EXISTEN MESES PROCESADOS"
        Exit Sub
    End If
    xMeses.Clear
    Do While Not RSMESES.EOF
        xMeses.AddItem Format(Month(RSMESES!MES), "00") & "/" & Year(RSMESES!MES) & " : " & AMESES(Month(RSMESES!MES)) & " DE " & Year(RSMESES!MES)
        RSMESES.MoveNext
    Loop
    Set RSMESES = Nothing
    xMeses.ListIndex = 0
End Sub
Private Sub LISTA_ITEMCHECK(ByVal Item As MSComctlLib.ListItem)
    Dim SNOMBOL As String, STRSQL As String, RSAUX As New ADODB.Recordset
    Dim CATMP As String
    SNOMBOL = Right(Item.KEY, Len(Item.KEY) - 1)
    
    If Item.Checked Then
        Screen.MousePointer = 11
        
        If chkDetadel.Value = 0 Then
            'EL PROCESO CONVENCIONAL
            STRSQL = "INSERT INTO  [##_TMPLSTADEL" & VGL_COMPUTER & "] " _
            & " (CODTRAB,ADELANTO,NOMBOL,PERIODO,INUMBOL)" _
            & " SELECT  CODTRAB, MONTO, ORIGEN AS NOMBOL,'" & Item.Text & "' AS PERIODO,CODIGO" _
            & " FROM ADEL2000 " _
            & " WHERE ORIGEN=" & SNOMBOL
        Else
            'PROCESO QUE SE VALE DE LA TABLA DETADEL
            If ExisteTablaAux("[##TMP_IE1" & VGL_COMPUTER & "]") Then DBSYSTEM.Execute "DROP TABLE [##TMP_IE1" & VGL_COMPUTER & "]"
            If ExisteTablaAux("[##TMP_IE2" & VGL_COMPUTER & "]") Then DBSYSTEM.Execute "DROP TABLE [##TMP_IE2" & VGL_COMPUTER & "]"
            If ExisteTablaAux("[##ADELDETALLES" & VGL_COMPUTER & "]") Then DBSYSTEM.Execute "DROP TABLE [##ADELDETALLES" & VGL_COMPUTER & "]"
            CATMP = "SELECT CODTRAB,SUM(MONTO)as MONTO,NOMBOL,MES into [##TMP_IE1" & VGL_COMPUTER & "] FROM DETADEL WHERE NOMBOL=" & SNOMBOL & "  AND IE=1  " & _
            " GROUP BY CODTRAB,MES,NOMBOL ORDER BY CODTRAB"
            DBSYSTEM.Execute CATMP
            CATMP = "SELECT CODTRAB,SUM(MONTO)as MONTO,NOMBOL,MES into [##TMP_IE2" & VGL_COMPUTER & "] FROM DETADEL WHERE NOMBOL=" & SNOMBOL & " AND IE=2 " & _
            " GROUP BY CODTRAB,MES,NOMBOL ORDER BY CODTRAB"
            DBSYSTEM.Execute CATMP
            CATMP = "select TMP1.CODTRAB,MONTO=(TMP1.MONTO-TMP2.MONTO) ,TMP1.NOMBOL,TMP1.MES  INTO [##ADELDETALLES" & VGL_COMPUTER & "] " & _
            "from [##TMP_IE1" & VGL_COMPUTER & "] TMP1 INNER JOIN [##TMP_IE2" & VGL_COMPUTER & "]  TMP2 ON TMP1.CODTRAB=TMP2.CODTRAB"
            DBSYSTEM.Execute CATMP
            
            STRSQL = "INSERT INTO  [##_TMPLSTADEL" & VGL_COMPUTER & "] " _
            & " (CODTRAB,ADELANTO,NOMBOL,PERIODO,INUMBOL)" _
            & " SELECT  TMP.CODTRAB, TMP.MONTO,TMP.NOMBOL ,'" & Item.Text & "' AS PERIODO,ADE.CODIGO" _
            & " FROM [##ADELDETALLES" & VGL_COMPUTER & "] TMP INNER JOIN ADEL2000 ADE ON TMP.CODTRAB=ADE.CODTRAB " _
            & " WHERE ADE.ORIGEN=" & SNOMBOL
        End If
        DBSYSTEM.Execute STRSQL
        DBSYSTEM.Execute "UPDATE A SET A.NOMBRES = W.NOMBRES FROM  [##_TMPLSTADEL" & VGL_COMPUTER & "]  A, " & REGSISTEMA.BASESQL & ".dbo.VWTRABAJ W WHERE A.CODTRAB = W.CODTRAB "
        RSAUX.Open "SELECT CODTRAB, SUM(MONTO) AS TOT1 FROM PAGOSCTA WHERE TIPO=1 AND TIPOBOLETA='A' AND CODNOMBOL=" & SNOMBOL & " GROUP BY CODTRAB", DBSYSTEM, adOpenStatic, adLockReadOnly
        Do While Not RSAUX.EOF
            DBSYSTEM.Execute "UPDATE  [##_TMPLSTADEL" & VGL_COMPUTER & "]  SET OTROSING=" & RSAUX!Tot1 & " WHERE CODTRAB='" & RSAUX!CODTRAB & "' AND NOMBOL=" & SNOMBOL
            RSAUX.MoveNext
        Loop
        RSAUX.Close
        RSAUX.Open "SELECT CODTRAB, SUM(MONTO) AS TOT1 FROM PAGOSCTA WHERE TIPO=2 AND TIPOBOLETA='A' AND CODNOMBOL=" & SNOMBOL & " GROUP BY CODTRAB", DBSYSTEM, adOpenStatic, adLockReadOnly
        Do While Not RSAUX.EOF
            DBSYSTEM.Execute "UPDATE  [##_TMPLSTADEL" & VGL_COMPUTER & "]  SET OTROSEGR=" & RSAUX!Tot1 & " WHERE CODTRAB='" & RSAUX!CODTRAB & "' AND NOMBOL=" & SNOMBOL
            RSAUX.MoveNext
        Loop
        Set RSAUX = Nothing
        DBSYSTEM.Execute "UPDATE  [##_TMPLSTADEL" & VGL_COMPUTER & "]  SET OTROSING=0 WHERE (OTROSING)IS NULL"
        DBSYSTEM.Execute "UPDATE  [##_TMPLSTADEL" & VGL_COMPUTER & "]  SET OTROSEGR=0 WHERE (OTROSEGR)IS NULL"
        DBSYSTEM.Execute "UPDATE  [##_TMPLSTADEL" & VGL_COMPUTER & "]  SET NETO=ADELANTO-OTROSEGR+OTROSING "
        Screen.MousePointer = 1
    Else
        DBSYSTEM.Execute "DELETE FROM  [##_TMPLSTADEL" & VGL_COMPUTER & "]  WHERE NOMBOL=" & SNOMBOL
    End If
    Set RSLISTA = Nothing
    Set RSLISTA = New ADODB.Recordset
    RSLISTA.Open " [##_TMPLSTADEL" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic
    'RSLISTA.REQUERY
    Set dgBoletas.DataSource = RSLISTA
    FORMATEARDG
    xArea.Text = ""
    xArea.Tag = ""
End Sub

Private Sub LMARCATODOS_Click()
    XMARCATODOS_CLICK
End Sub

Private Sub RSLISTA_MOVECOMPLETE(ByVal ADREASON As ADODB.EventReasonEnum, ByVal PERROR As ADODB.Error, ADSTATUS As ADODB.EventStatusEnum, ByVal PRECORDSET As ADODB.Recordset)
    On Error GoTo ERRMOVE
    If RSLISTA.EOF Then
        cmPagosBanco.Enabled = False
        cmBilletes.Enabled = False
        cmPrtUno.Enabled = False
        cmPrtTodos.Enabled = False
    Else
        cmPagosBanco.Enabled = True
        cmBilletes.Enabled = True
        cmPrtUno.Enabled = True
        cmPrtTodos.Enabled = True
    End If
    If ADREASON = adRsnMove Then
        If RSLISTA!Neto < 0 Then xError.Visible = True Else xError.Visible = False
    End If
    Exit Sub
ERRMOVE:
    Resume Next
End Sub

Private Sub SEL1_Click(INDEX As Integer)
    xArea.Text = ""
    xArea.Tag = ""
End Sub
Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case Trim(UCase(ButtonMenu.KEY))
        Case UCase("EliminaTodasBol")
            Dim MON As Integer
            Dim RS As New ADODB.Recordset
            If RSLISTA.EOF Then
                MsgBox "NO SE HAN ENCONTRADO BOLETAS DE ADELANTOS PARA ELIMINAR", vbCritical
                Exit Sub
            End If
            If MsgBox("Realmente desea eliminar todos los registros de adelantos", vbYesNo + vbInformation) = vbYes Then
                RSLISTA.MoveFirst
                Do While Not RSLISTA.EOF
                    DBSYSTEM.Execute "DELETE FROM " & REGSISTEMA.TABLAADEL & " WHERE CODIGO=" & RSLISTA!INUMBOL
                    DBSYSTEM.Execute "DELETE FROM  [##_TMPLSTADEL" & VGL_COMPUTER & "]  WHERE INUMBOL=" & RSLISTA!INUMBOL
                    
                    'ELIMINANDO LOS MOVIMIENTOS DE LOS CONCEPTOS QUE CORRESPONDEN A LOS ADELANTOS DETALLADOS
                    DBSYSTEM.Execute "" & _
                    " DELETE INGMOV2000 FROM  INGMOV2000 A,DETADEL B " & _
                    " WHERE A.CODNOMBOL=B.NOMBOL AND A.CODTRAB=B.CODTRAB AND " & _
                    " A.CONCEPTO=B.CODCONCEP AND " & _
                    " A.CODTRAB='" & RSLISTA!CODTRAB & "'  AND " & _
                    "  A.CODNOMBOL = " & RSLISTA!NOMBOL
                    
                    'ELIMINANDO LOS ADELANTOS DETALLADOS
                    DBSYSTEM.Execute "DELETE FROM DETADEL WHERE CODTRAB='" & RSLISTA!CODTRAB & "' AND NOMBOL=" & RSLISTA!NOMBOL
                    'ELIMINANDO LOS MOVIMIENTOS QUE SE ENCUENTREN EN ADELANTOS
                    Set RS = New ADODB.Recordset
                    RS.Open "SELECT * FROM MOVICTA M,PAGOSCTA P WHERE M.CODMOV=P.CODMOV AND P.CODNOMBOL=" & RSLISTA!NOMBOL & " AND P.CODTRAB='" & RSLISTA!CODTRAB & "' AND ( M.PORCQUINC <> 0 OR PROGRAMADO=1 )", DBSYSTEM
                    If RS.RecordCount > 0 Then
                        RS.MoveFirst
                        Do While Not RS.EOF
                            DBSYSTEM.Execute "DELETE FROM PAGOSCTA WHERE CODMOV=" & RS.Fields("CODMOV").Value & " AND TIPOBOLETA='A'"
                            MON = DevuelveValor("SELECT MONEDA FROM MOVICTA WHERE CODMOV=" & RS.Fields("CODMOV"), DBSYSTEM)
                            Call ACTSALDO(RS("CODMOV"), MON)
                            RS.MoveNext
                        Loop
                    End If
                    RSLISTA.MoveNext
                Loop
                RSLISTA.Requery
                FORMATEARDG
            End If
    End Select
End Sub

Private Sub XAREA_DblClick()
    Dim RSAUX As ADODB.Recordset
    Set RSAUX = New ADODB.Recordset
    If Sel1(0).Value Then 'SI ES POR AREAS
        RSAUX.Open "SELECT CODCCOSTO, NOMBRE FROM AREASTRAB ORDER BY CODCCOSTO", DBSYSTEM, adOpenStatic
    Else
        RSAUX.Open "SELECT CODCCOSTO, NOMBRE FROM CCOSTOS ORDER BY CODCCOSTO", DBSYSTEM, adOpenStatic
    End If
    If RSAUX.RecordCount = 0 Then
        MsgBox "NO SE HAN ENCONTRADO REGISTROS DE AREAS DE TRABAJO/CENTRO DE COSTO", vbCritical
        Set RSAUX = Nothing
        Exit Sub
    End If
    frmComun.CONECTAR RSAUX
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xArea.Text = RSAUX!CODCCOSTO & ": " & RSAUX!NOMBRE
        xArea.Tag = RSAUX!CODCCOSTO
        If ExisteTablaAux(" [##BOLSXAREA" & VGL_COMPUTER & "] ") Then
            DBSYSTEM.Execute "DROP TABLE  [##BOLSXAREA" & VGL_COMPUTER & "] "
        End If
        If Sel1(0).Value Then 'SI ES POR AREAS DE TRABAJO
            DBSYSTEM.Execute "SELECT CODTRAB INTO  [##BOLSXAREA" & VGL_COMPUTER & "]  FROM VWTRABAJ WHERE CODAREA LIKE '" & xArea.Tag & "%'"
        Else
            DBSYSTEM.Execute "SELECT CODTRAB INTO  [##BOLSXAREA" & VGL_COMPUTER & "]  FROM VWTRABAJ WHERE CODCCOSTO LIKE '" & xArea.Tag & "%'"
        End If
        frWait.Show 1
        DBSYSTEM.Execute "DELETE FROM  [##_TMPLSTADEL" & VGL_COMPUTER & "]  WHERE CODTRAB NOT IN (SELECT CODTRAB FROM  [##BOLSXAREA" & VGL_COMPUTER & "] )"
        RSLISTA.Requery
        FORMATEARDG
    End If
    Set RSAUX = Nothing
End Sub

Private Sub XAREA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        xArea.Text = ""
        xArea.Tag = ""
    End If
End Sub

Private Sub XMARCATODOS_CLICK()
    Dim XITEM As ListItem
    For Each XITEM In Lista.ListItems
        If Not XITEM.Checked Then
            XITEM.Checked = True
            LISTA_ITEMCHECK XITEM
        End If
    Next
End Sub

Private Sub XMESES_Click()
    If xMeses.ListIndex = -1 Then Exit Sub
    Dim sMes As String
    If ITSOPEN Then
        RSBOLE.Close
    End If
    ITSOPEN = True
    sMes = "01/" & Left(xMeses.Text, 2) & "/" & Mid(xMeses.Text, 4, 4)
    RSBOLE.Open "SELECT * FROM NOMBOL WHERE MES=" & DateSQL(sMes) & " ORDER BY FECHAINI, NOMBOL.NOMBRE", DBSYSTEM, adOpenKeyset
    If RSBOLE.RecordCount = 0 Then Exit Sub
    Dim xLista As ListItem
    RSBOLE.MoveFirst
    Lista.ListItems.Clear
    Do While Not RSBOLE.EOF
        Set xLista = Lista.ListItems.Add(, "C" & RSBOLE!Codigo, "ADELANTO: " & RSBOLE!NOMBRE, , 1)
        xLista.SubItems(1) = RSBOLE!FECHAINI
        xLista.SubItems(2) = RSBOLE!FECHAFIN
        xLista.Tag = RSBOLE!Codigo
        RSBOLE.MoveNext
    Loop
    RSBOLE.MoveFirst
    If ExisteTablaAux(" [##_TMPLSTADEL" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DELETE FROM  [##_TMPLSTADEL" & VGL_COMPUTER & "] "
    RSLISTA.Requery
    Set dgBoletas.DataSource = RSLISTA
    FORMATEARDG
End Sub

Public Sub FORMATEARDG()
    Dim RSAUX As New ADODB.Recordset
    RSAUX.Open "SELECT SUM(ADELANTO) AS TOTAL1,SUM(NETO) AS TOTAL3, SUM(OTROSING) AS TOTAL4, SUM(OTROSEGR) AS TOTAL5 FROM  [##_TMPLSTADEL" & VGL_COMPUTER & "] ", DBSYSTEM
    xSumIng.Caption = Format(IIf(IsNull(RSAUX!Total1), 0, RSAUX!Total1), "##,##0.00 ")
    xSumNet.Caption = Format(IIf(IsNull(RSAUX!TOTAL3), 0, RSAUX!TOTAL3), "##,##0.00 ")
    sum1.Caption = Format(IIf(IsNull(RSAUX!TOTAL4), 0, RSAUX!TOTAL4), "##,##0.00 ")
    sum2.Caption = Format(IIf(IsNull(RSAUX!TOTAL5), 0, RSAUX!TOTAL5), "##,##0.00 ")
    xCont.Caption = RSLISTA.RecordCount & " REGISTROS"
    Set RSAUX = Nothing
End Sub

Public Sub COMANDOTOOLBAR(COMANDO As String)
    Select Case UCase(COMANDO)
        Case "BUSCAR"
            If RSLISTA.EOF Then Exit Sub
            Dim RSLISTA2 As New ADODB.Recordset
            Set RSLISTA2 = RSLISTA.Clone
            frmComun.CONECTAR RSLISTA2
            frmComun.Show 1
            If VGUTIL(1) <> "" Then
                RSLISTA.MoveFirst
                RSLISTA.FIND "CODTRAB='" & VGUTIL(1) & "'"
            End If
            Set RSLISTA2 = Nothing
        Case "IMPRIMIR", "PRELIMINAR"
            If MsgBox("Desea Imprimir el Normal(Si) o el Detallado(No)", vbYesNo + vbQuestion) = vbYes Then
                With rptBoletas
                    .Reset
                    .WindowTitle = "PLAN0041 - REPORTE DE ADELANTOS DE REMUNERACIONES - NETOS A PAGAR"
                    .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
                    .ReportFileName = REGSISTEMA.REPORTES & "PLAN0041.RPT"
                    .Destination = crptToWindow
                    .StoredProcParam(0) = " [##_TMPLSTADEL" & VGL_COMPUTER & "] "
                    .WindowShowPrintBtn = True
                    .WindowShowRefreshBtn = True
                    .WindowShowSearchBtn = True
                    .WindowShowPrintSetupBtn = True
                    .WindowState = crptMaximized
                    .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
                    .Formulas(1) = "XRUC='RUC N° " & REGSISTEMA.RUC & "'"
                    .Formulas(2) = "''"
                    If .Status <> 2 Then .Action = 1
                End With
              Else
                If ExisteTablaAux("[##DETADEL" & VGL_COMPUTER & "]") Then DBSYSTEM.Execute "DROP TABLE  [##DETADEL" & VGL_COMPUTER & "] "
                DBSYSTEM.Execute "SELECT * INTO  [##DETADEL" & VGL_COMPUTER & "]  FROM DETADEL WHERE NOMBOL=" & Lista.SelectedItem.Tag
                If ExisteTablaAux(" [##TMPTABLA" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPTABLA" & VGL_COMPUTER & "] "
                DBSYSTEM.Execute "CREATE TABLE  [##TMPTABLA" & VGL_COMPUTER & "] (COD INT)"
                With rptBoletas
                    .Reset
                    .WindowTitle = "PLAN0091AUX - REPORTE DE ADELANTOS DE REMUNERACIONES - NETOS A PAGAR"
                    .Connect = "DSN=" & VGL_SERVER & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL & ""
                    .ReportFileName = REGSISTEMA.REPORTES & "PLAN0091AUX.RPT"
                    .Destination = crptToWindow
                    .StoredProcParam(0) = " [##TMPTABLA" & VGL_COMPUTER & "] "
                    .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
                    .Formulas(1) = "XRUC='" & REGSISTEMA.RUC & "'"
                    .Formulas(2) = "''"
                    .WindowShowPrintBtn = True
                    .WindowShowRefreshBtn = True
                    .WindowShowSearchBtn = True
                    .WindowShowPrintSetupBtn = True
                    .WindowState = crptMaximized
                    .SubreportToChange = "PLAN0091.rpt"
                    .Connect = "DSN=" & VGL_SERVER & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL & ""
                    .ParameterFields(0) = "@TABLATMP; [##DETADEL" & VGL_COMPUTER & "] ;TRUE"
                    '.Formulas(0) = "":  .Formulas(1) = "": .Formulas(2) = ""
                    If .Status <> 2 Then .Action = 1
                End With
            End If
        Case "ELIMINAR"
            If RSLISTA.EOF Then
                MsgBox "NO SE HAN ENCONTRADO BOLETAS DE ADELANTOS PARA ELIMINAR", vbCritical
                Exit Sub
            End If
            If MsgBox("REALMENTE DESEA ELIMINAR EL REGISTRO DE ADELANTO DE " & RSLISTA!NOMBRES, vbYesNo + vbInformation) = vbYes Then
                'If MsgBox("ADVERTENCIA SI BORRA EL REGISTRO SE ELIMINA LOS MOVIMIENTOS DE CUENTA CORRIENTE ENLAZADOS A ESTE ADELANTO", vbExclamation + vbOKCancel) = vbCancel Then Exit Sub
                DBSYSTEM.Execute "DELETE FROM " & REGSISTEMA.TABLAADEL & " WHERE CODIGO=" & RSLISTA!INUMBOL
                DBSYSTEM.Execute "DELETE FROM  [##_TMPLSTADEL" & VGL_COMPUTER & "]  WHERE INUMBOL=" & RSLISTA!INUMBOL
                
                'ELIMINANDO LOS MOVIMIENTOS DE LOS CONCEPTOS QUE CORRESPONDEN A LOS ADELANTOS DETALLADOS
                DBSYSTEM.Execute "" & _
                " DELETE INGMOV2000 FROM  INGMOV2000 A,DETADEL B " & _
                " WHERE A.CODNOMBOL=B.NOMBOL AND A.CODTRAB=B.CODTRAB AND " & _
                " A.CONCEPTO=B.CODCONCEP AND " & _
                " A.CODTRAB='" & RSLISTA!CODTRAB & "'  AND " & _
                "  A.CODNOMBOL = " & RSLISTA!NOMBOL
                
                'ELIMINANDO LOS ADELANTOS DETALLADOS
                DBSYSTEM.Execute "DELETE FROM DETADEL WHERE CODTRAB='" & RSLISTA!CODTRAB & "' AND NOMBOL=" & RSLISTA!NOMBOL
                
                Dim MON As Integer
                'ELIMINANDO LOS MOVIMIENTOS QUE SE ENCUENTREN EN ADELANTOS
                Dim RS As New ADODB.Recordset
                RS.Open "SELECT * FROM MOVICTA M,PAGOSCTA P WHERE M.CODMOV=P.CODMOV AND P.CODNOMBOL=" & RSLISTA!NOMBOL & " AND P.CODTRAB='" & RSLISTA!CODTRAB & "' AND ( M.PORCQUINC <> 0 OR PROGRAMADO=1 )", DBSYSTEM
                If RS.RecordCount > 0 Then
                    RS.MoveFirst
                    Do While Not RS.EOF
                        DBSYSTEM.Execute "DELETE FROM PAGOSCTA WHERE CODMOV=" & Str(RS.Fields("CODMOV").Value & " AND TIPOBOLETA='A'")
                        MON = DevuelveValor("SELECT MONEDA FROM MOVICTA WHERE CODMOV=" & RS.Fields("CODMOV"), DBSYSTEM)
                        Call ACTSALDO(RS("CODMOV"), MON)
                        RS.MoveNext
                    Loop
                End If
                
                RSLISTA.Requery
                FORMATEARDG
            End If
    End Select
End Sub

Public Sub CARGABOL()
    Dim RSAUX As New ADODB.Recordset
    Dim RSBOL As New ADODB.Recordset
    If RSLISTA.EOF Then Exit Sub
    If ExisteTabla(" [##TMPTRANS" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPTRANS" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "CREATE TABLE  [##TMPTRANS" & VGL_COMPUTER & "]  (CODIGO VARCHAR(6), DESCRIPCION VARCHAR(30), VALOR  Numeric(20,2) , TIPO BIT, ENLACE VARCHAR(8), FILA INT)"
    'JALAR LOS CONCEPTOS
    DBSYSTEM.Execute "INSERT INTO  [##TMPTRANS" & VGL_COMPUTER & "]  (CODIGO,DESCRIPCION,VALOR,TIPO,ENLACE,FILA) SELECT MOVICTA.CODMOV, DESCRIPCION, MONTO, 1 AS TIPO, ' ' AS ENLACE,11 AS FILA FROM MOVICTA, PAGOSCTA WHERE MOVICTA.CODMOV=PAGOSCTA.CODMOV AND MOVICTA.TIPOGRUPO=1 AND PAGOSCTA.CODTRAB='" & RSLISTA!CODTRAB & "' AND CODNOMBOL=" & RSLISTA!NOMBOL
    'JALAR LOS OTROS INGRESOS (CUENTAS CORRIENTES)
    DBSYSTEM.Execute "INSERT INTO  [##TMPTRANS" & VGL_COMPUTER & "]  SELECT CONCEPTOS.CODIGO, CONCEPTOS.NOMBRE AS DESCRIPCION, MONTO AS VALOR, CONCEPTOS.TIPO, CONCEPTOS.ENLACE, FILA FROM MOV" & Left(xMeses.Text, 2) & Mid(xMeses.Text, 4, 4) & " MOV, CONCEPTOS WHERE MOV.CONCEPTO=CONCEPTOS.CODIGO AND INUMBOL=" & RSLISTA!INUMBOL
    'JALAR LOS ADELANTOS DE PAGO
    DBSYSTEM.Execute "INSERT INTO  [##TMPTRANS" & VGL_COMPUTER & "]  SELECT CODIGO, '<ADELANTO DE PAGO>' AS DESCRIPCION,MONTO AS VALOR,2 AS TIPO,' ' AS ENLACE, 4 AS FILA FROM " & REGSISTEMA.TABLAADEL & " WHERE NOMBOL=" & RSLISTA!NOMBOL & " AND CODTRAB='" & RSLISTA!CODTRAB & "'"
    'JALAR LOS OTROS EGRESOS (CUENTAS CORRIENTES)
    DBSYSTEM.Execute "INSERT INTO  [##TMPTRANS" & VGL_COMPUTER & "]  (CODIGO,DESCRIPCION,VALOR,TIPO,ENLACE,FILA) SELECT MOVICTA.CODMOV, DESCRIPCION, MONTO, 2 AS TIPO, ' ' AS ENLACE,12 AS FILA FROM MOVICTA, PAGOSCTA WHERE MOVICTA.CODMOV=PAGOSCTA.CODMOV AND MOVICTA.TIPOGRUPO=2 AND PAGOSCTA.CODTRAB='" & RSLISTA!CODTRAB & "' AND CODNOMBOL=" & RSLISTA!NOMBOL
    RSAUX.Open "SELECT * FROM VWTRABAJ WHERE CODTRAB='" & RSLISTA!CODTRAB & "'", DBSYSTEM, adOpenStatic
    RSBOL.Open "RPTBOLETAS", DBSTARPLAN, adOpenKeyset, adLockOptimistic
    'ADICION DE LA BOLETA
    RSBOL.AddNew
    RSBOL!CODTRAB = RSLISTA!CODTRAB
    RSBOL!NOMBRES = RSAUX!NOMBRES
    RSBOL!CENTROCOSTO = RSAUX!CENTRO
    RSBOL!FECHAING = RSAUX!FECHAING
    RSBOL!PERIODO = RSLISTA!PERIODO
    RSBOL!TOTING = RSLISTA!INGRESOS
    RSBOL!TOTEGR = RSLISTA!EGRESOS
    RSBOL!BASICO = RSAUX!BASICO
    RSBOL!AFP = RSAUX!FONDOPENS & "-" & RSAUX!NOMBREAFP
    RSBOL!CARGO = RSAUX!CARGO
    RSBOL!FECHAING = RSAUX!FECHAING
    RSBOL!DOCUMENTO = RSAUX!TIPDOC & "-" & RSAUX!DOCIDEN
    RSBOL!CARNETSEG = RSAUX!CARNETSEG
    RSBOL!CUENTABANCO = RSAUX!BANCO & "-" & RSAUX!CTABANCO
    Dim IND1 As Byte, IND2 As Byte, IND3 As Byte, IND4 As Byte
    Set RSAUX = Nothing
    RSAUX.Open "SELECT * FROM  [##TMPTRANS" & VGL_COMPUTER & "]  ORDER BY TIPO, FILA", DBSYSTEM, adOpenStatic
    IND1 = 0
    IND2 = 0
    IND3 = 0
    IND4 = 0
    Do While Not RSAUX.EOF
        Select Case RSAUX!TIPO
            Case 0: IND1 = IND1 + 1
                    RSBOL.Fields("C" & IND1).Value = RSAUX!DESCRIPCION
                    RSBOL.Fields("INF" & IND1).Value = RSAUX!VALOR
            Case 1: IND1 = IND1 + 1
                    RSBOL.Fields("C" & IND1).Value = RSAUX!DESCRIPCION
                    RSBOL.Fields("I" & IND1).Value = RSAUX!VALOR
            Case 2: IND3 = IND3 + 1
                    RSBOL.Fields("R" & IND3).Value = RSAUX!DESCRIPCION
                    RSBOL.Fields("E" & IND3).Value = RSAUX!VALOR
            Case 3: IND4 = IND4 + 1
                    RSBOL.Fields("G" & IND4).Value = RSAUX!DESCRIPCION
                    RSBOL.Fields("A" & IND4).Value = RSAUX!VALOR
        End Select
        RSAUX.MoveNext
    Loop
    RSBOL.Update
    Set RSAUX = Nothing
    Set RSBOL = Nothing
End Sub


