VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frPlanBol 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planilla de Pago - Por Boletas de Remuneraciones"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9105
   FillStyle       =   2  'Horizontal Line
   Icon            =   "frPlanBol.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdPlanN 
      Caption         =   "Planilla &Normal"
      Height          =   375
      Left            =   7170
      TabIndex        =   23
      Top             =   3375
      Width           =   1755
   End
   Begin Crystal.CrystalReport Reporte 
      Left            =   3795
      Top             =   2430
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      Caption         =   "Factura"
      Height          =   1545
      Left            =   180
      TabIndex        =   11
      Top             =   4710
      Width           =   8835
      Begin AplisetControlText.Aplitext xconcepto 
         Height          =   285
         Left            =   1245
         TabIndex        =   18
         Top             =   1155
         Width           =   4485
         _ExtentX        =   7911
         _ExtentY        =   503
         Text            =   ""
      End
      Begin VB.TextBox xComentario 
         Height          =   975
         Left            =   5805
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   465
         Width           =   2880
      End
      Begin MSComCtl2.DTPicker xFecha 
         Height          =   300
         Left            =   4380
         TabIndex        =   17
         Top             =   810
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   529
         _Version        =   393216
         Format          =   61997057
         CurrentDate     =   36719
      End
      Begin AplisetControlText.Aplitext xFactura 
         Height          =   285
         Left            =   1245
         TabIndex        =   15
         Top             =   810
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   503
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xRazon 
         Height          =   285
         Left            =   1245
         TabIndex        =   13
         Top             =   465
         Width           =   4485
         _ExtentX        =   7911
         _ExtentY        =   503
         Text            =   ""
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Concepto"
         Height          =   195
         Left            =   135
         TabIndex        =   21
         Top             =   1185
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Comentario (Al final)"
         Height          =   195
         Left            =   5820
         TabIndex        =   19
         Top             =   225
         Width           =   1395
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Fact."
         Height          =   195
         Left            =   3240
         TabIndex        =   16
         Top             =   855
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Factura N°"
         Height          =   195
         Left            =   150
         TabIndex        =   14
         Top             =   870
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Razón Social"
         Height          =   195
         Left            =   135
         TabIndex        =   12
         Top             =   510
         Width           =   945
      End
   End
   Begin VB.CommandButton cmCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   7170
      TabIndex        =   10
      Top             =   4275
      Width           =   1755
   End
   Begin VB.CommandButton cmImprimir 
      Caption         =   "&Imprimir Text"
      Height          =   360
      Left            =   7155
      TabIndex        =   9
      Top             =   2700
      Visible         =   0   'False
      Width           =   1755
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4260
      Top             =   2505
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlanBol.frx":030A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Formulas de Facturación (Util para empresas prestadoras de Servicios de Personal)"
      Height          =   2505
      Left            =   180
      TabIndex        =   2
      Top             =   2145
      Width           =   6825
      Begin VB.CommandButton cmEliminar 
         Caption         =   "Eliminar &Formula"
         Height          =   375
         Left            =   3885
         TabIndex        =   8
         Top             =   1980
         Width           =   1335
      End
      Begin VB.CommandButton cmEditar 
         Caption         =   "&Editar Formula"
         Height          =   375
         Left            =   2415
         TabIndex        =   7
         Top             =   1980
         Width           =   1335
      End
      Begin VB.CommandButton cmNueva 
         Caption         =   "&Nueva Formula"
         Height          =   375
         Left            =   945
         TabIndex        =   6
         Top             =   1980
         Width           =   1335
      End
      Begin VB.CommandButton cmQuitar 
         Caption         =   "&Quitar"
         Height          =   375
         Left            =   5670
         TabIndex        =   5
         Top             =   870
         Width           =   1020
      End
      Begin VB.CommandButton cmAgregar 
         Caption         =   "&Agregar"
         Height          =   375
         Left            =   5670
         TabIndex        =   4
         Top             =   390
         Width           =   1020
      End
      Begin MSComctlLib.ListView Lista 
         Height          =   1545
         Left            =   165
         TabIndex        =   3
         Top             =   375
         Width           =   5430
         _ExtentX        =   9578
         _ExtentY        =   2725
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
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
            Text            =   "Codigo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   3951
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Valor"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.CommandButton cmOcultar 
      Caption         =   "&Ocultar Columna"
      Height          =   360
      Left            =   7170
      TabIndex        =   1
      Top             =   2250
      Width           =   1755
   End
   Begin MSDataGridLib.DataGrid DataPlan 
      Height          =   1935
      Left            =   180
      TabIndex        =   0
      Top             =   135
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   3413
      _Version        =   393216
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
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Planilla &Facturación"
      Height          =   375
      Left            =   7170
      TabIndex        =   22
      Top             =   3795
      Width           =   1755
   End
End
Attribute VB_Name = "frPlanBol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSPLAN As ADODB.Recordset
Private Sub CMAGREGAR_CLICK()
    Dim RSAUX As New ADODB.Recordset
    Dim XITEM As ListItem
    RSAUX.Open "FORMFACT", DBSYSTEM, adOpenStatic
    Dim CAD As String
    
    frmComun.CONECTAR RSAUX
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        Set XITEM = Lista.ListItems.Add(, , RSAUX!Codigo, , 1)
        XITEM.SubItems(1) = RSAUX!NOMBRE
        CAD = RSAUX!FORMULA
        RSAUX.Close
        On Error GoTo ERRFORMULA
        RSAUX.Open "SELECT SUM(" & CAD & ") AS TOTAL1 FROM  [##PRTCOSTOS" & VGL_COMPUTER & "] ", DBSTARPLAN, adOpenStatic
        If RSAUX.RecordCount = 0 Then
            MsgBox "Error de USUARIO: Error en la formula de acción", vbInformation
        Else
            XITEM.SubItems(2) = Format(RSAUX!Total1, "0.00")
        End If
    End If
    Set RSAUX = Nothing
    Exit Sub
ERRFORMULA:
    MsgBox "Error de USUARIO: Error en la formula de acción. Sucede en dos casos: los rubros referenciados no existen en este formato o está el algoritmo con errores", vbCritical
    Set RSAUX = Nothing
    Exit Sub
End Sub

Private Sub cmCerrar_Click()
    Unload Me
End Sub

Private Sub CmdImprimir_Click()
    Screen.MousePointer = 11
    Call CREARPLAN
    'Creacion del sub Reporte
    If ExisteTablaAux(" [##TMPCREGLO" & VGL_COMPUTER & "] ") Then DBSTARPLAN.Execute "DROP TABLE  [##TMPCREGLO" & VGL_COMPUTER & "] "
    DBSTARPLAN.Execute "CREATE TABLE  [##TMPCREGLO" & VGL_COMPUTER & "]  (CODIGO VARCHAR(20),NOMBRE VARCHAR(80),TOTAL  Numeric(20,2) )"
    If Lista.ListItems.Count > 0 Then
        Dim XITEM As ListItem
        For Each XITEM In Lista.ListItems
           DBSYSTEM.Execute "INSERT INTO  [##TMPCREGLO" & VGL_COMPUTER & "]  VALUES('" & Trim(XITEM.Text) & "','" & Trim(XITEM.SubItems(1)) & "'," & XITEM.SubItems(2) & ")"
        Next
    End If
    
    
    DBSTARPLAN.Execute "ARMAR_PLANILLA_NORMAL '" & VGL_COMPUTER & "'"
    With Reporte
        .Reset
        .WindowTitle = "PLAN0060.RPT -" & DataPlan.Caption
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0060.RPT"
        .Connect = "DSN=" & VGL_SERVER & ";User=sa;PWD=;DSQ=" & VGL_BASE & ""
        .StoredProcParam(0) = VGL_COMPUTER
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .Formulas(0) = "xCabeza='Planilla : " & frBolEmit.Lista.SelectedItem.Text & "'"
        .Formulas(1) = "xCliente='" & xRazon.Text & "'"
        .Formulas(2) = "xCorresp='" & xConcepto.Text & "'"
        .Formulas(3) = "xFecha=Date(" & Str(Year(xFecha)) & "," & Str(Month(xFecha)) & "," & Str(Day(xFecha)) & ")"
        .Formulas(4) = "xComentario='" & xComentario.Text & "'"
        .SubreportToChange = .GetNthSubreportName(0)
        .Connect = "DSN=" & VGL_SERVER & ";User=sa;PWD=;DSQ=" & VGL_BASE & ""
        .ParameterFields(0) = "@TABLATMP; [##TMPCREGLO" & VGL_COMPUTER & "] ;TRUE"
        '.StoredProcParam(0) = " [##TMPCREGLO"  & VGL_COMPUTER & "] "
        If .Status <> 2 Then .Action = 1
    End With
    Screen.MousePointer = 1
End Sub

Private Sub CmdPlanN_Click()
    Dim REG As Long
    CambiaPanelBD True
    Screen.MousePointer = 11
    Call CREARPLAN(REG)
    DBSTARPLAN.Execute "ARMAR_PLANILLA_NORMAL '" & VGL_COMPUTER & "'"
    With Reporte
        .Reset
        .Connect = "DSN=" & VGL_SERVER & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL & ""
        .WindowTitle = "PLAN0061.RPT -" & DataPlan.Caption
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0061.RPT"
        .StoredProcParam(0) = VGL_COMPUTER
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .Formulas(0) = "xCabeza='Planilla : " & frBolEmit.Lista.SelectedItem.Text & "'"
        .Formulas(1) = "xEmpresa='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(2) = "xRuc='" & REGSISTEMA.RUC & "'"
        .Formulas(3) = "xReg='" & Str(REG) & "'"
        If .Status <> 2 Then .Action = 1
    End With
    Screen.MousePointer = 1
    CambiaPanelBD False
End Sub

Private Sub CMEDITAR_CLICK()
    If Lista.ListItems.Count = 0 Then
        MsgBox "No existen registros para ser editados", vbCritical
        cmAgregar.SetFocus
        Exit Sub
    End If
    VPTAREA = "Editar"
    Load frEForFact
    frEForFact.xCodigo.Text = "" & Lista.SelectedItem.Text
    frEForFact.xNombre.Text = "" & Lista.SelectedItem.SubItems(1)
    frEForFact.xFormula.Text = "" & DevuelveValor("SELECT Formula FROM FormFact WHERE Codigo='" & Lista.SelectedItem.Text & "'", DBSYSTEM)
    frEForFact.Show 1
End Sub

Private Sub CMELIMINAR_CLICK()
    If Lista.ListItems.Count = 0 Then
        MsgBox "No existen registros para ser eliminados", vbCritical
        cmAgregar.SetFocus
        Exit Sub
    End If
    If MsgBox("Seguro de eliminar el registro seleccionado. Los cambios no se podrán deshacer", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    DBSYSTEM.Execute "DELETE FROM FORMFACT WHERE CODIGO='" & Lista.SelectedItem.Text & "'"
    MsgBox "Registro eliminado", vbInformation
End Sub

Private Sub CMIMPRIMIR_CLICK()
    Screen.MousePointer = 11
    If ExisteTablaAux(" [##PRTPLANILLA" & VGL_COMPUTER & "] ") Then DBSTARPLAN.Execute "DROP TABLE  [##PRTPLANILLA" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "CREATE TABLE  [##PRTPLANILLA" & VGL_COMPUTER & "]  (Fila VarChar(250))"
    RSPLAN.MoveFirst
    Dim X As Integer, CAD As String, XSTR As String, XNUMCHAR As Integer, xValor As Variant, XN As Byte, NCOUNT As Integer
    Dim RSAUX As ADODB.Recordset, CADTABLA As String
    Set RSAUX = New ADODB.Recordset
    
    'Proceso del encabezado de las planillas
    '---------------------------------------
    Set RSAUX = New ADODB.Recordset
    XSTR = " "
    For X = 0 To DataPlan.Columns.Count - 1
        If DataPlan.Columns(X).Visible Then
            Select Case RSPLAN.Fields(DataPlan.Columns(X).Caption).Type
                    Case adDate 'Si es por fecha
                        XNUMCHAR = 11
                    Case adSingle 'Si es numero simple
                        XNUMCHAR = 10
                    Case adVarChar
                        XNUMCHAR = RSPLAN.Fields(DataPlan.Columns(X).Caption).DefinedSize + 1
                    Case adUnsignedTinyInt
                        XNUMCHAR = 4
            End Select
            CAD = DataPlan.Columns(X).Caption
            XSTR = XSTR & Left(CAD & String(XNUMCHAR, " "), XNUMCHAR)
        End If
    Next
    CAD = XSTR
    Set RSAUX = Nothing
    NCOUNT = 0
    Screen.MousePointer = 11
    'Generacion de la cadena de impresion
    '------------------------------------
    Do While Not RSPLAN.EOF
        XSTR = ""
        For X = 0 To DataPlan.Columns.Count - 1
            xValor = RSPLAN.Fields(DataPlan.Columns(X).Caption).Value
            If DataPlan.Columns(X).Visible Then
                Select Case RSPLAN.Fields(DataPlan.Columns(X).Caption).Type
                    Case adDate 'Si es por fecha
                        If IsNull(xValor) Then xValor = "  /  /    "
                        XSTR = XSTR & " " & xValor
                    Case adSingle 'Si es numero simple
                        XSTR = XSTR & " " & Right("         " & Format$(xValor, "0.00"), 9)
                    Case adVarChar
                        XN = RSPLAN.Fields(DataPlan.Columns(X).Caption).DefinedSize
                        XSTR = XSTR & " " & Left(xValor & String(XN, " "), XN)
                    Case adUnsignedTinyInt
                        XSTR = XSTR & " " & Format(xValor, "00")
                    Case Else
                        'MsgBox "TIPO no encontrado: " & RSPLAN.Fields(DataPlan.Columns(X).Caption).Type, vbCritical
                End Select
            End If
        Next
        NCOUNT = NCOUNT + 1
        If NCOUNT = 1 Then
            DBSTARPLAN.Execute "INSERT INTO  [##PRTPLANILLA" & VGL_COMPUTER & "]  VALUES ('" & String(Len(XSTR), "_") & "')"
            'Incluir aqui la cabezar
            DBSTARPLAN.Execute "INSERT INTO  [##PRTPLANILLA" & VGL_COMPUTER & "]  VALUES ('" & CAD & "')"
            DBSTARPLAN.Execute "INSERT INTO  [##PRTPLANILLA" & VGL_COMPUTER & "]  VALUES ('" & String(Len(XSTR), "_") & "')"
        End If
        DBSTARPLAN.Execute "INSERT INTO  [##PRTPLANILLA" & VGL_COMPUTER & "]  VALUES ('" & XSTR & "')"
        RSPLAN.MoveNext
    Loop
    CAD = String(Len(XSTR), "_")
    DBSTARPLAN.Execute "INSERT INTO  [##PRTPLANILLA" & VGL_COMPUTER & "]  VALUES ('" & CAD & "')"
    
    'Proceso del Cálculo de TOTALes de la planilla
    '---------------------------------------------
    Set RSAUX = New ADODB.Recordset
    XSTR = ""
    For X = 0 To DataPlan.Columns.Count - 1
        If DataPlan.Columns(X).Visible Then
            Select Case RSPLAN.Fields(DataPlan.Columns(X).Caption).Type
                    Case adDate 'Si es por fecha
                        XNUMCHAR = 11
                    Case adSingle 'Si es numero simple
                        XNUMCHAR = 10
                    Case adVarChar
                        XNUMCHAR = RSPLAN.Fields(DataPlan.Columns(X).Caption).DefinedSize + 1
                    Case adUnsignedTinyInt
                        XNUMCHAR = 3
            End Select
            If RSPLAN.Fields(DataPlan.Columns(X).Caption).Type = adSingle Then
                RSAUX.Open "SELECT SUM(" & DataPlan.Columns(X).Caption & ") as TOTAL FROM  [##PRTCOSTOS" & VGL_COMPUTER & "] ", DBSTARPLAN, adOpenStatic
                If RSAUX.RecordCount = 0 Or IsNull(RSAUX!TOTAL) Then xValor = 0 Else xValor = RSAUX!TOTAL
                CAD = Right("         " & Format$(xValor, "0.00"), XNUMCHAR)
                RSAUX.Close
            Else
                CAD = String(XNUMCHAR, " ")
            End If
            XSTR = XSTR & CAD
        End If
    Next
    XSTR = "TOTAL General: " & Right(XSTR, Len(XSTR) - Len("TOTAL General: "))
    Set RSAUX = Nothing
    DBSTARPLAN.Execute "INSERT INTO  [##PRTPLANILLA" & VGL_COMPUTER & "]  VALUES ('" & XSTR & "')"
    CAD = String(Len(XSTR), "_")
    DBSTARPLAN.Execute "INSERT INTO ##PRTPLANILA VALUES ('" & CAD & "')"
    If ExisteTablaAux(" [##GLOSAPLAN" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##GLOSAPLAN" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "CREATE TABLE  [##GLOSAPLAN" & VGL_COMPUTER & "]  (GLOSA VARCHAR(500))"
    If Lista.ListItems.Count > 0 Then
        DBSTARPLAN.Execute "INSERT INTO  [##GLOSAPLAN" & VGL_COMPUTER & "]  (GLOSA) VALUES ('Facturación:" & Chr(13) & Chr(10) & "------------')", X
        Dim XITEM As ListItem
        For Each XITEM In Lista.ListItems
            DBSTARPLAN.Execute "UPDATE  [##GLOSAPLAN" & VGL_COMPUTER & "]  SET GLOSA=GLOSA + '" & Chr(13) & Chr(10) & Left(XITEM.SubItems(1) & String(25, " "), 25) & ":" & Right(String(10, " ") & XITEM.SubItems(2), 10) & "'"
        Next
    Else
        DBSTARPLAN.Execute "INSERT INTO  [##GLOSAPLAN" & VGL_COMPUTER & "]  (GLOSA) VALUES ('')"
    End If
  
    With Reporte
        frWait.Show 1
        .WindowTitle = "PLAN0012.RPT -" & DataPlan.Caption
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0012.RPT"
        .DataFiles(0) = App.PATH & "\BDAuxCom.mdb"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .Formulas(0) = "xTitulo='Planilla : " & frBolEmit.Lista.SelectedItem.Text & "'"
        .Formulas(1) = "xCabeza1='Cliente: " & xRazon.Text & "'"
        .Formulas(2) = "xCabeza2='Correspondiente a: " & xConcepto.Text & "'"
        .Formulas(3) = "xCabeza3='Fecha de Facturación: " & xFecha.Value & "'"
        .Formulas(4) = "xComentario='" & xComentario.Text & "'"
        Screen.MousePointer = 1
        If .Status <> 2 Then .Action = 1
    End With
End Sub
Private Sub CREARPLAN(Optional ByRef REG As Long)
    Dim RSTRABPLAN As New ADODB.Recordset
    Dim RSAUX As New ADODB.Recordset
    Dim I As Integer, ORDEN As Integer
    Dim CONC As String
    If ExisteTablaAux(" [##TMPCREPLAN" & VGL_COMPUTER & "] ") Then DBSTARPLAN.Execute "DROP TABLE  [##TMPCREPLAN" & VGL_COMPUTER & "] "
    DBSTARPLAN.Execute "CREATE TABLE  [##TMPCREPLAN" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8), CODCONCEP VARCHAR(20),DESCONCEP VARCHAR(50),ORDEN INT, MONTO  Numeric(20,2) )"
    RSAUX.Open " [##PRTCOSTOS" & VGL_COMPUTER & "] ", DBSTARPLAN, adOpenKeyset, adLockReadOnly
    REG = RSAUX.RecordCount
    RSTRABPLAN.Open " [##TMPCREPLAN" & VGL_COMPUTER & "] ", DBSTARPLAN, adOpenKeyset, adLockOptimistic
    If RSAUX.RecordCount = 0 Then
        MsgBox "No existe ningún registro para imprimir la planilla"
        Exit Sub
    End If
    RSAUX.MoveFirst
    Do While Not RSAUX.EOF
        For I = 11 To RSAUX.Fields.Count - 1
            ORDEN = ORDEN + 1
            If DataPlan.Columns.Item(RSAUX.Fields(I).Name).Visible And _
            RSAUX.Fields(I).Value > 0 Then
                RSTRABPLAN.AddNew
                RSTRABPLAN!CODTRAB = RSAUX!CODTRAB
                RSTRABPLAN!CODCONCEP = Trim(RSAUX.Fields(I).Name)
                CONC = DevuelveValor("SELECT NOMBRE FROM CONCEPTOS WHERE CODIGO='" & _
                                     Trim(RSAUX.Fields(I).Name) & "'", DBSYSTEM)
                If Trim(CONC) = "" Then
                    Select Case UCase(Trim(RSAUX.Fields(I).Name))
                        Case "OTROSINGR": CONC = "Otros Ingresos"
                        Case "TOTING": CONC = "Total Ingresos"
                        Case "OTROSEGRE": CONC = "Otros Egresos"
                        Case "ADELANTO": CONC = "Adelanto"
                        Case "TOTEGR": CONC = "Total Egresos"
                        Case "NETO": CONC = "Neto"
                    End Select
                End If
                RSTRABPLAN!DESCONCEP = CONC
                RSTRABPLAN!ORDEN = ORDEN
                RSTRABPLAN!MONTO = RSAUX.Fields(I).Value
                RSTRABPLAN.Update
            End If
        Next
        ORDEN = 0
        RSAUX.MoveNext
    Loop
End Sub
Private Sub CMNUEVA_Click()
    VPTAREA = "Nuevo"
    frEForFact.Show 1
End Sub

Private Sub cmOcultar_Click()
    If DataPlan.COL < 0 Then Exit Sub
    If DataPlan.Columns(DataPlan.COL).Caption = "NOMBRES" Then
        MsgBox "No se puede ocultar esta columna", vbInformation
        Exit Sub
    End If
    DataPlan.Columns(DataPlan.COL).Visible = False
End Sub

Private Sub CMQUITAR_CLICK()
    If Lista.ListItems.Count = 0 Then
        MsgBox "No existen registros para ser quitados de entre los seleccionados", vbCritical
        cmAgregar.SetFocus
        Exit Sub
    End If
    Lista.ListItems.Remove Lista.SelectedItem.INDEX
    If Lista.ListItems.Count = 0 Then
        MsgBox "La lista de formulas de facturación se encuentra vacía", vbInformation
    End If
End Sub
Private Sub Form_Load()
    'Proceso de Creación de la tabla temporal donde se almacenaran los datos
    CambiaPanelBD True
    If GetSetting(App.CompanyName, "Planillas", "Nando", "No") <> "Hola" Then On Error GoTo PasarError
    If ExisteTablaAux(" [##PRTCOSTOS" & VGL_COMPUTER & "] ") Then DBSTARPLAN.Execute "DROP TABLE  [##PRTCOSTOS" & VGL_COMPUTER & "] "
    Dim RSAUX As New ADODB.Recordset
    Dim strTabla As String, CAD As String, STRSQL As String
    
    Dim SNOMBOL As String
    SNOMBOL = Right(frBolEmit.Lista.SelectedItem.KEY, Len(frBolEmit.Lista.SelectedItem.KEY) - 1)
    DataPlan.Caption = frBolEmit.Lista.SelectedItem.Text
    Dim FMES As Date
    FMES = DevuelveValor("SELECT MES FROM NOMBOL WHERE CODIGO=" & SNOMBOL, DBSYSTEM)
    strTabla = Format(Month(FMES), "00") & Year(FMES)
    STRSQL = ""
    STRSQL = "CREATE TABLE  [##PRTCOSTOS" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8),NOMBRES VARCHAR(50),DOCIDEN VARCHAR(15),TIPOTRAB VARCHAR(2),FECHAING DATETIME,CCOSTO VARCHAR(25),AREA VARCHAR(25),CARGO VARCHAR(50),BASICO  Numeric(20,2) ,CARNETSEG VARCHAR(15),FONDOPENS VARCHAR(2)"

    RSAUX.Open "SELECT CODIGO, NOMBRE,TIPO FROM CONCEPTOS WHERE CODIGO IN (SELECT DISTINCT CONCEPTO FROM MOV" & strTabla & " WHERE CODNOMBOL IN " & VPTAREA & ") ORDER BY TIPO, FILA", DBSYSTEM, adOpenStatic
    Dim tmpTipo As Integer
    tmpTipo = 0
    While Not RSAUX.EOF
        If tmpTipo <> RSAUX!TIPO Then
            Select Case RSAUX!TIPO
                Case 2
                    STRSQL = STRSQL & ",OTROSINGR  Numeric(20,2) , TOTING  Numeric(20,2)  "
                Case 3
                    STRSQL = STRSQL & ",OTROSEGRE  Numeric(20,2) , ADELANTO  Numeric(20,2) , TOTEGR  Numeric(20,2) "
            End Select
            tmpTipo = RSAUX!TIPO
        End If
        STRSQL = STRSQL & "," & RSAUX!Codigo & "  Numeric(20,2) "
        RSAUX.MoveNext
    Wend
    STRSQL = STRSQL & ")"
    DBSYSTEM.Execute STRSQL
    'Llenado de los datos personales para la planilla
    Set RSAUX = Nothing
    Dim RSBOL As New ADODB.Recordset
    Set RSAUX = New ADODB.Recordset
    RSAUX.Open "SELECT * FROM VWTRABAJ WHERE CODTRAB IN (SELECT CODTRAB FROM  [##_TMPLSTBOL" & VGL_COMPUTER & "] )", DBSYSTEM, adOpenStatic
    With RSAUX
        While Not RSAUX.EOF
            DBSTARPLAN.Execute "INSERT INTO  [##PRTCOSTOS" & VGL_COMPUTER & "]  (CODTRAB,NOMBRES,DOCIDEN,TIPOTRAB,FECHAING,CCOSTO,AREA,CARGO,BASICO,CARNETSEG,FONDOPENS) VALUES ('" & !CODTRAB & "','" & !NOMBRES & "','" & !DOCIDEN & "','" & !TIPOTRAB & "'," & FechS(!FECHAING, Sqlf) & ",'" & !CENTRO & "','" & !NOMBREAREA & "','" & !CARGO & "'," & !BASICO & ",'" & !CARNETSEG & "','" & !FONDOPENS & "')"
            RSAUX.MoveNext
        Wend
        'DBSTARPLAN.Execute "UPDATE [##PRTCOSTOS" & VGL_COMPUTER & "] SET [##PRTCOSTOS" & VGL_COMPUTER & "].BASICO=TMPLSTBOL.BASICO FROM [##PRTCOSTOS" & VGL_COMPUTER & "] PRTCOSTOS,[##_TMPLSTBOL" & VGL_COMPUTER & "] TMPLSTBOL WHERE TMPLSTBOL.CODTRAB=PRTCOSTOS.CODTRAB"
    End With
    RSAUX.Close
    'Apertura del Grid
    Set RSPLAN = New ADODB.Recordset
    RSPLAN.Open " [##PRTCOSTOS" & VGL_COMPUTER & "] ", DBSTARPLAN, adOpenKeyset
    For tmpTipo = 0 To RSPLAN.Fields.Count - 1
        If RSPLAN.Fields(tmpTipo).Type = 6 Then
            DBSTARPLAN.Execute "UPDATE  [##PRTCOSTOS" & VGL_COMPUTER & "]  SET " & RSPLAN.Fields(tmpTipo).Name & "=0 WHERE (" & RSPLAN.Fields(tmpTipo).Name & ") IS NULL"
        End If
    Next
    'Llenado de los datos remunerativos para la planilla de trabajo
    If Not ExisteCampo("TOTING", " [##PRTCOSTOS" & VGL_COMPUTER & "] ", DBSYSTEM) Then
        DBSTARPLAN.Execute "ALTER TABLE  [##PRTCOSTOS" & VGL_COMPUTER & "]  ADD TOTING  Numeric(20,2) "
    End If
    If Not ExisteCampo("TOTEGR", " [##PRTCOSTOS" & VGL_COMPUTER & "] ", DBSYSTEM) Then
        DBSTARPLAN.Execute "ALTER TABLE  [##PRTCOSTOS" & VGL_COMPUTER & "]  ADD TOTEGR  Numeric(20,2) "
    End If
    If Not ExisteCampo("NETO", " [##PRTCOSTOS" & VGL_COMPUTER & "] ", DBSYSTEM) Then
        DBSTARPLAN.Execute "ALTER TABLE  [##PRTCOSTOS" & VGL_COMPUTER & "]  ADD NETO  Numeric(20,2) "
    End If
    '************************************////////////
    RSAUX.Open "SELECT CODTRAB,TOTING,TOTEGR,CONCEPTO,MONTO,BASICO FROM BOL" & strTabla & " BOL,MOV" & strTabla & " MOV WHERE BOL.INUMBOL=MOV.INUMBOL AND BOL.CODNOMBOL IN " & VPTAREA & " ORDER BY CODTRAB", DBSYSTEM, adOpenStatic
    CAD = ""
    While Not RSAUX.EOF
        If CAD <> RSAUX!CODTRAB Then
            DBSTARPLAN.Execute "UPDATE  [##PRTCOSTOS" & VGL_COMPUTER & "]  SET TOTING=" & IIf(IsNull(RSAUX!TOTING), 0, RSAUX!TOTING) & ",TOTEGR=" & IIf(IsNull(RSAUX!TOTEGR), 0, RSAUX!TOTEGR) & " WHERE CODTRAB='" & RSAUX!CODTRAB & "'"
            CAD = RSAUX!CODTRAB
        End If
        DBSTARPLAN.Execute "UPDATE  [##PRTCOSTOS" & VGL_COMPUTER & "]  SET BASICO=" & RSAUX!BASICO & "," & RSAUX!CONCEPTO & "=" & RSAUX!CONCEPTO & "+" & RSAUX!MONTO & " WHERE CODTRAB='" & RSAUX!CODTRAB & "'"
        RSAUX.MoveNext
    Wend
    '************************************//////////////
    DBSTARPLAN.Execute "UPDATE  [##PRTCOSTOS" & VGL_COMPUTER & "]  SET NETO=TOTING-TOTING"
    RSAUX.Close
    'Carga de los adelantos ya cobrados
    RSAUX.Open "SELECT CODTRAB, MONTO FROM ADEL2000 WHERE NOMBOL IN " & VPTAREA, DBSYSTEM, adOpenStatic
    If Not ExisteCampo("Adelanto", " [##PRTCOSTOS" & VGL_COMPUTER & "] ", DBSYSTEM) Then
        DBSYSTEM.Execute "ALTER TABLE  [##PRTCOSTOS" & VGL_COMPUTER & "]  ADD ADELANTO  Numeric(20,2) "
        DBSYSTEM.Execute "UPDATE  [##PRTCOSTOS" & VGL_COMPUTER & "]  SET ADELANTO=0"
    End If
    Do While Not RSAUX.EOF
        DBSYSTEM.Execute "UPDATE  [##PRTCOSTOS" & VGL_COMPUTER & "]  SET ADELANTO=ADELANTO+" & RSAUX!MONTO & " WHERE CODTRAB='" & RSAUX!CODTRAB & "'"
        RSAUX.MoveNext
    Loop
    RSAUX.Close
    RSAUX.Open "SELECT CODTRAB, MONTO, TIPO FROM PAGOSCTA WHERE CODNOMBOL In " & VPTAREA, DBSYSTEM, adOpenStatic
    Do While Not RSAUX.EOF
        If RSAUX!TIPO = 1 Then
            'si es tipo Ingreso
            DBSYSTEM.Execute "UPDATE  [##PRTCOSTOS" & VGL_COMPUTER & "]  SET OTROSING=OTROSING+" & RSAUX!MONTO & " WHERE CODTRAB='" & RSAUX!CODTRAB & "'"
        Else
            DBSYSTEM.Execute "UPDATE  [##PRTCOSTOS" & VGL_COMPUTER & "]  SET OTROSEGRE=OTROSEGRE+" & RSAUX!MONTO & " WHERE CODTRAB='" & RSAUX!CODTRAB & "'"
        End If
        RSAUX.MoveNext
    Loop
    RSAUX.Close
    RSPLAN.Requery
    Set DataPlan.DataSource = RSPLAN
    For tmpTipo = 0 To RSPLAN.Fields.Count - 1
        If RSPLAN.Fields(tmpTipo).Type = 6 Then
            DataPlan.Columns(RSPLAN.Fields(tmpTipo).Name).NumberFormat = "##,##0.00 "
            DataPlan.Columns(RSPLAN.Fields(tmpTipo).Name).Alignment = dbgRight
            DataPlan.Columns(RSPLAN.Fields(tmpTipo).Name).Width = 950
        End If
    Next
    Set RSAUX = Nothing
    DBSYSTEM.Execute "CREATE INDEX CODTRAB ON  [##PRTCOSTOS" & VGL_COMPUTER & "]  (CODTRAB)"
    CambiaPanelBD False
    Exit Sub
PasarError:
    MsgBox "Inconsistencia: " & ERR.Description, vbInformation, "MS SQL: " & ERR.Number
    Resume Next
    Resume
End Sub
