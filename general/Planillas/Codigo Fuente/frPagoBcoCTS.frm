VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frPagoBcoCTS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pagos por Banco"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6540
   Icon            =   "frPagoBcoCTS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   6540
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Exportación"
      Height          =   330
      Left            =   165
      TabIndex        =   16
      Top             =   6165
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CommandButton cmActualizarCtas 
      Caption         =   "&Actualizar Ctas. Banco"
      Height          =   390
      Left            =   4305
      TabIndex        =   15
      Top             =   1495
      Width           =   1950
   End
   Begin AplisetControlText.Aplitext xDolar 
      Height          =   300
      Left            =   2925
      TabIndex        =   14
      Top             =   6135
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   529
      MaxLength       =   12
      Text            =   "0"
      TipoDato        =   "N"
   End
   Begin MSComCtl2.DTPicker xFecha 
      Height          =   300
      Left            =   2925
      TabIndex        =   12
      Top             =   5805
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   529
      _Version        =   393216
      Format          =   62062593
      CurrentDate     =   36827
   End
   Begin VB.CommandButton cmCustodia 
      Caption         =   "&Custodia"
      Height          =   390
      Left            =   4710
      TabIndex        =   10
      Top             =   6075
      Visible         =   0   'False
      Width           =   1605
   End
   Begin Crystal.CrystalReport Reporte 
      Left            =   2925
      Top             =   1860
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3465
      Top             =   1860
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
            Picture         =   "frPagoBcoCTS.frx":044A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmPlanilla 
      Caption         =   "&Planilla C.T.S."
      Height          =   390
      Left            =   4305
      TabIndex        =   9
      Top             =   935
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.CommandButton cmListado 
      Caption         =   "&Listado al Banco"
      Height          =   390
      Left            =   4305
      TabIndex        =   8
      Top             =   375
      Width           =   1950
   End
   Begin VB.CommandButton cmCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   390
      Left            =   4305
      TabIndex        =   7
      Top             =   2055
      Width           =   1950
   End
   Begin VB.CommandButton cmQuitar 
      Caption         =   "&Quitar Trabajador"
      Height          =   390
      Left            =   4710
      TabIndex        =   3
      Top             =   5580
      Width           =   1605
   End
   Begin MSDataGridLib.DataGrid dgLista 
      Height          =   2625
      Left            =   180
      TabIndex        =   2
      Top             =   2880
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   4630
      _Version        =   393216
      Appearance      =   0
      BackColor       =   14869218
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
   Begin MSComctlLib.ListView Bancos 
      Height          =   2385
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   4207
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Banco"
         Object.Width           =   4304
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor del Dolar"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1665
      TabIndex        =   13
      Top             =   6188
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de Abono"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1650
      TabIndex        =   11
      Top             =   5850
      Width           =   1185
   End
   Begin VB.Label xNumTrabs 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " 0 Trabajadores"
      ForeColor       =   &H80000005&
      Height          =   195
      Left            =   165
      TabIndex        =   6
      Top             =   5550
      Width           =   1110
   End
   Begin VB.Label l1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Total"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1650
      TabIndex        =   5
      Top             =   5550
      Width           =   405
   End
   Begin VB.Label xTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00 "
      Height          =   300
      Left            =   2925
      TabIndex        =   4
      Top             =   5490
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Entidades Bancarias"
      Height          =   195
      Left            =   135
      TabIndex        =   1
      Top             =   135
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000003&
      BackStyle       =   1  'Opaque
      Height          =   3840
      Left            =   75
      Top             =   2775
      Width           =   6390
   End
End
Attribute VB_Name = "frPagoBcoCTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSPAGOS As New ADODB.Recordset
Dim XITEM As ListItem

Private Sub BANCOS_ITEMCHECK(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked Then
        DBSYSTEM.Execute "INSERT INTO  [##ULTPAGOS" & VGL_COMPUTER & "]  SELECT *, GETDATE() AS FECHANACIMIENTO FROM  [##PAGOSXBANCO" & VGL_COMPUTER & "]  WHERE BANCO='" & Item.Text & "'"
    Else
        DBSYSTEM.Execute "DELETE FROM  [##ULTPAGOS" & VGL_COMPUTER & "]  WHERE BANCO='" & Item.Text & "'"
    End If
    FORMATEARDG
End Sub

Private Sub CMACTUALIZARCTAS_Click()
    If RSPAGOS.RecordCount = 0 Then
        MsgBox "NO EXISTEN REGISTROS PARA ACTUALIZAR", vbCritical
        Exit Sub
    End If
    RSPAGOS.MoveFirst
    Do While Not RSPAGOS.EOF
        DBSYSTEM.Execute "UPDATE PLANCTS SET CTABANCO='" & Trim(DevuelveValor("SELECT CTACTS FROM TRABAJADORES WHERE CODTRAB='" & RSPAGOS!CODTRAB & "'", DBSYSTEM)) & "', BANCO='" & DevuelveValor("SELECT BANCOCTS FROM TRABAJADORES WHERE CODTRAB='" & RSPAGOS!CODTRAB & "'", DBSYSTEM) & "' WHERE CODIGO=" & VPTRASPRM & " AND CODTRAB='" & RSPAGOS!CODTRAB & "'"
        RSPAGOS.MoveNext
    Loop
    DBSYSTEM.Execute "UPDATE PLANCTS SET BANCO='NONE' WHERE CTABANCO='' OR (CTABANCO)IS NULL"
    If MsgBox("LOS CAMBIOS SE PRESENTARAN LA PROXIMA VEZ QUE INGRESE A ESTA VENTANA. DESEA SALIR DE ESTA VENTANA PARA QUE VUELVA A INGRESAR", vbYesNo + vbQuestion) = vbYes Then Unload Me Else RSPAGOS.MoveFirst
End Sub

Private Sub cmCerrar_Click()
    Unload Me
End Sub

Private Sub CMIMPRIMIR_CLICK()
    If RSPAGOS.RecordCount = 0 Then
        MsgBox "NO EXISTEN REGISTROS PARA IMPRIMIR", vbCritical
        Exit Sub
    End If
    With Reporte
        frWait.Show 1
        .WindowTitle = "PLAN0026 - REPORTE DE PAGOS POR BANCO"
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0026.RPT"
        .DataFiles(0) = App.PATH & "\BDAUXCOM.MDB"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "XRUC='RUC N° " & REGSISTEMA.RUC & "'"
        If .Status <> 2 Then .Action = 1
    End With
End Sub

Private Sub CMCUSTODIA_CLICK()
    If RSPAGOS.EOF Or RSPAGOS.RecordCount = 0 Then
        MsgBox "NO EXISTEN REGISTROS PARA PONER EN CUSTODIA DE PAGO DE C.T.S.", vbCritical
        Exit Sub
    End If
End Sub

Private Sub CMLISTADO_CLICK()
    If RSPAGOS.RecordCount = 0 Then
        MsgBox "NO EXISTEN REGISTROS PARA IMPRIMIR", vbCritical
        Exit Sub
    End If
    DBSYSTEM.Execute "UPDATE  [##ULTPAGOS" & VGL_COMPUTER & "]  SET  [##ULTPAGOS" & VGL_COMPUTER & "] .FECHANACIMIENTO = B.FECHANAC FROM  [##ULTPAGOS" & VGL_COMPUTER & "]  A, TRABAJADORES B WHERE A.CODTRAB = B.CODTRAB"
    With Reporte
        .WindowTitle = "PLAN0048 - REPORTE DE PAGOS POR BANCO - C.T.S."
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0048.RPT"
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
        .StoredProcParam(0) = " [##BANCOS" & VGL_COMPUTER & "] "
        .StoredProcParam(1) = " [##ULTPAGOS" & VGL_COMPUTER & "] "
        .StoredProcParam(2) = "CODBANCO"
        .StoredProcParam(3) = "BANCO"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "XRUC='RUC N° " & REGSISTEMA.RUC & "'"
        .Formulas(2) = "XFECHA='" & xFecha.Value & "'"
        .Formulas(3) = "XVALORDOLAR=" & xDolar.Text
        If .Status <> 2 Then .Action = 1
    End With
End Sub

Private Sub CMQUITAR_CLICK()
    If RSPAGOS.EOF Or RSPAGOS.RecordCount = 0 Then
        MsgBox "NO EXISTEN REGISTROS A ELIMINAR", vbCritical
        Exit Sub
    End If
    If MsgBox("DESEA ELIMINAR EL REGISTRO DE " & RSPAGOS!NOMBRES, vbYesNo + vbQuestion) = vbNo Then Exit Sub
    DBSYSTEM.Execute "DELETE FROM ULTPAGOS WHERE CODTRAB='" & RSPAGOS!CODTRAB & "'"
    FORMATEARDG
End Sub
Private Sub Command1_Click()
 BcoCTSCred.Show 1
End Sub
Private Sub DGLISTA_HEADCLICK(ByVal COLINDEX As Integer)
    RSPAGOS.Sort = dgLista.Columns(COLINDEX).DataField
End Sub

Private Sub Form_Load()
    If ExisteTablaAux(" [##BANCOS" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##BANCOS" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "SELECT * INTO  [##BANCOS" & VGL_COMPUTER & "]  FROM BANCOS"
    DBSYSTEM.Execute "CREATE Index CODBANCO ON  [##BANCOS" & VGL_COMPUTER & "]  (CODBANCO)"
    RSPAGOS.Open "SELECT DISTINCT BANCOS.CODBANCO, BANCOS.NOMBRE FROM  [##BANCOS" & VGL_COMPUTER & "]  BANCOS,  [##PAGOSXBANCO" & VGL_COMPUTER & "]  PAGOSXBANCO WHERE PAGOSXBANCO.BANCO=BANCOS.CODBANCO ORDER BY BANCOS.NOMBRE", DBSYSTEM, adOpenStatic
    Do While Not RSPAGOS.EOF
        Set XITEM = Bancos.ListItems.Add(, , RSPAGOS!CODBANCO, , 1)
        XITEM.SubItems(1) = RSPAGOS!NOMBRE
        RSPAGOS.MoveNext
    Loop
    If ExisteTablaAux(" [##ULTPAGOS" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##ULTPAGOS" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "SELECT * INTO  [##ULTPAGOS" & VGL_COMPUTER & "]  FROM  [##PAGOSXBANCO" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "DELETE FROM  [##ULTPAGOS" & VGL_COMPUTER & "] "
    Set RSPAGOS = Nothing
    DBSYSTEM.Execute "ALTER TABLE  [##ULTPAGOS" & VGL_COMPUTER & "]  ADD FECHANACIMIENTO DATETIME"
    RSPAGOS.Open " [##ULTPAGOS" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenStatic
    xDolar.Text = MDIPrincipal.BarraEstado.Panels("Dolar").Text
    xFecha.Value = Date
    FORMATEARDG
End Sub

Public Sub FORMATEARDG()
    RSPAGOS.Requery
    Set dgLista.DataSource = RSPAGOS
    With dgLista
        .Columns("CODTRAB").Visible = False
        .Columns("TIPDOC").Visible = False
        .Columns("DOCIDEN").Visible = False
        .Columns("NETO").Alignment = dbgRight
        .Columns("NETO").NumberFormat = "##,##0.00 "
        .Columns("NOMBRES").Width = 2610
    End With
    xNumTrabs.Caption = " " & RSPAGOS.RecordCount & " TRABAJADORES"
    Dim RSAUX As New ADODB.Recordset
    RSAUX.Open "SELECT SUM(NETO) AS TOTAL FROM  [##ULTPAGOS" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenStatic
    If IsNull(RSAUX!TOTAL) Then xTotal.Caption = "0.00 " Else xTotal.Caption = Format(RSAUX!TOTAL, "0.00 ")
    Set RSAUX = Nothing
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSPAGOS = Nothing
End Sub

Public Function SOLOUNBANCO() As Boolean
    Dim xLista As ListItem
    Dim XVECES As Byte
    XVECES = 0
    If RSPAGOS.RecordCount = 0 Then
        SOLOUNBANCO = False
        Exit Function
    End If
    For Each xLista In Bancos.ListItems
        If xLista.Checked Then XVECES = XVECES + 1
    Next
    If XVECES <> 1 Then
        MsgBox "ERROR DE USUARIO: DEBE SELECCIONAR UNA/SOLO UNA ENTIDAD BANCARIA PARA PODER INGRESAR A ESTA OPCIÓN", vbCritical
        SOLOUNBANCO = False
    Else
        SOLOUNBANCO = True
    End If
End Function

