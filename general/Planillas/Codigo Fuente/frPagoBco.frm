VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frPagoBco 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pagos por Banco"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6420
   Icon            =   "frPagoBco.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   6420
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport Reporte 
      Left            =   2460
      Top             =   2775
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   405
      Top             =   2955
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
            Picture         =   "frPagoBco.frx":044A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmConsolidar 
      Caption         =   "Consolidar"
      Height          =   390
      Left            =   4305
      TabIndex        =   12
      Top             =   2175
      Width           =   1950
   End
   Begin VB.CommandButton cmContinental 
      Caption         =   "&Banco Continental"
      Height          =   390
      Left            =   4305
      TabIndex        =   11
      Top             =   1380
      Width           =   1950
   End
   Begin VB.CommandButton cmCredito 
      Caption         =   "&Banco de Crédito"
      Height          =   390
      Left            =   4305
      TabIndex        =   10
      Top             =   870
      Width           =   1950
   End
   Begin VB.CommandButton cmWiese 
      Caption         =   "&TeleWiese Empresarial"
      Height          =   390
      Left            =   4305
      TabIndex        =   9
      Top             =   375
      Width           =   1950
   End
   Begin VB.CommandButton cmCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   390
      Left            =   4965
      TabIndex        =   8
      Top             =   5790
      Width           =   1320
   End
   Begin VB.CommandButton cmImprimir 
      Caption         =   "&Imprimir"
      Height          =   390
      Left            =   3510
      TabIndex        =   4
      Top             =   5790
      Width           =   1320
   End
   Begin VB.CommandButton cmQuitar 
      Caption         =   "Quitar Trabajador"
      Height          =   390
      Left            =   135
      TabIndex        =   3
      Top             =   5790
      Width           =   1605
   End
   Begin MSDataGridLib.DataGrid dgLista 
      Height          =   2625
      Left            =   150
      TabIndex        =   2
      Top             =   2760
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   4630
      _Version        =   393216
      BackColor       =   -2147483633
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
      LabelEdit       =   1
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
   Begin VB.Label xNumTrabs 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " 0 Trabajadores"
      Height          =   300
      Left            =   150
      TabIndex        =   7
      Top             =   5415
      Width           =   1605
   End
   Begin VB.Label l1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total"
      Height          =   300
      Left            =   1770
      TabIndex        =   6
      Top             =   5415
      Width           =   1125
   End
   Begin VB.Label xTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00 "
      Height          =   300
      Left            =   2910
      TabIndex        =   5
      Top             =   5415
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
End
Attribute VB_Name = "frPagoBco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSPAGOS As New ADODB.Recordset
Dim XITEM As ListItem
Private Sub BANCOS_ITEMCHECK(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked Then
        DBSTARPLAN.Execute "INSERT INTO  [##ULTPAGOS" & VGL_COMPUTER & "]  SELECT * FROM  [##PAGOSXBANCO" & VGL_COMPUTER & "]  WHERE BANCO='" & Item.Text & "'"
    Else
        DBSTARPLAN.Execute "DELETE FROM  [##ULTPAGOS" & VGL_COMPUTER & "]  WHERE BANCO='" & Item.Text & "'"
    End If
    FORMATEARDG
End Sub

Private Sub cmCerrar_Click()
    Unload Me
End Sub

Private Sub CMCONSOLIDAR_Click()
    CambiaPanelBD True
    DBSTARPLAN.Execute "SELECT CODTRAB, NOMBRES, SUM(NETO) AS NETO2,TIPDOC, DOCIDEN,CTABANCO,BANCO INTO ##ULTPAGOS2 FROM  [##ULTPAGOS" & VGL_COMPUTER & "]  GROUP BY CODTRAB, NOMBRES,TIPDOC, DOCIDEN,CTABANCO,BANCO"
    DBSTARPLAN.Execute "UPDATE ##ULTPAGOS2 SET NETO2=NETO2"
    DBSTARPLAN.Execute "DROP TABLE  [##ULTPAGOS" & VGL_COMPUTER & "] "
    DBSTARPLAN.Execute "SELECT CODTRAB, NOMBRES, NETO2 AS NETO,TIPDOC, DOCIDEN,CTABANCO,BANCO INTO  [##ULTPAGOS" & VGL_COMPUTER & "]  FROM ##ULTPAGOS2"
    DBSTARPLAN.Execute "DROP TABLE ##ULTPAGOS2"
    FORMATEARDG
    CambiaPanelBD False
End Sub

Private Sub CMCONTINENTAL_Click()
    If Not SOLOUNBANCO() Then Exit Sub
    Load frBancoContinental
    frBancoContinental.xTotalAbonos.Caption = Val(xNumTrabs.Caption) & " "
    frBancoContinental.xTotalSoles.Caption = xTotal.Caption
    frBancoContinental.Show 1
End Sub

Private Sub CMCREDITO_Click()
    If SOLOUNBANCO() Then frBcoCredito.Show 1
End Sub

Private Sub CMIMPRIMIR_CLICK()
    If RSPAGOS.RecordCount = 0 Then
        MsgBox "NO EXISTEN REGISTROS PARA IMPRIMIR", vbCritical
        Exit Sub
    End If
    CambiaPanelBD True
    DBSTARPLAN.Execute "[ASISTMP2] ' [##ULTPAGOS" & VGL_COMPUTER & "] ',' [##BANCOS" & VGL_COMPUTER & "] ','BANCO','CODBANCO'"
    With Reporte
        .Reset
        .WindowTitle = "PLAN0026 - REPORTE DE PAGOS POR BANCO"
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0026.RPT"
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=" & VGL_BASE & ""
        .StoredProcParam(0) = " [##ULTPAGOS" & VGL_COMPUTER & "]"
        .StoredProcParam(1) = " [##BANCOS" & VGL_COMPUTER & "]"
        .StoredProcParam(2) = "BANCO"
        .StoredProcParam(3) = "CODBANCO"
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
    CambiaPanelBD False
End Sub

Private Sub CMQUITAR_CLICK()
    If RSPAGOS.EOF Or RSPAGOS.RecordCount = 0 Then
        MsgBox "NO EXISTEN REGISTROS A ELIMINAR", vbCritical
        Exit Sub
    End If
    If MsgBox("DESEA ELIMINAR EL REGISTRO DE " & RSPAGOS!NOMBRES, vbYesNo + vbQuestion) = vbNo Then Exit Sub
    DBSYSTEM.Execute "DELETE FROM  [##ULTPAGOS" & VGL_COMPUTER & "]  WHERE CODTRAB='" & RSPAGOS!CODTRAB & "'"
    FORMATEARDG
End Sub

Private Sub CMWIESE_Click()
    If SOLOUNBANCO() Then frWiese.Show 1
End Sub

Private Sub DGLISTA_HEADCLICK(ByVal COLINDEX As Integer)
    RSPAGOS.Sort = dgLista.Columns(COLINDEX).DataField
End Sub

Private Sub Form_Load()
    If ExisteTablaAux(" [##BANCOS" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##BANCOS" & VGL_COMPUTER & "] "
    DBSTARPLAN.Execute "SELECT * INTO  [##BANCOS" & VGL_COMPUTER & "]  FROM " & REGSISTEMA.BASESQL & ".dbo.BANCOS"
    DBSTARPLAN.Execute "CREATE INDEX CODBANCO ON  [##BANCOS" & VGL_COMPUTER & "]  (CODBANCO)"
    RSPAGOS.Open "SELECT DISTINCT A.CODBANCO, A.NOMBRE FROM  [##BANCOS" & VGL_COMPUTER & "]  A,  [##PAGOSXBANCO" & VGL_COMPUTER & "]  B WHERE B.BANCO=A.CODBANCO ORDER BY A.NOMBRE", DBSTARPLAN, adOpenStatic
    Do While Not RSPAGOS.EOF
        Set XITEM = Bancos.ListItems.Add(, , RSPAGOS!CODBANCO, , 1)
        XITEM.SubItems(1) = RSPAGOS!NOMBRE
        RSPAGOS.MoveNext
    Loop
    If ExisteTablaAux(" [##ULTPAGOS" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##ULTPAGOS" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "SELECT * INTO  [##ULTPAGOS" & VGL_COMPUTER & "]  FROM  [##PAGOSXBANCO" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "DELETE FROM  [##ULTPAGOS" & VGL_COMPUTER & "] "
    Set RSPAGOS = Nothing
    RSPAGOS.Open " [##PAGOSXBANCO" & VGL_COMPUTER & "] ", DBSTARPLAN, adOpenKeyset, adLockOptimistic
    Do While Not RSPAGOS.EOF
        DBSYSTEM.Execute "UPDATE  [##PAGOSXBANCO" & VGL_COMPUTER & "]  SET NETO=" & Round(IIf(IsNull(RSPAGOS!Neto), 0, RSPAGOS!Neto), 2) & " WHERE NETO=" & IIf(IsNull(RSPAGOS!Neto), 0, RSPAGOS!Neto)
        RSPAGOS.MOVE 1
    Loop
    Set RSPAGOS = Nothing
    RSPAGOS.Open " [##ULTPAGOS" & VGL_COMPUTER & "] ", DBSTARPLAN, adOpenStatic
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
    xNumTrabs.Caption = " " & RSPAGOS.RecordCount & " Trabajadores"
    Dim RSAUX As New ADODB.Recordset
    RSAUX.Open "SELECT SUM(NETO) AS TOTAL FROM  [##ULTPAGOS" & VGL_COMPUTER & "] ", DBSTARPLAN, adOpenStatic
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

