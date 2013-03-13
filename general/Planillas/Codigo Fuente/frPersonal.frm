VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frPersonal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Panel de Trabajadores"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10350
   Icon            =   "frPersonal.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   10350
   Tag             =   "Panel de Administración de Trabajadores"
   Begin VB.CommandButton cmdEventos 
      Caption         =   "Eventos"
      Height          =   315
      Left            =   5160
      TabIndex        =   15
      Top             =   5880
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Estudios"
      Height          =   315
      Left            =   6780
      TabIndex        =   14
      Top             =   5880
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton CmdImpFotoCh 
      Caption         =   "Im&primir Fotocheck"
      Height          =   315
      Left            =   8430
      TabIndex        =   13
      Top             =   5880
      Width           =   1785
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frPersonal.frx":08CA
      Left            =   8280
      List            =   "frPersonal.frx":08CC
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   105
      Width           =   1950
   End
   Begin Crystal.CrystalReport RptTrab 
      Left            =   3750
      Top             =   2625
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   450
      Left            =   6285
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   ""
      Top             =   3315
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.PictureBox Ole1 
      DataField       =   "Foto"
      DataSource      =   "Data1"
      Height          =   1305
      Left            =   6300
      ScaleHeight     =   1245
      ScaleWidth      =   1530
      TabIndex        =   10
      Top             =   1920
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.OptionButton Option2 
      Caption         =   "&Centros de Costo"
      Height          =   210
      Left            =   4980
      TabIndex        =   5
      Top             =   180
      Width           =   1605
   End
   Begin VB.OptionButton Option1 
      Caption         =   "&Areas de Trabajo"
      Height          =   210
      Left            =   3180
      TabIndex        =   4
      Top             =   180
      Value           =   -1  'True
      Width           =   1605
   End
   Begin MSDataGridLib.DataGrid dgPersonal 
      CausesValidation=   0   'False
      Height          =   5280
      Left            =   3180
      TabIndex        =   0
      Top             =   480
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   9313
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
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
      Caption         =   "Trabajadores"
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
         MarqueeStyle    =   3
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox chDepende 
      Caption         =   "Incluir Trabajadores dependientes"
      Height          =   255
      Left            =   90
      TabIndex        =   3
      Top             =   180
      Width           =   3060
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5445
      Top             =   2940
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
            Picture         =   "frPersonal.frx":08CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPersonal.frx":11AA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tView 
      Height          =   5295
      Left            =   45
      TabIndex        =   1
      Top             =   480
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   9340
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.TextBox SqlCad 
      Height          =   315
      Left            =   4005
      TabIndex        =   6
      Top             =   1245
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.TextBox SqlSel 
      Height          =   315
      Left            =   4005
      TabIndex        =   7
      Top             =   1575
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox SqlText 
      Height          =   285
      Left            =   4005
      TabIndex        =   8
      Top             =   1920
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.TextBox SqlWhere 
      Height          =   285
      Left            =   4020
      TabIndex        =   9
      Top             =   2250
      Visible         =   0   'False
      Width           =   1830
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   390
      Left            =   60
      TabIndex        =   16
      Top             =   5850
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   688
      ButtonWidth     =   3307
      ButtonHeight    =   582
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Otros Reportes        "
            Object.ToolTipText     =   "Click aquí para más reportes de trabajadores"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   6
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "IngEgr"
                  Text            =   "Listado de Ingresos y Egresos de Trabajadores"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "xTiTra"
                  Text            =   "Reporte por Tipo de Trabajador"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "AUXTRAB"
                  Text            =   "Reporte Auxiliar de Trabajadores"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RPCC"
                  Text            =   "Relacion de Personal por Centro de Costo"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "TRAB2"
                  Text            =   "Reporte General de Trabajadores 2"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Key             =   "PERSO"
                  Text            =   "Personalizar"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Ver:"
      Height          =   195
      Left            =   7770
      TabIndex        =   11
      Top             =   165
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "No Existen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5475
      TabIndex        =   2
      Top             =   1815
      Width           =   1305
   End
   Begin VB.Image Image2 
      Height          =   870
      Left            =   4830
      Picture         =   "frPersonal.frx":200A
      Stretch         =   -1  'True
      Top             =   960
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   885
      Left            =   4410
      Picture         =   "frPersonal.frx":28D4
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   885
   End
End
Attribute VB_Name = "frPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSPERSONAL As New ADODB.Recordset
Dim REGACT As REGWIN
Dim INTO As String
Dim FILTRADO As Boolean
Private Sub CMDEVENTOS_CLICK()
    If Not REGSISTEMA.VALRRHH Then
        MsgBox "No esta permitido el uso de esta opcion por no incluirse en la licencia", vbInformation
        Exit Sub
    End If
    If VAR_SHOW = 0 Then
        If RSPERSONAL.RecordCount Then
            Dim X As Integer
            X = 0
            On Error Resume Next
            DBADMINPER.Execute "UPDATE TRABAJADORES SET APEPAT=APEPAT WHERE CODTRAB='" & RSPERSONAL!CODTRAB & "'", X
            FrmEventos.Label9.Caption = dgPersonal.Columns(1).Text
            FrmEventos.Aplitext1.Text = dgPersonal.Columns(0).Text
            VAR_SHOW = 2
            FrmEventos.Show
        Else
            MsgBox "No existen registros de trabajadores", vbInformation
            Exit Sub
        End If
    End If
End Sub
Private Sub CMDIMPFOTOCH_CLICK()
    IMPRIMIRFOTOCHECKS
End Sub
Private Sub Command1_Click()
    If Not REGSISTEMA.VALRRHH Then
        MsgBox "No esta permitido el uso de esta opcion por no incluirse en la licencia", vbInformation
        Exit Sub
    End If
    If VAR_SHOW = 0 Then
        If RSPERSONAL.RecordCount Then
            Dim X As Integer
            X = 0
            On Error Resume Next
            DBADMINPER.Execute "UPDATE TRABAJADORES SET APEPAT=APEPAT WHERE CODTRAB='" & RSPERSONAL!CODTRAB & "'", X
            frEstCurrTrab.Text2.Text = dgPersonal.Columns(1).Text
            frEstCurrTrab.Text3.Text = dgPersonal.Columns(0).Text
            VAR_SHOW = 1
            frEstCurrTrab.Show
        Else
            MsgBox "No existen registros, ingrese el trabajador", vbInformation, "INFORMACION"
            Exit Sub
        End If
    End If
End Sub
Private Sub CHDEPENDE_CLICK()
    TVIEW_NODECLICK tView.SelectedItem
End Sub
Private Sub Combo1_Click()
    If Combo1.ListIndex = 8 Then
        FILTRADO = False
    Else
        FILTRADO = True
    End If
    TVIEW_NODECLICK tView.SelectedItem
End Sub

Private Sub DGPERSONAL_DBLCLICK()
    COMANDOTOOLBAR "EDITAR"
End Sub

Private Sub DGPERSONAL_HEADCLICK(ByVal COLINDEX As Integer)
    RSPERSONAL.Sort = dgPersonal.Columns(COLINDEX).DataField
End Sub
Private Sub DGPERSONAL_ROWCOLCHANGE(LASTROW As Variant, ByVal LASTCOL As Integer)
If VAR_MODO_EDIT = False Then
    If VAR_SHOW = 1 Then
            If RSPERSONAL.RecordCount Then
                frEstCurrTrab.Text2.Text = dgPersonal.Columns(1).Text
                frEstCurrTrab.Text3.Text = dgPersonal.Columns(0).Text
            Else
                MsgBox "No existen registros, ingresa al trabajador", vbInformation, "INFORMACION"
                Exit Sub
            End If
    ElseIf VAR_SHOW = 2 Then
      If RSPERSONAL.RecordCount Then
            FrmEventos.Label9.Caption = dgPersonal.Columns(1).Text
            FrmEventos.Aplitext1.Text = dgPersonal.Columns(0).Text
            FrmEventos.Show
      Else
            MsgBox "No existen registros, ingresa al trabajador", vbInformation, "INFORMACION"
            Exit Sub
      End If
    End If
End If
End Sub
Private Sub Form_Activate()
    ActivarTools REGACT
End Sub

Private Sub Form_Load()
    CambiaPanelBD True
    Data1.DatabaseName = App.PATH & "\BDAUXCOM.MDB"
    Data1.RecordSource = "TMPFOTOCHECKS"
    Data1.Refresh
    With REGACT
        .BUSCAR = True
        .EDITAR = True
        .ELIMINAR = True
        .FILTRAR = True
        .IMPRIMIR = True
        .NUEVO = True
        .PRELIMINAR = False
    End With
    For X = 0 To 8
        Combo1.AddItem ARRSITUACION(X)
    Next
    Combo1.AddItem "TODOS LOS TRABAJADORES"
    Combo1.AddItem "TODOS LOS ACTIVOS"
    Combo1.AddItem "NO CALCULADOS"
    Combo1.AddItem "CALCULADOS"
    CARGAARBOL
    FILTRADO = False
    RSPERSONAL.Open "SELECT * FROM VWTRABAJ ORDER BY NOMBREs", DBSYSTEM, adOpenKeyset, adLockOptimistic 'VWTRABAJ
    Select Case Combo1.ListIndex
        Case 9: SqlSel.Text = "SELECT * FROM VWTRABAJ NOCALCULO=-1"
        Case 10: SqlSel.Text = "SELECT * FROM VWTRABAJ NOCALCULO=0"
        Case Else
            SqlSel.Text = "SELECT * FROM VWTRABAJ " & IIf(FILTRADO, " WHERE SITUACIÓN=" & Combo1.ListIndex, "")
    End Select
       
    Set dgPersonal.DataSource = RSPERSONAL
    Combo1.ListIndex = 8
    CambiaPanelBD False
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    If RSPERSONAL.State <> 0 Then RSPERSONAL.Close
    Set RSPERSONAL = Nothing
End Sub

Public Sub COMANDOTOOLBAR(ByVal COMANDO As String)
    Select Case UCase(COMANDO)
        Case "NUEVO"
            VPTAREA = "NUEVO"
            frTrab.Command1.Enabled = True
            frTrab.Show 1
            If RSPERSONAL.State <> 0 Then RSPERSONAL.Requery
        Case "EDITAR"
            If RSPERSONAL.EOF Then Exit Sub
            VPTAREA = RSPERSONAL!CODTRAB
            CambiaPanelBD True
            Load frTrab
            frTrab.Command1.Enabled = False
            CambiaPanelBD False
            frTrab.Show 1
            RSPERSONAL.Requery
        Case "ELIMINAR"
            If RSPERSONAL.EOF Then Exit Sub
            Dim RSDEL As New ADODB.Recordset
            RSDEL.Open "SELECT * FROM MOVICTA WHERE CODTRAB='" & RSPERSONAL!CODTRAB & "'", DBSYSTEM, adOpenStatic
            If RSDEL.RecordCount > 0 Then
                MsgBox "El trabajador presenta dependencias en otros registros, los cuales pueden ser adelantos, descuentos o boletas", vbCritical
                Exit Sub
            End If
            CambiaPanelBD True
            If MsgBox("Realmente desea eliminar al trabajador seleccionado", vbInformation + vbYesNo) = vbNo Then Exit Sub
            DBSYSTEM.Execute "DELETE FROM TRABAJADORES WHERE CODTRAB='" & RSPERSONAL!CODTRAB & "'"
            DBADMINPER.Execute "DELETE FROM ESTUDIOS WHERE COD_TRAB='" & RSPERSONAL!CODTRAB & "'"
            DBADMINPER.Execute "DELETE FROM EVENTOS WHERE COD_TRABAJADOR='" & RSPERSONAL!CODTRAB & "'"
            DBADMINPER.Execute "DELETE FROM IDIOMAS WHERE COD_TRAB='" & RSPERSONAL!CODTRAB & "'"
            DBADMINPER.Execute "DELETE FROM LABORAL WHERE COD_TRAB='" & RSPERSONAL!CODTRAB & "'"
            If VAR_SHOW = 1 Then
                Unload frEstCurrTrab
            ElseIf VAR_SHOW = 2 Then
                Unload FrmEventos
            End If
            RSPERSONAL.Requery
            CambiaPanelBD False
        Case "BUSCAR"
            If RSPERSONAL.EOF Then Exit Sub
            CambiaPanelBD True
            Dim RSAUX As New ADODB.Recordset
            Set RSAUX = RSPERSONAL.Clone
            frmComun.CONECTAR RSAUX
            CambiaPanelBD False
            frmComun.Show 1
            If VGUTIL(1) <> "" Then
                RSPERSONAL.MoveFirst
                RSPERSONAL.FIND "CODTRAB='" & VGUTIL(1) & "'"
            End If
            Set RSAUX = Nothing
        Case "FILTRAR"
            MsgBox "Esta opcion ha sido reemplazada. La presente version del sistema ya no admite esta funcion", vbInformation
        Case "IMPRIMIR"
            CambiaPanelBD True
            With RptTrab
                .Reset
                .ReportFileName = REGSISTEMA.REPORTES & "\PLAN0035.RPT"
                .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=" & VGL_BASE & ""
                .StoredProcParam(0) = REGSISTEMA.BASESQL & ".dbo.VWTRABAJ"
                .Destination = crptToWindow
                .WindowState = crptMaximized
                .WindowShowPrintBtn = True
                .WindowShowRefreshBtn = True
                .WindowShowSearchBtn = True
                .WindowShowPrintSetupBtn = True
                .WindowTitle = "PLAN0035 - LISTADO GENERAL DE TRABAJADORES"
                .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
                .PrintReport
            End With
            CambiaPanelBD False
    End Select
End Sub

Public Sub CARGAARBOL()
    tView.Nodes.Clear
    Dim RSCCOSTO As New ADODB.Recordset
    INTO = ""
    If Option1.Value Then 'SI ES POR AREAS
        RSCCOSTO.Open "SELECT *  FROM AREASTRAB ORDER BY CODCCOSTO", DBSYSTEM, adOpenKeyset, adLockOptimistic
    Else
        RSCCOSTO.Open "SELECT *  FROM CCOSTOS ORDER BY CODCCOSTO", DBSYSTEM, adOpenKeyset, adLockOptimistic
    End If
    If RSCCOSTO.RecordCount = 0 Then
        Set RSCCOSTO = Nothing
        MsgBox "El archivo de areas de trabajo o centros de costos se encuentra vacia, por favor inicialize los centros de costos, vbCritical"
        Me.cmdEventos.Enabled = False
        Me.Command1.Enabled = False
        Exit Sub
    End If
    Dim XNODE As NODE
    Dim COD As String
    Set XNODE = tView.Nodes.Add(, , "RAIZ", IIf(Option1.Value, "TODAS LAS AREAS DE TRABAJO", "TODOS LOS CENTROS DE COSTO"), 1)
    With RSCCOSTO
        Do While Not .EOF
            If Len(!CODCCOSTO) = 2 Or InStr(!CODCCOSTO, ".") = 0 Then
                COD = "RAIZ"
            Else
               X = 1
               Do While Not X = 0
                  X = InStr(X + 1, !CODCCOSTO, ".")
                  If X <> 0 Then Y = X
               Loop
               COD = "C" & Mid(!CODCCOSTO, 1, Y - 1)
            End If
             Set XNODE = tView.Nodes.Add(COD, 4, "C" & !CODCCOSTO, !NOMBRE, 2)
            .MoveNext
        Loop
    End With
    tView.Nodes("RAIZ").Expanded = True
    tView.Nodes("RAIZ").Selected = True
    Set RSCCOSTO = Nothing
End Sub

Private Sub OPTION1_CLICK()
    CARGAARBOL
    TVIEW_NODECLICK tView.SelectedItem
End Sub

Private Sub OPTION2_Click()
    CARGAARBOL
    TVIEW_NODECLICK tView.SelectedItem
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim X As Integer
    Select Case UCase(ButtonMenu.KEY)
        Case "INGEGR"
            frmIngreCese.Show
        Case "XTITRA"
            CambiaPanelBD True
            Screen.MousePointer = vbHourglass
            If ExisteTablaAux(" [##TMPTIPTRAB" & VGL_COMPUTER & "] ") Then DBAUXCOM.Execute "DROP TABLE  [##TMPTIPTRAB" & VGL_COMPUTER & "] "
            SQL = "SELECT VWTRABAJ.CODTRAB, VWTRABAJ.NOMBRES, VWTRABAJ.CENTRO, VWTRABAJ.NOMBREAREA, VWTRABAJ.FECHAING, VWTRABAJ.BASICO, TIPOSTRAB.DESCRIP INTO  [##TMPTIPTRAB" & VGL_COMPUTER & "]  " & _
                    " FROM TIPOSTRAB INNER JOIN VWTRABAJ ON TIPOSTRAB.TIPTRAB = VWTRABAJ.TIPOTRAB "
            DBSYSTEM.Execute SQL
                            For Y = 1 To X
                                frWait.Show
                            Next
                            For Y = 1 To X
                            Next
                            For Y = 1 To X
                            Next
                With RptTrab
                    .Reset
                    .WindowTitle = "PLAN0088 - REPORTE DE TRABAJADORES POR TIPO"
                    .ReportFileName = REGSISTEMA.REPORTES & "PLAN0088.RPT"
                    .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
                    .StoredProcParam(0) = " [##TMPTIPTRAB" & VGL_COMPUTER & "] "
                    .Destination = crptToWindow
                    .WindowState = crptMaximized
                    .WindowShowPrintBtn = True
                    .WindowShowRefreshBtn = True
                    .WindowShowSearchBtn = True
                    .WindowShowPrintSetupBtn = True
                    .Formulas(0) = "EMPRESA='" & REGSISTEMA.EMPRESA & "'"
                    .Formulas(1) = "RUC='" & REGSISTEMA.RUC & "'"
                    If .Status <> 2 Then .Action = 1
                End With
                CambiaPanelBD False
                Screen.MousePointer = vbNormal
        Case "AUXTRAB"
            CambiaPanelBD True
            Screen.MousePointer = vbHourglass
                With RptTrab
                    .Reset
                    .WindowTitle = "PANEL DE TRABAJADORES "
                    .ReportFileName = REGSISTEMA.REPORTES & "PANEL.RPT"
                    .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
                    .StoredProcParam(0) = REGSISTEMA.BASESQL & ".dbo.VWTRABAJ"
                    .Destination = crptToWindow
                    .WindowState = crptMaximized
                    .WindowShowPrintBtn = True
                    .WindowShowRefreshBtn = True
                    .WindowShowSearchBtn = True
                    .WindowShowPrintSetupBtn = True
                    .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
                    .Formulas(1) = "XRUC='" & REGSISTEMA.RUC & "'"
                    If .Status <> 2 Then .Action = 1
                End With
                CambiaPanelBD False
                Screen.MousePointer = vbNormal
        Case "RPCC"
            CambiaPanelBD True
            Screen.MousePointer = vbHourglass
            If ExisteTablaAux(" [##TMPRELCC" & VGL_COMPUTER & "] ") Then DBAUXCOM.Execute "DROP TABLE  [##TMPRELCC" & VGL_COMPUTER & "] "
            DBSYSTEM.Execute "SELECT CCOSTOS.CODCCOSTO, CCOSTOS.NOMBRE AS CENTRO , TRABAJADORES.CODTRAB, TRABAJADORES.APEPAT, TRABAJADORES.APEMAT, TRABAJADORES.NOMBRE, TRABAJADORES.FECHAING, AREASTRAB.NOMBRE AS AREA , TRABAJADORES.CARGO INTO  [##TMPRELCC" & VGL_COMPUTER & "] " & _
                            " FROM AREASTRAB, CCOSTOS, TRABAJADORES WHERE CCOSTOS.CODCCOSTO = TRABAJADORES.CCOSTO AND AREASTRAB.CODCCOSTO = TRABAJADORES.AREA;", X
                For Y = 1 To X
                    frWait.Show
                Next
                For Y = 1 To X
                Next
                For Y = 1 To X
                Next
                With RptTrab
                    .Reset
                    .WindowTitle = "PERSONAL POR CENTRO DE COSTO"
                    .ReportFileName = REGSISTEMA.REPORTES & "PERSONALCC.RPT"
                    .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
                    .StoredProcParam(0) = " [##TMPRELCC" & VGL_COMPUTER & "] "
                    .Destination = crptToWindow
                    .WindowState = crptMaximized
                    .WindowShowPrintBtn = True
                    .WindowShowRefreshBtn = True
                    .WindowShowSearchBtn = True
                    .WindowShowPrintSetupBtn = True
                    .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
                    .Formulas(1) = "XRUC='" & REGSISTEMA.RUC & "'"
                    If .Status <> 2 Then .Action = 1
                End With
                CambiaPanelBD False
                Screen.MousePointer = vbNormal
            Case "TRAB2"
                CambiaPanelBD True
                Screen.MousePointer = 11
                If ExisteTablaAux(" [##_TMPTRABAJADORES" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##_TMPTRABAJADORES" & VGL_COMPUTER & "] "
                DBSYSTEM.Execute "SELECT TRABAJADORES.CODTRAB, TRABAJADORES.APEPAT+' '+TRABAJADORES.APEMAT+' '+TRABAJADORES.NOMBRE AS NOMBRES, CCOSTOS.NOMBRE AS CENTRO, CCOSTOS.CODCCOSTO, AREASTRAB.NOMBRE AS NOMBREAREA, AREASTRAB.CODCCOSTO AS CODAREA, TRABAJADORES.FECHAING, TRABAJADORES.DEPARTAMENTO, TRABAJADORES.CARGO, TRABAJADORES.BASICO, TRABAJADORES.NUMFICHA, TRABAJADORES.CARNETSEG, TRABAJADORES.FONDOPENS, TRABAJADORES.ASIGFAM, TRABAJADORES.FECHACESE, TRABAJADORES.CODIGOALT, TRABAJADORES.TIPOTRAB, TRABAJADORES.SITUACIÓN, TRABAJADORES.CODSCTR, TRABAJADORES.RUCEPS, AFPS.NOMBRE AS NOMBREAFP, TRABAJADORES.BANCO, TRABAJADORES.CTABANCO, TRABAJADORES.TIPDOC, TRABAJADORES.DOCIDEN, TRABAJADORES.DIRECCIÓN, TRABAJADORES.SEXO, TRABAJADORES.ESTADOCIVIL, TRABAJADORES.FECHANAC,TRABAJADORES.CUSPP" & _
                                    " INTO  [##_TMPTRABAJADORES" & VGL_COMPUTER & "]  " & _
                                    " FROM AREASTRAB INNER JOIN (CCOSTOS INNER JOIN (AFPS INNER JOIN TRABAJADORES ON AFPS.CODAFP = TRABAJADORES.FONDOPENS) ON CCOSTOS.CODCCOSTO = TRABAJADORES.CCOSTO) ON AREASTRAB.CODCCOSTO = TRABAJADORES.AREA " & _
                                    " ORDER BY TRABAJADORES.APEPAT+' '+TRABAJADORES.APEMAT+' '+TRABAJADORES.NOMBRE", X
                For Y = 1 To X
                    frWait.Show
                Next
                For Y = 1 To X
                Next
                For Y = 1 To X
                Next
                With RptTrab
                    .Reset
                    .WindowTitle = "REPORTE GENERAL DE TRABAJADORES"
                    .ReportFileName = REGSISTEMA.REPORTES & "GENERAL.RPT"
                    .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
                    .StoredProcParam(0) = " [##_TMPTRABAJADORES" & VGL_COMPUTER & "] "
                    .Destination = crptToWindow
                    .WindowState = crptMaximized
                    .WindowShowPrintBtn = True
                    .WindowShowRefreshBtn = True
                    .WindowShowSearchBtn = True
                    .WindowShowPrintSetupBtn = True
                    .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
                    .Formulas(1) = "XRUC='" & REGSISTEMA.RUC & "'"
                    If .Status <> 2 Then .Action = 1
                End With
                CambiaPanelBD False
                Screen.MousePointer = vbNormal
            Case "PERSO"
                    strTemp = UCase(RSPERSONAL.Source)
                    strTemp = Replace(strTemp, "SELECT", "")
                    strTemp = Replace(strTemp, "*", "")
                    strTemp = Replace(strTemp, "FROM", "")
                    strTemp = Replace(strTemp, "INTO", "")
                    strTemp = Replace(strTemp, INTO, "")
                    If ExisteTablaAux("[##_TMPTRABPERSONALIZADOTOTAL" & VGL_COMPUTER & "]") Then
                            DBSYSTEM.Execute "DROP TABLE [##_TMPTRABPERSONALIZADOTOTAL" & VGL_COMPUTER & "]"
                    End If
                    DBSYSTEM.Execute "SELECT * INTO [##_TMPTRABPERSONALIZADOTOTAL" & VGL_COMPUTER & "] FROM " & strTemp
                    
                    frmTrabPersonalizado.GET_TABLE = "" & strTemp
                    frmTrabPersonalizado.Show 1
                    
    End Select
End Sub

Private Sub TVIEW_NODECLICK(ByVal NODE As MSComctlLib.NODE)
    On Error GoTo ERRNODE
    Set RSPERSONAL = Nothing
    INTO = ""
    If tView.Nodes.Count = 0 Then Exit Sub
    If NODE.KEY = "RAIZ" Then
        Select Case Combo1.ListIndex
            Case 9
                RSPERSONAL.Open "SELECT * FROM VWTRABAJ WHERE SITUACIÓN<='2'", DBSYSTEM, adOpenKeyset, adLockOptimistic
                SqlSel = "SELECT * FROM  VWTRABAJ WHERE NOCALCULO=-1 "
            Case 10
                RSPERSONAL.Open "SELECT * FROM VWTRABAJ WHERE NOCALCULO=-1 ", DBSYSTEM, adOpenKeyset, adLockOptimistic
                SqlSel = "SELECT * FROM  VWTRABAJ WHERE NOCALCULO=-1 "
            Case 11
                RSPERSONAL.Open "SELECT * FROM VWTRABAJ WHERE NOCALCULO=0 ", DBSYSTEM, adOpenKeyset, adLockOptimistic
                SqlSel = "SELECT * FROM  VWTRABAJ WHERE NOCALCULO=0 "
            Case Else
                RSPERSONAL.Open "SELECT * FROM VWTRABAJ" & IIf(FILTRADO, " WHERE SITUACIÓN='" & Combo1.ListIndex & "'", ""), DBSYSTEM, adOpenKeyset, adLockOptimistic
                SqlSel = "SELECT * FROM  VWTRABAJ " & IIf(FILTRADO, " WHERE SITUACIÓN='" & Combo1.ListIndex & "'", "")
        End Select
        
    Else
        If chDepende.Value = 0 Then
            If Option1.Value Then
                Select Case Combo1.ListIndex
                    Case 9
                        SqlSel = "SELECT * " & INTO & " FROM VWTRABAJ WHERE CODAREA='" & Right(NODE.KEY, Len(NODE.KEY) - 1) & "' AND  SITUACIÓN<='2'"
                    Case 10
                        SqlSel = "SELECT * " & INTO & " FROM VWTRABAJ WHERE CODAREA='" & Right(NODE.KEY, Len(NODE.KEY) - 1) & "' AND NOCALCULO=-1 "
                    Case 11
                        SqlSel = "SELECT * " & INTO & " FROM VWTRABAJ WHERE CODAREA='" & Right(NODE.KEY, Len(NODE.KEY) - 1) & "' AND NOCALCULO=0"
                    Case Else
                        SqlSel = "SELECT * " & INTO & " FROM VWTRABAJ WHERE CODAREA='" & Right(NODE.KEY, Len(NODE.KEY) - 1) & "'" & IIf(FILTRADO, " AND SITUACIÓN='" & Combo1.ListIndex & "'", "")
                End Select
                RSPERSONAL.Open SqlSel.Text, DBSYSTEM, adOpenKeyset, adLockOptimistic
            Else
                Select Case Combo1.ListIndex
                    Case 9
                        SqlSel = "SELECT * " & INTO & " FROM VWTRABAJ WHERE CODCCOSTO='" & Right(NODE.KEY, Len(NODE.KEY) - 1) & "' AND SITUACIÓN<='2'"
                    Case 10
                        SqlSel = "SELECT * " & INTO & " FROM VWTRABAJ WHERE CODCCOSTO='" & Right(NODE.KEY, Len(NODE.KEY) - 1) & "' AND NOCALCULO=-1 "
                    Case 10
                        SqlSel = "SELECT * " & INTO & " FROM VWTRABAJ WHERE CODCCOSTO='" & Right(NODE.KEY, Len(NODE.KEY) - 1) & "' AND NOCALCULO=0 "
                    Case Else
                        SqlSel = "SELECT * " & INTO & " FROM VWTRABAJ WHERE CODCCOSTO='" & Right(NODE.KEY, Len(NODE.KEY) - 1) & "'" & IIf(FILTRADO, " AND SITUACIÓN='" & Combo1.ListIndex & "'", "")
                End Select
                RSPERSONAL.Open SqlSel.Text, DBSYSTEM, adOpenKeyset, adLockOptimistic
            End If
        Else
            If Option1.Value Then
                Select Case Combo1.ListIndex
                    Case 9
                        SqlSel = "SELECT * " & INTO & " FROM VWTRABAJ WHERE CODAREA LIKE '" & Right(NODE.KEY, Len(NODE.KEY) - 1) & "%' AND SITUACIÓN<='2'"
                    Case 10
                        SqlSel = "SELECT * " & INTO & " FROM VWTRABAJ WHERE CODAREA LIKE '" & Right(NODE.KEY, Len(NODE.KEY) - 1) & "%' AND NOCALCULO=-1 "
                    Case 11
                        SqlSel = "SELECT * " & INTO & " FROM VWTRABAJ WHERE CODAREA LIKE '" & Right(NODE.KEY, Len(NODE.KEY) - 1) & "%' AND NOCALCULO=0"
                    Case Else
                        SqlSel = "SELECT * " & INTO & " FROM VWTRABAJ WHERE CODAREA LIKE '" & Right(NODE.KEY, Len(NODE.KEY) - 1) & "%'" & IIf(FILTRADO, " AND SITUACIÓN='" & Combo1.ListIndex & "'", "")
                End Select
                RSPERSONAL.Open SqlSel.Text, DBSYSTEM, adOpenKeyset, adLockOptimistic
            Else
                Select Case Combo1.ListIndex
                    Case 9
                        SqlSel = "SELECT * " & INTO & " FROM VWTRABAJ WHERE CODCCOSTO LIKE '" & Right(NODE.KEY, Len(NODE.KEY) - 1) & "%' AND SITUACIÓN<='2'"
                    Case 10
                        SqlSel = "SELECT * " & INTO & " FROM VWTRABAJ WHERE CODCCOSTO LIKE '" & Right(NODE.KEY, Len(NODE.KEY) - 1) & "%' AND NOCALCULO=-1 "
                    Case 11
                        SqlSel = "SELECT * " & INTO & " FROM VWTRABAJ WHERE CODCCOSTO LIKE '" & Right(NODE.KEY, Len(NODE.KEY) - 1) & "%' AND NOCALCULO=0 "
                    Case Else
                        SqlSel = "SELECT * " & INTO & " FROM VWTRABAJ WHERE CODCCOSTO LIKE '" & Right(NODE.KEY, Len(NODE.KEY) - 1) & "%'" & IIf(FILTRADO, " AND SITUACIÓN='" & Combo1.ListIndex & "'", "")
                End Select
                
                RSPERSONAL.Open SqlSel.Text, DBSYSTEM, adOpenKeyset, adLockOptimistic
            End If
        End If
    End If
    Set dgPersonal.DataSource = RSPERSONAL
    If RSPERSONAL.RecordCount = 0 Then
        dgPersonal.Visible = False
    Else
        dgPersonal.Visible = True
        dgPersonal.Caption = "TRABAJADORES DE " & tView.SelectedItem.Text & " (" & RSPERSONAL.RecordCount & " EN TOTAL)"
    End If
    Exit Sub
ERRNODE:
    Exit Sub
End Sub

Public Sub IMPRIMIRFOTOCHECKS()
Dim CnAux As ADODB.Connection
    If RSPERSONAL.RecordCount = 0 Then
        MsgBox "No existen registros de trabajadores por imprimir", vbInformation
        Exit Sub
    End If
    CambiaPanelBD True
    Set CnAux = New ADODB.Connection
    CnAux.Open "Provider = Microsoft.Jet.OLEDB.3.51;Persist Security Info = False ; Data Source =" & App.PATH & "\BDAuxCom.mdb"
    CnAux.Execute "Delete From TMPFOTOCHECKS "
    
    Data1.DatabaseName = App.PATH & "\BDAuxCom.mdb"
    Data1.RecordSource = "TMPFOTOCHECKS"
    Data1.Refresh

    Dim XBOOK
    If dgPersonal.SelBookmarks.Count = 0 Then
        RSPERSONAL.MoveFirst
        Do While Not RSPERSONAL.EOF
            With Data1.Recordset
                .AddNew
                !CODTRAB = RSPERSONAL!CODTRAB
                !APELLIDO = Trim(DevuelveValor("SELECT APEPAT FROM TRABAJADORES WHERE CODTRAB='" & RSPERSONAL!CODTRAB & "'", DBSYSTEM)) & " " & Trim(DevuelveValor("SELECT APEMAT FROM TRABAJADORES WHERE CODTRAB='" & RSPERSONAL!CODTRAB & "'", DBSYSTEM))
                !NOMBRE = Trim(DevuelveValor("SELECT NOMBRE FROM TRABAJADORES WHERE CODTRAB='" & RSPERSONAL!CODTRAB & "'", DBSYSTEM))
                !CENTRO = Left(RSPERSONAL!CENTRO, 25)
                !CODCCOSTO = RSPERSONAL!CODCCOSTO
                !NOMBREAREA = RSPERSONAL!NOMBREAREA
                !CODAREA = RSPERSONAL!CODAREA
                !CARGO = RSPERSONAL!CARGO
                !TIPDOC = RSPERSONAL!TIPDOC
                !DOCIDEN = RSPERSONAL!DOCIDEN
                'FOTO
                If UCase(Dir$(REGSISTEMA.PATHFOTOS & "\" & RSPERSONAL!CODTRAB & ".FTE")) = UCase(RSPERSONAL!CODTRAB & ".FTE") Then
                    Set Ole1.Picture = LoadPicture(REGSISTEMA.PATHFOTOS & "\" & RSPERSONAL!CODTRAB & ".FTE")
                Else
                    Set Ole1.Picture = LoadPicture(REGSISTEMA.PATH & "\OBJBLANK.BMP")
                End If
                .Update
            End With
            RSPERSONAL.MoveNext
        Loop
    Else
        For Each XBOOK In dgPersonal.SelBookmarks
            RSPERSONAL.Bookmark = XBOOK
            With Data1.Recordset
                .AddNew
                !CODTRAB = RSPERSONAL!CODTRAB
                !APELLIDO = Trim(DevuelveValor("SELECT APEPAT FROM TRABAJADORES WHERE CODTRAB='" & RSPERSONAL!CODTRAB & "'", DBSYSTEM)) & " " & Trim(DevuelveValor("SELECT APEMAT FROM TRABAJADORES WHERE CODTRAB='" & RSPERSONAL!CODTRAB & "'", DBSYSTEM))
                !NOMBRE = Trim(DevuelveValor("SELECT NOMBRE FROM TRABAJADORES WHERE CODTRAB='" & RSPERSONAL!CODTRAB & "'", DBSYSTEM))
                !CENTRO = Left(RSPERSONAL!CENTRO + String(25, " "), 25)
                !CODCCOSTO = RSPERSONAL!CODCCOSTO
                !NOMBREAREA = RSPERSONAL!NOMBREAREA
                !CODAREA = RSPERSONAL!CODAREA
                !CARGO = RSPERSONAL!CARGO
                !TIPDOC = RSPERSONAL!TIPDOC
                !DOCIDEN = RSPERSONAL!DOCIDEN
                If UCase(Dir$(REGSISTEMA.PATHFOTOS & "\" & RSPERSONAL!CODTRAB & ".FTE")) = UCase(RSPERSONAL!CODTRAB & ".FTE") Then
                    Set Ole1.Picture = LoadPicture(REGSISTEMA.PATHFOTOS & "\" & RSPERSONAL!CODTRAB & ".FTE")
                Else
                    Set Ole1.Picture = LoadPicture(REGSISTEMA.PATH & "\OBJBLANK.BMP")
                End If
                .Update
            End With
        Next
    End If
    RptTrab.Reset
    RptTrab.ReportFileName = REGSISTEMA.REPORTES & "PLAN0024.RPT"
    RptTrab.DataFiles(0) = App.PATH & "\BDAUXCOM.MDB"
    RptTrab.Destination = crptToWindow
    RptTrab.WindowState = crptMaximized
    RptTrab.WindowShowPrintBtn = True
    RptTrab.WindowShowRefreshBtn = True
    RptTrab.WindowShowSearchBtn = True
    RptTrab.WindowShowPrintSetupBtn = True
    RptTrab.WindowTitle = "PLAN0024 - FOTOCHECKS "
    RptTrab.Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
    A = InputBox("DEBERÁ INGRESAR UNA FECHA DE VENCIMIENTO PARA EL FOTOCHECK", "FOTOCHECK - FECHA DE VENCIMIENTO", "" & Date)
    RptTrab.Formulas(1) = "VECIMIENTO='VENCE: " & A & "'"
    RptTrab.DiscardSavedData = True
    If RptTrab.Status <> 2 Then RptTrab.Action = 1
    Screen.MousePointer = 1
    CambiaPanelBD False
End Sub


