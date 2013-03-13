VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frPlans 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planillas de Remuneraciones Procesadas"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8475
   Icon            =   "frPlans.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   8475
   Tag             =   "Panel de Planillas de Remuneraciones"
   Begin VB.Frame Frame1 
      Caption         =   "Otro Filtro"
      Height          =   900
      Left            =   6060
      TabIndex        =   11
      Top             =   4710
      Width           =   2280
      Begin VB.OptionButton Option1 
         Caption         =   "Declarados en PDT"
         Height          =   210
         Left            =   210
         TabIndex        =   13
         Top             =   330
         Value           =   -1  'True
         Width           =   1770
      End
      Begin VB.OptionButton Option2 
         Caption         =   "No declarados en PDT"
         Height          =   225
         Left            =   195
         TabIndex        =   12
         Top             =   570
         Width           =   2025
      End
   End
   Begin VB.CommandButton cmImprimir 
      Caption         =   "&Imprimir con Filtro"
      Height          =   390
      Left            =   6375
      TabIndex        =   10
      Top             =   5700
      Width           =   1650
   End
   Begin VB.Frame Frame3 
      Height          =   540
      Left            =   75
      TabIndex        =   9
      Top             =   5550
      Width           =   2520
   End
   Begin VB.Frame Frame2 
      Caption         =   "Filtro por Centro de Costo"
      Height          =   1395
      Left            =   2700
      TabIndex        =   7
      Top             =   4695
      Width           =   3300
      Begin VB.ComboBox xNivel 
         Height          =   315
         ItemData        =   "frPlans.frx":0442
         Left            =   1650
         List            =   "frPlans.frx":0455
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   945
         Width           =   1545
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Detallar Centros de Costos"
         Height          =   225
         Left            =   135
         TabIndex        =   14
         Top             =   645
         Value           =   1  'Checked
         Width           =   2340
      End
      Begin AplisetControlText.Aplitext xArea 
         Height          =   285
         Left            =   135
         TabIndex        =   8
         Top             =   270
         Width           =   3060
         _ExtentX        =   5398
         _ExtentY        =   503
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nivel de Impresión:"
         Height          =   210
         Left            =   135
         TabIndex        =   15
         Top             =   990
         Width           =   1350
      End
   End
   Begin Crystal.CrystalReport RptPlan 
      Left            =   3270
      Top             =   2265
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ProgressBar Prog 
      Height          =   150
      Left            =   105
      TabIndex        =   5
      Top             =   5085
      Visible         =   0   'False
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   3735
      Top             =   2250
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   25
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":0486
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":0D2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":1182
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":15D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":192A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":2206
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":2AE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":413E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":445E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":47B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":508E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":596A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":5DBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":6112
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":6466
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":68BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":7196
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":75EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":7EC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":831A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":866E
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":89C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":97E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":A0C2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3075
      Top             =   2265
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   25
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":A3DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":A832
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":AC86
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":B0DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":B52E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":B882
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":C15E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":CA3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":E096
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":E3B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":E70A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":EFE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":F8C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":FD16
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":1006A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":103BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":10812
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":110EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":11542
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":11E1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":12272
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":125C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":1291A
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":1373E
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frPlans.frx":1401A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LProcs 
      Height          =   4245
      Left            =   75
      TabIndex        =   1
      Top             =   360
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   7488
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList2"
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nombre de Función"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView LPlans 
      Height          =   4245
      Left            =   2655
      TabIndex        =   0
      Top             =   360
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   7488
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList2"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Mes"
         Text            =   "Mes"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Fecha"
         Text            =   "Fecha de Proceso"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Autor"
         Text            =   "Autor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   "NumTrab"
         Text            =   "Versión"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lProg 
      AutoSize        =   -1  'True
      Caption         =   "Espere, ejecutando tareas"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   105
      TabIndex        =   6
      Top             =   4845
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Label CCosto 
      BackColor       =   &H80000010&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Todos los Centros de Costos"
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   3285
      TabIndex        =   4
      Top             =   90
      Width           =   5055
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Filtro"
      Height          =   270
      Left            =   2670
      TabIndex        =   3
      Top             =   90
      Width           =   600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Funciones Activas"
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   90
      TabIndex        =   2
      Top             =   90
      Width           =   2565
   End
End
Attribute VB_Name = "frPlans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim XITEM As ListItem
Dim RSPLANS As ADODB.Recordset
Dim REGACT As REGWIN
Public ASAKD As String

Private Sub CCOSTO_DblClick()
    Dim RSCCOSTOS As New ADODB.Recordset
    RSCCOSTOS.Open "SELECT CODCCOSTO,NOMBRE FROM CCOSTOS ORDER BY CODCCOSTO", DBSYSTEM, adOpenKeyset, adLockOptimistic
    frmComun.CONECTAR RSCCOSTOS
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        CCosto.Caption = VGUTIL(1) & " :  " & VGUTIL(2)
        CCosto.Tag = VGUTIL(1)
    End If
    Set RSCCOSTOS = Nothing
End Sub
Private Sub CMIMPRIMIR_CLICK()
            'IMPRIMIR CON FILTRO
            If LPlans.ListItems.Count = 0 Then
                MsgBox "NO EXISTEN PLANILLAS PROCESADAS. IMPOSIBLE ELIMINAR", vbCritical
                Exit Sub
            End If
            Screen.MousePointer = 11
            If ExisteTablaAux(" [##PLANILLA" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##PLANILLA" & VGL_COMPUTER & "] "
            DBSYSTEM.Execute "SELECT * INTO  [##PLANILLA" & VGL_COMPUTER & "]  FROM PLAN2000"
            If xArea.Text <> "" Then
                DBSYSTEM.Execute "DELETE FROM PLAN2000 WHERE CCOSTO NOT LIKE '" & xArea.Tag & "%'"
            End If
            If Option1.Value Then
                DBSYSTEM.Execute "DELETE FROM PLAN2000 WHERE CODTRAB NOT IN (SELECT CODTRAB FROM TRABAJADORES WHERE NOPDT=0)"
            Else
                DBSYSTEM.Execute "DELETE FROM PLAN2000 WHERE CODTRAB IN (SELECT CODTRAB FROM TRABAJADORES WHERE NOPDT=0)"
            End If
            Dim RSAUX As New ADODB.Recordset, RSCCOSTOS As New ADODB.Recordset
            Dim XSTR As String, XANT As String
            If xNivel.ListIndex <> -1 Then
                RSAUX.Open "SELECT DISTINCT CCOSTO FROM PLAN2000", DBSYSTEM, adOpenStatic, adLockOptimistic
                RSCCOSTOS.Open "SELECT CODCCOSTO,NOMBRE FROM CCOSTOS", DBSYSTEM, adOpenStatic, adLockReadOnly
                If RSCCOSTOS.EOF Or RSCCOSTOS.RecordCount = 0 Then
                    MsgBox "NO SE HAN ENCONTRADO CENTROS DE COSTOS, EL SISTEMA BLOQUEARA EL PROCESO", vbCritical
                    cmImprimir.Enabled = False
                    Set RSAUX = Nothing
                    Set RSCCOSTOS = Nothing
                    Exit Sub
                End If
                Set RSAUX.ActiveConnection = Nothing
                Do While Not RSAUX.EOF
                    XANT = RSAUX!CCosto
                    XSTR = Getcad(".", xNivel.ListIndex + 1, RSAUX!CCosto)
                    If XANT <> XSTR Then
                        RSCCOSTOS.MoveFirst
                        RSCCOSTOS.FIND "CODCCOSTO='" & XSTR & "'"
                        If RSCCOSTOS.EOF Then
                            DBSYSTEM.Execute "UPDATE PLAN2000 SET CCOSTO='X_X_X',CENTROCOSTO='* SIN CENTRO DE COSTO *' WHERE CCOSTO='" & RSAUX!CCosto & "'"
                        Else
                            DBSYSTEM.Execute "UPDATE PLAN2000 SET CCOSTO='" & XSTR & "',CENTROCOSTO='" & RSCCOSTOS!NOMBRE & "' WHERE CCOSTO='" & RSAUX!CCosto & "'"
                        End If
                    End If
                    RSAUX.MoveNext
                Loop
                Set RSAUX = Nothing
                Set RSCCOSTOS = Nothing
            End If
            With RptPlan
                Dim STRFILE As String
                STRFILE = DevNomRep(Trim(REGSISTEMA.RUC), Trim(REGSISTEMA.USER), FILEPLANILLA)
                STRFILE = Left(STRFILE, 4) & "F" & Right(STRFILE, 7)
                .Connect = "DSN=" & VGL_SERVER & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL & ""
                If UCase(Dir$(REGSISTEMA.REPORTES & STRFILE)) <> UCase(STRFILE) Then
                    MsgBox "NO SE ENCUENTRA EL ARCHIVO DE REPORTE PARA ESTA PLANILLA " & _
                           "COMUNIQUESE CON ENTERPRISE", vbInformation
                    Screen.MousePointer = 1
                    Exit Sub
                Else
                    .ReportFileName = REGSISTEMA.REPORTES & STRFILE
                End If
                .Destination = crptToWindow
                .WindowShowPrintBtn = True
                .WindowShowRefreshBtn = True
                .WindowShowSearchBtn = True
                .WindowShowPrintSetupBtn = True
                .WindowState = crptMaximized
                .WindowTitle = .ReportFileName
                .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & IIf(xArea.Text = "", "", ": " & xArea.Text) & "'"
                .Formulas(1) = "XRUC='RUC N° " & REGSISTEMA.RUC & "'"
                .Formulas(2) = "XMES='CORRESPONDIENTE AL MES DE " & LPlans.SelectedItem.Text & "'"
                .Formulas(3) = "XDIRECCION='" & DevuelveValor("SELECT DIRECCIÓN FROM EMPRESA", DBSYSTEM) & "'"
                If RptPlan.Status <> 2 Then .Action = 1
            End With
            DBSYSTEM.Execute "DROP TABLE PLAN2000"
            DBSYSTEM.Execute "SELECT * INTO PLAN2000 FROM PLANILLA IN '" & App.PATH & "\BDAUXCOM.MDB'"
            DBSYSTEM.Execute "DROP TABLE PLANILLA"
            Screen.MousePointer = 1
End Sub

Private Sub Form_Activate()
    ActivarTools REGACT
End Sub

Private Sub Form_Load()
    With REGACT
        .BUSCAR = False
        .EDITAR = False
        .ELIMINAR = True
        .FILTRAR = True
        .IMPRIMIR = True
        .NUEVO = True
        .PRELIMINAR = True
    End With
    CARGARDEFAULT
    CARGAPLANS
End Sub


Public Sub CARGARDEFAULT()
    LProcs.ListItems.Clear
    Set XITEM = LProcs.ListItems.Add(, "SYS_REPROC", "PROCESAR PLANILLA", 7, 7)
    Set XITEM = LProcs.ListItems.Add(, "SYS_VER", "VER PLANILLA", 24, 24)
    Set XITEM = LProcs.ListItems.Add(, "SYS_RESUMEN", "RESUMEN DE PLANILLA", 13, 13)
    Set XITEM = LProcs.ListItems.Add(, "SYS_GRAFICA", "GRAFICA ESTADÍSTICA", 14, 14)
    Set XITEM = LProcs.ListItems.Add(, "SYS_PDTREMU", "EXPORTA PDT REMUNERACIONES", 23, 23)
    Set XITEM = LProcs.ListItems.Add(, "SYS_PDTSCTR", "EXPORTA PDT S.C.T.R.", 23, 23)
    Set XITEM = LProcs.ListItems.Add(, "SYS_PDTTRAB", "TRABAJADORES PDT SUNAT", 23, 23)
    Set XITEM = LProcs.ListItems.Add(, "SYS_PDTDERE", "Exporta DerechoHabientes al PDT", 23, 23)
    Set XITEM = LProcs.ListItems.Add(, "SYS_AFP", "PLANILLAS AFP", 25, 25)
    Set XITEM = LProcs.ListItems.Add(, "SYS_BACKUP", "COPIA DE SEGURIDAD", 11, 11)
    LProcs.ColumnHeaders(1).Width = 2505.2674
End Sub

Public Sub CARGAPLANS()
    Dim xCad As String
    Set RSPLANS = New ADODB.Recordset
    RSPLANS.Open "SELECT * FROM PLANILLAS ORDER BY MES", DBSYSTEM, adOpenStatic
    LPlans.ListItems.Clear
    With RSPLANS
        Do While Not .EOF
            xCad = "M" & Format(Month(!MES), "00") & Year(!MES)
            Set XITEM = LPlans.ListItems.Add(, xCad, AMESES(Month(!MES)) & " DE " & Year(!MES), 1, 1)
            XITEM.SubItems(1) = !MES
            XITEM.SubItems(2) = !AUTOR
            XITEM.SubItems(3) = Format(12, "0.00")
            .MoveNext
        Loop
    End With
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSPLANS = Nothing
End Sub

Private Sub LPROCS_DblClick()
    On Error GoTo ERRPROCS
    If LPlans.ListItems.Count = 0 Then
        MsgBox "NO EXISTEN PERIODOS PARA LAS FUNCIONES DE PLANILLA", vbInformation
        Exit Sub
    End If
    Select Case LProcs.SelectedItem.KEY
        Case "SYS_GRAFICA"
            frCfgGraf.Show 1
        Case "SYS_REPROC"
            'REPROCESO DE PLANILLAS DE REMUNERACIONES. JALA TODOS LOS REGISTROS
            REPROCESARPLANILLA
        Case "SYS_VER"
            VPTAREA = LPlans.SelectedItem.Text
            Load frVerPlan
            frVerPlan.CARGAPLAN "SELECT * FROM " & REGSISTEMA.TABLAPLAN & " WHERE MES=" & DateSQL(FechaMMAAAA(Right(LPlans.SelectedItem.KEY, 6)))
            frVerPlan.Show
        Case "SYS_RESUMEN"
            VPTAREA = "PLANILLAS"
            frSuma.Show 1
        Case "SYS_AFP"
            VPTRASPRM = ""
            VPTAREA = LPlans.SelectedItem.SubItems(1)
            frCCAR.Show 1
            If VPTRASPRM <> "" Then frPlanAFP.Show
        Case "SYS_PDTREMU"
            PDTSUNAT
        Case "SYS_PDTSCTR"
            PDTSUNATSCTR
        Case "SYS_PDTTRAB"
            PDTTRABAJADORES
        Case "SYS_PDTDERE"
            PDTDerecho
    End Select
    Exit Sub
ERRPROCS:
    MsgBox "ERROR EN FUNCIONES DE PLANILLAS: " & ERR.Description & ":" & ERR.Number
    Screen.MousePointer = 1
    Resume Next
    Resume
    Exit Sub
End Sub

Public Sub COMANDOTOOLBAR(ByVal COMANDO As String)
    Select Case UCase(COMANDO)
        Case "NUEVO"
            REPROCESARPLANILLA
        Case "ELIMINAR"
            If LPlans.ListItems.Count = 0 Then
                MsgBox "NO EXISTEN PLANILLAS PROCESADAS. IMPOSIBLE ELIMINAR", vbCritical
                Exit Sub
            End If
            If MsgBox("SEGURO DE ELIMINAR LA PLANILLA SELECCIONADA", vbYesNo + vbQuestion) = vbNo Then Exit Sub
            DBSYSTEM.Execute "DELETE FROM PLANILLAS WHERE MES=" & DateSQL(FechaMMAAAA(Right(LPlans.SelectedItem.KEY, 6)))
            DBSYSTEM.Execute "DELETE FROM " & REGSISTEMA.TABLAPLAN & " WHERE MES=" & DateSQL(FechaMMAAAA(Right(LPlans.SelectedItem.KEY, 6)))
            CARGAPLANS
        Case "IMPRIMIR", "PRELIMINAR"
            If LPlans.ListItems.Count = 0 Then
                MsgBox "NO EXISTEN PLANILLAS PROCESADAS. IMPOSIBLE ELIMINAR", vbCritical
                Exit Sub
            End If
            With RptPlan
                .Reset
                Dim STRFILE As String
                STRFILE = DevNomRep(Trim(REGSISTEMA.RUC), Trim(REGSISTEMA.USER), FILEPLANILLA)
                If STRFILE = "" Then
                    MsgBox "EL REGISTRO DE INFORMACIÓN GENERAL DE EMPRESAS NO ES VÁLIDO", vbCritical
                    Exit Sub
                End If
                If UCase(Dir$(REGSISTEMA.REPORTES & STRFILE)) <> UCase(STRFILE) Then
                    MsgBox "NO SE HA ENCONTRADO EL REPORTE.. SE TRATARA DE USAR EL PREDETERMINADO DEL SISTEMA", vbInformation, "FALTA: " & STRFILE
                    Exit Sub
                Else
                    .ReportFileName = REGSISTEMA.REPORTES & STRFILE
                End If
                
                .Connect = "DSN=" & VGL_SERVER & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL & ""
                .SelectionFormula = "{PLAN2000.MES} = DATE(" & Right(LPlans.SelectedItem.KEY, 4) & "," & Mid(LPlans.SelectedItem.KEY, 2, 2) & ",01)"
                .Destination = IIf(UCase(COMANDO) = "PRELIMINAR", crptToWindow, crptToPrinter)
                .WindowShowPrintBtn = True
                .WindowShowRefreshBtn = True
                .WindowShowSearchBtn = True
                .WindowShowPrintSetupBtn = True
                .WindowState = crptMaximized
                .WindowTitle = .ReportFileName & " : PLANILLA DE REMUNERACIONES"
                .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
                .Formulas(1) = "XRUC='RUC N° " & REGSISTEMA.RUC & "'"
                .Formulas(2) = "XMES='CORRESPONDIENTE AL MES DE " & LPlans.SelectedItem.Text & "'"
                .Formulas(3) = "XDIRECCION='" & DevuelveValor("SELECT DIRECCIÓN FROM EMPRESA", DBSYSTEM) & "'"
                On Error GoTo ERRPLANFORMATO
                If RptPlan.Status <> 2 Then .Action = 1
            End With
    End Select
    Exit Sub
ERRPLANFORMATO:
    MsgBox "EL FORMATO DE PLANILLA NO ES EL CORRECTO O SE ENCUENTRA DAÑADO. COMUNICARSE CON ENTERPRISE SOLUTIONS", vbCritical
    Resume Next
End Sub

Public Sub REPROCESARPLANILLA()
    If Not ExisteCampo("CUSPP", "PLAN2000", DBSYSTEM) Then
        MsgBox "DEBERA VOLVER GENERAR LA TABLA DE PLANILLAS. PULSE SOBRE EL BOTON GENERAR TABLA", vbInformation
        frColPL.Show
        Exit Sub
    End If
    If Not ExisteCampo("TOTING", "PLAN2000", DBSYSTEM) Then
        MsgBox "DEBERA VOLVER GENERAR LA TABLA DE PLANILLAS. PULSE SOBRE EL BOTON GENERAR TABLA", vbInformation
        frColPL.Show
        Exit Sub
    End If
    If Not ExisteCampo("TOTEGR", "PLAN2000", DBSYSTEM) Then
        MsgBox "DEBERA VOLVER GENERAR LA TABLA DE PLANILLAS. PULSE SOBRE EL BOTON GENERAR TABLA", vbInformation
        frColPL.Show
        Exit Sub
    End If
    If Not ExisteCampo("NETO", "PLAN2000", DBSYSTEM) Then
        MsgBox "DEBERA VOLVER GENERAR LA TABLA DE PLANILLAS. PULSE SOBRE EL BOTON GENERAR TABLA", vbInformation
        frColPL.Show
        Exit Sub
    End If
    
    Dim xCont As Boolean
    xCont = False
    If Not COMPRUEBAPLAN Then Exit Sub
    Dim VMES As Date
    Dim RSMESES As New ADODB.Recordset
    Dim VARCODE As Long
    RSMESES.Open "EMPRESA", DBSYSTEM, adOpenStatic, adLockReadOnly
    If RSMESES.RecordCount = 0 Then
        MsgBox "SE HA ENCONTRADO UN PROBLEMA EN LA DEFINICIÓN DE LA TABLA EMPRESA", vbCritical
        Set RSMESES = Nothing
        Exit Sub
    Else
        If Not IsNull(RSMESES!ADELPLAN) Then
            REGSISTEMA.COLPLANADEL = RSMESES!ADELPLAN
        End If
    End If
    Set RSMESES = Nothing
    If Not ExisteCampo(REGSISTEMA.COLPLANADEL, "PLAN2000", DBSYSTEM) Then
        MsgBox "NO SE ENCUENTRA EL CAMPO CORRESPONDIENTE A ADELANTOS DE REMUNERACIONES. LOS ADELANTOS DE REMUNERACIONES DEBEN DE SER ALMACENADOS EN UNA COLUMNA DE PLANILLA, LA CUAL NO EXISTE. DEFINA LA COLUMNA EN EL PANEL DE CONFIGURACIÓN DEL SISTEMA", vbInformation
        Exit Sub
    End If
    RSMESES.Open "SELECT MESACTIVO, NOMBRE FROM MESESACT ORDER BY MESACTIVO", DBSYSTEM, adOpenKeyset, adLockOptimistic
    frmComun.CONECTAR RSMESES
    frmComun.Show 1
    Set RSMESES = Nothing
    'SI ES CONTINUAR
    If VGUTIL(1) <> "" Then
        If MsgBox("DESEA PROCESAR LA PLANILLA CORRESPONDIENTE AL MES DE " & AMESES(Month(VGUTIL(1))) & " DE " & Year(VGUTIL(1)), vbYesNo + vbQuestion) = vbNo Then Exit Sub
        Dim REGPLAN As TYPEREGPLAN
        With REGPLAN
            .AUTOR = REGSISTEMA.USER
            .DATABASE = "MASTER" 'SIGNIFICA QUE ES LA QUE ESTÁ ACTIVA
            .FECHA = Date
            .MES = VGUTIL(1)
            .TABLABOL = "BOL" & Format(Month(.MES), "00") & Year(.MES)
            .TABLAMOV = "MOV" & Format(Month(.MES), "00") & Year(.MES)
        End With
        'REGISTRO DE LA PLANILLA EN LA BASE DE DATOS : PLANILLAS
        If RSPLANS.RecordCount > 0 Then
            RSPLANS.MoveFirst
            RSPLANS.FIND "MES=" & DateSQL(REGPLAN.MES)
            If Not RSPLANS.EOF Then
                If MsgBox("YA EXISTE UNA PLANILLA DEL MES INDICADO: " & AMESES(Month(VGUTIL(1))) & " DE " & Year(VGUTIL(1)) & ", DESEA REEMPLAZARLA", vbYesNo + vbQuestion) = vbNo Then Exit Sub
            Else
                DBSYSTEM.Execute "INSERT INTO PLANILLAS (MES, DESCRIPCIÓN, TIPO, AUTOR, FECHA, ALMACEN, PROTEGIDA) VALUES (" & DateSQL(REGPLAN.MES) & ", 'MES DE " & AMESES(Month(VGUTIL(1))) & " DE " & Year(VGUTIL(1)) & "',0,'" & REGSISTEMA.USER & "', " & DateSQL(Date) & ",'" & REGPLAN.DATABASE & "',0)"
            End If
        Else
            DBSYSTEM.Execute "INSERT INTO PLANILLAS (MES, DESCRIPCIÓN, TIPO, AUTOR, FECHA, ALMACEN, PROTEGIDA) VALUES (" & DateSQL(REGPLAN.MES) & ", 'MES DE " & AMESES(Month(VGUTIL(1))) & " DE " & Year(VGUTIL(1)) & "',0,'" & REGSISTEMA.USER & "', " & DateSQL(Date) & ",'" & REGPLAN.DATABASE & "',0)"
        End If
        DBSYSTEM.Execute "DELETE FROM " & REGSISTEMA.TABLAPLAN & " WHERE MES=" & DateSQL(REGPLAN.MES)
        'PROCESO DE APERTURA DE TABLAS DE BOLETAS Y MOVIMIENTOS
        Prog.Visible = True
        lProg.Visible = True
        Dim RSBOLS As New ADODB.Recordset
        Dim RSPLAN2 As New ADODB.Recordset
        Dim RSMOVS As New ADODB.Recordset
        '---------------------------------------------------------------
        RSPLAN2.Open "SELECT MES,TIPOPLANILLA,CODTRAB,NOMBRES,TIPOTRAB,FECHAING,SITUACION,CCOSTO,CENTROCOSTO,DEPARTAMENTO,CARGO,BASICO,FONDOPENS,FECHACESE,CODSCTR,EPS,INUMBOL,CARNETSEG,CUSPP,VACINI,VACFIN FROM " & REGSISTEMA.TABLAPLAN & " WHERE MES=" & DateSQL(REGPLAN.MES), DBSYSTEM, adOpenDynamic, adLockOptimistic
        RSBOLS.Open "SELECT BOL.CODNOMBOL, A.CODTRAB, A.NOMBRES, INUMBOL, TIPOPLAN, TOTING, TOTEGR, A.TIPOTRAB, A.FECHAING, SITUACIÓN, BOL.CCOSTO, A.CENTRO, A.DEPARTAMENTO, A.CARGO, BOL.BASICO, BOL.CODAFP, A.FECHACESE, A.CODSCTR, A.RUCEPS, BOL.XREDONDEO FROM VWTRABAJ A, " & REGPLAN.TABLABOL & " BOL WHERE BOL.CODTRAB=A.CODTRAB AND BOL.CODNOMBOL IN (SELECT CODIGO FROM NOMBOL WHERE MES=" & DateSQL(REGPLAN.MES) & " ) ORDER BY NOMBRES", DBSYSTEM, adOpenStatic
        lProg.Caption = "1.- ASIGNANDO REMUNERACIONES"
        Prog.Max = RSBOLS.RecordCount + 1
        Prog.Value = 0
        Do While Not RSBOLS.EOF
            Prog.Value = Prog.Value + 1
            If Not RSPLAN2.EOF Then
                RSPLAN2.MoveFirst
                RSPLAN2.FIND "CODTRAB='" & RSBOLS!CODTRAB & "'"
            End If
            If RSPLAN2.EOF Then
                RSPLAN2.AddNew
                RSPLAN2!MES = REGPLAN.MES
                RSPLAN2!TIPOPLANILLA = RSBOLS!TIPOPLAN
                RSPLAN2!CODTRAB = Trim(RSBOLS!CODTRAB)
                RSPLAN2!NOMBRES = Trim(Left(RSBOLS!NOMBRES & String(35, " "), 35))
                RSPLAN2!TIPOTRAB = Trim(RSBOLS!TIPOTRAB)
                RSPLAN2!FECHAING = CDate(RSBOLS!FECHAING)
                RSPLAN2!SITUACION = Trim(RSBOLS!SITUACIÓN)
                RSPLAN2!CCosto = Trim(RSBOLS!CCosto)
                RSPLAN2!CENTROCOSTO = Trim(Left(RSBOLS!CENTRO & String(25, " "), 25))
                RSPLAN2!DEPARTAMENTO = Trim(RSBOLS!DEPARTAMENTO)
                RSPLAN2!CARGO = Trim(RSBOLS!CARGO)
                RSPLAN2!BASICO = RSBOLS!BASICO
                RSPLAN2!FONDOPENS = Trim(RSBOLS!CODAFP)
                If Not IsNull(RSBOLS!FECHACESE) Then
                     RSPLAN2!FECHACESE = CDate(RSBOLS!FECHACESE)
                End If
                RSPLAN2!CODSCTR = Trim(RSBOLS!CODSCTR)
                RSPLAN2!EPS = Trim(IIf(RSBOLS!RUCEPS = "", " ", RSBOLS!RUCEPS))
                If DevuelveValor("SELECT CARNETSEG FROM TRABAJADORES WHERE CODTRAB='" & RSBOLS!CODTRAB & "'", DBSYSTEM) = "" Then
                    RSPLAN2!CARNETSEG = " "
                Else
                   RSPLAN2!CARNETSEG = "" & DevuelveValor("SELECT CARNETSEG FROM TRABAJADORES WHERE CODTRAB='" & RSBOLS!CODTRAB & "'", DBSYSTEM)
                End If
                If DevuelveValor("SELECT CUSPP FROM TRABAJADORES WHERE CODTRAB='" & RSBOLS!CODTRAB & "'", DBSYSTEM) = "" Then
                    RSPLAN2!CUSPP = ""
                Else
                    RSPLAN2!CUSPP = "" & DevuelveValor("SELECT CUSPP FROM TRABAJADORES WHERE CODTRAB='" & RSBOLS!CODTRAB & "'", DBSYSTEM)
                End If
            End If
            RSPLAN2!INUMBOL = RSBOLS!INUMBOL
            If Not IsNull(DevuelveValor("SELECT CODIGO FROM HISTOVAC WHERE CERRADO=1 AND NOMBOL=" & RSBOLS!CODNOMBOL & " AND CODTRAB='" & RSBOLS!CODTRAB & "'", DBSYSTEM)) Then
                VARCODE = DevuelveValor("SELECT CODIGO FROM HISTOVAC WHERE CERRADO=1 AND NOMBOL=" & RSBOLS!CODNOMBOL & " AND CODTRAB='" & RSBOLS!CODTRAB & "'", DBSYSTEM)
                If VARCODE <> 0 Then
                    RSPLAN2!VACINI = DevuelveValor("SELECT FECHAINI FROM HISTOVAC WHERE CODIGO=" & VARCODE, DBSYSTEM)
                    RSPLAN2!VACFIN = DevuelveValor("SELECT FECHAFIN FROM HISTOVAC WHERE CODIGO=" & VARCODE, DBSYSTEM)
                End If
            End If
            RSPLAN2.Update
            DBSYSTEM.Execute "UPDATE PLAN2000 SET TOTING=(SELECT CASE WHEN TOTING IS NULL THEN 0 ELSE TOTING END ) + " & IIf(IsNull(RSBOLS!TOTING), 0, RSBOLS!TOTING) & " , " & _
                           " TOTEGR =(SELECT CASE WHEN TOTEGR IS NULL  THEN 0 ELSE TOTEGR END ) + " & IIf(IsNull(RSBOLS!TOTEGR), 0, RSBOLS!TOTEGR) & "," & _
                           " REDONDEO = (SELECT CASE WHEN REDONDEO IS NULL THEN 0 ELSE REDONDEO END ) + " & IIf(IsNull(RSBOLS!XREDONDEO), 0, RSBOLS!XREDONDEO) & "" & _
                           " WHERE INUMBOL=" & RSBOLS!INUMBOL & " AND MES=" & DateSQL(REGPLAN.MES)
            '"UPDATE PLAN2000 SET TOTING=IIF(ISNULL(TOTING),0,TOTING)+" & RSBOLS!TOTING & ",TOTEGR=IIF(ISNULL(TOTEGR),0,TOTEGR)+" & RSBOLS!TOTEGR & ",REDONDEO=IIF(ISNULL(REDONDEO),0,REDONDEO)+" & RSBOLS!xRedondeo & " WHERE INUMBOL=" & RSBOLS!INUMBOL & " AND MES=" & DateSQL(REGPLAN.MES)
            RSMOVS.Open "SELECT COLPLANILLA, MONTO FROM " & REGPLAN.TABLAMOV & " MOV, CONCEPTOS WHERE MOV.CONCEPTO=CONCEPTOS.CODIGO AND INUMBOL=" & RSBOLS!INUMBOL, DBSYSTEM, adOpenStatic
            Do While Not RSMOVS.EOF
                If Trim(RSMOVS!COLPLANILLA) <> "" Then DBSYSTEM.Execute "UPDATE PLAN2000 SET " & Trim$(RSMOVS!COLPLANILLA) & "=(SELECT CASE WHEN " & Trim$(RSMOVS!COLPLANILLA) & " IS NULL THEN 0 ELSE " & Trim$(RSMOVS!COLPLANILLA) & " END)+" & RSMOVS!MONTO & " WHERE INUMBOL=" & RSBOLS!INUMBOL & " AND MES=" & DateSQL(REGPLAN.MES)
                RSMOVS.MoveNext
            Loop
            RSMOVS.Close
            RSBOLS.MoveNext
        Loop
        Set RSMOVS = Nothing
        If ExisteCampo("NETO", "PLAN2000", DBSYSTEM) Then
            DBSYSTEM.Execute "UPDATE PLAN2000 SET NETO=TOTING-TOTEGR WHERE MES=" & DateSQL(REGPLAN.MES)
        Else
            DBSYSTEM.Execute "UPDATE PLAN2000 SET NETOPAGO=TOTING-TOTEGR WHERE MES=" & DateSQL(REGPLAN.MES)
        End If
        '---------------------------------------------------------------
        'COLOCANDO CERO A LOS VALORES NULOS
        RSBOLS.Close
        RSBOLS.Open "COLUMPL", DBSYSTEM, adOpenStatic
        Do While Not RSBOLS.EOF
            DBSYSTEM.Execute "UPDATE " & REGSISTEMA.TABLAPLAN & " SET " & RSBOLS!Codigo & " =0 WHERE " & RSBOLS!Codigo & " IS NULL"
            RSBOLS.MoveNext
        Loop
        '---------------------------------------------------------------
        'ASIGNACIÓN DE LOS ADELANTOS DE PAGO
        RSBOLS.Close
        RSBOLS.Open "SELECT BOL.INUMBOL, MONTO FROM " & REGPLAN.TABLABOL & " BOL, " & REGSISTEMA.TABLAADEL & " ADEL WHERE BOL.CODTRAB=ADEL.CODTRAB AND ADEL.NOMBOL IN (SELECT CODIGO FROM NOMBOL WHERE MES=" & DateSQL(REGPLAN.MES) & ")", DBSYSTEM, adOpenStatic
        lProg.Caption = "2.- ASIGNANDO ADELANTOS DE PAGO"
        Prog.Max = RSBOLS.RecordCount + 1
        Prog.Value = 0
        Do While Not RSBOLS.EOF
            Prog.Value = Prog.Value + 1
            DBSYSTEM.Execute "UPDATE " & REGSISTEMA.TABLAPLAN & " SET " & REGSISTEMA.COLPLANADEL & "= " & REGSISTEMA.COLPLANADEL & "+" & IIf(IsNull(RSBOLS!MONTO), 0, RSBOLS!MONTO) & " WHERE INUMBOL=" & RSBOLS!INUMBOL & " AND MES=" & DateSQL(REGPLAN.MES)
            RSBOLS.MoveNext
        Loop
        '---------------------------------------------------------------
        'ASIGNANDO LOS VALORES DE CUENTAS CORRIENTES
        RSBOLS.Close
        RSBOLS.Open "SELECT PAGOSCTA.CODTRAB, TIPOPLAN, BOL.INUMBOL, PLANILLA, MONTO FROM " & REGPLAN.TABLABOL & " BOL, PAGOSCTA, MOVICTA, CTAGRUPO WHERE BOL.CODTRAB=PAGOSCTA.CODTRAB AND PAGOSCTA.CODMOV=MOVICTA.CODMOV AND MOVICTA.CODGRUPO=CTAGRUPO.CODGRUPO AND PAGOSCTA.CODNOMBOL IN (SELECT CODIGO FROM NOMBOL WHERE MES=" & DateSQL(REGPLAN.MES) & ")", DBSYSTEM, adOpenStatic
        lProg.Caption = "3.- ASIGNANDO CUENTAS CORRIENTES"
        Prog.Max = RSBOLS.RecordCount + 1 'SE AGREGA UNO PORQUE AVECES NO EXISTEN REGISTROS Y EL MAX NO PUEDE SER IGUAL A MIN
        Prog.Value = 0
        Do While Not RSBOLS.EOF
            Prog.Value = Prog.Value + 1
            If ExisteCampo(RSBOLS!PLANILLA, "PLAN2000", DBSYSTEM) Then
                DBSYSTEM.Execute "UPDATE " & REGSISTEMA.TABLAPLAN & " SET " & RSBOLS!PLANILLA & "=" & RSBOLS!PLANILLA & "+" & RSBOLS!MONTO & " WHERE INUMBOL=" & RSBOLS!INUMBOL & " AND MES=" & DateSQL(REGPLAN.MES)
            Else
                MsgBox "NO EXISTE EL CAMPO DE PLANILLA " & RSBOLS!PLANILLA & ". ERROR EN LA CONFIGURACIÓN DE CUENTAS CORRIENTES. NO SE HAN CARGADO LOS DATOS", vbInformation
            End If
            RSBOLS.MoveNext
        Loop
        '---------------------------------------------------------------
        'TAREA CUMPLIDA, CERRANDO TABLAS
        Set RSBOLS = Nothing
        Set RSPLAN2 = Nothing
        Set RSMOVS = Nothing
        CARGAPLANS
        Prog.Visible = False
        lProg.Visible = False
        MsgBox "SE HA PROCESADO SATISFACTORIAMENTE LA PLANILLA DE REMUNERACIONES DE: " & AMESES(Month(VGUTIL(1))) & " DE " & Year(VGUTIL(1)), vbInformation
    End If
End Sub

Private Sub XAREA_DblClick()
    Dim RSAUX As ADODB.Recordset
    Set RSAUX = New ADODB.Recordset
    RSAUX.Open "SELECT CODCCOSTO, NOMBRE FROM CCOSTOS ORDER BY CODCCOSTO", DBSYSTEM, adOpenStatic
    If RSAUX.RecordCount = 0 Then
        MsgBox "NO SE HAN ENCONTRADO REGISTROS DE CENTRO DE COSTO", vbCritical
        Set RSAUX = Nothing
        Exit Sub
    End If
    frmComun.CONECTAR RSAUX
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xArea.Text = RSAUX!CODCCOSTO & " - " & RSAUX!NOMBRE
        xArea.Tag = RSAUX!CODCCOSTO
    End If
    Set RSAUX = Nothing
End Sub

Private Sub XAREA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        xArea.Text = ""
        xArea.Tag = ""
    End If
End Sub

Public Sub PDTSUNAT()
    If DevuelveValor("SELECT PDTTRIBUTO FROM EMPRESA", DBSYSTEM) = "" Then
        MsgBox "Debe Colocar el tributo de 5ta Categoria en configuración del sistema", vbExclamation
        Exit Sub
    End If
    Screen.MousePointer = 11
    Dim xFile As String
    xFile = "BOL" & Mid(LPlans.SelectedItem.SubItems(1), 4, 2) & Right(LPlans.SelectedItem.SubItems(1), 4)
    Dim RSBOLS As ADODB.Recordset
    Set RSBOLS = New ADODB.Recordset
    RSBOLS.Open "SELECT APEPAT, APEMAT, NOMBRE, BOL.CODTRAB, CODAFP, TIPDOC, DOCIDEN, HORASTRAB, ESSALUDVIDA,SUMASALUD, SUMAIES, SUMARENTA, SUMAAFP, RENTA5TA,RUCEPS FROM " & xFile & " BOL, TRABAJADORES WHERE TRABAJADORES.CODTRAB=BOL.CODTRAB AND TRABAJADORES.NOPDT = 0", DBSYSTEM, adOpenStatic
    If ExisteTablaAux(" [##PDTSUNAT" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##PDTSUNAT" & VGL_COMPUTER & "] "
    'TEMPORAL PARA ALMACENAR LA EXPORTACIÓN DEL PDT SUNAT FORM. 0600
    DBSYSTEM.Execute "CREATE TABLE  [##PDTSUNAT" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8), NOMBRES VARCHAR(100),TIPDOC VARCHAR(2), DOCIDEN VARCHAR(15), DIASTRAB  Numeric(20,2) , REMUIES  Numeric(20,2) , REMUPENSION  Numeric(20,2) , REMUSALUD  Numeric(20,2) , REMUARTISTAS  Numeric(20,2) , REMU5TA  Numeric(20,2) , TRIBUTO5TA  Numeric(20,2) , ESVIDA  Numeric(20,2) )"
    Dim RSPDT As ADODB.Recordset
    Set RSPDT = New ADODB.Recordset
    RSPDT.Open " [##PDTSUNAT" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic
    Do While Not RSBOLS.EOF
        If RSPDT.RecordCount <> 0 Then
            RSPDT.MoveFirst
            RSPDT.FIND "CODTRAB='" & RSBOLS!CODTRAB & "'"
            If RSPDT.EOF Then 'SI EN CASO NO EXISTE, ENTONCES HAY QUE AGREGAR
                RSPDT.AddNew
                RSPDT!CODTRAB = RSBOLS!CODTRAB
                RSPDT!NOMBRES = Trim(RSBOLS!ApePat) & " " & Trim(RSBOLS!ApeMat) & " " & Trim(RSBOLS!NOMBRE)
                RSPDT!TIPDOC = "" & RSBOLS!TIPDOC
                RSPDT!DOCIDEN = "" & RSBOLS!DOCIDEN
            End If
        Else
            RSPDT.AddNew
            RSPDT!CODTRAB = RSBOLS!CODTRAB
            RSPDT!NOMBRES = Trim(RSBOLS!ApePat) & " " & Trim(RSBOLS!ApeMat) & " " & Trim(RSBOLS!NOMBRE)
            RSPDT!TIPDOC = "" & RSBOLS!TIPDOC
            RSPDT!DOCIDEN = "" & RSBOLS!DOCIDEN
        End If
        RSPDT!DIASTRAB = IIf(IsNull(RSPDT!DIASTRAB), 0, RSPDT!DIASTRAB) + Round(RSBOLS!HORASTRAB / 8, 0)
        RSPDT!REMUIES = IIf(IsNull(RSPDT!REMUIES), 0, RSPDT!REMUIES) + RSBOLS!SUMAIES
        RSPDT!REMUPENSION = IIf(IsNull(RSPDT!REMUPENSION), 0, RSPDT!REMUPENSION) + IIf(RSBOLS!CODAFP = "ON", RSBOLS!SUMAAFP, 0)
        If Trim(RSBOLS!RUCEPS) = "" Then RSPDT!REMUSALUD = IIf(IsNull(RSPDT!REMUSALUD), 0, RSPDT!REMUSALUD) + RSBOLS!SUMASALUD Else RSPDT!REMUSALUD = 0
        RSPDT!REMUARTISTAS = 0
        RSPDT!REMU5TA = IIf(IsNull(RSPDT!REMU5TA), 0, RSPDT!REMU5TA) + RSBOLS!SUMARENTA
        RSPDT!TRIBUTO5TA = IIf(IsNull(RSPDT!TRIBUTO5TA), 0, RSPDT!TRIBUTO5TA) + RSBOLS!RENTA5TA
        RSPDT!ESVIDA = IIf(RSBOLS!ESSALUDVIDA, 2, 0)
        RSPDT.Update
        RSBOLS.MoveNext
    Loop
    '/*ANTIGUO
'    DBSYSTEM.Execute "UPDATE  [##PDTSUNAT"  & VGL_COMPUTER & "]  SET REMU5TA=0 WHERE TRIBUTO5TA=0"
'    Screen.MousePointer = 1
'    If RSPDT.RecordCount = 0 Then
'        MsgBox "NO SE HAN AGREGADO REGISTROS PARA EL PDT SUNAT, POSIBLEMENTE LA PLANILLA SE ENCUENTRE VACIA O LA RELACIÓN CON LOS PERIODOS DE PAGOS DE REMUNERACIONES SEA DIFERENTE", vbCritical
'    Else
'        frPDTSunatSCTR.Show 1
'    End If
'    Set RSPDT = Nothing
'    Set RSBOLS = Nothing
    '/ANTIGUO*
    
    RSBOLS.Close
    RSBOLS.Open "SELECT CODTRAB," & DevuelveValor("SELECT PDTTRIBUTO FROM EMPRESA", DBSYSTEM) & " FROM PLAN2000 WHERE MES=" & DateSQL(CDate("01/" & Mid(LPlans.SelectedItem.SubItems(1), 4, 2) & "/" & Right(LPlans.SelectedItem.SubItems(1), 4))) & " AND " & DevuelveValor("SELECT PDTTRIBUTO FROM EMPRESA", DBSYSTEM) & "<>0", DBSYSTEM, adOpenStatic, adLockReadOnly
    Do While Not RSBOLS.EOF
        RSPDT.MoveFirst
        RSPDT.FIND "CODTRAB='" & RSBOLS!CODTRAB & "'"
        If Not RSPDT.EOF Then
            DBAUXCOM.Execute "UPDATE  [##PDTSUNAT" & VGL_COMPUTER & "]  SET TRIBUTO5TA=" & RSBOLS(1) & " WHERE CODTRAB='" & RSBOLS!CODTRAB & "'"
        End If
        RSBOLS.MoveNext
    Loop
    Set RSBOLS = Nothing
    
    DBAUXCOM.Execute "UPDATE  [##PDTSUNAT" & VGL_COMPUTER & "]  SET REMU5TA=0 WHERE TRIBUTO5TA=0"
    Screen.MousePointer = 1
    If RSPDT.RecordCount = 0 Then
        MsgBox "NO SE HAN AGREGADO REGISTROS PARA EL PDT SUNAT, POSIBLEMENTE LA PLANILLA SE ENCUENTRE VACIA O LA RELACIÓN CON LOS PERIODOS DE PAGOS DE REMUNERACIONES SEA DIFERENTE", vbCritical
    Else
        frPDTSunatSCTR.Show 1
    End If
    Set RSPDT = Nothing
End Sub

Public Sub PDTSUNATSCTR()
    Screen.MousePointer = 11
    Dim xFile As String
    xFile = "BOL" & Mid(LPlans.SelectedItem.SubItems(1), 4, 2) & Right(LPlans.SelectedItem.SubItems(1), 4)
    Dim RSBOLS As ADODB.Recordset
    Set RSBOLS = New ADODB.Recordset
    RSBOLS.Open "SELECT APEPAT, APEMAT, TRABAJADORES.NOMBRE, BOL.CODTRAB, TIPDOC, DOCIDEN, SUMASCTR, CENTROSAR.NOMBRE AS CARNOM, TASA,CORRELATIVO, RUC FROM " & xFile & " BOL, TRABAJADORES, CENTROSAR WHERE TRABAJADORES.CODTRAB=BOL.CODTRAB AND CODSCTR=CODCAR AND TRABAJADORES.NOPDT = 0 AND CENTROSAR.CODCAR<>'NONE'", DBSYSTEM, adOpenStatic
    If ExisteTablaAux("##PDTSUNATSCTR") Then DBSYSTEM.Execute "DROP TABLE ##PDTSUNATSCTR"
    'TEMPORAL PARA ALMACENAR LA EXPORTACIÓN DEL PDT SUNAT FORM. 0600
    DBSYSTEM.Execute "CREATE TABLE ##PDTSUNATSCTR(CODTRAB VARCHAR(8), NOMBRES VARCHAR(100),TIPDOC VARCHAR(2), DOCIDEN VARCHAR(15), RUC VARCHAR(11),CORRELATIVO INT,TASA  Numeric(20,2) ,REMUSCTR  Numeric(20,2) )"
    Dim RSPDT As ADODB.Recordset
    Set RSPDT = New ADODB.Recordset
    RSPDT.Open "##PDTSUNATSCTR", DBSYSTEM, adOpenKeyset, adLockOptimistic
    Do While Not RSBOLS.EOF
        If RSPDT.RecordCount <> 0 Then
            RSPDT.MoveFirst
            RSPDT.FIND "CODTRAB='" & RSBOLS!CODTRAB & "'"
            If RSPDT.EOF Then 'SI EN CASO NO EXISTE, ENTONCES HAY QUE AGREGAR
                RSPDT.AddNew
                RSPDT!CODTRAB = RSBOLS!CODTRAB
                RSPDT!NOMBRES = Trim(RSBOLS!ApePat) & " " & Trim(RSBOLS!ApeMat) & " " & Trim(RSBOLS!NOMBRE)
                RSPDT!TIPDOC = "" & RSBOLS!TIPDOC
                RSPDT!DOCIDEN = "" & RSBOLS!DOCIDEN
            End If
        Else
            RSPDT.AddNew
            RSPDT!CODTRAB = RSBOLS!CODTRAB
            RSPDT!NOMBRES = Trim(RSBOLS!ApePat) & " " & Trim(RSBOLS!ApeMat) & " " & Trim(RSBOLS!NOMBRE)
            RSPDT!TIPDOC = "" & RSBOLS!TIPDOC
            RSPDT!DOCIDEN = "" & RSBOLS!DOCIDEN
        End If
        RSPDT!RUC = "" & RSBOLS!RUC
        RSPDT!CORRELATIVO = RSBOLS!CORRELATIVO
        RSPDT!TASA = RSBOLS!TASA
        RSPDT!REMUSCTR = IIf(IsNull(RSPDT!REMUSCTR), 0, RSPDT!REMUSCTR) + RSBOLS!SUMASCTR
        RSPDT.Update
        RSBOLS.MoveNext
    Loop
    Screen.MousePointer = 1
    If RSPDT.RecordCount = 0 Then
        MsgBox "NO SE HAN AGREGADO REGISTROS PARA EL PDT SUNAT SCTR, POSIBLEMENTE LA PLANILLA SE ENCUENTRE VACIA O LA RELACIÓN CON LOS PERIODOS DE PAGOS DE REMUNERACIONES SEA DIFERENTE", vbCritical
    Else
        frPDTSunat.Show 1
    End If
    Set RSPDT = Nothing
    Set RSBOLS = Nothing
End Sub

Public Function COMPRUEBAPLAN() As Boolean
    Dim RSRUBROS As New ADODB.Recordset
    COMPRUEBAPLAN = False
    RSRUBROS.Open "CONCEPTOS", DBSYSTEM, adOpenKeyset, adLockOptimistic
    If RSRUBROS.EOF Then
        MsgBox "LOS CONCEPTOS DE REMUNERACIONES NO SE HAN DEFINIDO AÚN", vbInformation
        Set RSRUBROS = Nothing
        Exit Function
    End If
    Dim X As Long, Z As Byte
    X = 0
    Do While Not RSRUBROS.EOF
        If Trim(RSRUBROS!COLPLANILLA) <> "" Then
            DBSYSTEM.Execute "UPDATE COLUMPL SET TIPO=TIPO WHERE CODIGO='" & Trim(RSRUBROS!COLPLANILLA) & "'", X
            If X = 0 Then
                Z = MsgBox("EL CONCEPTO DE REMUNERACIÓN " & RSRUBROS!NOMBRE & " PRESENTA COMO COLUMNA DE PLANILLA EL CÓDIGO " & RSRUBROS!COLPLANILLA & " EL CUAL NO EXISTE DENTRO DE LA BASE DE DATOS. DESEA DEPURAR EL CONCEPTO DE REMUNERACIÓN", vbQuestion + vbYesNoCancel)
                If Z = vbCancel Or Z = vbNo Then Exit Function
                If Z = vbYes Then
                    VPTAREA = "EDITAR"
                    VPCODTMP = RSRUBROS!Codigo
                    Load frECnpt
                    frECnpt.cmCancela.Enabled = False
                    frECnpt.Show 1
                End If
            End If
        End If
        RSRUBROS.MoveNext
    Loop
    Set RSRUBROS = Nothing
    COMPRUEBAPLAN = True
End Function

Public Sub PDTTRABAJADORES()
    If LPlans.ListItems.Count = 0 Then
        MsgBox "NO SE HAN ENCONTRADO PLANILLAS", vbInformation
        Exit Sub
    End If
    On Error GoTo Err1
    Dim xFile As String, CADPDT As String
    frSelDir.Show 1
    If VPTAREA = "" Then Exit Sub
    If Right(VPTAREA, 1) <> "\" Then VPTAREA = VPTAREA & "\"
    xFile = VPTAREA & "\" & REGSISTEMA.RUC & ".ASE"
    If Dir$(xFile) <> "" Then
        If MsgBox("YA EXISTE EN ESTA RUTA UN ARCHIVO CORRESPONDIENTE AL PDT TRABAJADORES, DESEA UD. REEMPLAZAR EL ARCHIVO POR EL NUEVO QUE ESTÁ PROCESANDO", vbYesNo + vbQuestion) = vbNo Then Exit Sub
        Kill xFile
    End If
    
    Dim RSTRABS As New ADODB.Recordset
    If MsgBox("DESEA UD. LA IMPORTACIÓN DE DATOS DE LOS TRABAJADORES QUE SOLO HAYAN TENIDO PARTICIPACIÓN EN LA PLANILLA ACTUAL. SI ESCOJE NO ENTONCES SE IMPORTARÁ TODOS LOS TRABAJADORES ACTIVOS Y CESADOS DURANTE ESTE PERIODO SELECCIONADO", vbYesNo + vbQuestion) = vbYes Then
        RSTRABS.Open "SELECT * FROM TRABAJADORES WHERE SITUACIÓN<'2' AND CODTRAB IN (SELECT CODTRAB FROM PLAN2000 WHERE MES=" & DateSQL(LPlans.SelectedItem.SubItems(1)) & ")", DBSYSTEM, adOpenStatic, adLockReadOnly
    Else
        RSTRABS.Open "SELECT * FROM TRABAJADORES WHERE SITUACIÓN<'2'", DBSYSTEM, adOpenStatic, adLockReadOnly
    End If
    If RSTRABS.RecordCount = 0 Then
        MsgBox "NO SE HAN ENCONTRADO TRABAJADORES PARA SER EXPORTADOS AL PDT SUNAT", vbInformation
        Exit Sub
    End If
    Open xFile For Append As #1
    Dim ASit
    ASit = Array("10", "11", "12", "13", "14", "15", "18", "19")
    With RSTRABS
        Do While Not RSTRABS.EOF
            CADPDT = ""
            CADPDT = !TIPDOC & "|" & IIf(IsNull(!DOCIDEN), "", !DOCIDEN) & "|" & IIf(IsNull(!ApePat), "", !ApePat) & "|" & IIf(IsNull(!ApeMat), "", !ApeMat) & "|" & IIf(IsNull(!NOMBRE), "", !NOMBRE) & "|" & IIf(IsNull(!FechaNac), "", !FechaNac) & "|" & IIf(!Sexo = 1, "1", "2") & "|" & SoloNumeros(IIf(IsNull(!TELEFONO), "", !TELEFONO)) & "|" & IIf(IsNull(!FECHAING), "", !FECHAING) & "|" & ASit(!SITUACIÓN) & "|" & IIf(IsNull(!TIPOTRAB), "", !TIPOTRAB) & "|" & IIf(!SITUACIÓN < 2, "", !FECHACESE) & "|" & Trim(IIf(IsNull(!RUCEPS), "", !RUCEPS)) & "|" & IIf(!ESSALUDVIDA, "1", "0") & "|" & IIf(!FONDOPENS = "ON", "2", "1") & "|" & IIf(!CODSCTR = "NONE", "0", "1") & "|" & IIf(IsNull(!FECHAIAFP), "", !FECHAIAFP) & "|||||||||"
            Print #1, CADPDT
            RSTRABS.MoveNext
        Loop
    End With
    Close #1
    Set RSTRABS = Nothing
    MsgBox "PROCESO COMPLETADO. INGRESE AL PDT SUNAT Y ESCOJA LA OPCIÓN IMPORTAR DEL MENÚ DECLARACIONES, DENTRO DEL MÓDULO 0600 DDJJ RETENCIONES Y CONTRIBUCIONES - REMUNERACIONES", vbInformation
    Exit Sub
Err1:
    MsgBox ERR.Description
    Exit Sub
End Sub

Public Sub PDTDerecho()
    If LPlans.ListItems.Count = 0 Then
        MsgBox "No se han encontrado planillas", vbInformation
        Exit Sub
    End If
    On Error GoTo Err1
    Dim xFile As String, CADPDT As String
    frSelDir.Show 1
    If VPTAREA = "" Then Exit Sub
    If Right(VPTAREA, 1) <> "\" Then VPTAREA = VPTAREA & "\"
    xFile = VPTAREA & "\" & REGSISTEMA.RUC & ".der"
    If Dir$(xFile) <> "" Then
        If MsgBox("Ya existe en esta ruta un archivo correspondiente al PDT DerechoHabientes, Desea Ud. reemplazar el archivo por el nuevo que está procesando", vbYesNo + vbQuestion) = vbNo Then Exit Sub
        Kill xFile
    End If
    
    Dim RSTRABS As New ADODB.Recordset
    If MsgBox("Desea Ud. la importación de datos de los DerechHabientes que solo hayan tenido participación en la planilla actual. Si escoje NO entonces se importará todos los trabajadores Activos y Cesados durante este periodo seleccionado", vbYesNo + vbQuestion) = vbYes Then
'        RsTrabs.Open "SELECT Trabajadores.TipDoc, Trabajadores.DocIden, Familiar.TipoDoc, Familiar.NumDoc, Familiar.ApePat, Familiar.ApeMat, Familiar.Nombre, Familiar.FechaNac, Familiar.Sexo, Familiar.Carta, Familiar.Situacion, Familiar.MotivoBaja, Familiar.DocIncap, Familiar.IDP, Familiar.NombreVia, Familiar.Numero, Familiar.Interior, Familiar.Zona, Familiar.Referencia, Familiar.TipoVia, Familiar.TipoZona, Familiar.Ubigeo " & _
'                    " FROM Familiar INNER JOIN Trabajadores ON Familiar.CodTrab = Trabajadores.Codtrab " & _
'                    " WHERE Trabajadores.Situación<'2' AND Trabajadores.CodTrab IN (SELECT CodTrab FROM Plan2000 WHERE Mes=" & DateSQL(LPlans.SelectedItem.SubItems(1)) & ")", DbSystem, adOpenStatic, adLockReadOnly
        RSTRABS.Open "SELECT TRABAJADORES.TIPDOC, TRABAJADORES.DOCIDEN, FAMILIAR.TIPODOC, FAMILIAR.NUMDOC, FAMILIAR.APEPAT, FAMILIAR.APEMAT, FAMILIAR.NOMBRE, FAMILIAR.FECHANAC, FAMILIAR.SEXO, FAMILIAR.VINCULO, FAMILIAR.CARTA, FAMILIAR.SITUACION, FAMILIAR.MOTIVOBAJA, FAMILIAR.DOCINCAP, FAMILIAR.IDP, FAMILIAR.NOMBREVIA, FAMILIAR.NUMERO, FAMILIAR.INTERIOR, FAMILIAR.ZONA, FAMILIAR.REFERENCIA, FAMILIAR.TIPOVIA, FAMILIAR.TIPOZONA, FAMILIAR.UBIGEO FROM FAMILIAR INNER JOIN TRABAJADORES ON FAMILIAR.CODTRAB = TRABAJADORES.CODTRAB WHERE TRABAJADORES.SITUACIÓN<'2' AND TRABAJADORES.CODTRAB IN (SELECT CODTRAB FROM PLAN2000 WHERE MES=" & DateSQL(LPlans.SelectedItem.SubItems(1)) & ")", DBSYSTEM, adOpenStatic, adLockReadOnly
    Else
        RSTRABS.Open "SELECT TRABAJADORES.TIPDOC, TRABAJADORES.DOCIDEN, FAMILIAR.TIPODOC, FAMILIAR.NUMDOC, FAMILIAR.APEPAT, FAMILIAR.APEMAT, FAMILIAR.NOMBRE, FAMILIAR.FECHANAC, FAMILIAR.SEXO, FAMILIAR.CARTA, FAMILIAR.SITUACION, FAMILIAR.MOTIVOBAJA, FAMILIAR.DOCINCAP, FAMILIAR.IDP, FAMILIAR.NOMBREVIA, FAMILIAR.NUMERO, FAMILIAR.INTERIOR, FAMILIAR.ZONA, FAMILIAR.REFERENCIA, FAMILIAR.TIPOVIA, FAMILIAR.TIPOZONA, FAMILIAR.UBIGEO " & _
                    " FROM FAMILIAR INNER JOIN TRABAJADORES ON FAMILIAR.CODTRAB = TRABAJADORES.CODTRAB " & _
                    " WHERE TRABAJADORES.SITUACIÓN<'2'", DBSYSTEM, adOpenStatic, adLockReadOnly
    End If
    If RSTRABS.RecordCount = 0 Then
        MsgBox "No se han encontrado DerechoHabientes para ser exportados al PDT Sunat", vbInformation
        Exit Sub
    End If
    Open xFile For Append As #1
    Dim ARRTIPDOC
    ARRTIPDOC = Array("01", "02", "03", "04", "07", "08", "10", "11")
    
    
    With RSTRABS
        Do While Not RSTRABS.EOF
            CADPDT = ""
            CADPDT = !TIPDOC & "|" & IIf(IsNull(!DOCIDEN), "", !DOCIDEN) & "|" & ARRTIPDOC(!TIPODOC) & "|" & IIf(IsNull(!NUMDOC), "", !NUMDOC) & "|" & IIf(IsNull(!ApePat), "", !ApePat) & "|" & IIf(IsNull(!ApeMat), "", !ApeMat) & "|" & IIf(IsNull(!NOMBRE), "", !NOMBRE) & "|" & IIf(IsNull(!FechaNac), "", !FechaNac) & "|" & IIf(!Sexo = 1, "2", "1") & "|" & IIf(IsNull(!VINCULO), "", !VINCULO + 1) & "|" & "" & "|" & Trim(!CARTA) & "|" & (!SITUACION) + 10 & "|" & IIf(!MOTIVOBAJA = 0, "", !MOTIVOBAJA + 1) & "|" & Trim(IIf(IsNull(!DOCINCAP), "", !DOCINCAP)) & "|" & IIf(!IDP, "1", "0") & "|" & IIf(IsNull(!NOMBREVIA), "", !NOMBREVIA) & "|" & IIf(IsNull(!Numero), "", !Numero) & "|" & IIf(IsNull(!INTERIOR), "", !INTERIOR) & "|" & IIf(IsNull(!ZONA), "", !ZONA) & "|" & IIf(IsNull(!REFERENCIA), "", !REFERENCIA) & "|" & IIf(!TIPOVIA = 0, "", Format(!TIPOVIA + 1, "00")) & "|" & IIf(!TIPOZONA = 0, "", Format(!TIPOZONA + 1, "00")) & "|" & IIf(IsNull(!UBIGEO), "", !UBIGEO) & "|"
            Print #1, CADPDT
            RSTRABS.MoveNext
        Loop
    End With
    Close #1
    Set RSTRABS = Nothing
    MsgBox "Proceso completado. Ingrese al PDT Sunat y escoja la Opción Importar del Menú Declaraciones, dentro del Módulo 0600 DDJJ Retenciones y Contribuciones - Remuneraciones", vbInformation
    Exit Sub
Err1:
    MsgBox ERR.Description
    Resume Next
    Resume
    Exit Sub
End Sub
