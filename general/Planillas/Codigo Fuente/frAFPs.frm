VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frAFPs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fondos de Pensiones"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   Icon            =   "frAFPs.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   7080
   Tag             =   "Panel de Fondo de Pensiones"
   Begin Crystal.CrystalReport RptAFP 
      Left            =   2790
      Top             =   1230
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5280
      Top             =   2610
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
            Picture         =   "frAFPs.frx":030A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvAFPs 
      Height          =   2985
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5745
      _ExtentX        =   10134
      _ExtentY        =   5265
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Aport.Obli."
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Seguro"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Tope Seguro"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Rem. Aseg."
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "AFP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6000
      TabIndex        =   1
      Top             =   765
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6255
      Picture         =   "frAFPs.frx":0626
      Top             =   150
      Width           =   480
   End
End
Attribute VB_Name = "frAFPs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim REGACT As REGWIN

Private Sub Form_Activate()
    ActivarTools REGACT
End Sub

Private Sub Form_Load()
    With REGACT
        .BUSCAR = False
        .EDITAR = True
        .ELIMINAR = True
        .FILTRAR = False
        .IMPRIMIR = True
        .NUEVO = True
        .PRELIMINAR = True
    End With
    RS.Open "AFPS", DBSYSTEM, adOpenKeyset, adLockOptimistic
    CARGADATOS
End Sub

Public Sub CARGADATOS()
    Dim XITEM As ListItem
    lvAFPs.ListItems.Clear
    If RS.RecordCount = 0 Then Exit Sub
    RS.MoveFirst
    Do While Not RS.EOF
        Set XITEM = lvAFPs.ListItems.Add(, "C" & RS!CODAFP, RS!CODAFP, 1, 1)
        XITEM.SubItems(1) = "" & RS!NOMBRE
        XITEM.SubItems(2) = Format(RS!APOROBLI, "0.00")
        XITEM.SubItems(3) = Format(RS!SEGURO, "0.00")
        XITEM.SubItems(4) = Format(RS!TOPESEGURO, "0.00")
        XITEM.SubItems(5) = Format(RS!COMISIONRA, "0.00")
        RS.MoveNext
    Loop
End Sub

Public Function EXISTE(ByVal xCod As String) As Boolean
    If xCod = "" Then
        EXISTE = True
        Exit Function
    End If
    RS.MoveFirst
    RS.FIND "CODAFP='" & xCod & "'"
    If RS.EOF Then
        EXISTE = False
    Else
        EXISTE = True
    End If
End Function

Private Sub FORM_UNLOAD(CANCEL As Integer)
    RS.Close
    Set RS = Nothing
End Sub

Public Sub COMANDOTOOLBAR(COMANDO As String)
    Select Case UCase(COMANDO)
        Case "NUEVO"
            VPTAREA = "NUEVO"
            frEAFP.Show 1
            RS.Requery
            CARGADATOS
        Case "EDITAR"
            VPTAREA = "EDITAR"
            frEAFP.Show 1
            RS.Requery
            CARGADATOS
        Case "ELIMINAR"
            If lvAFPs.ListItems.Count = 0 Then Exit Sub
            X = 0
            DBSYSTEM.Execute "UPDATE TRABAJADORES SET CUSPP=CUSPP WHERE FONDOPENS='" & lvAFPs.SelectedItem.Text & "'", X
            If X > 0 Then
                MsgBox "No se puede eliminar la afp seleccionada pues contiene trabajadores afiliados", vbCritical
                Exit Sub
            End If
            If MsgBox("Seguro de eliminar el registro de afp seleccionado", vbQuestion + vbYesNo) = vbYes Then
                DBSYSTEM.Execute "DELETE FROM AFPS WHERE CODAFP='" & lvAFPs.SelectedItem.Text & "'"
            End If
            RS.Requery
            CARGADATOS
        Case "IMPRIMIR", "PRELIMINAR"
            Screen.MousePointer = 11
            With RptAFP
                .Reset
                .ReportFileName = REGSISTEMA.REPORTES & "PLAN0001.RPT"
                .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=" & VGL_BASE & ""
                .StoredProcParam(0) = "TRABAJADORES"
                .StoredProcParam(1) = "AFPS"
                .StoredProcParam(2) = "FONDOPENS"
                .StoredProcParam(3) = "CODAFP"
                .StoredProcParam(4) = REGSISTEMA.BASESQL
                .Destination = IIf(UCase(COMANDO) = "PRELIMINAR", crptToWindow, crptToWindow)
                .WindowTitle = "PLAN0001 - Listado de Afiliados por Fondos de Pensiones"
                .WindowState = crptMaximized
                .WindowShowPrintBtn = True
                .WindowShowRefreshBtn = True
                .WindowShowSearchBtn = True
                .WindowShowPrintSetupBtn = True
                .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
                .Formulas(1) = "XRUC='" & REGSISTEMA.RUC & "'"
                .Formulas(2) = "XHORA='" & Format(Time, "HH:MM") & "'"
                If RptAFP.Status <> 2 Then .Action = 1
            End With
        Case Else
            MsgBox "No disponible para este modulo"
    End Select
        Screen.MousePointer = 1
End Sub

Private Sub LVAFPS_DBLCLICK()
    VPTAREA = "EDITAR"
    frEAFP.Show 1
    RS.Requery
    CARGADATOS
End Sub
