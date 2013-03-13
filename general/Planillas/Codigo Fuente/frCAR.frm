VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frCAR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Centros de Alto Riesgo"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   Icon            =   "frCAR.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   6150
   Tag             =   "Panel de Centros de Alto Riesgo"
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   90
      Top             =   2310
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
            Picture         =   "frCAR.frx":08CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvCAR 
      Height          =   3285
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   5794
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "codigo"
         Text            =   "Codigo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "nombre"
         Text            =   "Nombre"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   "tasa"
         Text            =   "Tasa"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   "correlativo"
         Text            =   "Correlativo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "ruc"
         Text            =   "RUC"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frCAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As ADODB.Recordset
Dim REGACT As REGWIN

Private Sub CMBORRAR_CLICK()
    If lvCAR.SelectedItem.INDEX = -1 Then
        MsgBox "DEBE SELECCIONAR UN ELEMENTO", vbCritical
        Exit Sub
    End If
    If lvCAR.SelectedItem.Text = "NONE" Then
        MsgBox "NO SE PUEDE ELIMINAR EL ELEMENTO NONE POR SER PARTE DEL SISTEMA", vbCritical
        Exit Sub
    End If
    If MsgBox("REALMENTE DESEA ELEMINAR EL REGISTRO SELECCIONADO. AL ELIMINARLO TAMBIÉN SE ELIMINARÁN AUTOMÁTICAMENTE TODAS REFERENCIAS DEL MISMO EN OTROS ARCHIVOS DEL SISTEMA", vbInformation + vbYesNo) = vbNo Then Exit Sub
    DBSYSTEM.Execute ("DELETE FROM CENTROSAR WHERE CODCAR='" & lvCAR.SelectedItem.Text & "'")
    CARGADATOS
End Sub

Private Sub CMEDITAR_CLICK()
    If lvCAR.SelectedItem.Text = "NONE" Then
        MsgBox "NO SE PUEDE EDITAR EL ELEMENTO NONE POR SER UN REGISTRO DE SOLO LECTURA (FORMA PARTE DEL PROGRAMA DE CONTROL DEL SISTEMA", vbCritical
        Exit Sub
    End If
    VPTAREA = "EDITAR"
    frEdCAR.Show 1
    CARGADATOS
End Sub

Private Sub CMNUEVO_CLICK()
    VPTAREA = "NUEVO"
    frEdCAR.Show 1
    CARGADATOS
End Sub

Private Sub CMSALIR_CLICK()
    Unload Me
End Sub

Private Sub FORM_ACTIVATE()
    lvCAR.SetFocus
    ActivarTools REGACT
End Sub

Private Sub FORM_LOAD()
    Call ALTERRUC
    Set RS = DBSYSTEM.Execute("CENTROSAR")
    CARGADATOS
    With REGACT
        .BUSCAR = False
        .EDITAR = True
        .ELIMINAR = True
        .FILTRAR = False
        .IMPRIMIR = True
        .NUEVO = True
        .PRELIMINAR = True
    End With
End Sub

Public Sub CARGADATOS()
    Dim XLIST As ListItem
    lvCAR.ListItems.Clear
    RS.Requery
    If RS.RecordCount = 0 Then Exit Sub
    RS.MoveFirst
    Do While Not RS.EOF
        Set XLIST = lvCAR.ListItems.Add(, RS!CODCAR, RS!CODCAR, 1, 1)
        XLIST.SubItems(1) = "" & RS!NOMBRE
        XLIST.SubItems(2) = Format(RS!TASA, "0.00")
        XLIST.SubItems(3) = "" & RS!CORRELATIVO
        XLIST.SubItems(4) = "" & RS!RUC
        RS.MoveNext
    Loop
    lvCAR.ColumnHeaders(3).Width = 854.9292
    lvCAR.ColumnHeaders(4).Width = 1049.953
    lvCAR.ColumnHeaders(5).Width = 1019.906
    lvCAR.Refresh
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    RS.Close
    Set RS = Nothing
End Sub

Public Sub COMANDOTOOLBAR(ByVal COMANDO As String)
    Select Case UCase(COMANDO)
        Case "NUEVO": CMNUEVO_CLICK
        Case "ELIMINAR": CMBORRAR_CLICK
        Case "EDITAR": CMEDITAR_CLICK
    End Select
End Sub
Private Sub ALTERRUC()
Dim RS As New ADODB.Recordset
    RS.Open "SELECT TOP 0 RUC FROM CENTROSAR", DBSYSTEM
    If RS.Fields(0).DefinedSize = 8 Then
        DBSYSTEM.Execute "ALTER TABLE CENTROSAR ALTER COLUMN RUC VARCHAR(11)"
    End If
End Sub
