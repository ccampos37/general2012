VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frAreas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Areas de Trabajo"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   Icon            =   "frAreas.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   7365
   Tag             =   "Centros de Costos para Planillas"
   Begin VB.TextBox SqlCad 
      Height          =   330
      Left            =   6570
      TabIndex        =   2
      Top             =   2880
      Visible         =   0   'False
      Width           =   750
   End
   Begin MSComctlLib.ListView lvCCostos 
      Height          =   3765
      Left            =   30
      TabIndex        =   0
      Top             =   150
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   6641
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "RUC"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "FechaIng"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6570
      Top             =   3300
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
            Picture         =   "frAreas.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frAreas.frx":0D1E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6510
      Picture         =   "frAreas.frx":1172
      Top             =   210
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Areas de Trabajo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      Left            =   6255
      TabIndex        =   1
      Top             =   915
      Width           =   1005
   End
End
Attribute VB_Name = "frAreas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim REGACT As REGWIN

Private Sub FORM_ACTIVATE()
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
    RS.Open "SELECT * FROM AREASTRAB ORDER BY CODCCOSTO", DBSYSTEM, adOpenDynamic, adLockOptimistic
    CARGAREGS
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RS = Nothing
End Sub

Public Sub CARGAREGS()
    'CARGA LOS REGISTROS Y LOS COLOCA AL TREEVIEW COMO AL LISVIEW
    Dim VRELAT As String
    Dim XLIST As ListItem
    lvCCostos.ListItems.Clear
    If RS.RecordCount = 0 Or RS.EOF Then Exit Sub
    RS.MoveFirst
    Do While Not RS.EOF
        If Len(RS!CODCCOSTO) = 2 Then VRELAT = "MAIN" Else VRELAT = "C" & Left(RS!CODCCOSTO, Len(RS!CODCCOSTO) - 3)
        Set XLIST = lvCCostos.ListItems.Add(, "C" & RS!CODCCOSTO, RS!CODCCOSTO, 1, IIf(InStr(RS!CODCCOSTO, ".") > 0, 1, 2))
        XLIST.SubItems(1) = RS!NOMBRE
        XLIST.SubItems(2) = "" & RS!RUC
        XLIST.SubItems(3) = "" & RS!FECHAING
        RS.MoveNext
    Loop
End Sub

Public Function EXISTECODIGO(ByVal xCod As String) As Boolean
    Dim RSAUX As New ADODB.Recordset
    RSAUX.Open "SELECT CODCCOSTO FROM AREASTRAB WHERE CODCCOSTO='" & xCod & "'", DBSYSTEM, adOpenStatic
    If RSAUX.EOF Or RSAUX.RecordCount = 0 Then
        EXISTECODIGO = False
    Else
        EXISTECODIGO = True
    End If
End Function

Public Sub COMANDOTOOLBAR(COMANDO As String)
    Select Case UCase(COMANDO)
        Case "NUEVO"
            VPTAREA = "NUEVO"
            frEArea.Show 1
            RS.Requery
            CARGAREGS
        Case "EDITAR"
            If lvCCostos.ListItems.Count = 0 Then Exit Sub
            VPTAREA = "EDITAR"
            frEArea.Show 1
            RS.Requery
            CARGAREGS
        Case "ELIMINAR"
            If lvCCostos.ListItems.Count = 0 Then
                MsgBox "Debe seleccionar o existir un elemento para poder eliminarlo", vbCritical
                Exit Sub
            End If
            Dim RsTmp As New ADODB.Recordset
            RsTmp.Open "SELECT * FROM AREASTRAB WHERE CODCCOSTO LIKE '" & lvCCostos.SelectedItem.Text & "%'", DBSYSTEM, adOpenKeyset, adLockReadOnly
            If RsTmp.RecordCount >= 2 Then
                MsgBox "No se puede eliminar porque existen elementos enlazados al codigo actual", vbCritical
                Exit Sub
            End If
            RsTmp.Close
            RsTmp.Open "SELECT CODTRAB FROM TRABAJADORES WHERE CCOSTO='" & lvCCostos.SelectedItem.Text & "'", DBSYSTEM, adOpenStatic
            If RsTmp.RecordCount > 0 Then
                MsgBox "No se puede eliminar el elemento seleccionado, pues contiene trabajadores ligados a esta", vbCritical
                Set RsTmp = Nothing
                Exit Sub
            End If
            Set RsTmp = Nothing
            If MsgBox("Realmente desea eliminar el elemento seleccionado", vbInformation + vbYesNo) = vbYes Then
                DBSYSTEM.Execute ("DELETE FROM AREASTRAB WHERE CODCCOSTO='" & lvCCostos.SelectedItem.Text & "'")
                RS.Requery
                CARGAREGS
            End If
        Case "IMPRIMIR": IMPRIMIR
        Case Else
            MsgBox "Funcion no permitida", vbCritical
    End Select
End Sub
Private Sub IMPRIMIR()
    Dim X As Long
    Screen.MousePointer = 11
    On Error GoTo ERRADO
    'SE CREA UNA TABLA AUXILIAR AREASTRAB CON UN CAMPO NIVEL
    If ExisteTablaAux(" [##TMPAREASTRAB" & VGL_COMPUTER & "] ") Then DBAUXCOM.Execute "DROP TABLE  [##TMPAREASTRAB" & VGL_COMPUTER & "] "
    Dim RUTA As String
    SqlCad.Text = "" & _
        "SELECT *, 0 AS NIVEL " & _
        "INTO  [##TMPAREASTRAB" & VGL_COMPUTER & "]  " & _
        " FROM AREASTRAB "
    Screen.MousePointer = 11
    DBSYSTEM.Execute SqlCad.Text, X
    Screen.MousePointer = 1
    If X = 0 Then
        MsgBox "Mensaje del sistema: " & _
        " NO SE ENCONTRARÓN REGISTROS ", vbInformation
        Exit Sub
    End If
   
    
    'SE CUENTA LA CANTIDAD DE PUNTOS EN EL CODIGO DE AREA Y DEPENDIENDO DE ESO
    'SE ACTUALIZA EL CAMPO NIVEL RECORRIENDO TODOS LOS REGISTROS DEL AREA
    Dim RSTEMP As New ADODB.Recordset
    RSTEMP.Open " [##TMPAREASTRAB" & VGL_COMPUTER & "] ", DBAUXCOM, adOpenDynamic, adLockOptimistic
    RSTEMP.MoveFirst
    Do While Not RSTEMP.EOF
        RSTEMP!NIVEL = BusCad(".", RSTEMP!CODCCOSTO)
        RSTEMP.Update
        RSTEMP.MoveNext
    Loop
    Set RSTEMP = Nothing
    
    'SE SELECCIONAN LOS NIVELES EXISTENTES
    RSTEMP.Open "SELECT DISTINCT NIVEL FROM  [##TMPAREASTRAB" & VGL_COMPUTER & "]  ORDER BY NIVEL ", _
    DBAUXCOM, adOpenDynamic, adLockOptimistic
        
    Dim I As Integer
    ''LLENAR EL COMBO NIVEL DE
    RSTEMP.MoveFirst
    For I = 1 To RSTEMP.RecordCount
        FrRepNivAr.CmbNivel.AddItem (Format(RSTEMP!NIVEL + 1, 0) & " -> NIVEL")
        RSTEMP.MoveNext
    Next
    FrRepNivAr.CmbNivel.ListIndex = 0
    Set RSTEMP = Nothing
    
    FrRepNivAr.Show 1
    
    Exit Sub
ERRADO: MsgBox "Por favor espere un momento"
        Screen.MousePointer = 1
End Sub



