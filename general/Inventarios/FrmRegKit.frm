VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmRegKit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualiza Registro de Kits"
   ClientHeight    =   5370
   ClientLeft      =   1080
   ClientTop       =   2250
   ClientWidth     =   7440
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   7440
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   6555
      Top             =   4650
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton command5 
      Caption         =   "&Reporte"
      Height          =   675
      Left            =   4440
      Picture         =   "FrmRegKit.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4485
      Width           =   775
   End
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   7185
      Begin VB.ComboBox CmbOrden 
         Height          =   315
         ItemData        =   "FrmRegKit.frx":0442
         Left            =   5445
         List            =   "FrmRegKit.frx":044C
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   225
         Width           =   1575
      End
      Begin VB.TextBox TxFiltro 
         Height          =   300
         Left            =   1080
         TabIndex        =   6
         Text            =   "TxFiltro"
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label33 
         Caption         =   "Orden :"
         Height          =   255
         Left            =   4485
         TabIndex        =   9
         Top             =   285
         Width           =   855
      End
      Begin VB.Label Label32 
         Caption         =   "Buscar  :"
         Height          =   285
         Left            =   150
         TabIndex        =   8
         Top             =   255
         Width           =   930
      End
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Height          =   675
      Left            =   2385
      Picture         =   "FrmRegKit.frx":0469
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4500
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.CommandButton CmdModi 
      Caption         =   "&Modificar"
      Height          =   675
      Left            =   2355
      Picture         =   "FrmRegKit.frx":08AB
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4500
      Width           =   775
   End
   Begin VB.CommandButton CmdEli 
      Caption         =   "&Eliminar"
      Height          =   675
      Left            =   3420
      Picture         =   "FrmRegKit.frx":0CED
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4500
      Width           =   775
   End
   Begin VB.CommandButton CmdIng 
      Caption         =   "&Crear"
      Height          =   675
      Left            =   1320
      Picture         =   "FrmRegKit.frx":112F
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4515
      Width           =   775
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   675
      Left            =   5430
      Picture         =   "FrmRegKit.frx":1571
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4485
      Width           =   775
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3375
      Left            =   120
      TabIndex        =   10
      Top             =   870
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   5953
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
            LCID            =   3082
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
            LCID            =   3082
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
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   7095
      Begin VB.TextBox TxCantidad 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   360
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   2760
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   360
         Width           =   1695
      End
      Begin MSFlexGridLib.MSFlexGrid Salida 
         Height          =   2880
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   5080
         _Version        =   393216
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label3"
         Height          =   300
         Left            =   1800
         TabIndex        =   17
         Top             =   720
         Width           =   5055
      End
      Begin VB.Label Label2 
         Caption         =   "Descripción     :"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Código             :"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmRegKit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adodc1 As ADODB.Recordset
Dim cSel1 As ADODB.Recordset
Dim cSql1 As String, CSQL2 As String, cCod As String
Dim nT As Integer       'Ingreso,Modificación
Dim nCom As Integer, nTra As Integer, nCursor As Integer

Private Sub OculObj03(nTipo As Boolean) ' Todos los datos
Frame1.Visible = nTipo
End Sub
Private Sub OculObj04(nTipo As Boolean) ' Botones principales
CmdIng.Visible = nTipo
CmdModi.Visible = nTipo
'CmdEli.Visible = nTipo
CmdSalir.Visible = nTipo
End Sub
Private Sub OculObj05(nTipo As Boolean)  'Orden y Filtro
Frame5.Visible = nTipo
Label32.Visible = nTipo
TxFiltro.Visible = nTipo
Label33.Visible = nTipo
CmbOrden.Visible = nTipo
End Sub
Private Sub OculObj06(nTipo As Boolean)  'Datagrid
DataGrid1.Visible = nTipo
End Sub
Private Sub CmbOrden_Click()             ' Ordenar por
Dim cD As String
nCom = CmbOrden.ListIndex
Set adodc1 = New ADODB.Recordset
cD = "SELECT DISTINCT(ACODIGO),ADESCRI,AUNIDAD FROM KITS,MAEART WHERE ACODIGO = CODKIT"

Select Case nCom
Case 0
            cD = cD & " ORDER BY ACODIGO"
Case 1
            cD = cD & " ORDER BY ADESCRI"
End Select
adodc1.Open cD, VGCNx, adOpenStatic
TxFiltro = ""
Set DataGrid1.DataSource = adodc1
Set_Data
If DataGrid1.Visible Then DataGrid1.SetFocus
End Sub

Private Sub CmdEli_Click()              ' Elimina
Dim CSQL2 As String, nN As Integer
Dim I As Integer
On Error GoTo EliErr
If Salida.Visible Then
  If Salida.Rows = 1 Then
    MsgBox "No hay registros para Eliminar", vbInformation, "Información"
    Exit Sub
  End If
  If MsgBox("Desea Eliminar el Registro?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
    I = Salida.RowSel
      cSql1 = "Delete from KITS where CODkit = '" & Text1 & "' and codart = '" & Salida.TextMatrix(I, 1) & "' "
      VGCNx.Execute cSql1
    If Salida.Rows > 2 Then
        Salida.RemoveItem I
    Else
        Salida.Clear
        Salida.Rows = 1
        Salida.Row = 0
        Set_Flex
        CmdSalir.SetFocus
    End If
  
  End If
    
Else
If adodc1.RecordCount > 0 Then
    cCod = adodc1("ACODIGO")
    cSql1 = "Select STCODIGO from STKART where STCODIGO = '" & cCod & "' and STSKDIS > 0"
    Set cSel1 = New ADODB.Recordset
    cSel1.Open cSql1, VGCNx, adOpenStatic
    If cSel1.RecordCount > 0 Then          ' vGAlmacen
        MsgBox "El artículo tiene Movimientos de Almacén con Cantidad Disponible mayor a Cero, no se puede Eliminar", vbInformation, "Mensaje"
        cSel1.Close
        Exit Sub
    End If
    cSel1.Close
        
    If MsgBox("   Desea Eliminar " & Chr(10) & "" & Mid(adodc1("ADESCRI"), 1, 25) & "", vbQuestion + vbOKCancel, "Mensaje") = vbOK Then
        nN = Pos_Dato(adodc1)
        
        cSql1 = "Select CodArt,CodKit,CANART from Kits where CODKIT = '" & Text1 & "'"
        Set cSel1 = New ADODB.Recordset
        cSel1.Open cSql1, VGCNx, adOpenStatic
        Do While Not cSel1.EOF
           CSQL2 = "Update STKART Set STSKDIS=STSKDIS+" & cSel1("CanArt") & " where StAlma='" & VGAlma & "' and StCodigo='" & cSel1("CODART") & "'"
           VGCNx.Execute CSQL2
           cSel1.MoveNext
        Loop
        cSel1.Close
                
        cSql1 = "Delete from KITS where CODkit = '" & cCod & "'"
        CSQL2 = "Delete from STKART where STCODIGO = '" & cCod & "'"
      
        nTra = 1
        VGCNx.BeginTrans
        VGCNx.Execute cSql1
        VGCNx.Execute CSQL2
        VGCNx.CommitTrans
        
        nTra = 0
        
        adodc1.Requery
        If nN <> 0 Then adodc1.AbsolutePosition = nN
    End If
Else
    MsgBox "No existen registros", vbInformation, "Mensaje"
End If
Set_Data
DataGrid1.SetFocus
End If
Exit Sub
EliErr:
    MsgBox Err.Description
    If nTra = 1 Then VGCNx.RollbackTrans
End Sub

Private Sub CmdGrabar_Click()           ' Grabar
Dim I As Integer
On Error GoTo GrabErr

If Trim(Text1) = "" Then
    MsgBox "Ingrese Código", vbInformation, "Mensaje"
    Text1.SetFocus: Exit Sub
End If

If MsgBox("Es correcta la Información", vbQuestion + vbOKCancel, "Mensaje") = vbOK Then
    If nT = 1 Then      'Ingreso
        If codigo(Text1) = False Then
            If Existe(1, Text1, "kits", "codkit", False) Then
                MsgBox "Código de Artículo ya existe", vbInformation, "Mensaje"
                Text1.SetFocus: Exit Sub
            End If
        End If
        For I = 1 To Salida.Rows - 1
            'If Salida.TextMatrix(I, 0) = ">>" Then
                CSQL2 = "Insert Into Kits(CODART,CODKIT,CANART) Values " & _
                        "('" & Salida.TextMatrix(I, 1) & "','" & Text1 & "'," & IIf(Trim(Salida.TextMatrix(I, 3)) = "" Or Salida.TextMatrix(I, 3) = "0", 1, Salida.TextMatrix(I, 3)) & ")"
                
                nTra = 1
                VGCNx.BeginTrans
                VGCNx.Execute CSQL2
                VGCNx.CommitTrans
                
                CSQL2 = "Update STKART Set STSKDIS=STSKDIS-" & Salida.TextMatrix(I, 3) & " where StAlma='" & VGAlma & "' and StCodigo='" & Salida.TextMatrix(I, 1) & "'"
                VGCNx.Execute CSQL2
                
                nTra = 0
            'End If
        Next I
    ElseIf nT = 2 Then     'Modificar             Trim(Mid(Combo1.text, 1, 1))
        For I = 1 To Salida.Rows - 1
            'If Salida.TextMatrix(I, 0) = ">>" Then
                If Existe(1, Text1, "Kits", "Codkit", False, Salida.TextMatrix(I, 1), "codart") = False Then
                    CSQL2 = "Update STKART Set STSKDIS=STSKDIS-" & Salida.TextMatrix(I, 3) & " where StAlma='" & VGAlma & "' and StCodigo='" & Salida.TextMatrix(I, 1) & "'"
                    VGCNx.Execute CSQL2
                    
                    CSQL2 = "Insert Into Kits(CODART,CODKIT,CANART) Values " & _
                            "('" & Salida.TextMatrix(I, 1) & "','" & Text1 & "'," & IIf(Trim(Salida.TextMatrix(I, 3)) = "" Or Salida.TextMatrix(I, 3) = "0", 1, Salida.TextMatrix(I, 3)) & ")"
                Else
                 cSql1 = "Select CodArt,CodKit,CANART from Kits where CODART = '" & Salida.TextMatrix(I, 1) & "' and CodKit='" & Text1 & "'"
                 Set cSel1 = New ADODB.Recordset
                 cSel1.Open cSql1, VGCNx, adOpenStatic
                 If cSel1.RecordCount > 0 Then
                    CSQL2 = "Update STKART Set STSKDIS=STSKDIS+" & cSel1("CanArt") & " where StAlma='" & VGAlma & "' and StCodigo='" & Salida.TextMatrix(I, 1) & "'"
                             VGCNx.Execute CSQL2
                 End If
                 cSel1.Close
                             
                    CSQL2 = "Update STKART Set STSKDIS=STSKDIS-" & Salida.TextMatrix(I, 3) & " where StAlma='" & VGAlma & "' and StCodigo='" & Salida.TextMatrix(I, 1) & "'"
                    VGCNx.Execute CSQL2
                    
                    CSQL2 = "Update Kits Set CODART = '" & Salida.TextMatrix(I, 1) & "'," & _
                            "CANART = " & IIf(Trim(Salida.TextMatrix(I, 3)) = "" Or Salida.TextMatrix(I, 3) = "0", 1, Salida.TextMatrix(I, 3)) & "  Where CODKIT = '" & Text1 & "' and CODART = '" & Salida.TextMatrix(I, 1) & "'"
                End If
                nTra = 1
                VGCNx.BeginTrans
                VGCNx.Execute CSQL2
                VGCNx.CommitTrans
                
                
                
                nTra = 0
            'End If
        Next I
    End If
    adodc1.Requery
    Set_Data
    adodc1.Find "ACODIGO = '" & Text1 & "'"
End If

If nT = 1 Then
    Limpiar
    Text1.SetFocus
ElseIf nT = 2 Then
    CmdSalir_Click
End If

Exit Sub
GrabErr:
    MsgBox Err.Description
    If nTra = 1 Then VGCNx.RollbackTrans
End Sub

Private Sub CmdIng_Click()      'Ingresar
nT = 1
Me.Caption = "Ingreso de Registro de Kits"
OculObj04 (False)
OculObj05 (False)
OculObj06 (False)
OculObj02 (True)
OculObj03 (True)
Limpiar
Text1.Enabled = True
Text1.SetFocus
End Sub

Private Sub CmdModi_Click()     'Modificar
If adodc1.RecordCount > 0 Then
    nT = 2
    Me.Caption = "Modificación de Registros de Kits"
    OculObj04 (False)
    OculObj05 (False)
    OculObj06 (False)
    OculObj02 (True)
    OculObj03 (True)
    Limpiar
    cCod = adodc1("ACODIGO")
    Text1.Enabled = False
    CmdGrabar.Visible = True
    Mostrar (cCod)
    Salida.SetFocus
Else
    MsgBox "No existen registros", vbInformation, "Mensaje"
End If
End Sub

Private Sub CmdSalir_Click()    'Salida principal del formulario
If nT = 1 Or nT = 2 Then
    Me.Caption = "Actualiza Registro de Kits"
    OculObj02 (False)
    OculObj03 (False)
    OculObj04 (True)
    OculObj05 (True)
    OculObj06 (True)
    InhabObj (True)
    CmdGrabar.Enabled = True
    CmdGrabar.Visible = False
    nT = 0
    DataGrid1.SetFocus
Else
    Unload Me
End If
End Sub

Private Sub command5_Click()
CrystalReport1.WindowTitle = "InvKits -- Control de Inventarios"
CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "invkits.rpt"
Ubi_Tab CrystalReport1
CrystalReport1.DiscardSavedData = True
CrystalReport1.Destination = crptToWindow
CrystalReport1.WindowTitle = " Control de Inventarios"
CrystalReport1.WindowShowPrintBtn = True
CrystalReport1.WindowShowRefreshBtn = True
CrystalReport1.WindowShowSearchBtn = True
CrystalReport1.WindowShowPrintSetupBtn = True
CrystalReport1.formulas(0) = "emp ='" & VGparametros.RucEmpresa & "'"
CrystalReport1.Action = 1

End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyBack Then
 If Len(TxFiltro) - 1 > 0 Then
  TxFiltro = Left(TxFiltro, Len(TxFiltro) - 1)
 Else
  TxFiltro = ""
 End If
 KeyAscii = 0
ElseIf KeyAscii <> 13 Then
 TxFiltro = TxFiltro & Chr(KeyAscii)
End If
End Sub

Private Sub Salida_DblClick()
VGRegEnt = 1: VGForm1 = 4
FormAyuArt.Show 1
Set_Flex
End Sub

Private Sub Form_Activate()
TxFiltro = ""
CmbOrden.ListIndex = 0
If DataGrid1.Visible Then DataGrid1.SetFocus
End Sub

Private Sub Form_Load()
central Me         'Centra Formulario
Init_ControlDataGrid DataGrid1
TxCantidad.ZOrder (0)
Limpiar
OculObj03 (False)
OculObj04 (True)
OculObj05 (True)
OculObj06 (True)
Set adodc1 = New ADODB.Recordset
adodc1.Open "SELECT distinct(ACODIGO),ADESCRI,AUNIDAD FROM KITS,MAEART WHERE ACODIGO = CODKIT  ORDER BY ACODIGO", VGCNx, adOpenStatic, adLockReadOnly
adodc1.Requery
Set DataGrid1.DataSource = adodc1
Set_Data
DataGrid1.Refresh
End Sub
Private Sub Limpiar()       'Limpia variables
Text1 = "": Label3 = "": TxCantidad = 1: TxCantidad.Visible = False
Set_Flex
Salida.Rows = 1
End Sub

Private Sub Salida_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then Salida_DblClick
End Sub

Private Sub Salida_KeyPress(KeyAscii As Integer)
Alinear

'If Salida.Col = 3 Then
'    TxCantidad.Visible = True
'    TxCantidad.SetFocus
'
'    If KeyAscii <> 13 And KeyAscii <> 27 And KeyAscii <> 9 And KeyAscii <> 8 Then
'        TxCantidad.text = TxCantidad.text & Chr(KeyAscii)
'    End If
'
'    TxCantidad.SelStart = Len(TxCantidad.text)
'    TxCantidad.SelLength = 0
If (Salida.Col = 3 And Val(Salida.TextMatrix(Salida.Row, 3)) >= 0) Then
        If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 46 Then
            TxCantidad.FontName = Salida.CellFontName
            TxCantidad.FontSize = Salida.CellFontSize
            TxCantidad.Width = Salida.CellWidth
            TxCantidad.Height = Salida.CellHeight
            TxCantidad.Left = Salida.Left + Salida.CellLeft
            TxCantidad.Top = Salida.Top + Salida.CellTop
            TxCantidad.Visible = True
            TxCantidad = Chr(KeyAscii)
            TxCantidad.SelStart = 1
            TxCantidad.SetFocus
        End If
 End If
'End If
End Sub

Private Sub Text1_DblClick()
Dim Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset
Adodc2.Open "Select ACODIGO, ADESCRI,AUNIDAD from MaeArt", VGCNx, adOpenStatic, adLockOptimistic
frmReferencia.Conectar Adodc2, "Select ACODIGO, ADESCRI,AUNIDAD from MaeArt"
frmReferencia.Label1.Caption = "Artículos"
frmReferencia.Show vbModal
Adodc2.Close
If vGUtil(1) <> "" Then Text1 = (vGUtil(1))
If vGUtil(2) <> "" Then Label3 = (vGUtil(2))
'Salida.AddItem "" & vbTab & Trim(Text1) & vbTab & Adodc2("ACodigo")
End Sub


Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then Text1_DblClick
End Sub
Private Sub TxCantidad_KeyPress(KeyAscii As Integer)
If NumPto(KeyAscii) Then
    Select Case KeyAscii
      Case Is = 13
         Salida.text = TxCantidad.text
         TxCantidad.Visible = False
         TxCantidad.text = ""
      Case Is = 27
         TxCantidad.Visible = False
         TxCantidad.text = ""
    End Select
Else
    KeyAscii = 0
End If
End Sub

Private Sub TxFiltro_Change()
If adodc1.RecordCount > 0 Then
    If Trim(TxFiltro) <> "" Then
        nCursor = adodc1.Bookmark
        adodc1.AbsolutePosition = 1
        adodc1.MoveFirst
        
        If CmbOrden.ListIndex = 0 Then
            adodc1.Find "ACODIGO like '" & Trim(UCase(TxFiltro)) & "*'"
        ElseIf CmbOrden.ListIndex = 1 Then
            adodc1.Find "ADESCRI like '" & Trim(UCase(TxFiltro)) & "*'"
        End If
        If adodc1.EOF Then adodc1.AbsolutePosition = nCursor
    End If
End If
End Sub

Private Sub Text1_GotFocus()
Enfoque Text1
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(Text1) = "" Then
        MsgBox "Ingrese el Código del articulo", vbInformation, "Mensaje"
        Text1.SetFocus
    Else
        If codigo(Text1) = False Then
            If Existe(1, Text1, "Kits", "CodKit", False) = False Then
                Salida.SetFocus
            Else
                MsgBox "El Código ya existe", vbInformation, "Mensaje"
                Text1.SetFocus
            End If
        Else
            MsgBox "El Código no existe,Tiene que registrarlo en la Tabla de articulo", vbInformation, "Mensaje"
            Text1.SetFocus
        End If
    End If
Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub

Private Sub Mostrar(cC1 As String) 'Muestra los datos
Dim cSqlM As String, cSelM As ADODB.Recordset
If Trim(cC1) = "" Then
    MsgBox "No hay registros para mostrar", vbInformation, "Mensaje"
    Exit Sub
End If
Salida.Visible = False
cSqlM = "Select CodKit,CodArt,CanArt From kits,MaeArt Where codkit = Acodigo AND codkit = '" & cC1 & "' "
Set cSelM = New ADODB.Recordset
cSelM.Open cSqlM, VGCNx, adOpenStatic
If cSelM.RecordCount > 0 Then
    Text1 = cSelM("codkit")
    Label3 = Devolver_Dato(1, Text1, "MaeArt", "Acodigo", False, "Adescri")
    Do While Not cSelM.EOF
        Salida.AddItem (" " & vbTab & cSelM("CodArt") & vbTab & Devolver_Dato(1, cSelM("CODART"), "MaeArt", "Acodigo", False, "Adescri") & vbTab & cSelM("CanArt"))
        cSelM.MoveNext
        If cSelM.EOF Then Exit Do
    Loop
Else
    MsgBox "No existe registro", vbInformation, "Mensaje"
    CmdSalir_Click
End If
Salida.Visible = True
cSelM.Close
End Sub

Private Sub InhabObj(nTipo As Boolean) ' Habilita e Inhabilita los objetos
Text1.Enabled = nTipo
End Sub

Private Sub Set_Data()
DataGrid1.Columns(0).Locked = True
DataGrid1.Columns(0).WrapText = True
DataGrid1.Columns(0).Caption = "   CODIGO"
DataGrid1.Columns(1).Caption = "       DESCRIPCION"
DataGrid1.Columns(2).Caption = "   UNIDAD"
DataGrid1.Columns(0).Width = 1800
DataGrid1.Columns(1).Width = 3800
DataGrid1.Columns(2).Width = 1200
End Sub
Private Sub OculObj02(nTipo As Boolean)  'Grabar y salir
CmdGrabar.Visible = nTipo
CmdSalir.Visible = nTipo
End Sub

Private Sub Set_Flex()
Salida.FormatString = "^Seleccion|  Codigo|   Descripcion|  Cantidad "
Salida.Row = 0
Salida.ColWidth(0) = 910
Salida.ColWidth(1) = 1200
Salida.ColWidth(2) = 3250
Salida.ColWidth(3) = 800
Salida.ColAlignment(1) = 1
End Sub
Sub Alinear()
TxCantidad.Width = Salida.CellWidth
TxCantidad.Left = Salida.CellLeft + Salida.Left
TxCantidad.Top = Salida.CellTop + Salida.Top
TxCantidad.Height = Salida.CellHeight
End Sub

