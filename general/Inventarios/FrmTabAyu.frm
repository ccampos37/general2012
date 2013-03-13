VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmArTabAyu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "r"
   ClientHeight    =   3270
   ClientLeft      =   2280
   ClientTop       =   2655
   ClientWidth     =   6585
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   6585
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   135
      TabIndex        =   13
      Top             =   2070
      Width           =   6255
      Begin VB.CommandButton CmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5340
         Picture         =   "FrmTabAyu.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   675
         Left            =   4290
         Picture         =   "FrmTabAyu.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton CmdRep 
         Caption         =   "&Reporte"
         Height          =   675
         Left            =   3240
         Picture         =   "FrmTabAyu.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton CmdEli 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   2205
         Picture         =   "FrmTabAyu.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton CmdModi 
         Caption         =   "&Modificar"
         Height          =   675
         Left            =   1200
         Picture         =   "FrmTabAyu.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton CmdIng 
         Caption         =   "&Ingreso"
         Height          =   675
         Left            =   120
         Picture         =   "FrmTabAyu.frx":154A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   775
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   5160
      Top             =   2340
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Height          =   2010
      Left            =   180
      TabIndex        =   12
      Top             =   60
      Width           =   6135
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "FrmTabAyu.frx":198C
         Height          =   1350
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   2381
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
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
            DataField       =   "TCLAVE"
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
            DataField       =   "TDESCRI"
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
            ScrollBars      =   2
            BeginProperty Column00 
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   4275.213
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   195
      TabIndex        =   9
      Top             =   330
      Visible         =   0   'False
      Width           =   6135
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   2055
         MaxLength       =   2
         TabIndex        =   5
         Top             =   285
         Width           =   1275
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   2055
         MaxLength       =   40
         TabIndex        =   6
         Text            =   "Text2"
         Top             =   645
         Width           =   3735
      End
      Begin VB.Label Lb1 
         Caption         =   "Lb1"
         Height          =   255
         Left            =   375
         TabIndex        =   11
         Top             =   315
         Width           =   1515
      End
      Begin VB.Label Lb2 
         Caption         =   "Descripción              :"
         Height          =   255
         Left            =   375
         TabIndex        =   10
         Top             =   675
         Width           =   1635
      End
   End
End
Attribute VB_Name = "FrmArTabAyu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public G_nOpc As Integer  'Del Menu
Public G_cTabla As String  'Del Menu
Dim adodc1 As ADODB.Recordset
Dim cSel1 As ADODB.Recordset
Dim nOpc As Integer
Dim cTabla As String, cSql1 As String
Dim CSQL2 As String, cClave As String
Dim nTra As Integer, nTra2 As Integer
Dim nOperador As Byte
Dim cTitulo As String
Dim cMensaje As String

Private Sub CmdEli_Click()              'Eliminar
Dim nPosi As Integer
On Error GoTo EliErr
If adodc1.RecordCount > 0 Then
    cSql1 = "Delete from Tabayu Where tcod = '" & cTabla & "' and tclave = '" & adodc1("TCLAVE") & "' "

    If MsgBox("Seguro de Eliminar ?", vbQuestion + vbOKCancel, "Inventarios") = vbOK Then
        nPosi = Pos_Dato(adodc1)
        nTra = 1
        VGCNx.BeginTrans
        VGCNx.Execute cSql1
        VGCNx.CommitTrans
        nTra = 0: adodc1.Requery
        DataGrid1.Refresh
        
        If nPosi <> 0 Then adodc1.AbsolutePosition = nPosi
    End If
    If DataGrid1.Visible Then DataGrid1.SetFocus
Else
    MsgBox "No existe ningún registro para Eliminar", vbInformation, "Inventarios"
End If
Exit Sub
EliErr:
    MsgBox Err.Description
    If nTra = 1 Then VGCNx.RollbackTrans
End Sub

Private Sub CmdGrabar_Click()          ' Grabar
On Error GoTo GrabErr

If nOperador = 1 Then                  ' Si es Ingreso
    If Trim(Text1(0)) = "" Then
        MsgBox "Ingrese Código", vbInformation, "Mensaje"
        Text1(0).SetFocus: Exit Sub
    Else
        If ValiDocu(cTabla, Trim(Text1(0))) = False Then
            Text1(0).SetFocus
            Exit Sub
        End If
    End If
    If Trim(Text1(1)) = "" Then
        MsgBox "Ingrese Descripción", vbInformation, "Mensaje"
        Text1(1).SetFocus: Exit Sub
    End If
    
    CSQL2 = "Insert Into TabAyu (tcod,tclave,tdescri)"
    CSQL2 = CSQL2 & " Values ('" & cTabla & "','" & Text1(0) & "','" & SupCadSQL(Text1(1)) & "')"
ElseIf nOperador = 2 Then               'Si es Modificación
    CSQL2 = "Update TabAyu Set tdescri = '" & SupCadSQL(Text1(1)) & "' "
    CSQL2 = CSQL2 & "  Where tcod = '" & cTabla & "' and tclave = '" & Text1(0) & "'"
End If

nTra = 1
VGCNx.BeginTrans
VGCNx.Execute CSQL2
VGCNx.CommitTrans
nTra = 0
adodc1.Requery
 If adodc1.RecordCount > 0 Then adodc1.MoveFirst
 Do While Not adodc1.EOF
        If adodc1("tcod") = cTabla And adodc1("tclave") = Text1(0) Then Exit Do
        adodc1.MoveNext
        If adodc1.EOF Then Exit Do
 Loop
If nOperador = 1 Then
    OculObj (True)
    Text1(0) = "": Text1(1) = ""
    Text1(0).SetFocus
ElseIf nOperador = 2 Then
    OculObj (False)
    nOperador = 0
    DataGrid1.SetFocus
End If
'CarObj
Exit Sub
GrabErr:
    MsgBox Err.Description
    If nTra = 1 Then VGCNx.RollbackTrans
End Sub

Private Sub CmdIng_Click()          'Ingreso
OculObj (True)
Frame2.Caption = "Ingreso"
nOperador = 1
LLenar_Label
Text1(0).Enabled = True: Text1(0).SetFocus
End Sub

Private Sub CmdModi_Click()      'Modificación
If adodc1.RecordCount > 0 Then
    nOperador = 2
    Frame2.Caption = "Modificación"
    cClave = adodc1("tclave")
    cSql1 = "Select * from Tabayu where tcod = '" & cTabla & "' and tclave = '" & cClave & "'"
    
    Set cSel1 = New ADODB.Recordset
    cSel1.Open cSql1, VGCNx, adOpenStatic

    If cSel1.RecordCount > 0 Then
        OculObj (True)
        LLenar_Label
        If Not IsNull(cSel1("tclave")) Then Text1(0) = cSel1("tclave")
        If Not IsNull(cSel1("tdescri")) Then Text1(1) = cSel1("tdescri")
        Text1(0).Enabled = False
        Text1(1).SetFocus
    Else
        MsgBox "El registro ha sido Eliminado", vbInformation, "Inventarios"
    End If
    cSel1.Close
Else
    MsgBox "No existe ningún registro para modificar", vbInformation, "Inventarios"
End If
End Sub

Private Sub CmdRep_Click() 'REPORTE
Dim CADENA As String
Dim cNomRepor  As String
Dim cTituloReport As String

Select Case nOpc
Case 7
      'Me.Caption = "Tabla de Distritos"
      cNomRepor = "distrito.RPT"
      cTituloReport = "Reporte Distritos"
Case 9
      'Me.Caption = "Tabla de Zonas"
      cNomRepor = ".RPT"
      cTituloReport = "Reporte "
Case 10
      'Me.Caption = "Tabla de Giro de Proveedor"
      cNomRepor = "giroproveedor.RPT"
      cTituloReport = "Reporte de Giro de Proveedores"
Case 11
      'Me.Caption = "Tabla de Tarjeta de Crédito"
      cNomRepor = ".RPT"
      cTituloReport = "Reporte "
Case 14
      'Me.Caption = "Tabla de Territorios"
      cNomRepor = ".RPT"
      cTituloReport = "Reporte "
Case 15
      'Me.Caption = "Tabla de Rutas"
      cNomRepor = ".RPT"
      cTituloReport = "Reporte "
Case 16
      'Me.Caption = "Tabla de Segmentos"
      cNomRepor = ".RPT"
      cTituloReport = "Reporte "
Case 17
      'Me.Caption = "Tabla de Ubicación de Segmentos"
      cNomRepor = ".RPT"
      cTituloReport = "Reporte "
End Select

If Trim(cNomRepor) <> "" Then
    CrystalReport1.Reset
    CrystalReport1.WindowTitle = cTituloReport
    CrystalReport1.ReportFileName = VGParamSistem.RutaReport + "\" + cNomRepor
  
    CrystalReport1.Connect = VGcadenareport2
    CrystalReport1.StoredProcParam(0) = VGParamSistem.BDEmpresa
    
    CrystalReport1.formulas(0) = "emp ='" & VGparametros.RucEmpresa & "'"
    
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.WindowState = crptMaximized
    If CrystalReport1.Status <> 2 Then
        CrystalReport1.Action = 1
    End If
Else
    MsgBox "No existe el nombre del Reporte, verifique en Formatos", vbInformation, "Información"
    Exit Sub
End If


End Sub

Private Sub CmdSalir_Click()
If nOperador = 1 Or nOperador = 2 Then
    OculObj (False)
    If DataGrid1.Visible Then DataGrid1.SetFocus
    nOperador = 0
Else
    Unload Me
End If
End Sub

Private Sub Form_Activate()
If DataGrid1.Visible And DataGrid1.Enabled Then DataGrid1.SetFocus
End Sub

Private Sub Form_Load()
central Me                      'Centra Formulario
Init_ControlDataGrid DataGrid1
cTabla = G_cTabla
nOpc = G_nOpc
Set adodc1 = New ADODB.Recordset
cMensaje = ""
CarObj   'Carga el Adodc y el datagrid1
End Sub

Private Function ValiDocu(pCod As String, pClave As String) As Boolean
Dim csql As String, cRec As ADODB.Recordset

If Trim(pClave) = "" Then
   MsgBox "Falta ingresar Código de " & cMensaje, vbInformation, "Inventarios"
   ValiDocu = False
   Exit Function
End If
csql = "SELECT * FROM TABAYU WHERE TCOD = '" & pCod & "' AND TCLAVE = '" & pClave & "' ORDER BY TCLAVE"
Set cRec = New ADODB.Recordset
cRec.Open csql, VGCNx, adOpenStatic
If cRec.RecordCount > 0 Then
    MsgBox "Código de " & cMensaje & "  ya existe", vbInformation, "Inventarios"
    ValiDocu = False
    Exit Function
End If
ValiDocu = True: cRec.Close
End Function

Public Sub OculObj(bTip As Boolean)
Frame2.Visible = bTip
CmdIng.Enabled = Not bTip
CmdModi.Enabled = Not bTip
CmdEli.Enabled = Not bTip
CmdRep.Enabled = Not bTip
CmdIng.Enabled = Not bTip
Cmdgrabar.Enabled = bTip
Frame1.Visible = Not bTip
End Sub

Private Sub Text1_GotFocus(Index As Integer)
Enfoque Text1(Index)
End Sub

Private Sub CarObj()

Set adodc1 = New ADODB.Recordset

adodc1.Open "SELECT * FROM TABAYU WHERE TCOD = '" & cTabla & "' ORDER BY TCLAVE", VGCNx, adOpenStatic
Set DataGrid1.DataSource = adodc1
DataGrid1.Refresh

DataGrid1.Columns(0).Locked = True
DataGrid1.Columns(0).WrapText = True
DataGrid1.Columns(0).Caption = "Código"
DataGrid1.Columns(1).Caption = "Descripción"
DataGrid1.Columns(0).Locked = False
DataGrid1.Columns(0).WrapText = False

Select Case nOpc
Case 7
      Me.Caption = "Tabla de Distritos"
      cTitulo = "DISTRITOS"
      cMensaje = "Distrito"
Case 9
      Me.Caption = "Tabla de Zonas"
      cTitulo = "ZONAS DE VENTA"
      cMensaje = "Zonas de Venta"
Case 10
      Me.Caption = "Tabla de Giro de Proveedor"
      cTitulo = "GIRO DEL PROVEEDOR"
      cMensaje = "Giro del Proveedor"
Case 11
      Me.Caption = "Tabla de Tarjeta de Crédito"
      cTitulo = "TARJETA DE CREDITO"
      cMensaje = "Tarjetas de Credito"
Case 14
      Me.Caption = "Tabla de Territorios"
      cTitulo = "TERRITORIOS"
      cMensaje = "Territorios"
Case 15
      Me.Caption = "Tabla de Rutas"
      cTitulo = "RUTAS"
      cMensaje = "Rutas"
Case 16
      Me.Caption = "Tabla de Segmentos"
      cTitulo = "SEGMENTOS"
      cMensaje = "Segmentos"
Case 17
      Me.Caption = "Tabla de Ubicación de Segmentos"
      cTitulo = "UBICACION DE SEGMENTOS"
      cMensaje = "Ubicación de Segmentos"
End Select
If DataGrid1.Visible Then DataGrid1.SetFocus
End Sub

Private Sub LLenar_Label()
Select Case nOpc
Case 7
        Lb1 = "Distritos     :"
Case 9
        Lb1 = "Zona Venta    :"
Case 10
        Lb1 = "Giro Proveedor  :"
Case 11
        Lb1 = "Tarjeta Cred. :"
Case 14
        Lb1 = "Territorios   :"
Case 15
        Lb1 = "Rutas         :"
Case 16
        Lb1 = "Segementos    :"
Case 17
        Lb1 = "Ubic.Segmento :"
End Select
Text1(0) = "": Text1(1) = ""
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If Index = 0 Then
        If ValiDocu(cTabla, Trim(Text1(0))) = False Then
            Text1(0).SetFocus
            Exit Sub
        End If
    End If
   SendKeys "{tab}"
Else
    If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
    If Index = 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub
