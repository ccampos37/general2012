VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmArDocumento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipos de Documentos"
   ClientHeight    =   3690
   ClientLeft      =   1665
   ClientTop       =   1350
   ClientWidth     =   6585
   Icon            =   "FrmDocumento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   6585
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   150
      TabIndex        =   14
      Top             =   2430
      Width           =   6255
      Begin VB.CommandButton CmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5340
         Picture         =   "FrmDocumento.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   675
         Left            =   4290
         Picture         =   "FrmDocumento.frx":0D0C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton CmdRep 
         Caption         =   "&Reporte"
         Height          =   675
         Left            =   3255
         Picture         =   "FrmDocumento.frx":114E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton CmdEli 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   2205
         Picture         =   "FrmDocumento.frx":1590
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton CmdModi 
         Caption         =   "&Modificar"
         Height          =   675
         Left            =   1170
         Picture         =   "FrmDocumento.frx":19D2
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton CmdIng 
         Caption         =   "&Ingreso"
         Height          =   675
         Left            =   120
         Picture         =   "FrmDocumento.frx":1E14
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   775
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   5160
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      Height          =   1995
      Left            =   180
      TabIndex        =   10
      Top             =   180
      Visible         =   0   'False
      Width           =   6135
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   5190
         MaxLength       =   2
         TabIndex        =   16
         Text            =   "Text2"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   2175
         MaxLength       =   2
         TabIndex        =   7
         Text            =   "Text2"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   2175
         MaxLength       =   2
         TabIndex        =   5
         Top             =   510
         Width           =   555
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   2175
         MaxLength       =   30
         TabIndex        =   6
         Text            =   "Text2"
         Top             =   915
         Width           =   3735
      End
      Begin VB.Label Lb2 
         Caption         =   "Cod. Sunat       :"
         Height          =   255
         Index           =   2
         Left            =   3630
         TabIndex        =   17
         Top             =   1335
         Width           =   1440
      End
      Begin VB.Label Lb2 
         Caption         =   "Cod. Contable     :"
         Height          =   255
         Index           =   1
         Left            =   615
         TabIndex        =   15
         Top             =   1350
         Width           =   1440
      End
      Begin VB.Label Lb1 
         Caption         =   "Codigo                :  "
         Height          =   255
         Left            =   615
         TabIndex        =   12
         Top             =   540
         Width           =   1395
      End
      Begin VB.Label Lb2 
         Caption         =   "Descripción         :"
         Height          =   255
         Index           =   0
         Left            =   615
         TabIndex        =   11
         Top             =   945
         Width           =   1425
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2310
      Left            =   180
      TabIndex        =   13
      Top             =   60
      Width           =   6135
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "FrmDocumento.frx":2256
         Height          =   1890
         Left            =   180
         TabIndex        =   0
         Top             =   240
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   3334
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
            DataField       =   "TDO_TIPDOC"
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
            DataField       =   "TDO_DESCRI"
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
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3974.74
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmArDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adodc1 As ADODB.Recordset
Dim cSel1 As ADODB.Recordset
Dim nOpc As Integer
Dim cTabla As String, cSql1 As String
Dim CSQL2 As String, cClave As String
Dim nTra As Integer, nTra2 As Integer
Dim nOperador As Byte
Dim cTitulo As String

Private Sub CmdEli_Click()              'Eliminar
Dim nPosi As Integer
On Error GoTo EliErr
If adodc1.RecordCount > 0 Then
    cSql1 = "Delete from Tipo_Docu Where TDO_TIPDOC = '" & adodc1("TDO_TIPDOC") & "' "
    If MsgBox("Seguro de Eliminar ?", vbQuestion + vbOKCancel, "Sistema de Ventas") = vbOK Then
        If Existe(1, adodc1("TDO_TIPDOC"), "FacCab", "CFTD", False) Then
            MsgBox "El Documento no se puede eliminar, porque ya tiene un documento registrado", vbInformation, "Información"
        Else
            nPosi = Pos_Dato(adodc1)
            nTra = 1
            VGCNx.BeginTrans
            VGCNx.Execute cSql1
            VGCNx.CommitTrans
            nTra = 0: adodc1.Requery
            If nPosi <> 0 Then adodc1.AbsolutePosition = nPosi
        End If
    End If
    If DataGrid1.Visible Then DataGrid1.SetFocus
Else
    MsgBox "No existe ningún registro para Eliminar", vbInformation, "Sistema de Ventas"
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
        If Existe(1, Text1(0), "Tipo_Docu", "TDO_TIPDOC", False) Then
            MsgBox "El Código ya existe", vbInformation, "Sistema de Ventas"
            Text1(0).SetFocus: Exit Sub
        End If
        
    End If
 End If
    If Trim(Text1(1)) = "" Then
        MsgBox "Ingrese Descripción", vbInformation, "Mensaje"
        Text1(1).SetFocus: Exit Sub
    End If
    'If Trim(Text1(2)) = "" Then
    '    MsgBox "Ingrese Codigo Contable", vbInformation, "Mensaje"
    '    Text1(2).SetFocus: Exit Sub
    'End If
    
    'If Trim(Text1(3)) = "" Then
    '    MsgBox "Ingrese Codigo Sunat", vbInformation, "Mensaje"
    '    Text1(3).SetFocus: Exit Sub
    'End If
    
If nOperador = 1 Then                  ' Si es Ingreso
    CSQL2 = "Insert Into Tipo_Docu (TDO_TIPDOC,TDO_DESCRI,TDO_CODCON,TDO_CODSUN)"
    CSQL2 = CSQL2 & " Values ('" & Text1(0) & "','" & SupCadSQL(Text1(1)) & "','" & IIf(Trim(Text1(2)) <> "", Text1(2), "  ") & "','" & IIf(Trim(Text1(3)) <> "", Text1(3), "  ") & "')"
    
ElseIf nOperador = 2 Then               'Si es Modificación
    CSQL2 = "Update Tipo_Docu Set TDO_DESCRI = '" & SupCadSQL(Text1(1)) & "',TDO_CODCON = '" & IIf(Trim(Text1(2)) <> "", Text1(2), "  ") & "',TDO_CODSUN = '" & IIf(Trim(Text1(3)) <> "", Text1(3), "  ") & "'"
    CSQL2 = CSQL2 & "  Where TDO_TIPDOC = '" & Text1(0) & "'"
End If


nTra = 1
VGCNx.BeginTrans
VGCNx.Execute CSQL2
VGCNx.CommitTrans
nTra = 0
adodc1.Requery

adodc1.Find "TDO_TIPDOC = '" & Text1(0) & "'"

If nOperador = 1 Then
    OculObj (True)
    Limpiar
    Text1(0).SetFocus
ElseIf nOperador = 2 Then
    OculObj (False)
    nOperador = 0
    DataGrid1.SetFocus
End If
Exit Sub
GrabErr:
    MsgBox Err.Description
    If nTra = 1 Then VGCNx.RollbackTrans
End Sub

Private Sub CmdIng_Click()          'Ingreso
OculObj (True)
Frame2.Caption = "Ingreso de Tipos de Documentos "
Limpiar
nOperador = 1
Text1(0).Enabled = True: Text1(0).SetFocus
End Sub

Private Sub CmdModi_Click()      'Modificación
If adodc1.RecordCount > 0 Then
    Limpiar
    nOperador = 2
    Frame2.Caption = "Modificación de Tipo de Documentos"
    cSql1 = "Select * from Tipo_Docu where Tdo_TipDoc = '" & adodc1.Fields("Tdo_TipDoc") & "'"
    
    Set cSel1 = New ADODB.Recordset
    cSel1.Open cSql1, VGCNx, adOpenStatic

    If cSel1.RecordCount > 0 Then
        OculObj (True)
        If Not IsNull(cSel1("Tdo_TipDoc")) Then Text1(0) = cSel1("Tdo_TipDoc")
        If Not IsNull(cSel1("TDO_DESCRI")) Then Text1(1) = cSel1("TDO_DESCRI")
        If Not IsNull(cSel1("TDO_CODCON")) Then Text1(2) = cSel1("TDO_CODCON")
        If Not IsNull(cSel1("TDO_CODSUN")) Then Text1(3) = cSel1("TDO_CODSUN")
        Text1(0).Enabled = False
        Text1(1).SetFocus
    Else
        MsgBox "El registro ha sido Eliminado", vbInformation, "Sistema de Ventas"
    End If
    cSel1.Close
Else
    MsgBox "No existe ningún registro para modificar", vbInformation, "Sistema de Ventas"
End If
End Sub

Private Sub CmdRep_Click() 'REPORTE
   Dim CADENA As String
    Dim cNomRepor  As String

cNomRepor = "documentos.RPT"
If Trim(cNomRepor) <> "" Then
    CrystalReport1.Reset
    CrystalReport1.WindowTitle = "Reporte de Documentos"
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
central Me                       'Centra Formulario
Init_ControlDataGrid DataGrid1
CarObj                          'Carga el Adodc y el datagrid1
End Sub

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

adodc1.Open "SELECT * FROM TIPO_DOCU ORDER BY TDO_TIPDOC", VGCNx, adOpenStatic
Set DataGrid1.DataSource = adodc1
DataGrid1.Refresh

DataGrid1.Columns(0).Locked = True
DataGrid1.Columns(0).WrapText = True
DataGrid1.Columns(0).Caption = "Código"
DataGrid1.Columns(1).Caption = "Descripción"
DataGrid1.Columns(0).Locked = False
DataGrid1.Columns(0).WrapText = False

Me.Caption = "Tipos de Documentos"
cTitulo = "TIPOS DE DOCUMENTOS"
If DataGrid1.Visible Then DataGrid1.SetFocus
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
   
    If Trim(Text1(Index)) <> "" Then
      If Index = 0 Then
          If Existe(1, Text1(0), "Tipo_Docu", "TDO_TIPDOC", False) Then
             MsgBox "El Código ya existe", vbInformation, "Sistema de Ventas"
             Text1(0).SetFocus: Exit Sub
          End If
      End If
      
      If Index <> 3 Then
         Text1(Index + 1).SetFocus
      Else
         Cmdgrabar.SetFocus
      End If
      
    Else
       If Index = 0 Then
          MsgBox "Ingrese Código ", vbInformation, "Sistema de Ventas"
          
       ElseIf Index = 1 Then
          MsgBox "Ingrese Descripción", vbInformation, "Mensaje"
       ElseIf Index = 2 Then
          'MsgBox "Ingrese Codigo Contable", vbInformation, "Mensaje"
       End If
       Text1(Index).SetFocus: Exit Sub
    End If
Else
    If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
    If Index = 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub
Sub Limpiar()
Dim otext As TextBox

 For Each otext In Me.Text1
  otext.text = ""
 Next
End Sub
