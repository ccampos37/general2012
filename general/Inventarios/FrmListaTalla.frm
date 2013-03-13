VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmListaTalla 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de tallas"
   ClientHeight    =   3930
   ClientLeft      =   1665
   ClientTop       =   1350
   ClientWidth     =   7125
   Icon            =   "FrmListaTalla.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   7125
   Begin TabDlg.SSTab SSTab1 
      Height          =   2355
      Left            =   -30
      TabIndex        =   6
      Top             =   -15
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   4154
      _Version        =   393216
      TabOrientation  =   3
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Mant"
      TabPicture(0)   =   "FrmListaTalla.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Busqueda"
      TabPicture(1)   =   "FrmListaTalla.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame2 
         Height          =   2055
         Left            =   120
         TabIndex        =   9
         Top             =   15
         Visible         =   0   'False
         Width           =   6135
         Begin VB.TextBox txObserva 
            Height          =   285
            Left            =   2175
            MaxLength       =   50
            TabIndex        =   15
            Text            =   "Text2"
            Top             =   1395
            Width           =   3735
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   1
            Left            =   2175
            MaxLength       =   50
            TabIndex        =   11
            Text            =   "Text2"
            Top             =   915
            Width           =   3735
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   0
            Left            =   2175
            MaxLength       =   3
            TabIndex        =   10
            Top             =   510
            Width           =   540
         End
         Begin VB.Label Lb2 
            Caption         =   "Observación :"
            Height          =   255
            Index           =   1
            Left            =   600
            TabIndex        =   14
            Top             =   1440
            Width           =   1155
         End
         Begin VB.Label Lb2 
            Caption         =   "Descripción :"
            Height          =   255
            Index           =   0
            Left            =   615
            TabIndex        =   13
            Top             =   930
            Width           =   1155
         End
         Begin VB.Label Lb1 
            Caption         =   "Código :"
            Height          =   255
            Left            =   615
            TabIndex        =   12
            Top             =   540
            Width           =   1155
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   7
         Top             =   15
         Width           =   6165
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "FrmListaTalla.frx":0902
            Height          =   1755
            Left            =   210
            TabIndex        =   8
            Top             =   300
            Width           =   5835
            _ExtentX        =   10292
            _ExtentY        =   3096
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
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   "CODIGO"
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
               DataField       =   "DESCRIP"
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
            BeginProperty Column02 
               DataField       =   "OBSERVA"
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
                  ColumnWidth     =   794.835
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   975.118
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   3420.284
               EndProperty
            EndProperty
         End
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   15
      TabIndex        =   5
      Top             =   2325
      Width           =   6255
      Begin VB.CommandButton CmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   675
         Left            =   4785
         Picture         =   "FrmListaTalla.frx":0917
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   675
         Left            =   3750
         Picture         =   "FrmListaTalla.frx":0D59
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton CmdEli 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   2715
         Picture         =   "FrmListaTalla.frx":119B
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton CmdModi 
         Caption         =   "&Modifica"
         Height          =   675
         Left            =   1695
         Picture         =   "FrmListaTalla.frx":15DD
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton CmdIng 
         Caption         =   "&Ingreso"
         Height          =   675
         Left            =   660
         Picture         =   "FrmListaTalla.frx":1A1F
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   240
         Width           =   775
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   3945
      Top             =   2460
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "FrmListaTalla"
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
Dim cBase As String

Private Sub CmdEli_Click()              'Eliminar
Dim nPosi As Integer
On Error GoTo EliErr
If adodc1.RecordCount > 0 Then
    cSql1 = "Delete from TALLA Where CODIGO = '" & adodc1("CODIGO") & "' "
    If MsgBox("Seguro de Eliminar ?", vbQuestion + vbOKCancel, "Inventarios") = vbOK Then
        nPosi = Pos_Dato(adodc1)
        nTra = 1
        VGcnx.BeginTrans
        VGcnx.Execute cSql1
        VGcnx.CommitTrans
        nTra = 0: adodc1.Requery
        If nPosi <> 0 Then adodc1.AbsolutePosition = nPosi
    End If
    If DataGrid1.Visible Then DataGrid1.SetFocus
Else
    MsgBox "No existe ningún registro para Eliminar", vbInformation, "Inventarios"
End If
Exit Sub
EliErr:
    MsgBox Err.Description
    If nTra = 1 Then VGcnx.RollbackTrans
End Sub

Private Sub CmdGrabar_Click()          ' Grabar
On Error GoTo GrabErr

If nOperador = 1 Then                  ' Si es Ingreso
    If Trim(Text1(0)) = "" Then
        MsgBox "Ingrese Código", vbInformation, "Mensaje"
        Text1(0).SetFocus: Exit Sub
    Else
        If Existe(1, Text1(0), "TALLA", "CODIGO", False) Then
                MsgBox "El Código ya existe", vbInformation, "Inventarios"
                Text1(0).SetFocus: Exit Sub
        End If
     End If
End If
If Trim(Text1(1)) = "" Then
       MsgBox "Ingrese Descripción", vbInformation, "Mensaje"
       Text1(1).SetFocus: Exit Sub
End If
    
If nOperador = 1 Then                  ' Si es Ingreso
    CSQL2 = "Insert Into TALLA (CODIGO,DESCRIP,OBSERVA)"
    CSQL2 = CSQL2 & " Values ('" & Text1(0) & "','" & SupCadSQL(Text1(1)) & "','" & SupCadSQL(txObserva) & "')"
    
ElseIf nOperador = 2 Then               'Si es Modificación
    CSQL2 = "Update TALLA Set DESCRIP = '" & Trim(SupCadSQL(Text1(1))) & "',"
    CSQL2 = CSQL2 & "OBSERVA='" & Trim(SupCadSQL(txObserva)) & "'"
    CSQL2 = CSQL2 & "  Where CODIGO = '" & Text1(0) & "'"
End If

nTra = 1
VGcnx.BeginTrans
VGcnx.Execute CSQL2
VGcnx.CommitTrans
nTra = 0
adodc1.Requery

adodc1.Find "CODIGO = '" & Text1(0) & "'"

If nOperador = 1 Then
    SSTab1.Tab = 0
    OculObj (True)
    Limpiar
    Text1(0).SetFocus
ElseIf nOperador = 2 Then
    SSTab1.Tab = 1
    OculObj (False)
    nOperador = 0
    DataGrid1.SetFocus
End If

Exit Sub
GrabErr:
    MsgBox Err.Description
    If nTra = 1 Then VGcnx.RollbackTrans
End Sub

Private Sub CmdIng_Click()          'Ingreso
SSTab1.Tab = 0
OculObj (True)
Frame2.Caption = "Ingreso de Lista de tallas "
Limpiar
nOperador = 1

Text1(0).Enabled = True: Text1(0).SetFocus
End Sub

Private Sub CmdModi_Click()      'Modificación
If adodc1.RecordCount > 0 Then
    SSTab1.Tab = 0
    Limpiar
    nOperador = 2
    Frame2.Caption = "Modificación de Lista de tallas"
    
    cSql1 = "Select * from Talla where CODIGO = '" & adodc1.Fields("CODIGO") & "'"
    
    Set cSel1 = New ADODB.Recordset
    cSel1.Open cSql1, VGcnx, adOpenStatic
    
    If cSel1.RecordCount > 0 Then
        OculObj (True)
        If Not IsNull(cSel1("CODIGO")) Then Text1(0) = cSel1("CODIGO")
        If Not IsNull(cSel1("DESCRIP")) Then Text1(1) = cSel1("DESCRIP")
        If Not IsNull(cSel1("OBSERVA")) Then txObserva = cSel1("OBSERVA")
        
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

Private Sub CmdSalir_Click()
If nOperador = 1 Or nOperador = 2 Then
    SSTab1.Tab = 1
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
Me.Width = 6390: Me.Height = 3930
SSTab1.Tab = 1
central Me                      'Centra Formulario
Init_ControlDataGrid DataGrid1
CarObj                          'Carga el Adodc y el datagrid1
End Sub

Public Sub OculObj(bTip As Boolean)
Frame2.Visible = bTip
CmdIng.Enabled = Not bTip
CmdModi.Enabled = Not bTip
CmdEli.Enabled = Not bTip
CmdIng.Enabled = Not bTip
CmdGrabar.Enabled = bTip
Frame1.Visible = Not bTip
End Sub

Private Sub Text1_GotFocus(Index As Integer)
Enfoque Text1(Index)
End Sub

Private Sub CarObj()
Set adodc1 = New ADODB.Recordset

adodc1.Open "SELECT * FROM TALLA ORDER BY CODIGO", VGcnx, adOpenStatic
Set DataGrid1.DataSource = adodc1
DataGrid1.Refresh

DataGrid1.Columns(0).Locked = True
DataGrid1.Columns(0).WrapText = True
DataGrid1.Columns(0).Caption = "Código"
DataGrid1.Columns(1).Caption = "Descripción "
DataGrid1.Columns(2).Caption = "Observación "

DataGrid1.Columns(0).Locked = False
DataGrid1.Columns(0).WrapText = False

Me.Caption = "Lista de Tallas"
cTitulo = "Lista de Tallas "
If DataGrid1.Visible Then DataGrid1.SetFocus
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(Text1(Index)) <> "" Then
      If Index = 0 Then
            If Existe(1, Text1(0), "TALLA", "CODIGO", False) Then
                 MsgBox "El codigo ya existe ", vbInformation, "Inventarios"
                Text1(0).SetFocus: Exit Sub
            End If
          Text1(1).SetFocus: Exit Sub
      ElseIf Index = 1 Then
            CmdGrabar.SetFocus: Exit Sub
        End If
    Else
       If Index = 0 Then
          MsgBox "Ingrese Código ", vbInformation, "Inventarios"
       Else
          MsgBox "Ingrese Descripción ", vbInformation, "Inventarios"
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
txObserva = ""
End Sub
