VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FrmArTipoArticulo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo de Artículo"
   ClientHeight    =   3255
   ClientLeft      =   1665
   ClientTop       =   1350
   ClientWidth     =   6585
   Icon            =   "FrmTipoArti.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6585
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   165
      TabIndex        =   11
      Top             =   2055
      Width           =   6255
      Begin VB.CommandButton CmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   675
         Left            =   4785
         Picture         =   "FrmTipoArti.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   675
         Left            =   3750
         Picture         =   "FrmTipoArti.frx":0D0C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton CmdEli 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   2715
         Picture         =   "FrmTipoArti.frx":114E
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton CmdModi 
         Caption         =   "&Modifica"
         Height          =   675
         Left            =   1695
         Picture         =   "FrmTipoArti.frx":1590
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton CmdIng 
         Caption         =   "&Ingreso"
         Height          =   675
         Left            =   660
         Picture         =   "FrmTipoArti.frx":19D2
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
   Begin VB.Frame Frame1 
      Height          =   1965
      Left            =   180
      TabIndex        =   10
      Top             =   45
      Width           =   6150
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "FrmTipoArti.frx":1E14
         Height          =   1485
         Left            =   210
         TabIndex        =   12
         Top             =   300
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   2619
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
            DataField       =   "COD_TIPO"
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
            DataField       =   "DES_TIPO"
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
               ColumnWidth     =   2115.213
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3119.811
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1650
      Left            =   180
      TabIndex        =   7
      Top             =   165
      Visible         =   0   'False
      Width           =   6135
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   2175
         MaxLength       =   2
         TabIndex        =   3
         Top             =   510
         Width           =   540
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   2175
         MaxLength       =   32
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   900
         Width           =   3735
      End
      Begin VB.Label Lb1 
         Caption         =   "Código :"
         Height          =   255
         Left            =   615
         TabIndex        =   9
         Top             =   540
         Width           =   1155
      End
      Begin VB.Label Lb2 
         Caption         =   "Descripción :"
         Height          =   255
         Index           =   0
         Left            =   615
         TabIndex        =   8
         Top             =   930
         Width           =   1155
      End
   End
End
Attribute VB_Name = "FrmArTipoArticulo"
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
    cSql1 = "Delete from TIPO_ARTICULO Where COD_TIPO = '" & adodc1("COD_TIPO") & "' "
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
        If Existe(1, Text1(0), "TIPO_ARTICULO", "COD_TIPO", False) Then
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
    CSQL2 = "Insert Into Tipo_Articulo (COD_TIPO,DES_TIPO)"
    CSQL2 = CSQL2 & " Values ('" & Text1(0) & "','" & SupCadSQL(Text1(1)) & "')"
    
ElseIf nOperador = 2 Then               'Si es Modificación
    CSQL2 = "Update Tipo_Articulo Set DES_TIPO = '" & SupCadSQL(Text1(1)) & "' "
    CSQL2 = CSQL2 & "  Where COD_TIPO = '" & Text1(0) & "'"
End If

nTra = 1
VGcnx.BeginTrans
VGcnx.Execute CSQL2
VGcnx.CommitTrans
nTra = 0
adodc1.Requery

adodc1.Find "COD_TIPO = '" & Text1(0) & "'"

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
    If nTra = 1 Then VGcnx.RollbackTrans
End Sub

Private Sub CmdIng_Click()          'Ingreso
OculObj (True)
Frame2.Caption = "Ingreso de Tipo de Artículo "
Limpiar
nOperador = 1

Text1(0).Enabled = True: Text1(0).SetFocus
End Sub

Private Sub CmdModi_Click()      'Modificación
If adodc1.RecordCount > 0 Then
    Limpiar
    nOperador = 2
    Frame2.Caption = "Modificación de Tipo de Artículo"
    
    cSql1 = "Select * from Tipo_Articulo where COD_TIPO = '" & adodc1.Fields("COD_TIPO") & "'"
    
    Set cSel1 = New ADODB.Recordset
    cSel1.Open cSql1, VGcnx, adOpenStatic
    
    If cSel1.RecordCount > 0 Then
        OculObj (True)
        If Not IsNull(cSel1("COD_TIPO")) Then Text1(0) = cSel1("COD_TIPO")
        If Not IsNull(cSel1("DES_TIPO")) Then Text1(1) = cSel1("DES_TIPO")
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

adodc1.Open "SELECT * FROM TIPO_ARTICULO ORDER BY COD_TIPO", VGcnx, adOpenStatic
Set DataGrid1.DataSource = adodc1
DataGrid1.Refresh

DataGrid1.Columns(0).Locked = True
DataGrid1.Columns(0).WrapText = True
DataGrid1.Columns(0).Caption = "Código"
DataGrid1.Columns(1).Caption = "Descripción "
DataGrid1.Columns(0).Locked = False
DataGrid1.Columns(0).WrapText = False

Me.Caption = "Tipo de Artículos"
cTitulo = "TIPO DE ARTICULOS"
If DataGrid1.Visible Then DataGrid1.SetFocus
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(Text1(Index)) <> "" Then
      If Index = 0 Then
            If Existe(1, Text1(0), "TIPO_ARTICULO", "COD_TIPO", False) Then
                 MsgBox "La Cuenta Contable ya existe", vbInformation, "Inventarios"
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
End Sub
