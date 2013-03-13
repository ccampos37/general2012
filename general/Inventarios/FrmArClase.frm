VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FrmArClase 
   Caption         =   "Tipo de Clase"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6870
   LinkTopic       =   "Form2"
   ScaleHeight     =   3465
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   1650
      Left            =   360
      TabIndex        =   8
      Top             =   270
      Visible         =   0   'False
      Width           =   6060
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   2175
         MaxLength       =   30
         TabIndex        =   10
         Text            =   "Text2"
         Top             =   900
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   2175
         MaxLength       =   20
         TabIndex        =   9
         Top             =   510
         Width           =   1980
      End
      Begin VB.Label Lb2 
         Caption         =   "Descripción :"
         Height          =   255
         Index           =   0
         Left            =   615
         TabIndex        =   12
         Top             =   930
         Width           =   1155
      End
      Begin VB.Label Lb1 
         Caption         =   "Código :"
         Height          =   255
         Left            =   615
         TabIndex        =   11
         Top             =   540
         Width           =   1155
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1965
      Left            =   345
      TabIndex        =   6
      Top             =   60
      Width           =   6150
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "FrmArClase.frx":0000
         Height          =   1485
         Left            =   210
         TabIndex        =   7
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
            DataField       =   "COD_MARCA"
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
            DataField       =   "DESCRI_MARCA"
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
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   330
      TabIndex        =   0
      Top             =   2070
      Width           =   6255
      Begin VB.CommandButton CmdIng 
         Caption         =   "&Ingreso"
         Height          =   675
         Left            =   660
         Picture         =   "FrmArClase.frx":0015
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton CmdModi 
         Caption         =   "&Modifica"
         Height          =   675
         Left            =   1695
         Picture         =   "FrmArClase.frx":0457
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton CmdEli 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   2715
         Picture         =   "FrmArClase.frx":0899
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   675
         Left            =   3750
         Picture         =   "FrmArClase.frx":0CDB
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton CmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   675
         Left            =   4785
         Picture         =   "FrmArClase.frx":111D
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   775
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   4110
      Top             =   2475
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "FrmArClase"
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
    cSql1 = "Delete from MAEMARCA Where COD_MARCA = '" & adodc1("COD_MARCA") & "' "
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
        If Existe(1, Text1(0), "MAEMARCA", "COD_MARCA", False) Then
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
    CSQL2 = "Insert Into MAEMARCA (COD_MARCA,DESCRI_MARCA)"
    CSQL2 = CSQL2 & " Values ('" & Text1(0) & "','" & SupCadSQL(Text1(1)) & "')"
    
ElseIf nOperador = 2 Then               'Si es Modificación
    CSQL2 = "Update MAEMARCA Set DESCRI_MARCA = '" & SupCadSQL(Text1(1)) & "' "
    CSQL2 = CSQL2 & "  Where COD_MARCA = '" & Text1(0) & "'"
End If

nTra = 1
VGcnx.BeginTrans
VGcnx.Execute CSQL2
VGcnx.CommitTrans
nTra = 0
adodc1.Requery
adodc1.Find "COD_MARCA = '" & Text1(0) & "'"
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
Frame2.Caption = "Ingreso de Tipo de Clase de Artículo "
Limpiar
nOperador = 1

Text1(0).Enabled = True: Text1(0).SetFocus
End Sub

Private Sub CmdModi_Click()      'Modificación
If adodc1.RecordCount > 0 Then
    Limpiar
    nOperador = 2
    Frame2.Caption = "Modificación de Tipo de Clase"
    cSql1 = "Select * from MAEMARCA where COD_MARCA = '" & adodc1.Fields("COD_MARCA") & "'"
    
    Set cSel1 = New ADODB.Recordset
    cSel1.Open cSql1, VGcnx, adOpenStatic
    
    If cSel1.RecordCount > 0 Then
        OculObj (True)
        If Not IsNull(cSel1("COD_MARCA")) Then Text1(0) = cSel1("COD_MARCA")
        If Not IsNull(cSel1("DESCRI_MARCA")) Then Text1(1) = cSel1("DESCRI_MARCA")
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
Dim SQL As String
Set adodc1 = New ADODB.Recordset
On Error GoTo Err
adodc1.Open "SELECT * FROM MAEMARCA ORDER BY COD_MARCA", VGcnx, adOpenStatic
Set DataGrid1.DataSource = adodc1
DataGrid1.Refresh

DataGrid1.Columns(0).Locked = True
DataGrid1.Columns(0).WrapText = True
DataGrid1.Columns(0).Caption = "Código"
DataGrid1.Columns(1).Caption = "Descripción "
DataGrid1.Columns(0).Locked = False
DataGrid1.Columns(0).WrapText = False

Me.Caption = "Tipo de Clase de articulo"
cTitulo = "TIPO DE CLASE DE ARTICULO"
If DataGrid1.Visible Then DataGrid1.SetFocus
Exit Sub
Err:
  MsgBox Err.Description & Chr(13) & "Salir del Formulario", vbInformation, "Aviso"
  If Not ExisteElem(0, VGcnx, "MAEMARCA") Then
        SQL = " Create Table MAEMARCA (COD_MARCA Text(20),DESCRI_MARCA Text(20), " & _
        " CONSTRAINT Clave PRIMARY KEY (COD_MARCA))"
        VGcnx.Execute SQL
  End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(Text1(Index)) <> "" Then
      If Index = 0 Then
            If Existe(1, Text1(0), "MAEMARCA", "COD_MARCA", False) Then
                 MsgBox "El codigo  ya existe", vbInformation, "Inventarios"
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

