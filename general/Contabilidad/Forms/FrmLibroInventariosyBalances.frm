VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmLibroInventariosyBalances 
   Caption         =   "Libro de Inventarios y Balances"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11100
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   11100
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   960
      TabIndex        =   7
      Top             =   840
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   8055
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   9735
      Begin VB.Frame FrameCuentas 
         Height          =   6735
         Left            =   360
         TabIndex        =   1
         Top             =   960
         Width           =   9015
         Begin VB.Frame Frame2 
            Caption         =   "Opciones"
            Height          =   1215
            Left            =   2640
            TabIndex        =   2
            Top             =   5040
            Width           =   4215
            Begin VB.CommandButton Command1 
               Caption         =   "Impprimir"
               Height          =   615
               Left            =   360
               TabIndex        =   4
               Top             =   240
               Width           =   1575
            End
            Begin VB.CommandButton Command2 
               Caption         =   "Salir"
               Height          =   615
               Left            =   2400
               TabIndex        =   3
               Top             =   240
               Width           =   1575
            End
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   135
            Left            =   1320
            TabIndex        =   5
            Top             =   1560
            Width           =   30
            _ExtentX        =   53
            _ExtentY        =   238
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   4575
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   8070
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
   End
End
Attribute VB_Name = "FrmLibroInventariosyBalances"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rs1  As ADODB.Recordset
Private Sub Check1_Click()
Dim tipo As Integer
tipo = Check1.Value

   FrameCuentas.Visible = True
   SQL = " select formatocodigo,correlativo ,formatodescripcion2, reportegrupo  from ct_formatos where LibroSUNATCodigo =3 "
   SQL = SQL & " AND reporteGrupo>=1 order by correlativo "
   Set rs = VGCNx.Execute(SQL)
   Call LlenarLista(rs, tipo)

End Sub
Private Sub LlenarLista(ByRef rss As ADODB.Recordset, ByRef Tip As Integer)
 Dim i As Integer
 Dim itmX As ListItem
 
   ListView1.ColumnHeaders.Clear
   ListView1.ListItems.Clear
   ListView1.ColumnHeaders.Add , , "Tipo de Reporte", ListView1.Width / 1
   ListView1.View = lvwReport
   Do While Not rss.EOF
      Set itmX = ListView1.ListItems.Add(, , Str(rss!correlativo + 0) + "  " + rss!formatocodigo + "  " + rss!formatodescripcion2)
      ListView1.ListItems.Item(rss!correlativo + 0).Checked = Tip
      rss.MoveNext
   Loop
   
  End Sub

Private Sub Command1_Click()

If ExisteElem(0, VGCNx, VGcomputer) Then VGCNx.Execute (" drop table " & VGcomputer)
   VGCNx.Execute (" create table " & VGcomputer & "( formato varchar(20) , reportegrupo int )")
rs.MoveFirst
Do While Not rs.EOF
   If ListView1.ListItems.Item(rs!correlativo + 0).Checked = True Then
      SQL = " insert " & VGcomputer & "( formato, reportegrupo ) values ('" & rs!formatocodigo & "'," & rs!reportegrupo & ")"
      VGCNx.Execute (SQL)
   End If
  rs.MoveNext
Loop
Set rs = VGCNx.Execute(" select * from " & VGcomputer & "")
If rs.RecordCount > 0 Then
  Call imprimir
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
FrameCuentas.Visible = True
Check1.Value = 1
End Sub

Private Sub imprimir()
Dim aparam(6) As Variant
Dim aform(1) As Variant
aform(0) = "empresa='" & VGParametros.NomEmpresa & "'"
aparam(0) = VGCNx.DefaultDatabase
aparam(1) = VGParametros.empresacodigo
aparam(2) = VGParamSistem.Anoproceso
aparam(3) = VGParamSistem.Mesproceso
aparam(4) = "2"
rs.MoveFirst
Do While Not rs.EOF
   aparam(5) = rs!Formato
   If rs!reportegrupo = 1 Then
      Call ImpresionRptProc("ct_Libro03_03_06_12_13.rpt", aform, aparam, , "LIbro de Inventarios y Balances")
   ElseIf rs!reportegrupo = 2 Then
      Call ImpresionRptProc("ct_Libro03_11_14.rpt", aform, aparam, , "LIbro de Inventarios y Balances")
   ElseIf rs!reportegrupo = 3 Then
      Call ImpresionRptProc("ct_Libro03_11_14.rpt", aform, aparam, , "LIbro de Inventarios y Balances")
   End If
   
   rs.MoveNext
Loop


End Sub


