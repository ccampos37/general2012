VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmLibrosElectPrincipales 
   Caption         =   "Form1"
   ClientHeight    =   7770
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   ScaleHeight     =   7770
   ScaleWidth      =   10185
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   7215
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   9735
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   375
         Left            =   7080
         TabIndex        =   12
         Top             =   240
         Width           =   2295
      End
      Begin VB.Frame FrameCuentas 
         Height          =   6495
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   9015
         Begin VB.Frame Frame2 
            Caption         =   "Opciones"
            Height          =   1215
            Left            =   2640
            TabIndex        =   4
            Top             =   5040
            Width           =   4215
            Begin VB.CommandButton Command1 
               Caption         =   "Grabar"
               Height          =   615
               Left            =   360
               TabIndex        =   6
               Top             =   240
               Width           =   1575
            End
            Begin VB.CommandButton Command2 
               Caption         =   "Salir"
               Height          =   615
               Left            =   2400
               TabIndex        =   5
               Top             =   240
               Width           =   1575
            End
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   135
            Left            =   1320
            TabIndex        =   7
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
            TabIndex        =   8
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
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   4935
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Todos"
         CausesValidation=   0   'False
         Height          =   195
         Left            =   8040
         TabIndex        =   1
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "RUC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   6480
         TabIndex        =   10
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Empresa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   8880
      TabIndex        =   11
      Top             =   5520
      Width           =   1215
   End
End
Attribute VB_Name = "FrmLibrosElectPrincipales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim rs1 As New ADODB.Recordset
 Dim li_aRC As Integer
Private Sub LlenarLista()
 Dim I As Integer
 Dim itmX As ListItem
 Dim rs2 As New ADODB.Recordset
 ListView1.ColumnHeaders.Clear
   ListView1.ListItems.Clear
   ListView1.ColumnHeaders.Add , , "Presentacion de Libro", ListView1.Width / 1
   ListView1.View = lvwReport
   Set rs1 = Nothing
   Set rs1 = VGCNx.Execute("select * from ct_librossunatcorrelativos where categorialibro=1 ")
   rs1.MoveFirst
   I = 1
   Do While Not rs1.EOF
      Set itmX = ListView1.ListItems.Add(, , rs1!librocodigosunat + "  " + rs1!libroCorrelativodescripcion)
      If Check1.Value = 0 Then
         SQL = " select * from ct_librossunatxempresa where empresacodigo='" & VGParametros.empresacodigo & "'"
         SQL = SQL & " and librocodigosunat='" & rs1!librocodigosunat & "'"
         Set rs2 = VGCNx.Execute(SQL)
         If rs2.RecordCount = 0 Then
              ListView1.ListItems.Item(I + 0).Checked = 0
           Else
              ListView1.ListItems.Item(I + 0).Checked = rs2!estadoregistro
         End If
      Else
               ListView1.ListItems.Item(I + 0).Checked = 1
      End If
      I = I + 1
      rs1.MoveNext
   Loop
  End Sub

Private Sub Command1_Click()
Dim x As Integer
Dim libro As String
x = 1
rs1.MoveFirst
Do Until x > rs1.RecordCount
   libro = Left(ListView1.ListItems.Item(x).Text, 6)
   If libro = "xx" Then
   ElseIf libro = "050100" Then
      Call Generadiario("050100")
   ElseIf libro = "050200" Then
      Call GeneradiarioSimplificado("050200")
   ElseIf libro = "080100" Then
      Call GeneraCompras("080100")
   ElseIf libro = "140100" Then
      Call GeneraVentas("140100")
   End If
   x = x + 1
Loop
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Text1.Text = VGParametros.empresacodigo + " - " + VGParametros.NomEmpresa
Text2.Text = VGParametros.RucEmpresa
Call LlenarLista
End Sub
