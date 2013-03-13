VERSION 5.00
Begin VB.Form FrmLibroInventarios 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   8325
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   3975
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   7695
      Begin VB.ComboBox Combo4 
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "FrmLibroInventarios.frx":0000
         Left            =   1830
         List            =   "FrmLibroInventarios.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2040
         Width           =   2700
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   1455
         Left            =   600
         TabIndex        =   4
         Top             =   360
         Width           =   5535
         Begin VB.CheckBox Check2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Detalle de los movimeintos VALORIZADOS"
            Height          =   495
            Left            =   480
            TabIndex        =   6
            Top             =   840
            Width           =   3855
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Detalle de los movimientos FISICOS"
            Height          =   495
            Left            =   480
            TabIndex        =   5
            Top             =   360
            Width           =   3615
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   1320
         TabIndex        =   1
         Top             =   2640
         Width           =   3975
         Begin VB.CommandButton Command2 
            BackColor       =   &H0000C000&
            Caption         =   "Salir"
            Height          =   495
            Left            =   2160
            TabIndex        =   3
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Imprimir"
            Height          =   495
            Left            =   240
            TabIndex        =   2
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pto.  Venta :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   720
         TabIndex        =   8
         Top             =   2085
         Width           =   885
      End
   End
End
Attribute VB_Name = "FrmLibroInventarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim aparam(5) As Variant
Dim aform(1) As Variant

aform(0) = "empresa='" & VGParametros.NomEmpresa & "'"
aparam(0) = VGCNx.DefaultDatabase
aparam(1) = VGParametros.empresacodigo
aparam(2) = VGParamSistem.Anoproceso
aparam(3) = VGParamSistem.Mesproceso
aparam(4) = Left(Combo4.Text, 2)
If Check1.Value = 1 Then
    Call ImpresionRptProc("ct_LibroInventariosFisicos.rpt", aform, aparam, , Check1.Caption)
End If
If Check2.Value = 1 Then
    Call ImpresionRptProc("ct_LibroInventariosValorizado.rpt", aform, aparam, , Check2.Caption)
End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim rs As ADODB.Recordset

Set rs = VGCNx.Execute("select puntovtacodigo,puntovtadescripcion from vt_puntoventa ")
Combo4.Clear
Do While Not rs.EOF
    Combo4.AddItem rs(0) & " " & rs(1)
    rs.MoveNext
Loop

Combo4.ListIndex = 0

End Sub
