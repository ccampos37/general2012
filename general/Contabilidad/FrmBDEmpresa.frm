VERSION 5.00
Begin VB.Form FrmBDEmpresa 
   BackColor       =   &H00C00000&
   BorderStyle     =   0  'None
   Caption         =   "FrmBDEmpresa"
   ClientHeight    =   1650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   345
      Left            =   3120
      TabIndex        =   2
      Top             =   960
      Width           =   1185
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   345
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   1185
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1500
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   540
      Width           =   4305
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   90
      ScaleHeight     =   1125
      ScaleWidth      =   5895
      TabIndex        =   3
      Top             =   360
      Width           =   5925
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Empresa:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   210
         Width           =   1035
      End
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      Caption         =   "Seleccione empresa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   285
      Left            =   1170
      TabIndex        =   5
      Top             =   60
      Width           =   2055
   End
End
Attribute VB_Name = "FrmBDEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
    CargarParametros
    CargarParametrosContabilidad
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    LlenarListaempresas
End Sub
Private Sub LlenarListaempresas()
Dim REG1 As New ADODB.Recordset
Set REG1 = New ADODB.Recordset
Dim multiempresa As Integer
Set REG1 = VGCNx.Execute("Select * from co_multiempresas where empresacodigo<>'00'  ")
Combo2.Clear
If REG1.EOF Then Exit Sub
If REG1.BOF Then Exit Sub
Do While Not REG1.EOF
   Combo2.AddItem REG1.Fields("empresacodigo") + " " + REG1.Fields("empresadescripcion")
   REG1.MoveNext
Loop

REG1.MoveFirst
Combo2.ListIndex = 0
End Sub
Private Sub CargarParametros()
Dim rssql As New ADODB.Recordset
    Set rssql = VGCNx.Execute("select * from co_multiempresas where empresacodigo='" & VGParametros.empresacodigo & "'")
    If rssql.RecordCount > 0 Then
       VGParametros.RucEmpresa = ESNULO(rssql!empresaruc, "")
       VGParametros.NomEmpresa = rssql!empresadescripcion
       MDIPrincipal.StatusBar1.Panels(6).Text = "Base de Datos : " & VGCNx.DefaultDatabase & ""
       MDIPrincipal.StatusBar1.Panels(7).Text = "Empresa (" & RTrim(VGParametros.NomEmpresa) & ")"
 
    End If
End Sub
Public Sub Combo2_Click()
      VGParametros.empresacodigo = Left(Combo2.Text, 2)
      VGParametros.NomEmpresa = Right(Combo2.Text, Len(Combo2.Text) - 2)
      MDIPrincipal.Caption = "Sistema de Contabilidad - " & VGParametros.NomEmpresa
End Sub
