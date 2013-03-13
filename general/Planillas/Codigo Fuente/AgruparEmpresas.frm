VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form AgruparEmpresas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selección de Empresas"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   Icon            =   "AgruparEmpresas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   4920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2565
      TabIndex        =   5
      Top             =   3780
      Width           =   1350
   End
   Begin VB.CommandButton cmAcepta 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   975
      TabIndex        =   4
      Top             =   3795
      Width           =   1350
   End
   Begin VB.CommandButton cmEditar 
      Caption         =   "Editar"
      Height          =   360
      Left            =   405
      TabIndex        =   3
      Top             =   4740
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.CommandButton cmNueva 
      Caption         =   "&Nueva"
      Height          =   360
      Left            =   -75
      TabIndex        =   2
      Top             =   4740
      Visible         =   0   'False
      Width           =   405
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   285
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AgruparEmpresas.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LEmpresas 
      Height          =   2955
      Left            =   45
      TabIndex        =   1
      Top             =   720
      Width           =   4830
      _ExtentX        =   8520
      _ExtentY        =   5212
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID Empresa"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Ruta"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selección de Empresas para traslados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   450
      TabIndex        =   0
      Top             =   420
      Width           =   3255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   705
      Left            =   60
      Top             =   15
      Width           =   4830
   End
End
Attribute VB_Name = "AgruparEmpresas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsEmp As New ADODB.Recordset
Dim XITEM As ListItem

Private Sub CMCANCELAR_CLICK()
    End
End Sub
Private Sub Form_Load()
    On Error Resume Next
    CargaEmp
End Sub
Public Sub CargaEmp()
On Error Resume Next
    Set RsEmp = Nothing
    RsEmp.Open "SELECT * FROM EMPRESAS ORDER BY NOMBRE", DBSTARPLAN, adOpenStatic, adLockOptimistic
    RsEmp.Requery
    LEmpresas.ListItems.Clear
    Do While Not RsEmp.EOF
        Set XITEM = LEmpresas.ListItems.Add(, "R" & RsEmp!RUC, RsEmp!RUC, , 1)
        XITEM.SubItems(1) = RsEmp!NOMBRE
        XITEM.SubItems(2) = RsEmp!DIRMASTER
        RsEmp.MoveNext
    Loop
    RsEmp.Close
End Sub

