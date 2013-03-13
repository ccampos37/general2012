VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frPanEmp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selección de Empresas"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4905
   Icon            =   "frPanEmp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   4905
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmEditar 
      Caption         =   "Editar"
      Height          =   360
      Left            =   3105
      TabIndex        =   4
      Top             =   3855
      Width           =   1080
   End
   Begin VB.CommandButton cmNueva 
      Caption         =   "&Nueva"
      Height          =   360
      Left            =   1912
      TabIndex        =   3
      Top             =   3855
      Width           =   1080
   End
   Begin VB.CommandButton cmSelecc 
      Caption         =   "Seleccionar"
      Default         =   -1  'True
      Height          =   360
      Left            =   720
      TabIndex        =   2
      Top             =   3855
      Width           =   1080
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
            Picture         =   "frPanEmp.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LEmpresas 
      Height          =   2955
      Left            =   60
      TabIndex        =   1
      Top             =   810
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5212
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
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
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "BD SQL"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selección de Empresas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1095
      TabIndex        =   0
      Top             =   420
      Width           =   2490
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   765
      Left            =   60
      Top             =   15
      Width           =   4830
   End
End
Attribute VB_Name = "frPanEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsEmp As New ADODB.Recordset
Dim CnAux As New ADODB.Connection
Dim VCARGA As Boolean
Dim XITEM As ListItem

Private Sub CMEDITAR_CLICK()
    If UCase(Dir(App.PATH & "\PLANILLA.SQL")) <> "PLANILLA.SQL" Then
        MsgBox "UD. NO PRESENTA DERECHOS DE EDICIÓN DE EMPRESAS", vbInformation
        Exit Sub
    End If
    VPTAREA = ""
    If LEmpresas.ListItems.Count = 0 Then Exit Sub
    Load frAddEmp
    frAddEmp.xRuc.Text = LEmpresas.SelectedItem.Text
    frAddEmp.xNombre.Text = LEmpresas.SelectedItem.SubItems(1)
    frAddEmp.xDir.Caption = LEmpresas.SelectedItem.SubItems(2)
    frAddEmp.xDir.Enabled = False
    VPTRASPRM = LEmpresas.SelectedItem.Text
    frAddEmp.Show 1
    If VPTAREA = "OK" Then CargaEmp
End Sub

Private Sub CMNUEVA_Click()
    If UCase(Dir(App.PATH & "\PLANILLA.SQL")) <> "PLANILLA.SQL" Then
        MsgBox "UD. NO PRESENTA DERECHOS DE CREACIÓN DE EMPRESAS", vbInformation
        Exit Sub
    End If
    VPTAREA = ""
    VPTRASPRM = "NUEVA"
    frAddEmp.Show 1
    If VPTAREA = "OK" Then CargaEmp
End Sub

Private Sub CMSELECC_CLICK()
    If LEmpresas.ListItems.Count = 0 Then Exit Sub
    
    With REGSISTEMA
        .EMPRESA = LEmpresas.SelectedItem.SubItems(1)
        .RUC = Right(LEmpresas.SelectedItem.KEY, 11)
        .PATHEMPRESA = LEmpresas.SelectedItem.SubItems(2)
        .PATHFOTOS = .PATHEMPRESA & "\FOTOS"
        .BASESQL = UCase(LEmpresas.SelectedItem.SubItems(3))
    
    End With
    If UCase(GetSetting(App.CompanyName, "PLANILLAS", "Nando", "NO")) = "HOLA" Then
        Unload Me
        Exit Sub
    End If
    frPWD.Show 1
    
    Set CLMENU.Conexion = DBSTARPLAN
    Set CLMENU.MDIMenu = MDIPrincipal
    CLMENU.TablaMenu = "MENU"
    CLMENU.TabaMenuDet = "USUARIODET"
    CLMENU.TablaUsu = "USUARIO"
    CLMENU.CrearTablasSeguridad
End Sub

Private Sub Form_Load()
    VCARGA = False
    CargaEmp
    VCARGA = True
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
        XITEM.SubItems(3) = RsEmp!DIRALMACEN
        RsEmp.MoveNext
    Loop
    RsEmp.Close
End Sub
Private Sub FORM_QUERYUNLOAD(CANCEL As Integer, UNLOADMODE As Integer)
    If UNLOADMODE = 0 Then
        If MsgBox("DESEA REALMENTE SALIR DEL SISTEMA", vbQuestion + vbYesNo) = vbYes Then
            End
        Else
            CANCEL = 1
        End If
    End If
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RsEmp = Nothing
End Sub

Private Sub LEMPRESAS_DBLCLICK()
    CMSELECC_CLICK
End Sub

