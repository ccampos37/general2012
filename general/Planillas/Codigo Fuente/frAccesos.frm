VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form frAccesos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administrador de Accesos"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   Icon            =   "frAccesos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   6750
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Acceso a Procesos"
      Height          =   2460
      Left            =   120
      TabIndex        =   10
      Top             =   1365
      Visible         =   0   'False
      Width           =   5085
      Begin MSComctlLib.ListView Procs 
         Height          =   1905
         Left            =   180
         TabIndex        =   7
         Top             =   390
         Width           =   4740
         _ExtentX        =   8361
         _ExtentY        =   3360
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Proceso"
            Object.Width           =   20179
         EndProperty
      End
   End
   Begin VB.CommandButton cmEliminar 
      Caption         =   "&Eliminar"
      Height          =   390
      Left            =   5280
      TabIndex        =   4
      Top             =   1145
      Width           =   1410
   End
   Begin VB.CommandButton cmChangePassword 
      Caption         =   "&Modificar"
      Height          =   390
      Left            =   5280
      TabIndex        =   3
      Top             =   655
      Width           =   1410
   End
   Begin VB.CommandButton cmNuevo 
      Caption         =   "&Nuevo"
      Height          =   390
      Left            =   5295
      TabIndex        =   2
      Top             =   165
      Width           =   1410
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5370
      Top             =   3135
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frAccesos.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frAccesos.frx":0896
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   390
      Left            =   5280
      TabIndex        =   0
      Top             =   1635
      Width           =   1410
   End
   Begin VB.Frame Frame1 
      Caption         =   "Identificación"
      Height          =   1200
      Left            =   105
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   5085
      Begin VB.TextBox xPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   720
         Width           =   1770
      End
      Begin AplisetControlText.Aplitext xUser 
         Height          =   285
         Left            =   1680
         TabIndex        =   11
         Top             =   360
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   503
         MaxLength       =   10
         Text            =   ""
         TipoCodigo      =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Contraseña"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   765
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre de Usuario"
         Height          =   195
         Left            =   135
         TabIndex        =   6
         Top             =   405
         Width           =   1365
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   2655
      Left            =   90
      TabIndex        =   1
      Top             =   150
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   4683
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Usuario"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tipo"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frAccesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RS2 As New ADODB.Recordset
Dim XITEM As ListItem

Private Sub cmCerrar_Click()
    Unload Me
End Sub

Private Sub CMCHANGEPASSWORD_Click()
    If UCase(cmChangePassword.Caption) = "&CANCELAR" Then
        Lista.Visible = True
        Frame1.Visible = False
        Frame2.Visible = False
        cmNuevo.Caption = "&NUEVO"
        cmChangePassword.Caption = "&MODIFICAR"
        cmCerrar.Visible = True
        cmEliminar.Visible = True
    Else 'QUIERE DECIR QUE ES MODIFICAR
        Lista.Visible = False
        Frame1.Visible = True
        cmNuevo.Caption = "&GRABAR"
        cmChangePassword.Caption = "&CANCELAR"
        cmCerrar.Visible = False
        cmEliminar.Visible = False
        xUser.Locked = True
        cmNuevo.Tag = "A"
        xUser.Text = Lista.SelectedItem.KEY
    End If
End Sub

Private Sub CMELIMINAR_CLICK()
    If Lista.ListItems.Count = 0 Then
        MsgBox "NO EXISTEN DATOS POR ELIMINAR", vbCritical
        Exit Sub
    End If
    If Lista.SelectedItem.KEY = REGSISTEMA.USER Then
        MsgBox "NO SE PUEDE ELIMINAR EL MISMO USUARIO", vbCritical
        Exit Sub
    End If
    If MsgBox("SEGURO DE ELIMINAR EL USUARIO/ADMINISTRADOR SELECCIONADO", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    DBSTARPLAN.Execute "DELETE FROM USUARIOS WHERE USUARIO='" & Lista.SelectedItem.KEY & "'"
    CARGADATOS
End Sub

Private Sub CMNUEVO_CLICK()
    Dim PWD As String
    If UCase(cmNuevo.Caption) = "&NUEVO" Then
        Lista.Visible = False
        Frame1.Visible = True
        Frame2.Visible = True
        cmNuevo.Caption = "&GRABAR"
        cmChangePassword.Caption = "&CANCELAR"
        cmCerrar.Visible = False
        cmEliminar.Visible = False
        cmNuevo.Tag = "B"
        xUser.Locked = False
    Else
        If xUser.Text = "" Then
            MsgBox "FALTA NOMBRE DEL USUARIO", vbCritical
            Exit Sub
        End If
        If Len(xPass.Text) <= 4 Then
            MsgBox "LA CONTRASEÑA DEBE TENER MÁS DE 4 CARACTERES", vbCritical
            Exit Sub
        End If
        Lista.Visible = True
        Frame1.Visible = False
        Frame2.Visible = False
        PWD = PROCSIS.PrCodifica(Trim(xPass.Text), 5)
        If cmNuevo.Tag = "A" Then
            RS2.MoveFirst
            RS2.FIND "USUARIO='" & Lista.SelectedItem.KEY & "'"
            DBSTARPLAN.Execute "UPDATE USUARIOS SET CLAVE='" & PWD & "' WHERE USUARIO='" & xUser.Text & "'"
        Else
            DBSTARPLAN.Execute "INSERT INTO USUARIOS VALUES ('" & xUser.Text & "','" & PWD & "','1111','A')"
        End If
        cmNuevo.Caption = "&NUEVO"
        cmChangePassword.Caption = "&MODIFICAR"
        cmCerrar.Visible = True
        cmEliminar.Visible = True
    End If
    CARGADATOS
    Lista.Refresh
End Sub

Private Sub FORM_LOAD()
    CARGADATOS
End Sub

Public Sub CARGADATOS()
    If RS2.State <> 0 Then RS2.Close
    RS2.Open "USUARIOS", DBSTARPLAN
    Lista.ListItems.Clear
    Do While Not RS2.EOF
        Set XITEM = Lista.ListItems.Add(, RS2!USUARIO, RS2!USUARIO, , IIf(RS2!TIPO = "A", 1, 2))
        XITEM.SubItems(1) = IIf(RS2!TIPO = "A", "ADMINISTRADOR", "USUARIO")
        RS2.MoveNext
    Loop
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RS2 = Nothing
    Set RS2 = Nothing
End Sub
Private Sub OPTION1_CLICK()
    Frame2.Visible = False
End Sub

Private Sub OPTION2_Click()
    Frame2.Visible = True
End Sub

