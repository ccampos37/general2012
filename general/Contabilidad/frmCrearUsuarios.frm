VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frmCrearUsuarios 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4590
   ClientLeft      =   3540
   ClientTop       =   2325
   ClientWidth     =   8805
   Icon            =   "frmCrearUsuarios.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   8805
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Nuevo"
      Height          =   405
      Index           =   0
      Left            =   7305
      Picture         =   "frmCrearUsuarios.frx":058A
      TabIndex        =   16
      Top             =   1740
      Width           =   1350
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Grabar"
      Height          =   405
      Index           =   1
      Left            =   7305
      Picture         =   "frmCrearUsuarios.frx":09CC
      TabIndex        =   15
      Top             =   2295
      Width           =   1350
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "E&ditar"
      Height          =   405
      Index           =   2
      Left            =   7305
      Picture         =   "frmCrearUsuarios.frx":0E0E
      TabIndex        =   14
      Top             =   2835
      Width           =   1350
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Eliminar"
      Height          =   405
      Index           =   3
      Left            =   7305
      Picture         =   "frmCrearUsuarios.frx":1250
      TabIndex        =   13
      Top             =   3390
      Width           =   1350
   End
   Begin VB.CommandButton cmdBotones 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   405
      Index           =   5
      Left            =   7305
      Picture         =   "frmCrearUsuarios.frx":1692
      TabIndex        =   12
      Top             =   3945
      Width           =   1350
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   2925
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCrearUsuarios.frx":1AD4
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCrearUsuarios.frx":1BD0
            Key             =   "Abrir"
         EndProperty
      EndProperty
   End
   Begin VB.Frame FrameU 
      Caption         =   "Nuevo Usuario: "
      Height          =   4230
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   6855
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   5010
         MaxLength       =   20
         TabIndex        =   3
         Top             =   405
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   5025
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   765
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   2010
         MaxLength       =   10
         TabIndex        =   2
         Top             =   375
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   2010
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   780
         Width           =   1455
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   2640
         Left            =   240
         TabIndex        =   6
         Top             =   1470
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   4657
         _Version        =   393217
         Indentation     =   353
         Style           =   7
         ImageList       =   "ImageList1"
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Permisos de Usuario"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   255
         TabIndex        =   19
         Top             =   1230
         Width           =   6360
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre Usuario"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3570
         TabIndex        =   11
         Top             =   405
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Código Usuario"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   450
         TabIndex        =   10
         Top             =   405
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Password Usuario"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   450
         TabIndex        =   9
         Top             =   810
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Confirmar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   3570
         TabIndex        =   8
         Top             =   765
         Width           =   1335
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   855
         Left            =   240
         Top             =   300
         Width           =   6375
      End
   End
   Begin VB.Frame FrameU 
      Caption         =   "Lista de Usuarios: "
      Height          =   3975
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   6855
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3375
         Left            =   480
         TabIndex        =   0
         Top             =   360
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   5953
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
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
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            ScrollBars      =   2
            BeginProperty Column00 
               ColumnWidth     =   2055.118
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1769.953
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuarios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Index           =   0
      Left            =   7695
      TabIndex        =   17
      Top             =   960
      Width           =   945
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   7950
      Picture         =   "frmCrearUsuarios.frx":1CCC
      Top             =   285
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   8085
      Picture         =   "frmCrearUsuarios.frx":1FD6
      Top             =   585
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuarios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   7710
      TabIndex        =   18
      Top             =   975
      Width           =   945
   End
End
Attribute VB_Name = "frmCrearUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ADOREG1 As ADODB.Recordset
Dim ADOREG2 As ADODB.Recordset
Dim ADOUSU As ADODB.Recordset
Dim ADOMEN As ADODB.Recordset
Dim CCAD As String
Dim REGACTUAL As Integer
Dim NFRA As Integer
Dim NTIPO As Integer
Dim MNODE As NODE
Dim nI As Integer
Dim XFLAG As Boolean
Const NUMMAGICO As Integer = 10
Dim CLMENU As ClassMenu
Dim VPTAREA As String
Private Sub cmdBotones_Click(Index As Integer)
Dim NII As Integer
Dim TEMPI As Integer
Dim TEMPS As String
Select Case Index
 Case 0: 'NUEVO
         NTIPO = 1
         TreeView1.Refresh
         Call CARGA_VALOR(1, TreeView1.Nodes.Count, True, 1)
         TreeView1.Refresh
         If TreeView1.Nodes(1).Checked Then
            TreeView1.Nodes(1).Expanded = True
         Else
            TreeView1.Nodes(1).Expanded = False
         End If
         NFRA = 1
         Dim OTEXT As TextBox
         For Each OTEXT In Me.Text1
            OTEXT.Text = ""
         Next
         FrameU(1).Caption = "NUEVO USUARIO"
         FrameU(1).Visible = True
         FrameU(0).Visible = False
         
         BOTONES_SET False
         Text1(0).Enabled = True
         Text1(0).SetFocus
 Case 1: 'GRABAR
         Screen.MousePointer = 11
         If NTIPO = 1 Then
            Dim flag As Boolean
            flag = False
            'BUSCAR IGUAL CODIGO
            With ADOREG1
                If .RecordCount <> 0 Then
                    .MoveFirst
                    .Find "USU_CODIGO= '" & UCase(Text1(0).Text) & "'"
                    If Not .EOF Then
                        flag = True
                        Text1(0).Text = ""
                        MsgBox "EL USUARIO YA EXISTE:  INGRESE DE NUEVO", vbInformation, "INGRESO DE DATOS"
                        Text1(0).SetFocus
                    End If
                End If
            End With
            If Not flag Then
                If Text1(2).Text = "" Then
                    MsgBox "UD. NO HA INGRESADO SU PASSWORD", vbInformation, "INGRESO DE DATOS"
                    Text1(2).SetFocus
                ElseIf Text1(3).Text = "" Then
                    MsgBox "UD. NO HA INGRESADO SU CONFIRMACIÓN", vbInformation, "INGRESO DE DATOS"
                    Text1(3).SetFocus
                ElseIf Text1(2).Text = Text1(3).Text Then
                    'PASA
                    ADOREG2.AddNew
                    ADOREG2.Fields("USU_CODIGO") = UCase(Trim(Text1(0).Text))
                    ADOREG2.Fields("EMP_CODIGO") = VGParametros.RucEmpresa
                    ADOREG2.Fields("USU_PASSWORD") = CLMENU.CODIFICA(Trim(Text1(2).Text), NUMMAGICO) 'PASSWORD                    ADOREG2.UPDATEBATCH
                    If Trim(Text1(1).Text) <> "" Then ADOREG2.Fields("USU_NOMBRE") = UCase(Trim(Text1(1).Text))
                    ADOREG2.Update
                    ADOREG1.Requery
                    Call GRAB_MEN(UCase(Trim(Text1(0).Text)))
                    
                    FrameU(NFRA).Visible = False
                    FrameU(0).Visible = True
                    NFRA = 0
                    BOTONES_SET True
                Else
                    MsgBox "NOMBRE DE PASSWORD Y LA CONFIRMACIÓN NO COINCIDEN", vbInformation, "INGRESO DE DATOS"
                    Text1(2).Text = ""
                    Text1(3).Text = ""
                    Text1(2).SetFocus
                End If
            End If
        End If
        If NTIPO = 2 Then
            ADOREG2.Fields("USU_CODIGO") = UCase(Trim(Text1(0).Text))
            ADOREG2.Fields("EMP_CODIGO") = VGParametros.RucEmpresa
            ADOREG2.Fields("USU_PASSWORD") = CLMENU.CODIFICA(Trim(Text1(2).Text), NUMMAGICO)
            If Trim(Text1(1).Text) <> "" Then ADOREG2.Fields("USU_NOMBRE") = UCase(Trim(Text1(1).Text))
            ADOREG2.UpdateBatch
            ADOREG1.Requery
            ADOREG2.Requery
            
            Call GRAB_MEN(UCase(Trim(Text1(0).Text)))
            
            FrameU(1).Visible = False
            FrameU(0).Visible = True
            NFRA = 0
            BOTONES_SET True
         End If
         SETDATAGRID
         Screen.MousePointer = 1
         
 Case 2: 'EDITAR
         If ADOREG1.Bookmark Then
            Screen.MousePointer = 11
            NTIPO = 2
            ADOREG2.Bookmark = ADOREG1.Bookmark
            NFRA = 2
            Dim OTEXT1 As TextBox
            For Each OTEXT1 In Me.Text1
                OTEXT1.Text = ""
            Next
            Text1(0).Text = ADOREG2.Fields(0)
            Text1(2).Text = CLMENU.DECODIFICA(ADOREG2.Fields(2), NUMMAGICO)
            Text1(3).Text = CLMENU.DECODIFICA(ADOREG2.Fields(2), NUMMAGICO)
            If Not IsNull(ADOREG2.Fields("USU_NOMBRE")) Then Text1(1).Text = ADOREG2.Fields("USU_NOMBRE")
        
            Call EDIT_MEN(ADOREG2.Fields(0), VGParametros.RucEmpresa)
'            ME.CLS
'            CALL EDIT_MEN(ADOREG2.FIELDS(0), VGEMP_CODIGO)
            
            FrameU(1).Caption = "MODIFICAR USUARIO"
            FrameU(1).Visible = True
            FrameU(0).Visible = False
            BOTONES_SET False
            Text1(0).Enabled = False
            Text1(1).SetFocus
            'IF XFLAG THEN
            ' XFLAG = FALSE
            ' CMDBOTONES_Click 5
            ' CMDBOTONES_Click 2
            'ELSE
            ' XFLAG = TRUE
            'END IF
            Screen.MousePointer = 1
         Else
            MsgBox "DEBE SELECCIONAR UN REGISTRO PARA EDITARLO", vbInformation
            BOTONES_SET False
            cmdBotones_Click 5
         End If
       
 Case 3: 'ELIMINAR
          Dim OP As Integer
          OP = MsgBox("ESTA SEGURO QUE DESEA ELIMINAR EL REGISTRO ACTUAL ?", vbYesNo, "ELIMINACIÓN DE REGISTRO")
          If OP = vbYes Then
            ADOREG2.Bookmark = ADOREG1.Bookmark
            VGCNx.Execute "DELETE FROM " & CLMENU.TabaMenuDet & " WHERE USU_CODIGO = '" & ADOREG1("USU_CODIGO") & "' AND EMP_CODIGO = '" & VGParametros.RucEmpresa & "'"
            ADOUSU.Requery
            ADOREG2.Delete
            ADOREG2.UpdateBatch
            ADOREG2.Requery
            ADOREG1.Requery
            If ADOREG1.RecordCount = 0 Then
                BOTONES_INIT True
            Else
                BOTONES_SET True
            End If
          End If
          SETDATAGRID
          
 Case 5: 'SALIR , CANCELAR
         If cmdBotones(5).Caption = "&SALIR" Then
            Unload Me
         Else
            cmdBotones(5).Caption = "&SALIR"
            FrameU(1).Visible = False
            FrameU(0).Visible = True
            NFRA = 0
            If ADOREG1.RecordCount = 0 Then
                BOTONES_INIT True
            Else
                BOTONES_SET True
            End If
         End If
End Select
End Sub

Public Sub BOTONES_SET(flag As Boolean)
'FLAG=FALSE NUEVO; FLAG=TRUE .ETC...
 cmdBotones(0).Enabled = flag 'NUEVO
 cmdBotones(1).Enabled = Not flag 'GRABAR
 cmdBotones(2).Enabled = flag 'EDITAR
 cmdBotones(3).Enabled = flag 'ELIMINAR
 If flag Then
  cmdBotones(5).Caption = "&SALIR" 'SALIR
 Else
  cmdBotones(5).Caption = "&CANCELAR"
 End If
End Sub
Public Sub BOTONES_INIT(flag As Boolean)
'FLAG=FALSE NUEVO; FLAG=TRUE .ETC...
 cmdBotones(0).Enabled = flag 'NUEVO
 cmdBotones(1).Enabled = Not flag 'GRABAR
 cmdBotones(2).Enabled = Not flag 'EDITAR
 cmdBotones(3).Enabled = Not flag 'ELIMINAR
 cmdBotones(5).Caption = "&SALIR" 'SALIR
End Sub

Private Sub Command1_Click()
End Sub

Private Sub DATAGRID1_Click()
 REGACTUAL = IIf(IsNull(DataGrid1.Bookmark), 0, DataGrid1.Bookmark)
End Sub

Private Sub Form_Load()
    Set CLMENU = New ClassMenu
    Dim FRA As Frame
    Me.Width = 8895: Me.Height = 4965
    XFLAG = True
    Me.Caption = "USUARIOS - " & VGParametros.NomEmpresa

    Call ADOConectar
    If ADOREG1.RecordCount = 0 Then
        BOTONES_INIT True
    Else
        BOTONES_SET True
    End If
    SETDATAGRID
    
    For Each FRA In Me.FrameU: FRA.Visible = False: Next
    FrameU(0).Visible = True
    NFRA = 0
    NTIPO = 1
     
    ' CONFIGURA EL CONTROL TREEVIEW
    TreeView1.Sorted = False
    TreeView1.Checkboxes = True
    Set MNODE = TreeView1.Nodes.Add()
    MNODE.Text = "MENU"
    MNODE.Tag = VGCNx
    MNODE.Image = "Abrir"
    MNODE.Checked = True
    TreeView1.LabelEdit = False
    CARGAR_OPC
    
End Sub
Private Sub ADOConectar()
Set ADOREG1 = New ADODB.Recordset
Set ADOREG2 = New ADODB.Recordset
Set ADOUSU = New ADODB.Recordset
CLMENU.TablaUsu = "USUARIO"
CLMENU.TabaMenuDet = "ct_USUARIODET"
ADOREG1.CursorType = adOpenDynamic
ADOREG1.Open "SELECT USU_CODIGO,USU_NOMBRE FROM " & CLMENU.TablaUsu & " WHERE EMP_CODIGO=" & "'" & VGParametros.RucEmpresa & "'", VGCNx, adOpenStatic
ADOREG2.Open "SELECT * FROM " & CLMENU.TablaUsu & " WHERE EMP_CODIGO=" & "'" & VGParametros.RucEmpresa & "'", VGCNx, adOpenDynamic, adLockOptimistic
ADOUSU.Open "SELECT * FROM " & CLMENU.TabaMenuDet, VGCNx, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = ADOREG1
End Sub

Public Sub SETDATAGRID()
 DataGrid1.Refresh
 DataGrid1.Columns(0).Caption = "           CÓDIGO"
 DataGrid1.Columns(1).Caption = "                            NOMBRE"
 DataGrid1.Columns(0).Width = 1500
 DataGrid1.Columns(1).Width = 3700
End Sub

Private Sub Text1_GotFocus(Index As Integer)
With Text1(Index)
 .SelStart = 0
 .SelLength = Len(.Text)
End With
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If Index = 0 Then
        With ADOREG1
            If .RecordCount <> 0 Then
                .MoveFirst
                .Find "USU_CODIGO= '" & UCase(Text1(0).Text) & "'"
                If Not .EOF Then
                    Text1(0).Text = ""
                    MsgBox "EL USUARIO YA EXISTE:  INGRESE DE NUEVO", vbInformation, "INGRESO DE DATOS"
                    Text1(0).SetFocus: Exit Sub
                End If
            End If
        End With
    End If
            
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub
Private Sub CARGAR_OPC()
Dim INTINDEX01 ' VARIABLE PARA EL ÍNDICE DEL NODO ACTUAL.
Dim INTINDEX02 ' VARIABLE PARA EL ÍNDICE DEL NODO ACTUAL.
Dim INTINDEX03 ' VARIABLE PARA EL ÍNDICE DEL NODO ACTUAL.
CLMENU.TablaUsu = "USUARIO"
CLMENU.TabaMenuDet = "CT_UsuarioDet"
CLMENU.TablaMenu = "CT_MENU"
Set ADOMEN = New ADODB.Recordset
ADOMEN.Open "SELECT * FROM " & CLMENU.TablaMenu & " ORDER BY MEN_CODIGO", VGCNx, adOpenStatic

Do While Not ADOMEN.EOF
    If Len(ADOMEN("MEN_CODIGO")) = 2 Then
        ' AGREGA UN NODO AL TREEVIEW Y ESTABLECE SUS PROPIEDADES.
        Set MNODE = TreeView1.Nodes.Add(1, tvwChild, ADOMEN("MEN_CODIGO") & " ID", ADOMEN("MEN_DESCRI"), "Cerrar")
        MNODE.Tag = "MENU" ' IDENTIFICA LA TABLA.
        ' ESTABLECE EN LA VARIABLE INTINDEX LA PROPIEDAD Index DEL
        ' OBJETO NODE RECIÉN CREADO. USE ESTA VARIABLE PARA AGREGAR
        ' OBJETOS NODE HIJOS AL OBJETO NODE ACTUAL.
            
        INTINDEX01 = MNODE.Index
            
        ADOMEN.MoveNext
        If ADOMEN.EOF Then Exit Do
    End If
    If Len(ADOMEN("MEN_CODIGO")) = 4 Then
        Set MNODE = TreeView1.Nodes.Add(INTINDEX01, tvwChild)
        MNODE.Text = ADOMEN("MEN_DESCRI") ' TEXTO.
        MNODE.Key = ADOMEN("MEN_CODIGO") & " ID"  ' ID ÚNICO.
        MNODE.Image = "Cerrar"     ' IMAGEN DE IMAGELIST.
            
        INTINDEX02 = MNODE.Index
            
        ADOMEN.MoveNext
        If ADOMEN.EOF Then Exit Do
    End If
        
    If Len(ADOMEN("MEN_CODIGO")) = 6 Then
        Set MNODE = TreeView1.Nodes.Add(INTINDEX02, tvwChild)
        MNODE.Text = ADOMEN("MEN_DESCRI") ' TEXTO.
        MNODE.Key = ADOMEN("MEN_CODIGO") & " ID"   ' ID ÚNICO.
        MNODE.Image = "Cerrar"     ' IMAGEN DE IMAGELIST.
            
        INTINDEX03 = MNODE.Index
        
        ADOMEN.MoveNext
        If ADOMEN.EOF Then Exit Do
    End If
    
    If Len(ADOMEN("MEN_CODIGO")) = 8 Then
        Set MNODE = TreeView1.Nodes.Add(INTINDEX03, tvwChild)
        MNODE.Text = ADOMEN("MEN_DESCRI") ' TEXTO.
        MNODE.Key = ADOMEN("MEN_CODIGO") & " ID"   ' ID ÚNICO.
'        MNODE.Image = "CERRAR"     ' IMAGEN DE IMAGELIST.
            
        ADOMEN.MoveNext
        If ADOMEN.EOF Then Exit Do
    End If
Loop
End Sub
Private Sub TREEVIEW1_COLLAPSE(ByVal NODE As NODE)
    If NODE.Text = "MENU" Or NODE.Index > 1 Then
        NODE.Image = "Cerrar"
    End If
End Sub
Private Sub TREEVIEW1_EXPAND(ByVal NODE As NODE)
    If NODE.Text = "MENU" Or NODE.Index > 1 Then
        If TreeView1.Nodes(NODE.Index).Children > 0 Then
            If NODE.Checked = False Then
                NODE.Image = "Cerrar"
                NODE.Expanded = False
            Else
                NODE.Image = "Abrir"
            End If
        End If
        NODE.Sorted = False
    End If
End Sub

Private Sub CARGA_VALOR(NINI As Integer, NFIN As Integer, BFLAG As Boolean, NG As Integer, Optional cCod As String)
If NG = 1 Then
    For nI = NINI To NFIN
        TreeView1.Nodes.Item(nI).Checked = BFLAG
    Next nI
ElseIf NG = 2 Then
    For nI = NINI To NFIN
        If Mid(TreeView1.Nodes(nI).Key, 1, Len(Trim(cCod))) = Trim(cCod) And TreeView1.Nodes(nI).Key <> "060102" And TreeView1.Nodes(nI).Key <> "060404" Then
            TreeView1.Nodes.Item(nI).Checked = BFLAG
        End If
    Next nI
End If
End Sub

Private Sub TREEVIEW1_NODECHECK(ByVal NODE As MSComctlLib.NODE)
If NODE.Index = 1 Then
    If NODE.Root.Checked = True Then
        Call CARGA_VALOR(1, TreeView1.Nodes.Count, NODE.Root.Checked, 1)
        NODE.Expanded = True
    Else
        NODE.Expanded = False
    End If
Else
    If TreeView1.Nodes(NODE.Index).Children > 0 Then
        If NODE.Checked = False Then
            Call CARGA_VALOR(NODE.Index + 1, TreeView1.Nodes.Count, False, 2, Mid(TreeView1.Nodes(NODE.Index).Key, 1, InStr(TreeView1.Nodes(NODE.Index).Key, " ID")))
            NODE.Expanded = False
        Else
            Call CARGA_VALOR(NODE.Index + 1, TreeView1.Nodes.Count, True, 2, Mid(TreeView1.Nodes(NODE.Index).Key, 1, InStr(TreeView1.Nodes(NODE.Index).Key, " ID")))
            NODE.Expanded = True
        End If
    ElseIf TreeView1.Nodes(NODE.Index).Key <> "060102" And TreeView1.Nodes(NODE.Index).Key <> "060404" Then
             TreeView1.Nodes.Item(NODE.Index).Checked = False
    End If
End If
End Sub


Private Sub GRAB_MEN(cCod As String)
Dim CCAD As String
Dim NII As Integer
Dim NOP As Integer
NII = 2
ADOUSU.Requery
ADOMEN.MoveFirst
Do While Not ADOMEN.EOF
    If TreeView1.Nodes(1).Checked Then  'RAIZ
        If TreeView1.Nodes.Item(NII).Key = ADOMEN("MEN_CODIGO") & " ID" Then
            If ADOUSU.RecordCount <> 0 Then ADOUSU.MoveFirst
            NOP = 2
            ADOUSU.Filter = "USU_CODIGO = '" & cCod & "' AND EMP_CODIGO = '" & VGParametros.RucEmpresa & "' AND MEN_CODIGO = '" & ADOMEN("MEN_CODIGO") & "'"
            If ADOUSU.RecordCount <> 0 Then
                    NOP = 1
            Else
             ADOUSU.Filter = ""
            End If
            If NOP = 2 Then ADOUSU.AddNew
            ADOUSU("USU_CODIGO") = cCod
            ADOUSU("EMP_CODIGO") = VGParametros.RucEmpresa
            ADOUSU("MEN_CODIGO") = ADOMEN("MEN_CODIGO")
            If TreeView1.Nodes.Item(NII).Checked = True Then
                ADOUSU("MEN_HAB") = True
            Else
                ADOUSU("MEN_HAB") = False
            End If
            ADOUSU.Update
            ADOUSU.Requery
        End If
    Else
        NII = 0
        Exit Do
    End If
    ADOMEN.MoveNext
    NII = NII + 1
    If ADOMEN.EOF Then Exit Do
Loop
ADOUSU.Requery
If NII >= 2 Then
    MsgBox "SE HA GRABADO COMPLETAMENTE LAS OPCIONES ESCOGIDAS", vbInformation, "MENSAJE"
Else
    MsgBox "NO SE HA GRABADO LAS OPCIONES", vbInformation, "VERIFICAR"
End If
End Sub
Private Sub EDIT_MEN(CCODU As String, CCODE As String)
Dim ADOUSME As ADODB.Recordset
Dim NJ As Integer
Set ADOUSME = New ADODB.Recordset
For NJ = 1 To TreeView1.Nodes.Count
    TreeView1.Nodes(NJ).Checked = False
    TreeView1.Nodes(NJ).Expanded = False
Next NJ
FrameU(0).Visible = False
FrameU(1).Visible = True
ADOUSME.Open "SELECT MEN_CODIGO,MEN_HAB FROM " & UCase(CLMENU.TabaMenuDet) & " WHERE USU_CODIGO = '" & CCODU & "' AND EMP_CODIGO = '" & CCODE & "'", VGCNx, adOpenStatic

If Not ADOUSME.EOF Then
    TreeView1.Nodes(1).Checked = True 'RAIZ
    Do While Not ADOUSME.EOF
        For NJ = 2 To TreeView1.Nodes.Count
            If TreeView1.Nodes(NJ).Key = ADOUSME("MEN_CODIGO") & " ID" Then
                If ADOUSME("MEN_HAB") Then
                    TreeView1.Nodes(NJ).Checked = True
                Else
                    TreeView1.Nodes(NJ).Checked = False
                End If
                Exit For
            End If
        Next NJ
        ADOUSME.MoveNext
        If ADOUSME.EOF Then Exit Do
    Loop
Else
    TreeView1.Nodes(1).Checked = False
End If

If TreeView1.Nodes(1).Checked Then
    TreeView1.Nodes(1).Expanded = True
Else
    TreeView1.Nodes(1).Expanded = False
End If
TreeView1.Refresh
End Sub
Public Sub COMANDOTOOLBAR(COMANDO As String)
    Select Case UCase(COMANDO)
        Case "NUEVO"
            VPTAREA = "NUEVO"
            cmdBotones_Click 0
        Case "EDITAR"
            VPTAREA = "EDITAR"
            cmdBotones_Click 2
        Case "ELIMINAR"
            cmdBotones_Click 3
    End Select
End Sub

