VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCfgUsuario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuarios"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   Icon            =   "frmCfgUsuario.frx":0000
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Nuevo"
      Height          =   915
      Index           =   0
      Left            =   375
      Picture         =   "frmCfgUsuario.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4785
      Width           =   1260
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Grabar"
      Height          =   915
      Index           =   1
      Left            =   1695
      Picture         =   "frmCfgUsuario.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4785
      Width           =   1260
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "E&ditar"
      Height          =   915
      Index           =   2
      Left            =   3015
      Picture         =   "frmCfgUsuario.frx":114E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4785
      Width           =   1260
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Eliminar"
      Height          =   915
      Index           =   3
      Left            =   4335
      Picture         =   "frmCfgUsuario.frx":1590
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4785
      Width           =   1260
   End
   Begin VB.CommandButton cmdBotones 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   915
      Index           =   5
      Left            =   5640
      Picture         =   "frmCfgUsuario.frx":19D2
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4785
      Width           =   1260
   End
   Begin VB.Frame Frame0 
      BorderStyle     =   0  'None
      Height          =   4410
      Left            =   225
      TabIndex        =   15
      Top             =   210
      Width           =   6810
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4095
         Left            =   75
         TabIndex        =   0
         Top             =   150
         Width           =   6570
         _ExtentX        =   11589
         _ExtentY        =   7223
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4500
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   6735
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1920
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   1920
         MaxLength       =   8
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   5040
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   11
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   8
         Top             =   840
         Width           =   3015
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   -45
         Top             =   4620
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
               Picture         =   "frmCfgUsuario.frx":1E14
               Key             =   "Cerrar"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCfgUsuario.frx":1F10
               Key             =   "Abrir"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   2700
         Left            =   105
         TabIndex        =   16
         Top             =   1665
         Width           =   6450
         _ExtentX        =   11377
         _ExtentY        =   4763
         _Version        =   393217
         Indentation     =   882
         Style           =   7
         ImageList       =   "ImageList1"
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Confirmar            :"
         Height          =   255
         Index           =   2
         Left            =   3480
         TabIndex        =   14
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Password Usuario   :"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   13
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Código Usuario       :"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   12
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre Usuario  :"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   840
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmCfgUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Adoreg1 As ADODB.Recordset
Dim AdoReg2 As ADODB.Recordset
Dim AdoUsu As ADODB.Recordset
Dim ADOMen As ADODB.Recordset
Dim cCad As String
Dim RegActual As Integer
Dim nFra As Integer
Dim nTipo As Integer
Dim nI As Integer
Dim mNode As Node
Dim rsql As String

Private Sub Form_Load()
central Me
ADOConectar
' Init_ControlDataGrid DataGrid1
If Adoreg1.RecordCount = 0 Then
    Botones_Init True
Else
    Botones_Set True
End If
SetDataGrid

Frame1.Visible = False
Frame0.Visible = True
nFra = 0
nTipo = 1

' Configura el control TreeView
TreeView1.Sorted = False
TreeView1.Checkboxes = True
Set mNode = TreeView1.Nodes.Add()
mNode.text = "Menu"
mNode.Tag = VGconfig
mNode.Image = "Abrir"
mNode.Checked = True
TreeView1.LabelEdit = False
Cargar_Opc
End Sub

Private Sub ADOConectar()
Set Adoreg1 = New ADODB.Recordset
Set AdoReg2 = New ADODB.Recordset
Set AdoUsu = New ADODB.Recordset

'Adoreg1.CursorType = adOpenDynamic           NO USAR CAMBIAR BOOKMARK

Adoreg1.Open "Select USU_CODIGO,USU_NOMBRE from USUARIO ", VGconfig, adOpenStatic
AdoReg2.Open "Select * from USUARIO ", VGconfig, adOpenDynamic, adLockOptimistic
AdoUsu.Open "Select * From si_menuusuarios where tipodesistema=" & vgtipo & "", VGconfig, adOpenDynamic, adLockOptimistic
 
Set DataGrid1.DataSource = Adoreg1
End Sub

Public Sub SetDataGrid()
 DataGrid1.Refresh
 DataGrid1.Columns(0).Caption = "           Código"
 DataGrid1.Columns(1).Caption = "                            Nombre"
 DataGrid1.Columns(0).Width = 1800
 DataGrid1.Columns(1).Width = 4500
 DataGrid1.ScrollBars = dbgVertical
End Sub

Private Sub cmdBotones_Click(Index As Integer)
Dim nIi As Integer
Dim tempi As Integer
Dim temps As String
Select Case Index
 Case 0: 'Nuevo
         nTipo = 1
         TreeView1.Appearance = ccFlat
         Call Carga_Valor(1, TreeView1.Nodes.count, True, 1)
         TreeView1.Refresh
         If TreeView1.Nodes(1).Checked Then
            TreeView1.Nodes(1).Expanded = True
         Else
            TreeView1.Nodes(1).Expanded = False
         End If
         
         nFra = 1
         Dim otext As TextBox
         For Each otext In Me.Text1
            otext.text = ""
         Next
         Frame1.Caption = "Nuevo Usuario"
         Frame1.Visible = True
         Frame0.Visible = False
         Botones_Set False
         Text1(0).Enabled = True
         Text1(0).SetFocus
 
 Case 1: 'Grabar
         If Text1(0) = "" Then
            MsgBox "Ingrese el codigo de usuario", vbExclamation, mensaje1
            Text1(0).SetFocus
            Exit Sub
         End If
         If Text1(1) = "" Then
            MsgBox "Ingrese el nombre de usuario", vbExclamation, mensaje1
            Text1(1).SetFocus
            Exit Sub
         End If
         If Text1(2) = "" Then
            MsgBox "Ingrese el password del usuario", vbExclamation, mensaje1
            Text1(2).SetFocus
            Exit Sub
         End If
         If Text1(3) = "" Then
            MsgBox "Ingrese la confirmación del usuario", vbExclamation, mensaje1
            Text1(3).SetFocus
            Exit Sub
         End If
          If Text1(2) <> Text1(3) Then
            MsgBox "El password no coincide con la confirnación", vbExclamation, mensaje1
           Text1(3).SetFocus
            Exit Sub
         End If
         
         Screen.MousePointer = 11
         'Ingreso
         If nTipo = 1 Then
                Dim flag As Boolean
                flag = False
                'buscar igual codigo
                With Adoreg1
                    If .RecordCount <> 0 Then
                        .MoveFirst
                        .Find "USU_CODIGO= '" & UCase(Text1(0).text) & "'"
                        If Not .EOF Then
                            flag = True
                            Text1(0).text = ""
                            MsgBox "El Usuario ya existe:  Ingrese de nuevo", vbInformation, "Ingreso de Datos"
                            Text1(0).SetFocus
                        End If
                    End If
                End With
                If Not flag Then
                    If Text1(2).text = "" Then
                        MsgBox "Ud. No ha ingresado su Password", vbInformation, "Ingreso de Datos"
                        Text1(2).SetFocus
                    ElseIf Text1(3).text = "" Then
                        MsgBox "Ud. No ha ingresado su confirmación", vbInformation, "Ingreso de Datos"
                        Text1(3).SetFocus
                    ElseIf Text1(2).text = Text1(3).text Then
                        'pasa
                        AdoReg2.AddNew
                         AdoReg2.Fields("Usu_Codigo") = UCase(Trim(Text1(0).text))
                        AdoReg2.Fields("USU_PASSWORD") = CODIFICA(Trim(Text1(2).text), NUMMAGICO) 'password
                        AdoReg2.UpdateBatch
                        If Trim(Text1(1).text) <> "" Then AdoReg2.Fields("Usu_Nombre") = UCase(Trim(Text1(1).text))
                        AdoReg2.Update
                        Adoreg1.Requery
                        Call Grab_Men(UCase(Trim(Text1(0).text)))
                    
                        Frame0.Visible = True
                        nFra = 0
                        Botones_Set True
                Else
                        MsgBox "Nombre de Password y la Confirmación no coinciden", vbInformation, "Ingreso de datos"
                        Text1(2).text = ""
                        Text1(3).text = ""
                        Text1(2).SetFocus
                End If
            End If
        End If
         
        If nTipo = 2 Then
             
            AdoReg2.Fields("Usu_Codigo") = UCase(Trim(Text1(0).text))
            AdoReg2.Fields("USU_PASSWORD") = CODIFICA(Trim(Text1(2).text), NUMMAGICO)
            If Trim(Text1(1).text) <> "" Then AdoReg2.Fields("Usu_Nombre") = UCase(Trim(Text1(1).text))
            AdoReg2.UpdateBatch
            Adoreg1.Requery
            AdoReg2.Requery
       
            Call Grab_Men(UCase(Trim(Text1(0).text)))
            
            Frame1.Visible = False
            Frame0.Visible = True
            nFra = 0
            Botones_Set True
         End If
          Frame1.Visible = False
          Frame0.Visible = True
         SetDataGrid
         Screen.MousePointer = 1
         
 Case 2: 'Editar
         If Adoreg1.Bookmark Then
            Screen.MousePointer = 11
            nTipo = 2
            AdoReg2.Bookmark = Adoreg1.Bookmark
            nFra = 2
            Dim OTEXT1 As TextBox
            For Each OTEXT1 In Me.Text1
                OTEXT1.text = ""
            Next
            Text1(0).text = AdoReg2.Fields(0)
            Text1(2).text = DECODIFICA(AdoReg2.Fields(1), NUMMAGICO)
            Text1(3).text = DECODIFICA(AdoReg2.Fields(1), NUMMAGICO)
            If Not IsNull(AdoReg2.Fields("Usu_Nombre")) Then Text1(1).text = AdoReg2.Fields("Usu_Nombre")
            
            Frame1.Caption = "Modificar Usuario"
            Frame0.Visible = False
            Frame1.Visible = True
            TreeView1.Visible = True
            TreeView1.Refresh
            Botones_Set False
            Text1(0).Enabled = False
            Text1(1).SetFocus
            Call Edit_Men(AdoReg2.Fields(0))
            
            Screen.MousePointer = 1
         Else
            MsgBox "Debe seleccionar un Registro para editarlo", vbInformation
            Botones_Set False
            cmdBotones_Click 5
         End If
       
 Case 3: 'Eliminar
          Dim op As Integer
          op = MsgBox("Esta Seguro que desea Eliminar el registro actual ", vbQuestion + vbYesNo, "Eliminación de Registro")
          If op = vbYes Then
                AdoReg2.Bookmark = Adoreg1.Bookmark
                VGconfig.Execute "Delete From si_menuusuarios Where tipodesistema=" & vgtipo & " and  USU_CODIGO = '" & AdoReg2("USU_CODIGO") & "'"
                AdoReg2.Delete
                AdoReg2.UpdateBatch
                AdoReg2.Requery
                Adoreg1.Requery
                If Adoreg1.RecordCount = 0 Then
                    Botones_Init True
                Else
                    Botones_Set True
                End If
          End If
          SetDataGrid
          
 Case 5: 'Salir , Cancelar
         If cmdBotones(5).Caption = "&Salir" Then
            Unload Me
         Else
            cmdBotones(5).Caption = "&Salir"
            Frame1.Visible = False
            Frame0.Visible = True
            nFra = 0
            If Adoreg1.RecordCount = 0 Then
                Botones_Init True
            Else
                Botones_Set True
            End If
         End If
End Select
End Sub

Private Sub Text1_GotFocus(Index As Integer)
With Text1(Index)
 .SelStart = 0
 .SelLength = Len(.text)
End With
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If Index = 0 Then
        With Adoreg1
            If .RecordCount <> 0 Then
                .MoveFirst
                .Find "USU_CODIGO= '" & UCase(Text1(0).text) & "'"
                If Not .EOF Then
                    Text1(0).text = ""
                    MsgBox "El Usuario ya existe:  Ingrese de nuevo", vbInformation, "Ingreso de Datos"
                    Text1(0).SetFocus: Exit Sub
                End If
            End If
        End With
    End If
     If Index = 2 And Text1(2) = "" Then
          MsgBox "No ha ingresado el password", vbInformation, "Ingreso de Datos"
          Text1(2).SetFocus: Exit Sub
    End If
    If Index = 3 And Text1(3) = "" Then
          MsgBox "No ha confirmado el password", vbInformation, "Ingreso de Datos"
          Text1(3).SetFocus: Exit Sub
    End If
    If Trim(Text1(3)) <> "" Then
       If Text1(2) <> Text1(3) Then
              MsgBox "La confirmación del password no es mismo", vbInformation, "Ingreso de Datos"
              Text1(3).SetFocus: Exit Sub
       End If
    End If
    SendKeys "{tab}"
    KeyAscii = 0
End If
End Sub

Public Sub Botones_Set(flag As Boolean)
'flag=false Nuevo; flag=true .etc...
 cmdBotones(0).Enabled = flag 'Nuevo
 cmdBotones(1).Enabled = Not flag 'Grabar
 cmdBotones(2).Enabled = flag 'Editar
 cmdBotones(3).Enabled = flag 'Eliminar
 If flag Then
  cmdBotones(5).Caption = "&Salir" 'Salir
 Else
  cmdBotones(5).Caption = "&Cancelar"
 End If
End Sub

Public Sub Botones_Init(flag As Boolean)
'flag=false Nuevo; flag=true .etc...
 cmdBotones(0).Enabled = flag 'Nuevo
 cmdBotones(1).Enabled = Not flag 'Grabar
 cmdBotones(2).Enabled = Not flag 'Editar
 cmdBotones(3).Enabled = Not flag 'Eliminar
 cmdBotones(5).Caption = "&Salir" 'Salir
End Sub

Private Sub DataGrid1_Click()
     RegActual = IIf(IsNull(DataGrid1.Bookmark), 0, DataGrid1.Bookmark)
End Sub

Private Sub Cargar_Opc()
Dim intIndex01 ' Variable para el índice del nodo actual.
Dim intIndex02 ' Variable para el índice del nodo actual.
Dim intIndex03 ' Variable para el índice del nodo actual.
Dim bolVisibleP As Boolean
Dim bolVisibleS As Boolean

bolVisibleP = True
bolVisibleS = True
On Error GoTo Err1

Set ADOMen = New ADODB.Recordset
ADOMen.Open "SELECT * FROM si_menu where tipodesistema=" & vgtipo & " ORDER BY MEN_CODIGO", VGconfig, adOpenStatic

Do While Not ADOMen.EOF
    If Len(ADOMen("Men_Codigo")) = 2 And ADOMen("Men_Visible") Then
        ' Agrega un nodo al TreeView y establece sus propiedades.
        Set mNode = TreeView1.Nodes.Add(1, tvwChild, ADOMen("Men_Codigo") & " ID", ADOMen("Men_Descri"), "Cerrar")
        mNode.Tag = "Menu" ' Identifica la tabla.
        ' Establece en la variable intIndex la propiedad Index del
        ' objeto Node recién creado. Use esta variable para agregar
        ' objetos Node hijos al objeto Node actual.
            
        intIndex01 = mNode.Index
        bolVisibleP = True
        ADOMen.MoveNext
        If ADOMen.EOF Then Exit Do
         
    ElseIf Len(ADOMen("Men_Codigo")) = 2 And ADOMen("Men_Visible") = False Then
        bolVisibleP = False
        ADOMen.MoveNext
        If ADOMen.EOF Then Exit Do
    End If
        
    If Len(ADOMen("Men_Codigo")) = 4 And ADOMen("Men_Visible") And bolVisibleP Then
        Set mNode = TreeView1.Nodes.Add(intIndex01, tvwChild)
        mNode.text = ADOMen("Men_Descri") ' Texto.
        mNode.key = ADOMen("Men_Codigo") & " ID"  ' ID único.
        mNode.Image = "Cerrar"     ' Imagen de ImageList.
            
        intIndex02 = mNode.Index
        bolVisibleP = True
        bolVisibleS = True
        
        ADOMen.MoveNext
        If ADOMen.EOF Then Exit Do
        
    Else
        If Len(ADOMen("Men_Codigo")) = 4 And ADOMen("Men_Visible") = False Then
                bolVisibleS = False
                ADOMen.MoveNext
                If ADOMen.EOF Then Exit Do
                If Len(ADOMen("Men_Codigo")) > 4 Then
                        bolVisibleS = False
                ElseIf Len(ADOMen("Men_Codigo")) <= 4 Or bolVisibleP Then
                        bolVisibleS = True
                End If
        ElseIf Len(ADOMen("Men_Codigo")) = 4 And ADOMen("Men_Visible") And bolVisibleP = False Then
                ADOMen.MoveNext
                If ADOMen.EOF Then Exit Do
        End If
    End If
        
    If Len(ADOMen("Men_Codigo")) = 6 And ADOMen("Men_Visible") And bolVisibleP Then
        If bolVisibleS = True Then
            Set mNode = TreeView1.Nodes.Add(intIndex02, tvwChild)
            mNode.text = ADOMen("Men_Descri") ' Texto.
            mNode.key = ADOMen("Men_Codigo") & " ID"   ' ID único.
            mNode.Image = "Cerrar"     ' Imagen de ImageList.
            intIndex03 = mNode.Index
            bolVisibleS = True
        End If
        ADOMen.MoveNext
        If ADOMen.EOF Then Exit Do
        
    Else
        If Len(ADOMen("Men_Codigo")) = 6 And ADOMen("Men_Visible") = False Then
            bolVisibleS = False
            ADOMen.MoveNext
            If ADOMen.EOF Then Exit Do
            If Len(ADOMen("Men_Codigo")) > 6 Then
                   bolVisibleS = False
            ElseIf Len(ADOMen("Men_Codigo")) <= 6 Or bolVisibleP Then
                   bolVisibleS = True
            End If
        ElseIf Len(ADOMen("Men_Codigo")) = 6 And ADOMen("Men_Visible") And bolVisibleP = False Then
                ADOMen.MoveNext
                If ADOMen.EOF Then Exit Do
        End If
    End If
    
    If Len(ADOMen("Men_Codigo")) = 8 And ADOMen("Men_Visible") And bolVisibleP Then
        If bolVisibleS = True Then
            Set mNode = TreeView1.Nodes.Add(intIndex03, tvwChild)
            mNode.text = ADOMen("Men_Descri") ' Texto.
            mNode.key = ADOMen("Men_Codigo") & " ID"   ' ID único.
            mNode.Image = "Cerrar"     ' Imagen de ImageList.
        
            bolVisibleS = True
        End If
        ADOMen.MoveNext
        If ADOMen.EOF Then Exit Do
        
    Else
        If Len(ADOMen("Men_Codigo")) = 8 And ADOMen("Men_Visible") = False Then
            If ADOMen("Men_Visible") = False Then bolVisibleS = False
            ADOMen.MoveNext
            If ADOMen.EOF Then Exit Do
            If Len(ADOMen("Men_Codigo")) > 8 Then
                   bolVisibleS = False
            ElseIf Len(ADOMen("Men_Codigo")) <= 8 Or bolVisibleP Then
                    bolVisibleS = True
            End If
        ElseIf Len(ADOMen("Men_Codigo")) = 8 And ADOMen("Men_Visible") And bolVisibleP = False Then
                ADOMen.MoveNext
                If ADOMen.EOF Then Exit Do
        End If
    End If
Loop
Exit Sub
Err1:
Resume Next
End Sub

Private Sub TreeView1_Collapse(ByVal Node As Node)
    If Node.text = "Menu" Or Node.Index > 1 Then
        Node.Image = "Cerrar"
    End If
End Sub

Private Sub TreeView1_Expand(ByVal Node As Node)
    If Node.text = "Menu" Or Node.Index > 1 Then
        If TreeView1.Nodes(Node.Index).Children > 0 Then
            If Node.Checked = False Then
                Node.Image = "Cerrar"
                Node.Expanded = False
            Else
                Node.Image = "Abrir"
            End If
        End If
        Node.Sorted = False
    End If
End Sub

Private Sub Carga_Valor(nIni As Integer, nfin As Integer, bFlag As Boolean, nG As Integer, Optional cCod As String)
If nG = 1 Then
    For nI = nIni To nfin
        TreeView1.Nodes.item(nI).Checked = bFlag
    Next nI
ElseIf nG = 2 Then
    For nI = nIni To nfin
        If Mid(TreeView1.Nodes(nI).key, 1, Len(Trim(cCod))) = Trim(cCod) Then
            TreeView1.Nodes.item(nI).Checked = bFlag
        End If
    Next nI
End If
End Sub

Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
If Node.Index = 1 Then
    If Node.Root.Checked = True Then
        Call Carga_Valor(1, TreeView1.Nodes.count, Node.Root.Checked, 1)
        Node.Expanded = True
    Else
        Node.Expanded = False
    End If
Else
    If TreeView1.Nodes(Node.Index).Children > 0 Then
        If Node.Checked = False Then
            Call Carga_Valor(Node.Index + 1, TreeView1.Nodes.count, False, 2, Mid(TreeView1.Nodes(Node.Index).key, 1, InStr(TreeView1.Nodes(Node.Index).key, " ID")))
            Node.Expanded = False
        Else
            Call Carga_Valor(Node.Index + 1, TreeView1.Nodes.count, True, 2, Mid(TreeView1.Nodes(Node.Index).key, 1, InStr(TreeView1.Nodes(Node.Index).key, " ID")))
            Node.Expanded = True
        End If
    End If
End If
End Sub


Private Sub Grab_Men(cCod As String)
Dim cCad As String
Dim nIi As Integer
Dim nOp As Integer
nIi = 2

ADOMen.MoveFirst
'Wait.CargarWait ADOMen.RecordCount, 0
'Wait.PonLabel "Guardando Opciones ..."
Set AdoUsu = New ADODB.Recordset
rsql = "delete si_menuusuarios Where tipodesistema=" & vgtipo & " and Usu_Codigo = '" & cCod & "'"
Set AdoUsu = VGconfig.Execute(rsql)
Do While Not ADOMen.EOF
    If TreeView1.Nodes(1).Checked Then  'Raiz
        If TreeView1.Nodes.item(nIi).key = ADOMen("Men_Codigo") & " ID" And ADOMen("Men_Visible") Then
            nOp = 2
            rsql = "Select * From si_menuusuarios Where tipodesistema=" & vgtipo & " and Usu_Codigo = '" & cCod & "' and Men_Codigo = '" & ADOMen("Men_Codigo") & "'"
            AdoUsu.Open rsql, VGconfig, adOpenStatic
            If AdoUsu.RecordCount > 0 Then
                    nOp = 1
            End If
            AdoUsu.Close
            
            If nOp = 2 Then
                    rsql = "Insert Into si_menuusuarios (Usu_Codigo,tipodesistema,Men_Codigo,Men_Hab) Values ('" & cCod & "'," & vgtipo & ",'" & ADOMen("Men_Codigo") & "',"
                    If TreeView1.Nodes.item(nIi).Checked = True Then
                        rsql = rsql & "1" & ")"
                    Else
                        rsql = rsql & "0" & ")"
                    End If
            Else
                    rsql = "Update si_menuusuarios Set Men_Hab = "
                    If TreeView1.Nodes.item(nIi).Checked = True Then
                        rsql = rsql & "1"
                    Else
                        rsql = rsql & "0"
                    End If
                    rsql = rsql & " Where tipodesistema=" & vgtipo & " and Usu_Codigo = '" & cCod & "' and Men_Codigo = '" & ADOMen("Men_Codigo") & "'"
            End If
            VGconfig.Execute rsql
        End If
        nIi = nIi + 1
    Else
        nIi = 0
        Exit Do
    End If
    ADOMen.MoveNext
    'Wait.Inc
    If ADOMen.EOF Then Exit Do
Loop
'Unload Wait
If nIi >= 2 Then
    MsgBox "Se ha grabado completamente las opciones escogidas", vbInformation, "Mensaje"
Else
    MsgBox "No se ha grabado las opciones", vbInformation, "Verificar"
End If
End Sub

Private Sub Edit_Men(cCodU As String)
Dim ADOUsMe As ADODB.Recordset
Dim nJ As Integer
Set ADOUsMe = New ADODB.Recordset

For nJ = 1 To TreeView1.Nodes.count
    TreeView1.Nodes(nJ).Checked = False
    TreeView1.Nodes(nJ).Expanded = False
Next nJ

ADOUsMe.Open "SELECT * FROM si_menuusuarios  WHERE tipodesistema=" & vgtipo & " and USU_CODIGO = '" & cCodU & "' ", VGconfig, adOpenStatic
If ADOUsMe.RecordCount > 0 Then ADOUsMe.MoveFirst
If Not ADOUsMe.EOF Then
    TreeView1.Nodes(1).Checked = True 'Raiz
    Do While Not ADOUsMe.EOF
        For nJ = 2 To TreeView1.Nodes.count
            If TreeView1.Nodes(nJ).key = ADOUsMe("MEN_CODIGO") & " ID" Then
                If ADOUsMe("Men_Hab") Then
                    TreeView1.Nodes(nJ).Checked = True
                Else
                    TreeView1.Nodes(nJ).Checked = False
                End If
                Exit For
            End If
        Next nJ
        ADOUsMe.MoveNext
        If ADOUsMe.EOF Then Exit Do
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
