VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frmUsuario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuarios"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   Icon            =   "frmUsuario.frx":0000
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Nuevo"
      Height          =   675
      Index           =   0
      Left            =   645
      Picture         =   "frmUsuario.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2850
      Width           =   775
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Grabar"
      Height          =   675
      Index           =   1
      Left            =   1605
      Picture         =   "frmUsuario.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2850
      Width           =   775
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "E&ditar"
      Height          =   675
      Index           =   2
      Left            =   2565
      Picture         =   "frmUsuario.frx":114E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2850
      Width           =   775
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Eliminar"
      Height          =   675
      Index           =   3
      Left            =   3525
      Picture         =   "frmUsuario.frx":1590
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2850
      Width           =   775
   End
   Begin VB.CommandButton cmdBotones 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   675
      Index           =   5
      Left            =   4485
      Picture         =   "frmUsuario.frx":19D2
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2850
      Width           =   775
   End
   Begin VB.Frame Frame0 
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   240
      TabIndex        =   14
      Top             =   120
      Width           =   5055
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2415
         Left            =   0
         TabIndex        =   15
         Top             =   120
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   4260
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
      Height          =   2535
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   5175
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1920
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   1920
         MaxLength       =   5
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   1920
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   7
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Confirmar            :"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   13
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Password Usuario   :"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   12
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Código Usuario       :"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   11
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre Usuario  :"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   840
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db1 As ADODB.Connection

Dim Adoreg1 As ADODB.Recordset
Dim AdoReg2 As ADODB.Recordset
Dim cCad As String
Dim RegActual As Integer
Dim nFra As Integer
Dim nTipo As Integer
Dim nI As Integer


Private Sub Form_Load()
central Me
ADOConectar
Init_ControlDataGrid DataGrid1
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
End Sub


Private Sub ADOConectar()
Set db1 = New ADODB.Connection
Set Adoreg1 = New ADODB.Recordset
Set AdoReg2 = New ADODB.Recordset

'Adoreg1.CursorType = adOpenDynamic           NO USAR CAMBIAR BOOKMARK

With db1
  .CursorLocation = adUseClient
  .Provider = "Microsoft.Jet.OLEDB.3.51"
  .ConnectionString = "Data Source =" & cRuta2
  .Open
End With

Adoreg1.Open "Select usuariocodigo,usuarioNombre from USUARIO_INV where USUARIO_INV.EMP_CODIGO=" & "'" & VGCodEmpresa & "'", VGConfig, adOpenStatic
AdoReg2.Open "Select * from USUARIO_INV where USUARIO_INV.EMP_CODIGO=" & "'" & VGCodEmpresa & "'", VGConfig, adOpenDynamic, adLockOptimistic

 
Set DataGrid1.DataSource = Adoreg1
End Sub

Public Sub SetDataGrid()
 DataGrid1.Refresh
 DataGrid1.Columns(0).Caption = "           Código"
 DataGrid1.Columns(1).Caption = "                            Nombre"
 DataGrid1.Columns(0).Width = 1500
 DataGrid1.Columns(1).Width = 3700
End Sub




Private Sub cmdBotones_Click(Index As Integer)
Dim nIi As Integer
Dim tempi As Integer
Dim temps As String
Select Case Index
 Case 0: 'Nuevo
         nTipo = 1
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
            Exit Sub
         End If
         If Text1(1) = "" Then
            MsgBox "Ingrese el nombre de usuario", vbExclamation, mensaje1
            Exit Sub
         End If
         If Text1(2) = "" Then
            MsgBox "Ingrese el password del usuario", vbExclamation, mensaje1
            Exit Sub
         End If
         If Text1(3) = "" Then
            MsgBox "Ingrese la confirmación del usuario", vbExclamation, mensaje1
            Exit Sub
         End If
         Screen.MousePointer = 11
         If nTipo = 1 Then
            Dim flag As Boolean
            flag = False
            'buscar igual codigo
            With Adoreg1
                If .RecordCount <> 0 Then
                    .MoveFirst
                    .Find "usuariocodigo= '" & UCase(Text1(0).text) & "'"
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
                    AdoReg2.Fields("usuariocodigo") = UCase(Trim(Text1(0).text))
                    AdoReg2.Fields("Emp_Codigo") = VGCodEmpresa
                    AdoReg2.Fields("UsuarioPassword") = CODIFICA(Trim(Text1(2).text), NUMMAGICO) 'password                    ADOREG2.UpdateBatch
                    If Trim(Text1(1).text) <> "" Then AdoReg2.Fields("usuarioNombre") = UCase(Trim(Text1(1).text))
                    AdoReg2.Update
                    Adoreg1.Requery
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
            AdoReg2.Fields("usuariocodigo") = UCase(Trim(Text1(0).text))
            AdoReg2.Fields("Emp_Codigo") = VGCodEmpresa
            AdoReg2.Fields("UsuarioPassword") = CODIFICA(Trim(Text1(2).text), NUMMAGICO)
            If Trim(Text1(1).text) <> "" Then AdoReg2.Fields("usuarioNombre") = UCase(Trim(Text1(1).text))
            AdoReg2.UpdateBatch
            Adoreg1.Requery
            AdoReg2.Requery
       
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
            Text1(2).text = DECODIFICA(AdoReg2.Fields(2), NUMMAGICO)
            Text1(3).text = DECODIFICA(AdoReg2.Fields(2), NUMMAGICO)
            If Not IsNull(AdoReg2.Fields("usuarioNombre")) Then Text1(1).text = AdoReg2.Fields("usuarioNombre")
        
          
            
            Frame1.Caption = "Modificar Usuario"
            Frame1.Visible = True
            Frame0.Visible = False
            Botones_Set False
            Text1(0).Enabled = False
            Text1(1).SetFocus
            'If xFlag Then
            ' xFlag = False
            ' cmdBotones_Click 5
            ' cmdBotones_Click 2
            'Else
            ' xFlag = True
            'End If
            Screen.MousePointer = 1
         Else
            MsgBox "Debe seleccionar un Registro para editarlo", vbInformation
            Botones_Set False
            cmdBotones_Click 5
         End If
       
 Case 3: 'Eliminar
          Dim op As Integer
          op = MsgBox("Esta Seguro que desea Eliminar el registro actual ?", vbYesNo, "Eliminación de Registro")
          If op = vbYes Then
            AdoReg2.Bookmark = Adoreg1.Bookmark
           
           
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
                .Find "usuariocodigo= '" & UCase(Text1(0).text) & "'"
                If Not .EOF Then
                    Text1(0).text = ""
                    MsgBox "El Usuario ya existe:  Ingrese de nuevo", vbInformation, "Ingreso de Datos"
                    Text1(0).SetFocus: Exit Sub
                End If
            End If
        End With
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



