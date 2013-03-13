VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmCfgAdminist 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Administradores"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   Icon            =   "FrmCfgAdminist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   6735
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   300
      TabIndex        =   9
      Top             =   2520
      Width           =   6135
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Nuevo"
         Height          =   675
         Index           =   0
         Left            =   240
         Picture         =   "FrmCfgAdminist.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   775
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Grabar"
         Height          =   675
         Index           =   1
         Left            =   1440
         Picture         =   "FrmCfgAdminist.frx":0D0C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   775
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "E&ditar"
         Height          =   675
         Index           =   2
         Left            =   2640
         Picture         =   "FrmCfgAdminist.frx":114E
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   775
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Eliminar"
         Height          =   675
         Index           =   3
         Left            =   3840
         Picture         =   "FrmCfgAdminist.frx":1590
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   775
      End
      Begin VB.CommandButton cmdBotones 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   675
         Index           =   5
         Left            =   5040
         Picture         =   "FrmCfgAdminist.frx":19D2
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   775
      End
   End
   Begin VB.Frame FrameU 
      Caption         =   "Lista de Administradores: "
      Height          =   2415
      Index           =   0
      Left            =   300
      TabIndex        =   7
      Top             =   120
      Width           =   6135
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1935
         Left            =   480
         TabIndex        =   8
         Top             =   360
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   3413
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
               ColumnWidth     =   2055.118
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1769.953
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameU 
      Caption         =   "Nuevo Administrador: "
      Height          =   2415
      Index           =   1
      Left            =   300
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   3360
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   975
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   1
         Top             =   570
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   3360
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1395
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Confirmar:"
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   6
         Top             =   1395
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Password:"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   5
         Top             =   1005
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Código:"
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   4
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.Frame FrameU 
      Caption         =   "Modificar Usuario: "
      Height          =   2415
      Index           =   2
      Left            =   300
      TabIndex        =   15
      Top             =   120
      Width           =   6135
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   3360
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   18
         Top             =   1395
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   0
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   16
         Top             =   570
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   3360
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   17
         Top             =   975
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Código :"
         Height          =   255
         Index           =   5
         Left            =   1560
         TabIndex        =   21
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Password:"
         Height          =   255
         Index           =   6
         Left            =   1560
         TabIndex        =   20
         Top             =   1005
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Confirmar:"
         Height          =   255
         Index           =   7
         Left            =   1560
         TabIndex        =   19
         Top             =   1395
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmCfgAdminist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Adoreg1 As ADODB.Recordset
Dim AdoReg2 As ADODB.Recordset
Dim AdoMax As ADODB.Recordset

Dim RegActual As Integer
Dim nFra As Integer

Private Sub cmdBotones_Click(Index As Integer)
Dim tempi As Integer
Dim temps As String
Select Case Index
 Case 0: 'Nuevo
         FrameU(nFra).Visible = False
         FrameU(1).Visible = True
         nFra = 1
         Dim otext As TextBox
         For Each otext In Me.Text1
          otext.text = ""
         Next
         Botones_Set False
         Text1(0).SetFocus
 Case 1: 'Grabar
         If FrameU(1).Visible Then 'Nuevo
          Dim flag As Boolean
          flag = False
          'buscar igual codigo
          With AdoReg2
           If .RecordCount <> 0 Then
            .MoveFirst
            Do While Not .EOF
             If UCase(Text1(0).text) = .Fields(1) Then
              flag = True
              Text1(0).text = ""
              MsgBox "El Nombre ya existe:  Ingrese de nuevo", vbInformation, "Ingreso de Datos"
              Exit Do
             End If
             .MoveNext
            Loop
           End If
          End With
          If Not flag Then
           If Text1(1).text = "" Then
            MsgBox "Ud. No ha ingresado su Password", vbInformation, "Ingreso de Datos"
            Text1(1).SetFocus
           ElseIf Text1(2).text = "" Then
            MsgBox "Ud. No ha ingresado su confirmación", vbInformation, "Ingreso de Datos"
            Text1(2).SetFocus
           ElseIf Text1(1).text = Text1(2).text Then
            'pasa
            AdoReg2.AddNew
            AdoReg2.Fields(0) = GenerarCodigo
            AdoReg2.Fields(1) = UCase(Trim(Text1(0).text))
            AdoReg2.Fields(2) = CODIFICA(Trim(Text1(1).text), NUMMAGICO) 'password
            
            AdoReg2.UpdateBatch
            Adoreg1.Requery
            AdoReg2.Requery
            AdoMax.Requery
            
            FrameU(nFra).Visible = False
            FrameU(0).Visible = True
            nFra = 0
            Botones_Set True
           Else
            MsgBox "Nombre de Password y la Confirmación no coinciden", vbInformation, "Ingreso de datos"
            Text1(1).text = ""
            Text1(2).text = ""
            Text1(1).SetFocus
           End If
          End If
         End If
         
         If FrameU(2).Visible Then
          AdoReg2.Fields(1) = UCase(Trim(Text2(0).text))
          AdoReg2.Fields(2) = CODIFICA(Trim(Text2(1).text), NUMMAGICO)
          Adoreg1.UpdateBatch
          AdoReg2.UpdateBatch
          Adoreg1.Requery
          AdoReg2.Requery
          AdoMax.Requery
          FrameU(nFra).Visible = False
          FrameU(0).Visible = True
          nFra = 0
          Botones_Set True
         End If
         SetDataGrid
 Case 2: 'Editar
         If Adoreg1.Bookmark Then
          AdoReg2.Bookmark = Adoreg1.Bookmark
          FrameU(nFra).Visible = False
          FrameU(2).Visible = True
          nFra = 2
          Text2(0).text = AdoReg2.Fields(1)
          Text2(1).text = DECODIFICA(AdoReg2.Fields(2), NUMMAGICO)
          Text2(2).text = DECODIFICA(AdoReg2.Fields(2), NUMMAGICO)
          Botones_Set False
          Text2(0).SetFocus
         Else
          MsgBox "Debe seleccionar un Registro para editarlo", vbInformation
          Botones_Set False
          cmdBotones_Click 5
         End If
 Case 3: 'Eliminar
          Dim op As Integer
          op = MsgBox("Esta Seguro que desea Eliminar el registro actual", vbYesNo, "Eliminación de Registro")
          If op = vbYes Then
           AdoReg2.Bookmark = Adoreg1.Bookmark
           AdoReg2.Delete
           AdoReg2.UpdateBatch
           AdoReg2.Requery
           Adoreg1.Requery
           AdoMax.Requery
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
          FrameU(nFra).Visible = False
          FrameU(0).Visible = True
          nFra = 0
          If Adoreg1.RecordCount = 0 Then
           Botones_Init True
          Else
           Botones_Set True
          End If
         End If
End Select
End Sub

Public Sub Botones_Set(flag As Boolean)
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
' cmdBotones(4).Enabled = Not flag 'Buscar
 cmdBotones(5).Caption = "&Salir" 'Salir
End Sub
Private Sub DataGrid1_Click()
 RegActual = IIf(IsNull(DataGrid1.Bookmark), 0, DataGrid1.Bookmark)
End Sub

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
 Dim fra As Frame
 For Each fra In Me.FrameU
  fra.Visible = False
 Next
 FrameU(0).Visible = True
 nFra = 0
End Sub

Public Sub ADOConectar()
 Set Adoreg1 = New ADODB.Recordset
 Set AdoReg2 = New ADODB.Recordset
 Set AdoMax = New ADODB.Recordset
 
 Adoreg1.CursorType = adOpenDynamic
 Adoreg1.Open "Select ADM_NOMBRE from ADMINISTRADOR", VGConfig, adOpenStatic, adLockOptimistic
 AdoReg2.Open "Select * from ADMINISTRADOR", VGConfig, adOpenStatic, adLockOptimistic
 AdoMax.Open "Select MAX(ADM_CODIGO) from ADMINISTRADOR", VGConfig, adOpenStatic, adLockOptimistic
 Set DataGrid1.DataSource = Adoreg1
End Sub

Public Sub SetDataGrid()
 DataGrid1.Refresh
 DataGrid1.Columns(0).Caption = "Nombre"
 DataGrid1.Columns(0).Width = 5500
End Sub

Public Function GenerarCodigo() As String
 Dim Aux As Integer
 If AdoMax.RecordCount = 0 Then
  Aux = 0
 Else
  If Not IsNull(AdoMax.Fields(0)) Then
   Aux = Val(Mid(AdoMax.Fields(0), 2, 4))
  Else
   Aux = 0
  End If
 End If
 Aux = Aux + 1
 GenerarCodigo = Format(Aux, "A0000")
End Function


Private Sub Text1_GotFocus(Index As Integer)
With Text1(Index)
 .SelStart = 0
 .SelLength = Len(.text)
End With
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
 SendKeys "{tab}"
 KeyAscii = 0
End If
End Sub
Private Sub Text2_GotFocus(Index As Integer)
With Text2(Index)
 .SelStart = 0
 .SelLength = Len(.text)
End With
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
 SendKeys "{tab}"
 KeyAscii = 0
End If
End Sub

