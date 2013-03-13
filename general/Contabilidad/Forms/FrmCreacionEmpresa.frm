VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmCreacionEmpresa 
   Caption         =   "Creacion de empresas"
   ClientHeight    =   4695
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   6450
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   7435
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Lista de Empresas"
      TabPicture(0)   =   "FrmCreacionEmpresa.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "DataGrid1"
      Tab(0).Control(1)=   "Command1"
      Tab(0).Control(2)=   "Command2"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Datos de empresa"
      TabPicture(1)   =   "FrmCreacionEmpresa.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label4"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label5"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Ctr_Ayuempresa"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Text1"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Text2"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Text3"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Text4"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Command3"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Command4"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).ControlCount=   12
      Begin VB.CommandButton Command4 
         Caption         =   "Salir"
         Height          =   495
         Left            =   4320
         TabIndex        =   12
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Grabar"
         Height          =   495
         Left            =   2880
         TabIndex        =   11
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Salir"
         Height          =   495
         Left            =   -70680
         TabIndex        =   10
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   2040
         MaxLength       =   11
         TabIndex        =   5
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   4
         Top             =   2040
         Width           =   3615
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2040
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1560
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   2
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Adicionar"
         Height          =   495
         Left            =   -72120
         TabIndex        =   1
         Top             =   3480
         Width           =   1215
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuempresa 
         Height          =   315
         Left            =   2040
         TabIndex        =   14
         Top             =   480
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         XcodMaxLongitud =   3
         xcodwith        =   300
         NomTabla        =   "co_multiempresas"
         TituloAyuda     =   "Busqueda de Empresas"
         ListaCampos     =   "empresacodigo(1),empresadescripcion(1),agentederetencion(1)"
         XcodCampo       =   "empresacodigo"
         XListCampo      =   "empresadescripcion"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "empresacodigo,empresadescripcion,agentederetencion"
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2775
         Left            =   -74760
         TabIndex        =   15
         Top             =   480
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   4895
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
               LCID            =   10250
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
               LCID            =   10250
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
      Begin VB.Label Label5 
         Caption         =   "Empresa Base"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Ruc"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Direccion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Razon Social"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   1200
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmCreacionEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsql As New ADODB.Recordset
Dim rsql1 As New ADODB.Recordset
Dim codigo As String

Private Sub Command1_Click()
Text1 = codigo
SSTab1.Tab = 1
End Sub


Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
SQL = " insert co_multiempresas ( empresacodigo, empresadescripcion,empresadireccion, empresaruc )"
SQL = SQL & " values ( '" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "')"
Set rsql1 = VGCNx.Execute(SQL)
Call adicionaregistros
Call cargar
SSTab1.Tab = 0
End Sub
Private Sub adicionaregistros()
'  -------  plan de cuentas
If ExisteElem(0, VGCNx, "" & VGComputer & "") Then VGCNx.Execute ("drop table " & VGComputer & "")
SQL = "select * into " & VGComputer & " from ct_cuenta "
SQL = SQL & " where empresacodigo='" & Ctr_Ayuempresa.xclave & "'"
Set rsql1 = VGCNx.Execute(SQL)
'
SQL = "update " & VGComputer & " set empresacodigo='" & Text1 & "'"
Set rsql1 = VGCNx.Execute(SQL)
'
SQL = "insert ct_cuenta select * from  " & VGComputer & ""
Set rsql1 = VGCNx.Execute(SQL)
' --------------  asiento
SQL = " insert ct_asientocorre ( empresacodigo,asientocodigo,asientoanno )"
SQL = SQL & " select '" & Text1 & "',asientocodigo,'" & VGParamSistem.Anoproceso & "'"
SQL = SQL & " from ct_asientocorre where empresacodigo='" & Ctr_Ayuempresa.xclave & "'"
SQL = SQL & " and asientoanno='" & VGParamSistem.Anoproceso & "'"
Set rsql1 = VGCNx.Execute(SQL)
'
' --------------  libro
SQL = " insert ct_librocorre ( empresacodigo,librocodigo,libroanno )"
SQL = SQL & " select '" & Text1 & "',librocodigo,'" & VGParamSistem.Anoproceso & "'"
SQL = SQL & " from ct_librocorre where empresacodigo='" & Ctr_Ayuempresa.xclave & "'"
SQL = SQL & " and libroanno='" & VGParamSistem.Anoproceso & "'"
Set rsql1 = VGCNx.Execute(SQL)
' --------------  centro de costos
If ExisteElem(0, VGCNx, "" & VGComputer & "") Then VGCNx.Execute ("drop table " & VGComputer & "")
SQL = "select * into " & VGComputer & " from ct_centrocosto "
SQL = SQL & " where empresacodigo='" & Ctr_Ayuempresa.xclave & "'"
Set rsql1 = VGCNx.Execute(SQL)
'
SQL = "update " & VGComputer & " set empresacodigo='" & Text1 & "'"
Set rsql1 = VGCNx.Execute(SQL)
'
SQL = "insert ct_centrocosto select * from  " & VGComputer & ""
Set rsql1 = VGCNx.Execute(SQL)

Call cargar
SSTab1.Tab = 0
End Sub

Private Sub Command4_Click()
SSTab1.Tab = 0
End Sub

Private Sub Form_Load()
Ctr_Ayuempresa.conexion VGCNx
If VGParametros.multiempresas = False Then
   MsgBox (" Sistema no esta habilitado para varias empresas ")
   Exit Sub
End If
SSTab1.Tab = 0
Call cargar
End Sub
Private Sub cargar()
SQL = "select * from co_multiempresas"
Set rsql = VGCNx.Execute(SQL)
Set DataGrid1.DataSource = rsql
DataGrid1.Refresh
Set rsql1 = VGCNx.Execute("select codigo=max(empresacodigo)+1 from co_multiempresas where empresacodigo<'90'")
codigo = Format(rsql1!codigo, "00")
End Sub

