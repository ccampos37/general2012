VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form frmAsientos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asientos Contables"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6450
   Icon            =   "frmAsientos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   6450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   3503
      TabIndex        =   5
      Top             =   5100
      Width           =   1380
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   1568
      TabIndex        =   4
      Top             =   5100
      Width           =   1380
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   90
      TabIndex        =   0
      Top             =   75
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   8705
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Con&figuración"
      TabPicture(0)   =   "frmAsientos.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Registros"
      TabPicture(1)   =   "frmAsientos.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "Frame5"
      Tab(1).Control(2)=   "cmdExpo"
      Tab(1).Control(3)=   "Frame3"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Provisiones"
      TabPicture(2)   =   "frmAsientos.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2(0)"
      Tab(2).Control(1)=   "Frame2(1)"
      Tab(2).Control(2)=   "Frame2(2)"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Cuentas Contables"
      TabPicture(3)   =   "frmAsientos.frx":035E
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "DataGrid1"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Command7"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Command6"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Command5"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Command4"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).ControlCount=   5
      Begin VB.Frame Frame4 
         Caption         =   "Cuenta de Redondeo"
         Height          =   855
         Left            =   -74805
         TabIndex        =   55
         Top             =   3585
         Width           =   5895
         Begin AplisetControlText.Aplitext xcomp 
            Height          =   300
            Left            =   2055
            TabIndex        =   56
            Top             =   360
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   529
            Text            =   ""
         End
         Begin AplisetControlText.Aplitext Aplitext18 
            Height          =   300
            Left            =   4665
            TabIndex        =   71
            Top             =   360
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   529
            Text            =   ""
         End
         Begin VB.Label Label24 
            Caption         =   "Nº Cuenta para QUINCENA"
            Height          =   420
            Left            =   3375
            TabIndex        =   58
            Top             =   345
            Width           =   1245
         End
         Begin VB.Label Label22 
            Caption         =   "Nº Cuenta para Fin de MES"
            Height          =   420
            Left            =   870
            TabIndex        =   57
            Top             =   360
            Width           =   1155
         End
         Begin VB.Image Image7 
            Height          =   240
            Left            =   300
            Picture         =   "frmAsientos.frx":037A
            Top             =   270
            Width           =   240
         End
         Begin VB.Image Image8 
            Height          =   480
            Left            =   360
            Picture         =   "frmAsientos.frx":06BC
            Top             =   315
            Width           =   480
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Empresa de relacion en Contabilidad"
         Enabled         =   0   'False
         Height          =   915
         Left            =   -74820
         TabIndex        =   51
         Top             =   480
         Width           =   5910
         Begin VB.CommandButton cmdemp 
            Caption         =   "..."
            Height          =   315
            Left            =   5325
            TabIndex        =   53
            Top             =   345
            Width           =   435
         End
         Begin AplisetControlText.Aplitext xEmp 
            Height          =   300
            Left            =   960
            TabIndex        =   52
            Top             =   345
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   529
            Locked          =   -1  'True
            Text            =   ""
         End
         Begin VB.Label Lemp 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1500
            TabIndex        =   54
            Top             =   345
            Width           =   3795
         End
         Begin VB.Image Image10 
            Height          =   480
            Left            =   405
            Picture         =   "frmAsientos.frx":0AFE
            Top             =   450
            Width           =   480
         End
         Begin VB.Image Image9 
            Height          =   420
            Left            =   285
            Picture         =   "frmAsientos.frx":0E08
            Stretch         =   -1  'True
            Top             =   240
            Width           =   315
         End
      End
      Begin VB.CommandButton cmdExpo 
         Caption         =   "&Exportar Trabjadores"
         Enabled         =   0   'False
         Height          =   330
         Left            =   -74820
         TabIndex        =   50
         Top             =   4515
         Width           =   1980
      End
      Begin VB.Frame Frame3 
         Caption         =   "Configuración de Trabajadores"
         Height          =   2160
         Left            =   -74805
         TabIndex        =   49
         Top             =   1380
         Width           =   5895
         Begin VB.Frame Frame7 
            Caption         =   "Para Quincena"
            Height          =   1200
            Left            =   2400
            TabIndex        =   66
            Top             =   795
            Width           =   3405
            Begin AplisetControlText.Aplitext Aplitext1 
               Height          =   300
               Left            =   150
               TabIndex        =   67
               Top             =   705
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   529
               Text            =   ""
            End
            Begin AplisetControlText.Aplitext Aplitext17 
               Height          =   300
               Left            =   1740
               TabIndex        =   69
               Top             =   690
               Width           =   1560
               _ExtentX        =   2752
               _ExtentY        =   529
               Text            =   ""
            End
            Begin VB.Label Label26 
               Caption         =   "Neto Adelanto"
               Height          =   270
               Left            =   1755
               TabIndex        =   70
               Top             =   390
               Width           =   1275
            End
            Begin VB.Label Label25 
               Caption         =   "Cuenta Contable"
               Height          =   270
               Left            =   90
               TabIndex        =   68
               Top             =   420
               Width           =   1275
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Para Fin de Mes"
            Height          =   1200
            Left            =   225
            TabIndex        =   59
            Top             =   780
            Width           =   1980
            Begin AplisetControlText.Aplitext xcuenta 
               Height          =   300
               Left            =   165
               TabIndex        =   60
               Top             =   705
               Width           =   1485
               _ExtentX        =   2619
               _ExtentY        =   529
               Text            =   ""
            End
            Begin VB.Label Label23 
               Caption         =   "Cuenta Contable"
               Height          =   270
               Left            =   120
               TabIndex        =   61
               Top             =   405
               Width           =   1320
            End
         End
         Begin AplisetControlText.Aplitext xsubdi 
            Height          =   300
            Left            =   4590
            TabIndex        =   62
            Top             =   255
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   529
            Text            =   ""
         End
         Begin AplisetControlText.Aplitext xAnexo 
            Height          =   300
            Left            =   2070
            TabIndex        =   63
            Top             =   270
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            Text            =   ""
         End
         Begin VB.Image Image6 
            Height          =   480
            Left            =   225
            Picture         =   "frmAsientos.frx":114A
            Top             =   210
            Width           =   480
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Anexo"
            Height          =   195
            Left            =   885
            TabIndex        =   65
            Top             =   315
            Width           =   1035
         End
         Begin VB.Label Label21 
            Caption         =   "Subdiario Contable"
            Height          =   375
            Left            =   3270
            TabIndex        =   64
            Top             =   225
            Width           =   1320
         End
         Begin VB.Image Image5 
            Height          =   480
            Left            =   30
            Picture         =   "frmAsientos.frx":1454
            Top             =   195
            Width           =   480
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Provisiones de CTS"
         Height          =   1365
         Index           =   0
         Left            =   -74790
         TabIndex        =   36
         Top             =   510
         Width           =   5835
         Begin AplisetControlText.Aplitext Aplitext6 
            Height          =   285
            Left            =   4185
            TabIndex        =   37
            Top             =   945
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   503
            Text            =   ""
         End
         Begin AplisetControlText.Aplitext Aplitext5 
            Height          =   285
            Left            =   4185
            TabIndex        =   38
            Top             =   615
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   503
            Text            =   ""
         End
         Begin AplisetControlText.Aplitext Aplitext4 
            Height          =   285
            Left            =   1305
            TabIndex        =   39
            Top             =   945
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   503
            Text            =   ""
         End
         Begin AplisetControlText.Aplitext Aplitext3 
            Height          =   285
            Left            =   1305
            TabIndex        =   40
            Top             =   615
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   503
            Text            =   ""
         End
         Begin AplisetControlText.Aplitext Aplitext2 
            Height          =   285
            Left            =   1305
            TabIndex        =   41
            Top             =   285
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            Text            =   ""
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Subdiario"
            Height          =   195
            Left            =   255
            TabIndex        =   47
            Top             =   330
            Width           =   660
         End
         Begin VB.Label Label4 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2115
            TabIndex        =   46
            Top             =   285
            Width           =   3375
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta 1"
            Height          =   195
            Left            =   255
            TabIndex        =   45
            Top             =   660
            Width           =   645
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta 2"
            Height          =   195
            Left            =   255
            TabIndex        =   44
            Top             =   990
            Width           =   645
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta 3"
            Height          =   195
            Left            =   3270
            TabIndex        =   43
            Top             =   660
            Width           =   645
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta 4"
            Height          =   195
            Left            =   3270
            TabIndex        =   42
            Top             =   975
            Width           =   645
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Provisiones de Gratificaciones"
         Height          =   1365
         Index           =   1
         Left            =   -74790
         TabIndex        =   24
         Top             =   1950
         Width           =   5835
         Begin AplisetControlText.Aplitext Aplitext7 
            Height          =   285
            Left            =   4185
            TabIndex        =   25
            Top             =   945
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   503
            Text            =   ""
         End
         Begin AplisetControlText.Aplitext Aplitext8 
            Height          =   285
            Left            =   4185
            TabIndex        =   26
            Top             =   615
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   503
            Text            =   ""
         End
         Begin AplisetControlText.Aplitext Aplitext9 
            Height          =   285
            Left            =   1305
            TabIndex        =   27
            Top             =   945
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   503
            Text            =   ""
         End
         Begin AplisetControlText.Aplitext Aplitext10 
            Height          =   285
            Left            =   1305
            TabIndex        =   28
            Top             =   615
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   503
            Text            =   ""
         End
         Begin AplisetControlText.Aplitext Aplitext11 
            Height          =   285
            Left            =   1305
            TabIndex        =   29
            Top             =   285
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            Text            =   ""
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta 4"
            Height          =   195
            Left            =   3270
            TabIndex        =   35
            Top             =   975
            Width           =   645
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta 3"
            Height          =   195
            Left            =   3270
            TabIndex        =   34
            Top             =   660
            Width           =   645
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta 2"
            Height          =   195
            Left            =   255
            TabIndex        =   33
            Top             =   990
            Width           =   645
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta 1"
            Height          =   195
            Left            =   255
            TabIndex        =   32
            Top             =   660
            Width           =   645
         End
         Begin VB.Label Label13 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2115
            TabIndex        =   31
            Top             =   285
            Width           =   3375
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Subdiario"
            Height          =   195
            Left            =   255
            TabIndex        =   30
            Top             =   330
            Width           =   660
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Provisiones de Vacaciones"
         Height          =   1365
         Index           =   2
         Left            =   -74790
         TabIndex        =   12
         Top             =   3405
         Width           =   5835
         Begin AplisetControlText.Aplitext Aplitext12 
            Height          =   285
            Left            =   4185
            TabIndex        =   13
            Top             =   945
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   503
            Text            =   ""
         End
         Begin AplisetControlText.Aplitext Aplitext13 
            Height          =   285
            Left            =   4185
            TabIndex        =   14
            Top             =   615
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   503
            Text            =   ""
         End
         Begin AplisetControlText.Aplitext Aplitext14 
            Height          =   285
            Left            =   1305
            TabIndex        =   15
            Top             =   945
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   503
            Text            =   ""
         End
         Begin AplisetControlText.Aplitext Aplitext15 
            Height          =   285
            Left            =   1305
            TabIndex        =   16
            Top             =   615
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   503
            Text            =   ""
         End
         Begin AplisetControlText.Aplitext Aplitext16 
            Height          =   285
            Left            =   1305
            TabIndex        =   17
            Top             =   285
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            Text            =   ""
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta 4"
            Height          =   195
            Left            =   3270
            TabIndex        =   23
            Top             =   975
            Width           =   645
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta 3"
            Height          =   195
            Left            =   3270
            TabIndex        =   22
            Top             =   660
            Width           =   645
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta 2"
            Height          =   195
            Left            =   255
            TabIndex        =   21
            Top             =   990
            Width           =   645
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta 1"
            Height          =   195
            Left            =   255
            TabIndex        =   20
            Top             =   660
            Width           =   645
         End
         Begin VB.Label Label19 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2115
            TabIndex        =   19
            Top             =   285
            Width           =   3375
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Subdiario"
            Height          =   195
            Left            =   255
            TabIndex        =   18
            Top             =   330
            Width           =   660
         End
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Agregar"
         Height          =   360
         Left            =   900
         TabIndex        =   10
         Top             =   4170
         Width           =   1050
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Editar"
         Height          =   360
         Left            =   2115
         TabIndex        =   9
         Top             =   4170
         Width           =   1050
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Eliminar"
         Height          =   360
         Left            =   3315
         TabIndex        =   8
         Top             =   4170
         Width           =   1050
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Capturar"
         Height          =   360
         Left            =   4530
         TabIndex        =   7
         Top             =   4170
         Width           =   1050
      End
      Begin VB.Frame Frame1 
         Caption         =   "St@rsoft Contabilidad"
         Height          =   4230
         Left            =   -74790
         TabIndex        =   1
         Top             =   495
         Width           =   5805
         Begin VB.CheckBox xStarsoft 
            Caption         =   "Contamos con una instalación de St@rsoft Contabilidad"
            Height          =   225
            Left            =   930
            TabIndex        =   48
            Top             =   1875
            Width           =   4245
         End
         Begin VB.CommandButton cmValidar 
            Caption         =   "&Validar"
            Height          =   330
            Left            =   4170
            TabIndex        =   6
            Top             =   2190
            Width           =   990
         End
         Begin VB.CheckBox xImportaTrab 
            Caption         =   "Exportar Trabajadores como anexo al momento de crear un trabajador desde planilla"
            Enabled         =   0   'False
            Height          =   435
            Left            =   930
            TabIndex        =   3
            Top             =   3240
            Width           =   4605
         End
         Begin VB.Image Image4 
            Height          =   480
            Left            =   270
            Picture         =   "frmAsientos.frx":2296
            Top             =   3210
            Width           =   480
         End
         Begin VB.Label Label1 
            Caption         =   $"frmAsientos.frx":2B60
            Height          =   630
            Index           =   0
            Left            =   930
            TabIndex        =   2
            Top             =   720
            Width           =   4530
         End
         Begin VB.Image Image2 
            Height          =   480
            Left            =   360
            Picture         =   "frmAsientos.frx":2C20
            Top             =   900
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   165
            Picture         =   "frmAsientos.frx":2F2A
            Top             =   675
            Width           =   480
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3420
         Left            =   210
         TabIndex        =   11
         Top             =   660
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   6033
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
         Caption         =   "Cuentas Contables"
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
   End
End
Attribute VB_Name = "frmAsientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FLAG As Boolean
Dim RSCFGCONTA As New ADODB.Recordset
Dim sName As String

Private Sub Aplitext1_Click()

    If xStarsoft.Value = 0 Then Exit Sub
    Dim RSCUENTA As New ADODB.Recordset
    Set RSCUENTA = New ADODB.Recordset
    RSCUENTA.Open "SELECT PLANCTA_CODIGO,PLANCTA_DESCRIPCION FROM PLAN_CUENTA_NACIONAL WHERE PLANCTA_NIVEL=" & REGSISTEMA.scNivelCta, CONECTARDBSQL(xEmp.Text & "BDCONTABILIDAD"), adOpenKeyset, adLockReadOnly
    frmComun.CONECTAR RSCUENTA
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        Aplitext1.Text = VGUTIL(1)
    End If

End Sub

Private Sub Aplitext17_Click()
    If xStarsoft.Value = 0 Then Exit Sub
    Dim RSCUENTA As New ADODB.Recordset
    Set RSCUENTA = New ADODB.Recordset
    RSCUENTA.Open "SELECT PLANCTA_CODIGO,PLANCTA_DESCRIPCION FROM PLAN_CUENTA_NACIONAL WHERE PLANCTA_NIVEL=" & REGSISTEMA.scNivelCta, CONECTARDBSQL(xEmp.Text & "BDCONTABILIDAD"), adOpenKeyset, adLockReadOnly
    frmComun.CONECTAR RSCUENTA
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        Aplitext17.Text = VGUTIL(1)
    End If

End Sub

Private Sub Aplitext18_Click()
    If xStarsoft.Value = 0 Then Exit Sub
    Dim RSCUENTA As New ADODB.Recordset
    Set RSCUENTA = New ADODB.Recordset
    RSCUENTA.Open "SELECT PLANCTA_CODIGO,PLANCTA_DESCRIPCION FROM PLAN_CUENTA_NACIONAL WHERE PLANCTA_NIVEL=" & REGSISTEMA.scNivelCta, CONECTARDBSQL(xEmp.Text & "BDCONTABILIDAD"), adOpenKeyset, adLockReadOnly
    frmComun.CONECTAR RSCUENTA
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        Aplitext18.Text = VGUTIL(1)
    End If

End Sub

Private Sub cmdemp_Click()
    Dim sIniFile As String
    Dim RsEmp As New ADODB.Recordset
    Set RsEmp = New ADODB.Recordset
    Dim CNXAUX As New ADODB.Connection
    Set CNXAUX = New ADODB.Connection
    
    Set CNXAUX = CONECTARDBSQL("BDWENCO")
    RsEmp.Open "SELECT EMP_CODIGO,EMP_RAZON_NOMBRE FROM EMPRESA", CNXAUX, adOpenKeyset, adLockReadOnly
    
    frmComun.CONECTAR RsEmp
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xEmp.Text = VGUTIL(1)
        Lemp.Caption = VGUTIL(2)
        REGSISTEMA.scNivelCta = DevuelveValor("SELECT EMP_NIVEL FROM EMPRESA WHERE EMP_CODIGO='" & xEmp.Text & "'", CNXAUX)
    End If
End Sub
Private Sub cmValidar_Click()
    If Not EXISTECONTA Then
        MsgBox "No existe la base de datos principal de Contabilidad", vbExclamation
        FLAG = False
        xStarsoft.Value = 0
        Frame5.Enabled = False
      Else
        MsgBox "Si Valido Satisfactoriamente", vbInformation
    End If
End Sub
Private Sub Command2_Click()
    Dim SqlCad As String
    If Not VALIDAR Then Exit Sub
    SqlCad = "UPDATE CFGASIENTOS SET CHKCONTA=" & xStarsoft.Value & "," & _
             "RUTCONTA=''," & _
             "CHKCREATRAB=" & xImportaTrab.Value & "," & _
             "CODEMP='" & ESNULO(Trim(xEmp.Text), " ") & "'," & _
             "NOMEMP='" & ESNULO(Trim(Lemp.Caption), " ") & "'," & _
             "TIPANEX='" & ESNULO(Trim(xAnexo.Text), " ") & "'," & _
             "SUBDI='" & ESNULO(Trim(xsubdi.Text), " ") & "'," & _
             "COMP='" & ESNULO(Trim(xcomp.Text), " ") & "'," & _
             "CUENTA='" & ESNULO(Trim(xcuenta.Text), " ") & "'," & _
             "NETADEL='" & ESNULO(Trim(Aplitext17.Text), " ") & "'," & _
             "MONADEL='" & ESNULO(Trim(Aplitext1.Text), " ") & "'," & _
             "REDADEL='" & ESNULO(Trim(Aplitext18.Text), " ") & "'"
             
    DBSYSTEM.Execute SqlCad
    Call SETCFGCONTA
    Unload Me
End Sub
Private Function VALIDAR() As Boolean
    VALIDAR = True
    If xStarsoft.Value = 0 Then Exit Function
    If Trim(xEmp.Text) = "" Then
        MsgBox "Debe seleccionar la empresa", vbExclamation
        SSTab1.Tab = 1: cmdemp.SetFocus
        VALIDAR = False
        Exit Function
    End If
    If Trim(xAnexo.Text) = "" Then
        MsgBox "Colocar el tipo de anexo", vbExclamation
        SSTab1.Tab = 1: xAnexo.SetFocus
        VALIDAR = False
        Exit Function
    End If
    If Trim(xsubdi.Text) = "" Then
        MsgBox "Debe colocar el subdiario de Planillas", vbExclamation
        SSTab1.Tab = 1: xsubdi.SetFocus
        VALIDAR = False
        Exit Function
    End If
    If Trim(xcuenta.Text) = "" Then
        MsgBox "Debe colocar la cuenta general de Planillas", vbExclamation
        SSTab1.Tab = 1: xcuenta.SetFocus
        VALIDAR = False
        Exit Function
    End If
End Function

Private Sub COMMAND3_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    SSTab1.TabVisible(2) = False
    SSTab1.TabVisible(3) = False
    SSTab1.Tab = 0
    Call CARGARDATOS
End Sub

Private Sub TEXT1_CHANGE()

End Sub

Private Sub xAnexo_DblClick()
    If xStarsoft.Value = 0 Then Exit Sub
    Dim RSAUX As New ADODB.Recordset
    Set RSAUX = New ADODB.Recordset
    RSAUX.Open "TIPO_ANEXO", CONECTARDBSQL(xEmp.Text & "BDCONTABILIDAD"), adOpenKeyset, adLockReadOnly
    If RSAUX.RecordCount = 0 Then
        MsgBox "No se han encontrado Tipos de Anexo en Contabilidad", vbInformation
    Else
        Dim Str2 As String
        Str2 = RSAUX.Source
        frmComun.CONECTAR RSAUX
        frmComun.Show 1
        If VGUTIL(1) <> "" Then
            xAnexo.Text = VGUTIL(1)
        End If
    End If
    Set RSAUX = Nothing
End Sub

Private Sub xcomp_DblClick()
    If xStarsoft.Value = 0 Then Exit Sub
    Dim RSCUENTA As New ADODB.Recordset
    Set RSCUENTA = New ADODB.Recordset
    RSCUENTA.Open "SELECT PLANCTA_CODIGO,PLANCTA_DESCRIPCION FROM PLAN_CUENTA_NACIONAL WHERE PLANCTA_NIVEL=" & REGSISTEMA.scNivelCta, CONECTARDBSQL(xEmp.Text & "BDCONTABILIDAD"), adOpenKeyset, adLockReadOnly
    frmComun.CONECTAR RSCUENTA
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xcomp.Text = VGUTIL(1)
    End If
End Sub

Private Sub xcuenta_DblClick()
    If xStarsoft.Value = 0 Then Exit Sub
    Dim RSCUENTA As New ADODB.Recordset
    Set RSCUENTA = New ADODB.Recordset
    RSCUENTA.Open "SELECT PLANCTA_CODIGO,PLANCTA_DESCRIPCION FROM PLAN_CUENTA_NACIONAL WHERE PLANCTA_NIVEL=" & REGSISTEMA.scNivelCta, CONECTARDBSQL(xEmp.Text & "BDCONTABILIDAD"), adOpenKeyset, adLockReadOnly
    frmComun.CONECTAR RSCUENTA
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xcuenta.Text = VGUTIL(1)
    End If
End Sub

Private Sub xRutaWEnco_DblClick()
    If xStarsoft.Value = 0 Then Exit Sub
End Sub

Private Sub xStarsoft_Click()
    If xStarsoft.Value = 1 Then
        xImportaTrab.Enabled = True
        cmValidar.Enabled = True
        Frame5.Enabled = True
        cmdExpo.Enabled = True
    Else
        xImportaTrab.Enabled = False
        cmValidar.Enabled = False
        Frame5.Enabled = False
        cmdExpo.Enabled = False
    End If
End Sub

Private Sub xsubdi_DblClick()
    If xStarsoft.Value = 0 Then Exit Sub
    Dim RSAUX As New ADODB.Recordset
    Set RSAUX = New ADODB.Recordset
    RSAUX.Open "SELECT SUBDIAR_CODIGO,SUBDIAR_DESCRIPCION FROM  SUBDIARIOS", CONECTARDBSQL(xEmp.Text & "BDCONTABILIDAD"), adOpenKeyset, adLockReadOnly
    If RSAUX.RecordCount = 0 Then
        MsgBox "No se han encontrado Tipos de Anexo en Contabilidad", vbInformation
    Else
        Dim Str2 As String
        Str2 = RSAUX.Source
        frmComun.CONECTAR RSAUX
        frmComun.Show 1
        If VGUTIL(1) <> "" Then
            xsubdi.Text = VGUTIL(1)
        End If
    End If
    Set RSAUX = Nothing
End Sub
Private Sub CARGARDATOS()
    Set RSCFGCONTA = New ADODB.Recordset
    RSCFGCONTA.Open "CFGASIENTOS", DBSYSTEM, adOpenStatic, adLockReadOnly
    xStarsoft.Value = IIf(RSCFGCONTA("CHKCONTA"), 1, 0)
    xImportaTrab.Value = IIf(RSCFGCONTA("CHKCREATRAB"), 1, 0)
    xEmp.Text = Trim(RSCFGCONTA("CODEMP"))
    If xEmp.Text = "" Then
        Frame5.Enabled = False
      Else
        Frame5.Enabled = True
    End If
    Lemp.Caption = Trim(RSCFGCONTA("NOMEMP"))
    xAnexo.Text = Trim(RSCFGCONTA("TIPANEX"))
    xsubdi.Text = Trim(RSCFGCONTA("SUBDI"))
    xcomp.Text = Trim(RSCFGCONTA("COMP"))
    xcuenta.Text = Trim(RSCFGCONTA("CUENTA"))
    
    Aplitext1.Text = ESNULO(Trim(RSCFGCONTA("MONADEL")), "")
    Aplitext17.Text = ESNULO(Trim(RSCFGCONTA("NETADEL")), "")
    Aplitext18.Text = ESNULO(Trim(RSCFGCONTA("REDADEL")), "")
End Sub
