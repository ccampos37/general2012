VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmRequerimientos 
   Caption         =   "Generacion de Requerimientos"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   9128
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Consultas"
      TabPicture(0)   =   "FrmRequerimientos.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "DataGrid1"
      Tab(0).Control(1)=   "CmdSalir"
      Tab(0).Control(2)=   "cmdNue"
      Tab(0).Control(3)=   "CmdEli"
      Tab(0).Control(4)=   "cmdEdi"
      Tab(0).Control(5)=   "cmdImp"
      Tab(0).Control(6)=   "Data2"
      Tab(0).Control(7)=   "CrystalReport1"
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Ingresos "
      TabPicture(1)   =   "FrmRequerimientos.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Flex1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "FrameBotones"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Fradatos"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "fraTotales"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin VB.Frame fraTotales 
         Height          =   975
         Left            =   135
         TabIndex        =   29
         Top             =   4080
         Visible         =   0   'False
         Width           =   9825
         Begin VB.Label lblCom 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "892,760.00"
            Height          =   285
            Left            =   7080
            TabIndex        =   39
            Top             =   600
            Width           =   1110
         End
         Begin VB.Label lblIgv 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "892,760.00"
            Height          =   285
            Left            =   7080
            TabIndex        =   38
            Top             =   240
            Width           =   1110
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Compra :"
            Height          =   195
            Left            =   6360
            TabIndex        =   37
            Top             =   600
            Width           =   630
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "I.G.V.   :"
            Height          =   195
            Left            =   6360
            TabIndex        =   36
            Top             =   240
            Width           =   585
         End
         Begin VB.Label lblTot 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "892,760.00"
            Height          =   285
            Left            =   4200
            TabIndex        =   35
            Top             =   240
            Width           =   1230
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Total  :"
            Height          =   195
            Left            =   3600
            TabIndex        =   34
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblDes 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "892,760.00"
            Height          =   285
            Left            =   1680
            TabIndex        =   33
            Top             =   600
            Width           =   1110
         End
         Begin VB.Label lblImp 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "892,760.00"
            Height          =   285
            Left            =   1680
            TabIndex        =   32
            Top             =   240
            Width           =   1110
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Descuento :"
            Height          =   195
            Left            =   720
            TabIndex        =   31
            Top             =   600
            Width           =   870
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Importe      :"
            Height          =   195
            Left            =   720
            TabIndex        =   30
            Top             =   240
            Width           =   840
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3060
         Left            =   -74730
         TabIndex        =   6
         Top             =   525
         Width           =   9480
         _ExtentX        =   16722
         _ExtentY        =   5398
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
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
         Caption         =   "Reqerimientos pendientes de Atender"
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "tipoordencodigo"
            Caption         =   "T.Orden"
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
            DataField       =   "OC_CNUMORD"
            Caption         =   "        Número"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "OC_CRAZSOC"
            Caption         =   "                   Desc. Proveedor"
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
         BeginProperty Column03 
            DataField       =   "OC_DFECDOC"
            Caption         =   "    Emisión"
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
         BeginProperty Column04 
            DataField       =   "OC_CCODMON"
            Caption         =   "Mo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "#.##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "OC_NVENTA"
            Caption         =   "     Importe"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "EST_NOMBRE"
            Caption         =   "      Estado"
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
         BeginProperty Column07 
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
            MarqueeStyle    =   2
            ScrollBars      =   2
            AllowRowSizing  =   0   'False
            Size            =   273
            BeginProperty Column00 
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               Locked          =   -1  'True
               WrapText        =   -1  'True
               ColumnWidth     =   1365.165
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               WrapText        =   -1  'True
               ColumnWidth     =   3105.071
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               Locked          =   -1  'True
               WrapText        =   -1  'True
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               Locked          =   -1  'True
               WrapText        =   -1  'True
               ColumnWidth     =   434.835
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               Locked          =   -1  'True
               WrapText        =   -1  'True
               ColumnWidth     =   1019.906
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1379.906
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   14.74
            EndProperty
         EndProperty
      End
      Begin VB.Frame Fradatos 
         Height          =   1425
         Left            =   120
         TabIndex        =   20
         Top             =   975
         Width           =   9825
         Begin VB.TextBox txtObs 
            Height          =   288
            Left            =   2130
            TabIndex        =   22
            Top             =   945
            Width           =   7500
         End
         Begin VB.TextBox txtEntE 
            Height          =   288
            Left            =   3420
            MaxLength       =   50
            TabIndex        =   21
            Top             =   225
            Width           =   5295
         End
         Begin MSComCtl2.DTPicker txtEmi 
            Height          =   285
            Left            =   1005
            TabIndex        =   23
            Top             =   225
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Format          =   52428801
            CurrentDate     =   37015
         End
         Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_solicitante 
            Height          =   315
            Left            =   3405
            TabIndex        =   24
            Top             =   555
            Width           =   4110
            _ExtentX        =   7250
            _ExtentY        =   556
            XcodMaxLongitud =   3
            xcodwith        =   300
            NomTabla        =   "co_solicitantes"
            TituloAyuda     =   "Busqueda de Solicitante"
            ListaCampos     =   "solicitantecodigo(1),solicitantenombre(1)"
            XcodCampo       =   "solicitantecodigo"
            XListCampo      =   "solicitantenombre"
            ListaCamposDescrip=   "Codigo,Descripcion"
            ListaCamposText =   "solicitantecodigo,solicitantenombre"
         End
         Begin VB.Label Label12 
            Caption         =   "Observación :"
            Height          =   255
            Left            =   210
            TabIndex        =   28
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Emisión         :"
            Height          =   195
            Left            =   90
            TabIndex        =   27
            Top             =   240
            Width           =   990
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Entregar en   :"
            Height          =   195
            Left            =   2490
            TabIndex        =   26
            Top             =   240
            Width           =   1005
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Solicitante     :"
            Height          =   195
            Left            =   2490
            TabIndex        =   25
            Top             =   600
            Width           =   1005
         End
      End
      Begin VB.Frame Frame1 
         Height          =   636
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   9780
         Begin ctrlayuda_f.Ctr_Ayuda Ctrayu_tipoorden 
            Height          =   390
            Left            =   1170
            TabIndex        =   14
            Top             =   195
            Width           =   3390
            _ExtentX        =   5980
            _ExtentY        =   688
            XcodMaxLongitud =   11
            xcodwith        =   1100
            NomTabla        =   "co_tipodeorden"
            TituloAyuda     =   "Busqueda de Tipo de Orden"
            ListaCampos     =   "tipoordencodigo(1),tipoordendescripcion(1),tipoordennumeracion(2)"
            XcodCampo       =   "tipoordencodigo"
            XListCampo      =   "tipoordendescripcion"
            ListaCamposDescrip=   "Codigo,Descripcion"
            ListaCamposText =   "tipoordencodigo,tipoordendescripcion,tipoordennumeracion"
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "Tipo Orden     :"
            Height          =   192
            Left            =   96
            TabIndex        =   19
            Top             =   276
            Width           =   1032
         End
         Begin VB.Label lblNum 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   348
            Left            =   5340
            TabIndex        =   18
            Top             =   192
            Width           =   1560
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Número  :"
            Height          =   192
            Left            =   4656
            TabIndex        =   17
            Top             =   288
            Width           =   696
         End
         Begin VB.Label lblEst 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   348
            Left            =   7728
            TabIndex        =   16
            Top             =   204
            Width           =   1644
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Estado  :"
            Height          =   192
            Left            =   7080
            TabIndex        =   15
            Top             =   288
            Width           =   636
         End
      End
      Begin VB.Frame FrameBotones 
         Height          =   4575
         Left            =   10050
         TabIndex        =   7
         Top             =   480
         Width           =   1095
         Begin VB.CommandButton CmdSalir2 
            Caption         =   "&Salir"
            Height          =   615
            Left            =   195
            Picture         =   "FrmRequerimientos.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   3720
            Visible         =   0   'False
            Width           =   800
         End
         Begin VB.CommandButton cmdGra 
            Caption         =   "&Grabar"
            Height          =   630
            Left            =   165
            Picture         =   "FrmRequerimientos.frx":047A
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   2880
            Visible         =   0   'False
            Width           =   800
         End
         Begin VB.CommandButton cmdEdi2 
            Caption         =   "&Editar"
            Enabled         =   0   'False
            Height          =   630
            Left            =   135
            Picture         =   "FrmRequerimientos.frx":08BC
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   1200
            Visible         =   0   'False
            Width           =   800
         End
         Begin VB.CommandButton cmdEli2 
            Caption         =   "&Quitar"
            Enabled         =   0   'False
            Height          =   630
            Left            =   165
            Picture         =   "FrmRequerimientos.frx":0CFE
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   2040
            Visible         =   0   'False
            Width           =   800
         End
         Begin VB.CommandButton cmdNue2 
            Caption         =   "&Agregar"
            Height          =   630
            Left            =   120
            Picture         =   "FrmRequerimientos.frx":1140
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   360
            Visible         =   0   'False
            Width           =   800
         End
      End
      Begin VB.CommandButton CmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   675
         Left            =   -67920
         Picture         =   "FrmRequerimientos.frx":1582
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3750
         Width           =   775
      End
      Begin VB.CommandButton cmdNue 
         Caption         =   "&Nuevo"
         Height          =   675
         Left            =   -73185
         Picture         =   "FrmRequerimientos.frx":19C4
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3735
         Width           =   775
      End
      Begin VB.CommandButton CmdEli 
         Caption         =   "&Anular"
         Height          =   675
         Left            =   -70530
         Picture         =   "FrmRequerimientos.frx":1E06
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3765
         Width           =   775
      End
      Begin VB.CommandButton cmdEdi 
         Caption         =   "&Editar"
         Height          =   675
         Left            =   -71850
         Picture         =   "FrmRequerimientos.frx":2248
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3750
         Width           =   775
      End
      Begin VB.CommandButton cmdImp 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   -69240
         Picture         =   "FrmRequerimientos.frx":268A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3765
         Visible         =   0   'False
         Width           =   775
      End
      Begin VB.Data Data2 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -74760
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4125
         Visible         =   0   'False
         Width           =   1140
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Bindings        =   "FrmRequerimientos.frx":2ACC
         Left            =   -74760
         Top             =   3645
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex1 
         Height          =   1515
         Left            =   120
         TabIndex        =   40
         Top             =   2520
         Visible         =   0   'False
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   2672
         _Version        =   393216
         Cols            =   15
         FixedCols       =   0
         RowHeightMin    =   240
         BackColorSel    =   0
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   "^Código|Fab|Descripción|xUni|xCantidad|Uni.|Cantidad|PU|>Precio|>%Des|Igv|>Total|C1|C2"
         RowSizingMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   15
      End
   End
End
Attribute VB_Name = "FrmRequerimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Colex As New Collection
Dim adodc1 As ADODB.Recordset
Public VGvardllgen As dllgeneral.dll_general
Dim cSql1 As String
Dim nT As Integer       'Ingreso,Modificación,Ficha Tecnica
Dim cCod As String
Dim nTra As Integer
Dim Mensaje As String

Dim unum As String


Sub OculObj02(nTipo As Boolean)
    cmdGra.Visible = nTipo
    CmdSalir2.Visible = nTipo
End Sub

Sub OculObj03(nTipo As Boolean)
    Fradatos.Visible = nTipo
    fraTotales.Visible = nTipo
End Sub

Sub OculObj04(nTipo As Boolean)
    cmdNue.Visible = nTipo
    cmdEdi.Visible = nTipo
    CmdEli.Visible = nTipo
    cmdImp.Visible = nTipo
    CmdSalir.Visible = nTipo
End Sub

Sub OculObj06(nTipo As Boolean)
    DataGrid1.Visible = nTipo
End Sub

Sub Abre_Tabla_OCs()
    Dim strsql As String
    
    Set adodc1 = New ADODB.Recordset
    
    strsql = "SELECT * FROM co_cabordcompra a,co_estadoorden b WHERE a.estadooccodigo= b." & _
        "estadooccodigo and estadoocatendido<>1 ORDER BY oc_cnumord "
    adodc1.Open strsql, VGcnx, adOpenDynamic, adLockOptimistic
    
    Set DataGrid1.DataSource = adodc1
    
End Sub

Private Sub cmdEdi2_Click()
On Error GoTo Err
    With frmemisionOCdetalle
        .activado = False
        .CtrAyu_articulo.xclave = Flex1.TextMatrix(Flex1.Row, 0)
        .lblFab = Flex1.TextMatrix(Flex1.Row, 1)
        .CtrAyu_articulo.xnombre = Flex1.TextMatrix(Flex1.Row, 2)
        .lblUni = Flex1.TextMatrix(Flex1.Row, 3)
        .txtCan = Flex1.TextMatrix(Flex1.Row, 4)
        .txtCan.Enabled = True
        .tipo = Flex1.TextMatrix(Flex1.Row, 14)
        If Flex1.TextMatrix(Flex1.Row, 3) <> Flex1.TextMatrix(Flex1.Row, 5) Then
            .txtURe = Flex1.TextMatrix(Flex1.Row, 5)
            .txtRef = Flex1.TextMatrix(Flex1.Row, 6)
        Else
            .txtURe = ""
            .txtRef = ""
        End If
        If .txtURe <> "" Then .txtRef.Enabled = True
        .txtPUn = Flex1.TextMatrix(Flex1.Row, 7)
        .txtPDe = Flex1.TextMatrix(Flex1.Row, 9)
        .txtPIg = Flex1.TextMatrix(Flex1.Row, 10)
'        .Igv = .txtPIg
        .txtordfab = Flex1.TextMatrix(Flex1.Row, 12)
        .txtCo1 = Flex1.TextMatrix(Flex1.Row, 13)
        .CtrAyu_articulo.Enabled = False
        .activado = True
        .Calculo_Automatico
        .Show 1
        
        If Not .cancelado Then
            If .tipo = "S" Then
              .txtCan = 1
            End If
            Flex1.TextMatrix(Flex1.Row, 2) = .CtrAyu_articulo.xnombre
            Flex1.TextMatrix(Flex1.Row, 4) = .txtCan
            If .txtURe = "" Then
                Flex1.TextMatrix(Flex1.Row, 5) = .lblUni
                Flex1.TextMatrix(Flex1.Row, 6) = .txtCan
            Else
                Flex1.TextMatrix(Flex1.Row, 5) = .txtURe
                Flex1.TextMatrix(Flex1.Row, 6) = .txtRef
            End If
            Flex1.TextMatrix(Flex1.Row, 7) = .txtPUn
            Flex1.TextMatrix(Flex1.Row, 8) = .txtPUn
            Flex1.TextMatrix(Flex1.Row, 9) = .txtPDe
            Flex1.TextMatrix(Flex1.Row, 10) = .txtPIg
            Flex1.TextMatrix(Flex1.Row, 11) = Format(Flex1.TextMatrix(Flex1.Row, 6) * Flex1.TextMatrix(Flex1.Row, 8), "0.00")
            Flex1.TextMatrix(Flex1.Row, 12) = .txtordfab
            Flex1.TextMatrix(Flex1.Row, 13) = .txtCo1
            Calcula_Totales
        End If
        Flex1.SetFocus
        cmdNue2.SetFocus
    End With
 Exit Sub
Err:
    MsgBox Err.Description
 
End Sub

Private Sub CmdEli_Click()
    On Error GoTo EliErr
    
    If adodc1("oc_estadoorden") = 1 Or adodc1("oc_situacionorden") <> "0" Then
        Mensaje = "Imposible anular la Orden de compra en su estado actual"
        MsgBox Mensaje, vbCritical, "Mensaje"
        DataGrid1.SetFocus
        Exit Sub
    End If

    Dim strsql As String
    Dim voc As String
    
    Mensaje = "¿Está seguro que desea anular la Orden de compra?"
    If MsgBox(Mensaje, vbQuestion + vbYesNo + vbDefaultButton2, "Mensaje") = vbYes Then
        voc = adodc1("oc_cnumord")
        
        nTra = 1
        VGcnx.BeginTrans
        
        strsql = "UPDATE co_detordcompra SET oc_situacionorden=2  WHERE oc_cnumord='" & voc & "'"
        VGcnx.Execute strsql
        strsql = "UPDATE co_cabordcompra SET oc_estadoorden=1 WHERE oc_cnumord='" & voc & "'"
        VGcnx.Execute strsql

        VGcnx.CommitTrans
        nTra = 0
        
        If nTra = 0 Then
            adodc1.Requery
            adodc1.Find "oc_cnumord='" & voc & "'"
        End If
    End If
    DataGrid1.SetFocus
    Exit Sub
Exit Sub
    
Dim Adodc2 As ADODB.Recordset

    Mensaje = "¿Desea eliminar el documento " & adodc1("nrorequi") & "?"
    If MsgBox(Mensaje, vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        strsql = "DELETE * FROM requisd WHERE nrorequi='" & adodc1("nrorequi") & "'"
        
        nTra = 1
        VGcnx.BeginTrans
        VGcnx.Execute strsql
        VGcnx.CommitTrans
        nTra = 0
        
        If nTra = 0 Then
            adodc1.Delete
            adodc1.Update
        End If
        Estado_Botones
            
    End If
    If adodc1.RecordCount > 0 Then
        DataGrid1.SetFocus
    Else
        cmdNue.SetFocus
    End If
    Exit Sub

EliErr:
    MsgBox Err.Description
    If nTra = 1 Then VGcnx.RollbackTrans
End Sub

Private Sub CmdEli2_Click()
    If Tiene_Entregas Then
        Mensaje = "El artículo tiene cantidad entregada"
        MsgBox Mensaje, vbExclamation, "Advertencia"
    End If
    
    Mensaje = "¿Desea quitar el artículo seleccionado?"
    If MsgBox(Mensaje, vbQuestion + vbYesNo + vbDefaultButton2, "Mensaje") = vbYes Then
        If Flex1.Rows - 1 = 1 Then
            Dim I As Integer
            
            For I = 0 To 13
                Flex1.TextMatrix(1, I) = ""
            Next
        Else
            Flex1.RemoveItem Flex1.Row
        End If
        Calcula_Totales
        Estado_Items
    End If
End Sub

Private Sub cmdGra_Click()
    Dim SQLc As String
    Dim SQLd As String
    Dim Rs2 As New ADODB.Recordset
    Dim I As Integer
    Dim vFactor As Single, vCantid As Single
    Dim vPreuni As Single, vDscpor As Single
    Dim vDescto As Single, vIgv As Single
    Dim vIgvpor As Single, vPrenet As Single
    Dim vTotven As Single, vTotnet As Single
    Dim vURef As String, txtMon As String
    Dim txtEst As String, txtTip As Integer
    Dim txtPro As String, txtSol As String
    Dim lblPro As String, txtFor As String
    On Error GoTo GrabErr
    
    txtTip = 0
    If Trim(Ctrayu_tipoorden.xclave) = "" Then
       Mensaje = "Debe ingresar Código de Tipo de Orden"
       MsgBox Mensaje, vbExclamation, "Mensaje"
       Ctrayu_tipoorden.SetFocus
       Exit Sub
    End If
    
    If txtEmi > Date Then
       MsgBox "Fecha de emision no debe ser mayor a la fecha del Sistema", vbExclamation, "Error"
       Exit Sub
       txtEmi.SetFocus
    End If
       
   
    txtEst = ""
    txtSol = Trim(CtrAyu_solicitante.xclave)
    If txtSol = "" Then
        Mensaje = "Debe ingresar Solicitante"
        MsgBox Mensaje, vbExclamation, "Mensaje"
        CtrAyu_solicitante.SetFocus
        Exit Sub
    End If
    If Not cmdEli2.Enabled Then
        Mensaje = "Debe especificar artículos de la Orden de Compra"
        MsgBox Mensaje, vbExclamation, "Error"
        cmdNue2.SetFocus
        Exit Sub
    End If
    
    If nT = 1 Then
        Mensaje = "¿Desea ingresar la nueva Orden de Compra?"
    Else
        Mensaje = "¿Desea guardar los cambios realizados?"
    End If
    
    If MsgBox(Mensaje, vbQuestion + vbYesNo, "Mensaje") = vbYes Then
 '      nTra = 1
       VGcnx.BeginTrans
       unum = Format(Val(lblNum), "00000000000")

       If nT = 1 Then      'Ingreso
         'unum = Format(Devolver_Dato(1, , " & trim(ctrayu_tipoordencodigo) & ", "tipoordencodigo", False,
         '      "ctnnumero"), "00000000000")
         SQLc = "select tipoordennumeracion from co_tipodeorden where tipoordencodigo='" & Trim(Ctrayu_tipoorden.xclave) & "' "
         Set Rs2 = New ADODB.Recordset
         Rs2.Open SQLc, VGcnx, adOpenKeyset, adLockReadOnly
         unum = Rs2!tipoordennumeracion + 1
          
          SQLc = "UPDATE co_tipodeorden SET tipoordennumeracion=" & unum & _
                " WHERE tipoordencodigo='" & Trim(Ctrayu_tipoorden.xclave) & "' "
            VGcnx.Execute SQLc
           unum = Format(Val(unum), "00000000000")
           lblNum = unum
            SQLc = "INSERT INTO co_cabordcompra (tipoordencodigo,oc_cnumord,oc_dfecdoc,oc_ccodpro," & _
                "oc_crazsoc,oc_ccotiza,oc_ccodmon,oc_cforpag,oc_dfecent," & _
                "oc_cobserv,oc_csolict,oc_centreg,oc_estadoorden,estadooccodigo,oc_nimport,oc_ndescue," & _
                "oc_nigv,oc_nventa,oc_dfecact,oc_chora,oc_cusuari,oc_cconver) VALUES ('" & _
                Ctrayu_tipoorden.xclave & "','" & lblNum & "','" & txtEmi & "','" & txtPro & "','','" _
                & txtCot & "','" & txtMon & "','" & txtFor & "','" & _
                txtEnt & "','" & _
                SupCadSQL(txtObs) & "','" & txtSol & "','" & txtEntE & "',' ','0'," & _
                CDbl(lblImp) & "," & CDbl(lblDes) & "," & CDbl(lblIgv) & "," & CDbl(lblCom) & _
                ",'" & txtEmi.Value & "','" & Format(Time, "hh.mm.ss") & "','" & VGUsuario & _
                "','" & txtEst & "')"
            VGcnx.Execute SQLc
            
            For I = 1 To Flex1.Rows - 1
                vFactor = Val(Flex1.TextMatrix(I, 6))
                vCantid = Val(Flex1.TextMatrix(I, 4))
                If vCantid = 0 Then
                   vCantid = 1
                End If
                vPreuni = Val(Flex1.TextMatrix(I, 7))
                vDscpor = Val(Flex1.TextMatrix(I, 9))
                vDescto = IIf(vFactor > 0, vFactor / vCantid, 1) * vPreuni * vCantid * _
                    vDscpor / 100
                vIgvpor = Val(Flex1.TextMatrix(I, 10))
                vTotven = Val(Flex1.TextMatrix(I, 11))
                vIgv = (vTotven - vDescto) * vIgvpor / 100
                vPrenet = Val(Flex1.TextMatrix(I, 8)) * (1 - vDscpor / 100)
                SQLd = "INSERT INTO co_detordcompra (tipoordencodigo,oc_cnumord,oc_ccodpro,oc_dfecdoc,oc_citem," & _
                  "oc_ccodigo,oc_ccodref,oc_cdesref,oc_cunidad,oc_cuniref,oc_nfactor," & _
                  "oc_ncantid,oc_nsaldo,oc_npreuni,oc_ndscpor,oc_ndescto,oc_nigv,oc_nigvpor," & _
                  "oc_nprenet,oc_ntotven,oc_ntotnet,oc_situacionorden,ord_fabnum,oc_ccomen1, tipoarticulocodigo, " & _
                  "oc_ncanten)" & _
                  "VALUES ('" & Ctrayu_tipoorden.xclave & "','" & lblNum & "','" & txtPro & "','" & txtEmi _
                  & "','" & Format(I, "000") & "','" & _
                  Flex1.TextMatrix(I, 0) & "','" & Flex1.TextMatrix(I, 1) & "','" & _
                  Flex1.TextMatrix(I, 2) & "','" & Flex1.TextMatrix(I, 3) & "','" & _
                  Flex1.TextMatrix(I, 5) & "'," & vFactor & "," & vCantid & "," & vCantid & "," & _
                  vPreuni & "," & vDscpor & "," & vDescto & "," & vIgv & "," & _
                  vIgvpor & "," & vPrenet & "," & vTotven & "," & vTotven - vDescto + _
                  vIgv & ",'0','" & Flex1.TextMatrix(I, 12) & "','" & _
                  Flex1.TextMatrix(I, 13) & "','" & Flex1.TextMatrix(I, 14) & "',0)"
                VGcnx.Execute SQLd
                
                SQLd = "UPDATE maeart SET aprecom=" & Val(Flex1.TextMatrix(I, 8)) & _
                    ",acodpro='" & txtPro & "',afecven='" & txtEmi _
                    & "' WHERE acodigo='" & Flex1.TextMatrix(I, 0) & "'"
                VGcnx.Execute SQLd
            Next
        ElseIf nT = 2 Then     'Modificar
            SQLc = "UPDATE co_cabordcompra SET oc_dfecdoc='" & txtEmi & _
                "',oc_ccotiza='" & txtCot & "',oc_ccodmon='" & txtMon & "',oc_cforpag='" & _
                txtFor & "',oc_ntipcam=" & Val(txtTip) & ",oc_dfecent='" & _
                txtEnt & "',oc_cobserv='" & SupCadSQL(txtObs) & _
                "',oc_csolict='" & txtSol & "',oc_centreg='" & txtEntE & "',oc_nimport=" & _
                CDbl(lblImp) & ",oc_ndescue=" & CDbl(lblDes) & ",oc_nigv=" & CDbl(lblIgv) & _
                ",oc_nventa=" & CDbl(lblCom) & ",oc_dfecact='" & _
                txtEmi.Value & "',oc_chora='" & Format(Time, "hh.mm.ss") & "',oc_cusuari='" & _
                VGUsuario & "',oc_cconver='" & txtEst & "' WHERE tipoordencodigo='" & Ctrayu_tipoorden.xclave & "' and oc_cnumord='" & lblNum & "'"
            VGcnx.Execute SQLc
            
            SQLd = "DELETE co_detordcompra WHERE tipoordencodigo='" & Ctrayu_tipoorden.xclave & "' and oc_cnumord='" & lblNum & "'"
            VGcnx.Execute SQLd
            
            For I = 1 To Flex1.Rows - 1
                vURef = ""
                vFactor = 0
                If Flex1.TextMatrix(I, 3) <> Flex1.TextMatrix(I, 5) Then
                    vURef = Flex1.TextMatrix(I, 5)
                    vFactor = Val(Flex1.TextMatrix(I, 6))
                End If
                vCantid = Val(Flex1.TextMatrix(I, 4))
                vPreuni = Val(Flex1.TextMatrix(I, 7))
                vDscpor = Val(Flex1.TextMatrix(I, 9))
                vDescto = IIf(vFactor > 0, vFactor / vCantid, 1) * vPreuni * vCantid * _
                    vDscpor / 100
                vIgvpor = Val(Flex1.TextMatrix(I, 10))
                vTotven = Val(Flex1.TextMatrix(I, 11))
                vIgv = (vTotven - vDescto) * vIgvpor / 100
                vPrenet = Val(Flex1.TextMatrix(I, 8)) * (1 - vDscpor / 100)
                SQLd = "INSERT INTO co_detordcompra (tipoordencodigo,oc_cnumord,oc_ccodpro,oc_dfecdoc,oc_citem," & _
                    "oc_ccodigo,oc_ccodref,oc_cdesref,oc_cunidad,oc_cuniref,oc_nfactor," & _
                    "oc_ncantid,oc_npreuni,oc_ndscpor,oc_ndescto,oc_nigv,oc_nigvpor," & _
                    "oc_nprenet,oc_ntotven,oc_ntotnet,oc_situacionorden,ord_fabnum,oc_ccomen1,tipoarticulocodigo, " & _
                    "oc_ncanten,oc_nsaldo)" & _
                    "VALUES ('" & Ctrayu_tipoorden.xclave & "','" & lblNum & "','" & txtPro & "','" & txtEmi _
                    & "','" & Format(I, "000") & "','" & _
                    Flex1.TextMatrix(I, 0) & "','" & Flex1.TextMatrix(I, 1) & "','" & _
                    Flex1.TextMatrix(I, 2) & "','" & Flex1.TextMatrix(I, 3) & "','" & _
                    vURef & "'," & vFactor & "," & vCantid & "," & _
                    vPreuni & "," & vDscpor & "," & vDescto & "," & vIgv & "," & _
                    vIgvpor & "," & vPrenet & "," & vTotven & "," & vTotven - vDescto + _
                    vIgv & ",'0','" & Flex1.TextMatrix(I, 12) & "','" & _
                    Flex1.TextMatrix(I, 13) & "', '" & Flex1.TextMatrix(I, 14) & "',0,0)"
                VGcnx.Execute SQLd
                
                SQLd = "UPDATE maeart SET aprecom=" & Val(Flex1.TextMatrix(I, 8)) & _
                    ",acodpro='" & txtPro & "',afecven='" & txtEmi _
                    & "' WHERE acodigo='" & Flex1.TextMatrix(I, 0) & "'"
                VGcnx.Execute SQLd
            Next
        End If
        
        VGcnx.CommitTrans
        nTra = 0
        adodc1.Requery
        adodc1.Find "oc_cnumord='" & lblNum & "'"
        
        If nT = 1 Then
            unum = Format(Val(unum) + 1, "00000000000")
            lblNum = unum
            Limpiar
            Vacia_FlexGrid
            Estado_Items
            Calcula_Totales
            txtEmi = Date
            txtEnt = Date
            txtTip = "0.000"
                        
        Else
            CmdSalir2_Click
        End If
    
End If
Exit Sub
GrabErr:
    MsgBox Err.Description
 '  Resume
    If nTra = 1 Then VGcnx.RollbackTrans
End Sub

Private Sub cmdImp_Click()
Dim formulas(3) As String
Dim tipoorden As String
unum = adodc1("oc_cnumord")
tipoorden = adodc1("tipoordencodigo")
CrystalReport1.Reset
CrystalReport1.WindowTitle = "rptcoordencompra -- orden de compra"
   CrystalReport1.ReportFileName = cRutP & "al_rptordencompra.rpt"
    CrystalReport1.DiscardSavedData = True
     
     
     
    'CrystalReport1.LogOnServer "pdssql.dll", _
    '                            VGServer, _
    '                            VGBase3, _
    '                            VGBUsuario2, _
    '                            VGPassw
    '
     
     
    CrystalReport1.Connect = "DSN=" & VGServer & ";DSQ=" & VGBase3 & ";UID=" & VGBUsuario2 & ";PWD=" & VGPassw
    'CrystalReport1.Connect = "DSN=" & VGServer & ";DSQ=" & VGBase3 & ";UID=" & VGUsuario & ";PWD=" & VGPassw
    'CrystalReport1.Connect = "DSN=192.168.1.2;DSQ=MARFICE;UID=SA;PWD=administrador"
    
    
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    Dim letras As String
    letras = NUMLET(adodc1("oc_nventa"))
    If adodc1("oc_ccodmon") = "01" Then
      letras = letras + " Nuevos Soles "
     Else
      letras = letras + " Dolares Americanos "
    End If
    CrystalReport1.formulas(0) = "@emp ='" & VGNemp & "'"
    CrystalReport1.formulas(1) = "@ruc ='" & VGRUCEMP & "'"
    CrystalReport1.formulas(2) = "@letras ='" & letras & "'"
    CrystalReport1.StoredProcParam(0) = VGcnx.DefaultDatabase
    CrystalReport1.StoredProcParam(1) = tipoorden
   CrystalReport1.StoredProcParam(2) = unum
   If CrystalReport1.Status <> 2 Then
      CrystalReport1.Action = 1
   End If

End Sub

Private Sub cmdNue_Click()
 Dim cSqlM As String, cSelM As ADODB.Recordset
    nT = 1
    OculObj06 False
    OculObj04 False
    OculObj02 True
    OculObj03 True
    Proceso True
    lblImp = "0.00": lblTot = "0.00": lblIgv = "0.00"
    lblDes = "0.00": lblCom = "0.00"
    Frame1.Visible = True
    Fradatos.Visible = True
    Fradatos.Enabled = True
    cmdGra.Enabled = True
    CmdSalir2.Cancel = True
End Sub

Private Sub cmdEdi_Click()
    If adodc1("oc_estadoorden") = "A" Then
        Mensaje = "La Orden de compra ha sido anulada, no se permitirá modificaciones"
        MsgBox Mensaje, vbExclamation, "Advertencia"
        cmdNue2.Enabled = False
        cmdEdi2.Enabled = False
        cmdEli2.Enabled = False
        cmdGra.Enabled = False
        OculObj06 False
        OculObj04 False
        OculObj02 True
        Mostrar adodc1("oc_cnumord")
        OculObj03 True
        Proceso True
        Fradatos.Enabled = False
    Else
        nT = 2
        OculObj06 False
        OculObj04 False
        OculObj02 True
        Mostrar adodc1("oc_cnumord")
        OculObj03 True
        Proceso True
        Fradatos.Enabled = True
        Frame1.Visible = True
        cmdEdi2.Enabled = True
        cmdEli2.Enabled = True
        cmdGra.Enabled = True
        
        txtEmi.SetFocus
        CmdSalir2.Cancel = True
    End If
End Sub

Private Sub cmdNue2_Click()
    With frmemisionOCdetalle
        .activado = False
        .CtrAyu_articulo.xclave = ""
        .txtCan = "0.00"
        .txtPUn = "0.00"
        .txtPDe = "0.00"
        .txtPIg = "19.00"
        .txtordfab = ""
        .lblFab.Caption = ""
        .txtCo1 = ""
        .activado = True
       .Show 1
        
        If Not .cancelado Then
           If .tipo = "S" Then
              .txtCan = 1
            End If
            
            If Flex1.Rows - 1 = 1 Then
                If Flex1.TextMatrix(1, 0) = "" Then
                    Flex1.AddItem Trim(.CtrAyu_articulo.xclave) & vbTab & .lblFab.Caption & vbTab & Trim(.CtrAyu_articulo.xnombre) & vbTab & _
                        .lblUni & vbTab & .txtCan & vbTab & IIf(.txtURe = "", .lblUni, _
                        .txtURe) & vbTab & IIf(.txtURe = "", .txtCan, .txtRef) & vbTab & _
                        .txtPUn & vbTab & Format(Val(.lblPNe) + Val(.lblDes), "0.00") & _
                        vbTab & .txtPDe & vbTab & .txtPIg & vbTab & Format(IIf(.txtURe = "", .txtCan, .txtRef) * _
                        (Val(.lblPNe) + Val(.lblDes)), "0.00") & vbTab & .txtordfab & vbTab & _
                        .txtCo1 & vbTab & .tipo, 1
                    Flex1.Rows = 2
                Else
                    Flex1.AddItem Trim(.CtrAyu_articulo.xclave) & vbTab & .lblFab & vbTab & Trim(.CtrAyu_articulo.xnombre) & vbTab & _
                        .lblUni & vbTab & .txtCan & vbTab & IIf(.txtURe = "", .lblUni, _
                        .txtURe) & vbTab & IIf(.txtURe = "", .txtCan, .txtRef) & vbTab & _
                        .txtPUn & vbTab & Format(Val(.lblPNe) + Val(.lblDes), "0.00") & _
                        vbTab & .txtPDe & vbTab & .txtPIg & vbTab & Format(IIf(.txtURe = "", .txtCan, .txtRef) * _
                        (Val(.lblPNe) + Val(.lblDes)), "0.00") & vbTab & .txtordfab & vbTab & _
                        .txtCo1 & vbTab & .tipo
                    Flex1.Row = Flex1.Rows - 1
                End If
            Else
                Flex1.AddItem Trim(.CtrAyu_articulo.xclave) & vbTab & .lblFab & vbTab & Trim(.CtrAyu_articulo.xnombre) & vbTab & _
                    .lblUni & vbTab & .txtCan & vbTab & IIf(.txtURe = "", .lblUni, _
                    .txtURe) & vbTab & IIf(.txtURe = "", .txtCan, .txtRef) & vbTab & _
                    .txtPUn & vbTab & Format(Val(.lblPNe) + Val(.lblDes), "0.00") & _
                    vbTab & .txtPDe & vbTab & .txtPIg & vbTab & Format(IIf(.txtURe = "", .txtCan, .txtRef) * _
                    (Val(.lblPNe) + Val(.lblDes)), "0.00") & vbTab & .txtordfab & vbTab & _
                    .txtCo1 & vbTab & .tipo
                Flex1.Row = Flex1.Rows - 1
            End If
            
            Calcula_Totales
            Estado_Items
            Flex1.SetFocus
           cmdNue2.SetFocus
        Else
            Flex1.SetFocus
            cmdNue2.SetFocus
        End If
    End With
End Sub

Private Sub CmdSalir_Click()
    Unload frmReferencia
    Unload frmemisionOCdetalle
    Unload Me
End Sub

Private Sub CmdSalir2_Click()
    Limpiar
    Vacia_FlexGrid
    Estado_Items
    Estado_Botones
    OculObj02 False
    OculObj03 False
    OculObj04 True
    OculObj06 True
    Proceso False
    Frame1.Visible = False
    If adodc1.RecordCount > 0 Then
        DataGrid1.SetFocus
    Else
        cmdNue.SetFocus
    End If
    CmdSalir.Cancel = True
End Sub
Public Function SupCadSQL(S As String) As String
 Dim Aux As String
 If Not IsNull(S) Then
     Aux = Replace(S, "'", "''")
 End If
 SupCadSQL = Aux
 
End Function

Private Sub CtrAyu_tipoorden_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    Dim unum As String
    Set VGvardllgen = New dllgeneral.dll_general
    unum = VGvardllgen.ESNULO(ColecCampos("tipoordennumeracion").Value, "")
    unum = Format(Val(unum) + 1, "00000000000")
    lblNum = unum
    
End Sub


Private Sub CtrAyu_Proveedor_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    Set VGvardllgen = New dllgeneral.dll_general
    lblRuc.text = VGvardllgen.ESNULO(ColecCampos("prvcruc").Value, "")
End Sub
Private Sub CtrAyu_Proveedor_AlNoDevolverNada()
    lblRuc.text = ""
End Sub

Private Sub Form_Load()
    Formato_FlexGrid
    Call Ctrayu_tipoorden.Conexion(VGcnx): Ctrayu_tipoorden.Filtro = "(tipoordencodigo <>'00') "
    Call CtrAyu_solicitante.Conexion(cConexCom)
    
    OculObj02 False
    OculObj03 False
    OculObj04 True
    OculObj06 True
    
    txtEmi.Value = Date
    unum = ""
    Abre_Tabla_OCs
    Estado_Botones
    Frame1.Visible = False
    Load frmemisionOCdetalle
End Sub
Private Sub Reales_Positivos(k As Integer, t As TextBox)
Dim t1 As String
    k = Asc(UCase(Chr(k)))
    If k = 8 Then Exit Sub
    If k <> 45 And k <> 44 And k <> 32 And k <> 69 And k <> 43 Then
        t1 = Left(t, t.SelStart)
        t1 = t1 & Chr(k) & Right(t, Len(t) - Len(t1))
        If IsNumeric(t1) Then Exit Sub
    End If
    k = 0
    
End Sub

Public Function Existe(tipo As Integer, Cod As String, Tabla As String, Campo As String, Fecha As Boolean, Optional Cod2 As String, Optional cCampo2 As String, Optional Cod3 As String, Optional cCampo3 As String, Optional Cod4 As Boolean, Optional cCampo4 As String, Optional Cod5 As String, Optional cCampo5 As String) As Boolean
Dim cSel1 As ADODB.Recordset, cSL As String
Set cSel1 = New ADODB.Recordset

 If Fecha Then
        cSL = "Select * from " & Tabla & "  Where " & Campo & " =  '" & Cod & "'"
 Else
       If UCase(Tabla) = "PUNTO_VENTA" Then
                cSL = "Select * from " & Tabla & "  Where " & Campo & " =  '" & Cod & "'"
       Else
                cSL = "Select * from " & Tabla & "  Where " & Campo & " =  '" & Cod & "'"
       End If
       If Trim(Cod2) <> "" Then
            cSL = cSL & " And  " & cCampo2 & " =  '" & SupCadSQL(Cod2) & "'"
       End If
       If Trim(Cod3) <> "" Then
            cSL = cSL & " And  " & cCampo3 & " =  '" & SupCadSQL(Cod3) & "'"
       End If
       If Trim(cCampo4) <> "" Then
            If Cod4 = True Then
                cSL = cSL & " And  " & cCampo4
            Else
                cSL = cSL & " And  " & Not cCampo4
            End If
        End If
        If Trim(Cod5) <> "" Then
            cSL = cSL & " And  " & cCampo5 & " =  '" & Cod5 & "'"
        End If
 End If
 
Select Case tipo
Case 1 'Bd. Comun
            cSel1.Open cSL, VGcnx, adOpenStatic
Case 2 'Bd. Config
            cSel1.Open cSL, VGcnx, adOpenStatic
Case 3 'Bd. Contab
            cSel1.Open cSL, cConexCont, adOpenStatic
End Select

If cSel1.RecordCount > 0 Then
     Existe = True
Else
     Existe = False
End If
'csel1.Close
End Function

Sub Limpiar()

txtNSol = ""
txtCot = ""
Ctrayu_tipoorden.xclave = ""
Ctrayu_tipoorden.xnombre = ""
CtrAyu_Proveedor.xclave = ""
CtrAyu_Proveedor.xnombre = ""
CtrAyu_pago.xclave = ""
CtrAyu_pago.xnombre = ""
CtrAyu_solicitante.xclave = ""
CtrAyu_solicitante.xnombre = ""
CtrAyu_moneda.xclave = ""
CtrAyu_moneda.xnombre = ""
txtEntE = "": txtObs = ""
End Sub

Sub Mostrar(cC1 As String)
    Dim cSqlM As String, cSelM As ADODB.Recordset
    Dim k As Integer, I As Integer, vd As String
    Dim vpu As Single, txtPro As String
    Dim txtSol As String
    
    lblNum = cC1
    txtEmi = adodc1("oc_dfecdoc")
    txtEntE = adodc1("oc_centreg")
    CtrAyu_solicitante.xclave = adodc1("oc_csolict")
    txtSol = CtrAyu_solicitante.xclave
    CtrAyu_solicitante.xnombre = Devolver_Dato(1, txtSol, "co_solicitantes", "solicitantecodigo", False, "solicitantenombre")
    txtObs = adodc1("oc_cobserv")
    Ctrayu_tipoorden.xclave = adodc1("tipoordencodigo")
    
    cSqlM = "SELECT * FROM co_detordcompra WHERE oc_cnumord='" & cC1 & "' ORDER BY oc_citem"
    Set cSelM = New ADODB.Recordset
    
    cSelM.Open cSqlM, VGcnx, adOpenStatic
    cSelM.MoveFirst
    
    k = 0
    Do While Not cSelM.EOF
        k = k + 1
        If k = 1 Then
            vpu = IIf(cSelM("oc_nfactor") = 0, 1, cSelM("oc_nfactor") / cSelM("oc_ncantid"))
            Flex1.AddItem cSelM("oc_ccodigo") & vbTab & cSelM("oc_ccodref") & vbTab & _
                cSelM("oc_cdesref") & vbTab & cSelM("oc_cunidad") & vbTab & _
                Format(cSelM("oc_ncantid"), "0.00") & vbTab & IIf(cSelM("oc_cuniref") = "", _
                cSelM("oc_cunidad"), cSelM("oc_cuniref")) & vbTab & _
                IIf(cSelM("oc_cuniref") = "", Format(cSelM("oc_ncantid"), "0.00"), _
                Format(cSelM("oc_nfactor"), "0.00")) & vbTab & Format(cSelM("oc_npreuni"), _
                "0.00") & vbTab & Format(cSelM("oc_npreuni"), "0.00") & vbTab & _
                Format(cSelM("oc_ndscpor"), "0.00") & vbTab & _
                Format(cSelM("oc_nigvpor"), "0.00") & vbTab & _
                Format(cSelM("oc_ntotven"), "0.00") & vbTab & cSelM("ord_fabnum") & vbTab & _
                cSelM("oc_ccomen1") & vbTab & cSelM("tipoarticulocodigo"), 1
            Flex1.Rows = 2
        Else
            vpu = IIf(cSelM("oc_nfactor") = 0, 1, cSelM("oc_nfactor") / cSelM("oc_ncantid"))
            Flex1.AddItem cSelM("oc_ccodigo") & vbTab & cSelM("oc_ccodref") & vbTab & _
                cSelM("oc_cdesref") & vbTab & cSelM("oc_cunidad") & vbTab & _
                Format(cSelM("oc_ncantid"), "0.00") & vbTab & IIf(cSelM("oc_cuniref") = "", _
                cSelM("oc_cunidad"), cSelM("oc_cuniref")) & vbTab & _
                IIf(cSelM("oc_cuniref") = "", Format(cSelM("oc_ncantid"), "0.00"), _
                Format(cSelM("oc_nfactor"), "0.00")) & vbTab & Format(cSelM("oc_npreuni"), _
                "0.00") & vbTab & Format(cSelM("oc_npreuni"), "0.00") & vbTab & _
                Format(cSelM("oc_ndscpor"), "0.00") & vbTab & _
                Format(cSelM("oc_nigvpor"), "0.00") & vbTab & _
                Format(cSelM("oc_ntotven"), "0.00") & vbTab & cSelM("ord_fabnum") & vbTab & _
                cSelM("oc_ccomen1") & vbTab & cSelM("tipoarticulocodigo")
        End If
        cSelM.MoveNext
    Loop
    cSelM.Close
    Calcula_Totales
End Sub

Sub Estado_Botones()
    If adodc1.RecordCount > 0 Then
      '  cmdEdi.Enabled = True
      '  CmdEli.Enabled = True
        cmdImp.Enabled = True
    Else
       ' cmdEdi.Enabled = False
      '  CmdEli.Enabled = False
        cmdImp.Enabled = False
    End If
End Sub



Private Sub txtCot_GotFocus()
    Enfoque txtCot
End Sub

Private Sub txtCot_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtEntE.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub


Private Sub txtEmi_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  SendKeys "{TAB}"
End If
End Sub

Private Sub txtEmi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not ValidFecha(txtEmi) Then
            Mensaje = "Fecha No Válida"
            MsgBox Mensaje, vbExclamation, "Error"
            txtEmi.SetFocus
        Else
            txtEnt.SetFocus
        End If
    End If
End Sub

Function ValidFecha(vText As String) As String
Dim cTxtNew As String, ncnt As Integer
Dim cTxt As String, cTxtDig As String

cTxtDig = "": cTxtNew = ""
For ncnt = 1 To Len(vText)
      cTxt = Mid(vText, ncnt, 1)
      If cTxt = "/" Then
         cTxtNew = cTxtNew & Str(Val(cTxtDig)) & "/"
         cTxtDig = ""
      Else
         If cTxt <> "_" Then cTxtDig = cTxtDig & cTxt
      End If
Next
If cTxtDig <> "" Then cTxtNew = cTxtNew & Str(Val(cTxtDig))

If IsDate(cTxtNew) Then
   ValidFecha = Format(CDate(cTxtNew), "dd/mm/yyyy")
End If
End Function


Private Sub txtEnt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  SendKeys "{TAB}"
End If
End Sub

Private Sub txtEnt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not ValidFecha(txtEnt) Then
            Mensaje = "Fecha No Válida"
            MsgBox Mensaje, vbExclamation, "Error"
            txtEnt.SetFocus
        End If
    End If
End Sub

Private Sub txtEntE_GotFocus()
    Enfoque txtEntE
End Sub


Private Sub txtObs_GotFocus()
    Enfoque txtObs
End Sub

Private Sub txtObs_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmdEli2.Enabled Then
            Flex1.SetFocus
        Else
            cmdNue2.SetFocus
        End If
    End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub Enfoque(OBJ As Object)
  OBJ.SelStart = 0
  OBJ.SelLength = Len(OBJ)
End Sub

Sub Proceso(Estado As Boolean)
    Flex1.Visible = Estado
    cmdNue2.Visible = Estado
    cmdEdi2.Visible = Estado
    cmdEli2.Visible = Estado
End Sub

Sub Formato_FlexGrid()
    Flex1.ColWidth(0) = 1100
    Flex1.ColWidth(1) = 0
    Flex1.ColWidth(2) = 2800
    Flex1.ColWidth(3) = 0
    Flex1.ColWidth(4) = 0
    Flex1.ColWidth(5) = 450
    Flex1.ColWidth(6) = 900
    Flex1.ColWidth(7) = 0
    Flex1.ColWidth(8) = 1200
    Flex1.ColWidth(9) = 700
    Flex1.ColWidth(10) = 0
    Flex1.ColWidth(11) = 1200
    Flex1.ColWidth(12) = 0
    Flex1.ColWidth(13) = 0
    Flex1.ColWidth(14) = 5
    Flex1.ScrollBars = flexScrollBarHorizontal
End Sub

Sub Estado_Items()
    If Flex1.Rows - 1 = 1 Then
        If Flex1.TextMatrix(1, 0) = "" Then
            cmdEdi2.Enabled = False
            cmdEli2.Enabled = False
            cmdNue2.Enabled = True
            cmdNue2.SetFocus
        Else
            cmdEdi2.Enabled = True
            cmdEli2.Enabled = True
        End If
    Else
        cmdEdi2.Enabled = True
        cmdEli2.Enabled = True
    End If
End Sub

Sub Vacia_FlexGrid()
    Dim I As Integer
    
    Do While Flex1.Rows - 1 > 1
        Flex1.RemoveItem 1
    Loop
    
    For I = 0 To 14
        Flex1.TextMatrix(1, I) = ""
    Next
End Sub

Sub Calcula_Totales()
    Dim I As Integer
    Dim tV As Single, valor As Single
    Dim tD As Single, vDesc As Single
    Dim tI As Single, vIgv As Single
    
    With Flex1
        For I = 1 To Flex1.Rows - 1
            tV = Val(.TextMatrix(I, 11))
            valor = valor + tV
            tD = tV * Val(.TextMatrix(I, 9)) / 100
            vDesc = vDesc + tD
            tI = (tV - tD) * Val(.TextMatrix(I, 10)) / 100
            vIgv = vIgv + tI
        Next
    End With
    
    lblImp = Format(valor, "##,##0.0000")
    lblDes = Format(vDesc, "##,##0.0000")
    lblTot = Format(valor - vDesc, "#,##0.0000")
    lblIgv = Format(vIgv, "#,##0.00")
    lblCom = Format((valor - vDesc) + vIgv, "#,##0.00")
End Sub

Function Tiene_Entregas() As Boolean
    Dim Adodc2 As ADODB.Recordset
    
    Set Adodc2 = New ADODB.Recordset
    
    Adodc2.Open "SELECT * FROM co_detordcompra WHERE oc_cnumord='" & lblNum & "' AND oc_ccodigo='" & _
        Flex1.TextMatrix(Flex1.Row, 0) & "' AND oc_ncanten>0", VGcnx, adOpenStatic
    Tiene_Entregas = False
    If Adodc2.RecordCount > 0 Then Tiene_Entregas = True
End Function


