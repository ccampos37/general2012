VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmRegPlantilladeKits 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Plantillas de Kits"
   ClientHeight    =   5640
   ClientLeft      =   1080
   ClientTop       =   2250
   ClientWidth     =   9030
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   9030
   Begin VB.CommandButton CmdEliKits 
      Caption         =   "&Eliminar"
      Height          =   735
      Left            =   5040
      Picture         =   "FrmRegPlantilladeKits.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   4680
      Width           =   855
   End
   Begin VB.Frame Framedatos 
      Height          =   1575
      Left            =   960
      TabIndex        =   18
      Top             =   2280
      Width           =   4695
      Begin VB.CommandButton Command2 
         Caption         =   "Salir"
         Height          =   495
         Left            =   3240
         TabIndex        =   24
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   495
         Left            =   3240
         TabIndex        =   23
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Textporcentaje 
         Height          =   375
         Left            =   1320
         TabIndex        =   22
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Textunid 
         Height          =   375
         Left            =   1320
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "% del Costo"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Adicionar"
      Height          =   735
      Left            =   120
      Picture         =   "FrmRegPlantilladeKits.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4680
      Visible         =   0   'False
      Width           =   855
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   6552
      Top             =   4980
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton command5 
      Caption         =   "&Reporte"
      Height          =   735
      Left            =   4080
      Picture         =   "FrmRegPlantilladeKits.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Height          =   735
      Left            =   2280
      Picture         =   "FrmRegPlantilladeKits.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton CmdModi 
      Caption         =   "&Modificar"
      Height          =   735
      Left            =   2040
      Picture         =   "FrmRegPlantilladeKits.frx":0FD0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton CmdEli 
      Caption         =   "&Eliminar"
      Height          =   735
      Left            =   3120
      Picture         =   "FrmRegPlantilladeKits.frx":1412
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton CmdIng 
      Caption         =   "&Crear"
      Height          =   735
      Left            =   1080
      Picture         =   "FrmRegPlantilladeKits.frx":1854
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   675
      Left            =   7668
      Picture         =   "FrmRegPlantilladeKits.frx":1C96
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4608
      Width           =   775
   End
   Begin VB.Frame Frame5 
      Height          =   4476
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   8448
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3510
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Visible         =   0   'False
         Width           =   6150
         _ExtentX        =   10848
         _ExtentY        =   6191
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   19
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
            Name            =   "Arial"
            Size            =   7.5
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
               LCID            =   3082
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
               LCID            =   3082
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
      Begin VB.ComboBox CmbOrden 
         Height          =   288
         ItemData        =   "FrmRegPlantilladeKits.frx":20D8
         Left            =   4716
         List            =   "FrmRegPlantilladeKits.frx":20E2
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   252
         Width           =   1575
      End
      Begin VB.TextBox Txtarticulo 
         Height          =   300
         Left            =   1080
         TabIndex        =   6
         Text            =   "TxFiltro"
         Top             =   240
         Width           =   2655
      End
      Begin VB.Image Image1 
         Height          =   4092
         Left            =   6264
         Picture         =   "FrmRegPlantilladeKits.frx":20FF
         Stretch         =   -1  'True
         Top             =   216
         Width           =   2112
      End
      Begin VB.Label Label33 
         Caption         =   "Orden :"
         Height          =   252
         Left            =   3924
         TabIndex        =   9
         Top             =   288
         Width           =   852
      End
      Begin VB.Label Label32 
         Caption         =   "Buscar  :"
         Height          =   285
         Left            =   150
         TabIndex        =   8
         Top             =   255
         Width           =   930
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4488
      Left            =   72
      TabIndex        =   10
      Top             =   36
      Width           =   8820
      Begin VB.Frame Framearticulo 
         Caption         =   "Adiciona producto"
         Height          =   1575
         Left            =   840
         TabIndex        =   26
         Top             =   1920
         Visible         =   0   'False
         Width           =   5415
         Begin VB.CommandButton Command4 
            Caption         =   "Salir"
            Height          =   495
            Left            =   2760
            TabIndex        =   28
            Top             =   840
            Width           =   1455
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Aceptar"
            Height          =   495
            Left            =   600
            TabIndex        =   27
            Top             =   840
            Width           =   1455
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaArticulo 
            Height          =   375
            Left            =   240
            TabIndex        =   29
            Top             =   360
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   661
            XcodMaxLongitud =   0
            xcodwith        =   1200
            NomTabla        =   "maeart"
            ListaCampos     =   "acodigo(1),adescri(1)"
            XcodCampo       =   "acodigo"
            XListCampo      =   "adescri"
            ListaCamposDescrip=   "Codigo, Descripcion"
            ListaCamposText =   "acodigo,adescri"
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGridkit 
         Height          =   3015
         Left            =   360
         TabIndex        =   25
         Top             =   1200
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   5318
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Kits"
         Columns(0).DataField=   "codkit"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Producto"
         Columns(1).DataField=   "codart"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Cantidad"
         Columns(2).DataField=   "canart"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "porcentaje"
         Columns(3).DataField=   "porcentajevalor"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   4
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=4"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1984"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1905"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2249"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2170"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=2117"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2037"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=1958"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=1879"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   14215660
         RowDividerColor =   14215660
         RowSubDividerColor=   14215660
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(52)  =   "Named:id=33:Normal"
         _StyleDefs(53)  =   ":id=33,.parent=0"
         _StyleDefs(54)  =   "Named:id=34:Heading"
         _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(56)  =   ":id=34,.wraptext=-1"
         _StyleDefs(57)  =   "Named:id=35:Footing"
         _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(59)  =   "Named:id=36:Selected"
         _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(61)  =   "Named:id=37:Caption"
         _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(63)  =   "Named:id=38:HighlightRow"
         _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(65)  =   "Named:id=39:EvenRow"
         _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(67)  =   "Named:id=40:OddRow"
         _StyleDefs(68)  =   ":id=40,.parent=33"
         _StyleDefs(69)  =   "Named:id=41:RecordSelector"
         _StyleDefs(70)  =   ":id=41,.parent=34"
         _StyleDefs(71)  =   "Named:id=42:FilterBar"
         _StyleDefs(72)  =   ":id=42,.parent=33"
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1368
         MaxLength       =   20
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   288
         Width           =   1695
      End
      Begin VB.Image Image2 
         Height          =   4230
         Left            =   6510
         Picture         =   "FrmRegPlantilladeKits.frx":4EC2
         Stretch         =   -1  'True
         Top             =   180
         Width           =   2250
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1365
         TabIndex        =   14
         Top             =   645
         Width           =   4815
      End
      Begin VB.Label Label2 
         Caption         =   "Descripción     :"
         Height          =   372
         Left            =   144
         TabIndex        =   13
         Top             =   684
         Width           =   1812
      End
      Begin VB.Label Label1 
         Caption         =   "Código             :"
         Height          =   252
         Left            =   132
         TabIndex        =   11
         Top             =   288
         Width           =   1332
      End
   End
End
Attribute VB_Name = "FrmRegPlantilladeKits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adodc1 As ADODB.Recordset
Dim cSel1 As ADODB.Recordset
Dim SelM As ADODB.Recordset
Dim cSql1 As String, CSQL2 As String, cCod As String
Dim nT As Integer       'Ingreso,Modificación
Dim nCom As Integer, nTra As Integer, nCursor As Integer

Private Sub OculObj03(ntipo As Boolean) ' Todos los datos
Frame1.Visible = ntipo
End Sub
Private Sub OculObj04(ntipo As Boolean) ' Botones principales
CmdIng.Visible = ntipo
CmdModi.Visible = ntipo
CmdEli.Visible = ntipo
CmdSalir.Visible = ntipo
End Sub
Private Sub OculObj05(ntipo As Boolean)  'Orden y Filtro
Frame5.Visible = ntipo
Label32.Visible = ntipo
TxtArticulo.Visible = ntipo
Label33.Visible = ntipo
cmbOrden.Visible = ntipo
End Sub
Private Sub OculObj06(ntipo As Boolean)  'Datagrid
DataGrid1.Visible = ntipo
End Sub
Private Sub oculobj07(ntipo As Boolean)
If ntipo Then
   CmdEli.Visible = IIf(ntipo = True, False, ntipo)
  Else
  CmdEli.Visible = IIf(ntipo = False, True, ntipo)
End If
CmdEliKits.Visible = ntipo
End Sub
Private Sub CmbOrden_Click()             ' Ordenar por
Dim cD As String
nCom = cmbOrden.ListIndex
Set adodc1 = New ADODB.Recordset
cD = "SELECT DISTINCT(ACODIGO),ADESCRI,AUNIDAD FROM KITS,MAEART WHERE ACODIGO = CODKIT"

Select Case nCom
Case 0
            cD = cD & " ORDER BY ACODIGO"
Case 1
            cD = cD & " ORDER BY ADESCRI"
End Select
adodc1.Open cD, VGCNx, adOpenStatic
TxtArticulo = ""
Set DataGrid1.DataSource = adodc1
Set_Data
If DataGrid1.Visible Then DataGrid1.SetFocus
End Sub

Private Sub CmdEli_Click()              ' Elimina
Dim CSQL2 As String, nN As Integer
Dim I As Integer
On Error GoTo EliErr
  If MsgBox("Desea Eliminar el Registro?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        cCod = adodc1("ACODIGO")
        cSql1 = "Select STCODIGO from STKART where STCODIGO = '" & cCod & "' and STSKDIS > 0"
        Set cSel1 = New ADODB.Recordset
        cSel1.Open cSql1, VGCNx, adOpenStatic
        If cSel1.RecordCount > 0 Then          ' vGAlmacen
           MsgBox "El artículo tiene saldos con Cantidad Disponible mayor a Cero, no se puede Eliminar", vbInformation, "Mensaje"
           cSel1.Close
           Exit Sub
        End If
        cSel1.Close
        cSql1 = "Delete from KITS where CODkit = '" & cCod & "' and codart='" & SelM("codart") & "'"
        nTra = 1
        VGCNx.BeginTrans
        VGCNx.Execute cSql1
        VGCNx.CommitTrans
        
        nTra = 0
        
        adodc1.Requery
        If nN <> 0 Then adodc1.AbsolutePosition = nN
    End If
Mostrar (Text1)

Exit Sub
EliErr:
    MsgBox Err.Description
    If nTra = 1 Then VGCNx.RollbackTrans
    Exit Sub
    Resume
End Sub

Private Sub CmdEliKits_Click()
Dim CSQL2 As String, nN As Integer
Dim I As Integer
On Error GoTo EliErr
  If MsgBox("Desea Eliminar el Registro?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        cCod = adodc1("ACODIGO")
        cSql1 = "Select STCODIGO from STKART where STCODIGO = '" & cCod & "' and STSKDIS > 0"
        Set cSel1 = New ADODB.Recordset
        cSel1.Open cSql1, VGCNx, adOpenStatic
        If cSel1.RecordCount > 0 Then          ' vGAlmacen
           MsgBox "El artículo tiene saldos con Cantidad Disponible mayor a Cero, no se puede Eliminar", vbInformation, "Mensaje"
           cSel1.Close
           Exit Sub
        End If
        cSel1.Close
        cSql1 = "Delete from KITS where CODkit = '" & cCod & "' and codart='" & SelM("codart") & "'"
        nTra = 1
        VGCNx.BeginTrans
        VGCNx.Execute cSql1
        VGCNx.CommitTrans
        
        nTra = 0
        
        adodc1.Requery
        If nN <> 0 Then adodc1.AbsolutePosition = nN
    End If
Mostrar (Text1)

Exit Sub
EliErr:
    MsgBox Err.Description
    If nTra = 1 Then VGCNx.RollbackTrans
    Exit Sub
    Resume

End Sub

Private Sub CmdGrabar_Click()           ' Grabar
Dim I As Integer
Dim porc As Double
On Error GoTo GrabErr

If Trim(Text1) = "" Then
    MsgBox "Ingrese Código", vbInformation, "Mensaje"
    Text1.SetFocus: Exit Sub
End If

If MsgBox("Es correcta la Información", vbQuestion + vbOKCancel, "Mensaje") = vbOK Then
    If nT = 1 Then      'Ingreso
        If codigo(Text1) = False Then
            If Existe(1, Text1, "kits", "codkit", False) Then
                MsgBox "Código de Artículo ya existe", vbInformation, "Mensaje"
                Text1.SetFocus: Exit Sub
            End If
        End If
        SelM.MoveFirst
        porc = 0
        Do While Not SelM.EOF()
           porc = porc + SelM("porcentajevalor")
           SelM.MoveNext
        Loop
        If porc <> 100 Then
           MsgBox (" El total de porcentaje debe ser 100% ")
       '     Exit Sub
        End If
        SelM.MoveFirst
        Do While Not SelM.EOF()
           CSQL2 = "Insert Into Kits(CODART,CODKIT,CANART,porcentajevalor) Values " & _
                        "('" & SelM("codart") & "','" & SelM("codkit") & "'," & SelM("canart") & ",'" & SelM("porcentajevalor") & "')"
                
          VGCNx.BeginTrans
          VGCNx.Execute CSQL2
          VGCNx.CommitTrans
          SelM.MoveNext
        Loop
    ElseIf nT = 2 Then     'Modificar             Trim(Mid(Combo1.text, 1, 1))
        SelM.MoveFirst
        porc = 0
        Do While Not SelM.EOF()
           porc = porc + SelM("porcentajevalor")
           SelM.MoveNext
        Loop
        If porc <> 100 Then
           MsgBox (" El total de p[roceentaje debe ser 100% ")
           Exit Sub
        End If
        
        SelM.MoveFirst
        Do While Not SelM.EOF()
           CSQL2 = "update Kits set CANART='" & SelM("canart") & "',porcentajevalor=" & SelM("porcentajevalor") & ""
           CSQL2 = CSQL2 & " where codkit='" & SelM("codkit") & "' and codart='" & SelM("codart") & "'"
           VGCNx.BeginTrans
           VGCNx.Execute CSQL2
           VGCNx.CommitTrans
           SelM.MoveNext
        Loop
        
      
          nTra = 1

    End If
    adodc1.Requery
    Set_Data
    adodc1.Find "ACODIGO = '" & Text1 & "'"
End If

If nT = 1 Then
    Limpiar
    Text1.SetFocus
ElseIf nT = 2 Then
    CmdSalir_Click
End If

Exit Sub
GrabErr:
    MsgBox Err.Description
    If nTra = 1 Then VGCNx.RollbackTrans
    Exit Sub
    Resume
End Sub

Private Sub CmdIng_Click()      'Ingresar
nT = 1
Me.Caption = "Ingreso de Registro de Kits"
OculObj04 (False)
OculObj05 (False)
OculObj06 (False)
OculObj02 (True)
OculObj03 (True)
Limpiar
Text1.Enabled = True
Text1.SetFocus
End Sub

Private Sub CmdModi_Click()     'Modificar
If DataGrid1.Row = -1 Then Exit Sub

If ClsTock.ExisteEnStockAlmacenes(DataGrid1.Columns(0).text, VGCNx) Then
   MsgBox "Este Kit tiene Stock en uno de los Almacenes" & Chr(10) & "Desarmelo para poder Modificarlo..!", vbInformation, "Aviso.....!"
 '  Exit Sub
End If
If adodc1.RecordCount > 0 Then
    nT = 2
    Me.Caption = "Modificación de Registros de Kits"
    OculObj04 (False)
    OculObj05 (False)
    OculObj06 (False)
    OculObj02 (True)
    OculObj03 (True)
    oculobj07 (True)
    Limpiar
    cCod = adodc1("ACODIGO")
    Text1 = cCod
    Text1.Enabled = False
    Cmdgrabar.Visible = True
    Mostrar (cCod)
Else
    MsgBox "No existen registros", vbInformation, "Mensaje"
End If
End Sub

Private Sub CmdSalir_Click()    'Salida principal del formulario
If nT = 1 Or nT = 2 Then
    Me.Caption = "Actualiza Registro de Kits"
    OculObj02 (False)
    OculObj03 (False)
    OculObj04 (True)
    OculObj05 (True)
    OculObj06 (True)
    oculobj07 (False)
    InhabObj (True)
    Cmdgrabar.Enabled = True
    Cmdgrabar.Visible = False
    nT = 0
    DataGrid1.SetFocus
Else
    Unload Me
End If
End Sub

Private Sub Command1_Click()
 SelM("canart") = Textunid
 SelM("porcentajevalor") = Textporcentaje
 SelM.Update
 TDBGridkit.Refresh
 Framedatos.Visible = False
 End Sub

Private Sub Command2_Click()
Framedatos.Visible = False
End Sub

Private Sub Command3_Click()
SelM.AddNew
SelM("codkit") = Text1
SelM("codart") = Ctr_AyudaArticulo.xclave
Framearticulo.Visible = False
Framedatos.Visible = True
End Sub

Private Sub Command4_Click()
Framearticulo.Visible = False
End Sub

Private Sub command5_Click()
Dim arrform(2), arrparm(1) As Variant
On Error GoTo Imprime

arrparm(0) = VGParamSistem.BDEmpresa

arrform(0) = VGParamSistem.BDEmpresa
arrform(1) = VGparametros.RucEmpresa

Call ImpresionRptProc("invkits.rpt", arrform, arrparm, " ", "Control de Kits ")
Screen.MousePointer = 1

Exit Sub
Imprime:
Screen.MousePointer = 1
MsgBox "Err.Description", vbCritical, "Sistemas"


End Sub

Private Sub Command6_Click()
 VGRegEnt = 1: VGForm1 = 4
 Framearticulo.Visible = True
 
End Sub

Private Sub Form_Activate()
TxtArticulo = ""
cmbOrden.ListIndex = 0
If DataGrid1.Visible Then DataGrid1.SetFocus
Call Ctr_AyudaArticulo.conexion(VGCNx)
End Sub

Private Sub Form_Load()
central Me         'Centra Formulario
' Init_ControlDataGrid DataGrid1

Limpiar
OculObj03 (False)
OculObj04 (True)
OculObj05 (True)
OculObj06 (True)
Set adodc1 = New ADODB.Recordset
adodc1.Open "SELECT distinct(ACODIGO),ADESCRI,AUNIDAD FROM KITS,MAEART WHERE ACODIGO = CODKIT  ORDER BY ACODIGO", VGCNx, adOpenStatic, adLockReadOnly
adodc1.Requery
Framedatos.Visible = False
Set DataGrid1.DataSource = adodc1
Set_Data
DataGrid1.Refresh
End Sub
Private Sub Limpiar()       'Limpia variables
Text1 = "": Label3 = "":
Set_Flex
End Sub

Private Sub TDBGridkit_Click()
 Textunid = ESNULO(SelM("canart"), 0)
 Textporcentaje = ESNULO(SelM("porcentajevalor"), 0)
Framedatos.Visible = True
End Sub



Private Sub txtarticulo_Change()
If adodc1.RecordCount > 0 Then
    If Trim(TxtArticulo) <> "" Then
        nCursor = adodc1.Bookmark
        adodc1.AbsolutePosition = 1
        adodc1.MoveFirst
        
        If cmbOrden.ListIndex = 0 Then
            adodc1.Find "ACODIGO like '" & Trim(UCase(TxtArticulo)) & "*'"
        ElseIf cmbOrden.ListIndex = 1 Then
            adodc1.Find "ADESCRI like '" & Trim(UCase(TxtArticulo)) & "*'"
        End If
        If adodc1.EOF Then adodc1.AbsolutePosition = nCursor
    End If
End If
End Sub

Private Sub Text1_GotFocus()
Enfoque Text1
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(Text1) = "" Then
        MsgBox "Ingrese el Código del articulo", vbInformation, "Mensaje"
        Text1.SetFocus
    Else
        If codigo(Text1) = False Then
            If Existe(1, Text1, "Kits", "CodKit", False) = False Then
            Else
                MsgBox "El Código ya existe", vbInformation, "Mensaje"
                Text1.SetFocus
            End If
        Else
            MsgBox "El Código no existe,Tiene que registrarlo en la Tabla de articulo", vbInformation, "Mensaje"
            Text1.SetFocus
        End If
        
        If ClsTock.ArticuloConMovimiento(Text1, VGCNx) Then
           MsgBox "Este Articulo no Puede ser un Kit, ya tiene Movimientos", vbInformation, "Error al Seleccionar Articulo"
           Text1 = ""
           Label3 = ""
           Command6.Enabled = False
        Else
           Label3 = Devolver_Dato(1, Text1, "MAEART", "ACODIGO", False, "ADESCRI")
           Command6.Enabled = True
        End If
        
    End If
    Mostrar (Trim(Text1))
Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub

Private Sub Mostrar(cC1 As String) 'Muestra los datos
Dim cSqlM As String
If Trim(cC1) = "" Then
    MsgBox "No hay registros para mostrar", vbInformation, "Mensaje"
    Exit Sub
End If
cSqlM = "Select CodKit,CodArt,CanArt,porcentajevalor From kits,MaeArt Where codkit = Acodigo AND codkit = '" & cC1 & "' "
Set SelM = New ADODB.Recordset
SelM.Open (cSqlM), VGCNx, adOpenDynamic, adLockBatchOptimistic

TDBGridkit.DataSource = SelM
End Sub

Private Sub InhabObj(ntipo As Boolean) ' Habilita e Inhabilita los objetos
Text1.Enabled = ntipo
End Sub

Private Sub Set_Data()
DataGrid1.Columns(0).Locked = True
DataGrid1.Columns(0).WrapText = True
DataGrid1.Columns(0).Caption = "   CODIGO"
DataGrid1.Columns(1).Caption = "       DESCRIPCION"
DataGrid1.Columns(2).Caption = "   UNIDAD"
DataGrid1.Columns(0).Width = 950
DataGrid1.Columns(1).Width = 4000
DataGrid1.Columns(2).Width = 900
End Sub
Private Sub OculObj02(ntipo As Boolean)  'Grabar y salir
Cmdgrabar.Visible = ntipo
CmdSalir.Visible = ntipo
Command6.Visible = ntipo
End Sub

Private Sub Set_Flex()
End Sub


