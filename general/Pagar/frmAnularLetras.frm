VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmAnularLetras 
   Caption         =   "Anulación de Letras"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7605
   ScaleWidth      =   10275
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   7275
      Left            =   75
      TabIndex        =   5
      Top             =   75
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   12832
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "LETRAS"
      TabPicture(0)   =   "frmAnularLetras.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame Frame1 
         Height          =   4095
         Left            =   150
         TabIndex        =   22
         Top             =   2910
         Width           =   9735
         Begin TrueOleDBGrid70.TDBGrid TDBGDetalleDoc 
            Height          =   2085
            Left            =   60
            TabIndex        =   28
            Top             =   465
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   3678
            _LayoutType     =   0
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
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
            DeadAreaBackColor=   12632256
            RowDividerColor =   12632256
            RowSubDividerColor=   12632256
            DirectionAfterEnter=   1
            MaxRows         =   250000
            ViewColumnCaptionWidth=   0
            ViewColumnWidth =   0
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=176,.bold=0,.fontsize=825,.italic=0"
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
            _StyleDefs(44)  =   "Named:id=33:Normal"
            _StyleDefs(45)  =   ":id=33,.parent=0"
            _StyleDefs(46)  =   "Named:id=34:Heading"
            _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(48)  =   ":id=34,.wraptext=-1"
            _StyleDefs(49)  =   "Named:id=35:Footing"
            _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(51)  =   "Named:id=36:Selected"
            _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(53)  =   "Named:id=37:Caption"
            _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(55)  =   "Named:id=38:HighlightRow"
            _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(57)  =   "Named:id=39:EvenRow"
            _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(59)  =   "Named:id=40:OddRow"
            _StyleDefs(60)  =   ":id=40,.parent=33"
            _StyleDefs(61)  =   "Named:id=41:RecordSelector"
            _StyleDefs(62)  =   ":id=41,.parent=34"
            _StyleDefs(63)  =   "Named:id=42:FilterBar"
            _StyleDefs(64)  =   ":id=42,.parent=33"
         End
         Begin VB.Frame Frame3 
            Height          =   585
            Left            =   60
            TabIndex        =   24
            Top             =   2490
            Width           =   9600
            Begin VB.TextBox Text1 
               BackColor       =   &H00C0FFFF&
               Enabled         =   0   'False
               ForeColor       =   &H00404040&
               Height          =   285
               Index           =   3
               Left            =   8100
               MaxLength       =   10
               TabIndex        =   25
               Top             =   180
               Width           =   1425
            End
            Begin VB.Label Label2 
               Caption         =   "TOTAL"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   3
               Left            =   7380
               TabIndex        =   26
               Top             =   240
               Width           =   675
            End
         End
         Begin VB.Frame Frame4 
            Height          =   930
            Left            =   4080
            TabIndex        =   23
            Top             =   3045
            Width           =   1980
            Begin VB.CommandButton cmdBotones 
               Caption         =   "&Cancelar"
               Height          =   690
               Index           =   12
               Left            =   1050
               Picture         =   "frmAnularLetras.frx":001C
               Style           =   1  'Graphical
               TabIndex        =   4
               Top             =   180
               Width           =   855
            End
            Begin VB.CommandButton cmdBotones 
               Caption         =   "&Acepta"
               Height          =   690
               Index           =   11
               Left            =   90
               Picture         =   "frmAnularLetras.frx":045E
               Style           =   1  'Graphical
               TabIndex        =   3
               Top             =   180
               Width           =   870
            End
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "Detalle de Documentos Referenciados"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   180
            TabIndex        =   27
            Top             =   210
            Width           =   9405
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2745
         Left            =   150
         TabIndex        =   6
         Top             =   285
         Width           =   9735
         Begin MSMask.MaskEdBox MBox1 
            Height          =   285
            Index           =   0
            Left            =   1740
            TabIndex        =   7
            Top             =   -330
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12648447
            Enabled         =   0   'False
            MaxLength       =   6
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "d/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   3
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   8520
            TabIndex        =   8
            Top             =   210
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12648447
            AutoTab         =   -1  'True
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Cliente 
            Height          =   345
            Left            =   1320
            TabIndex        =   0
            Top             =   900
            Width           =   8235
            _ExtentX        =   14526
            _ExtentY        =   609
            XcodMaxLongitud =   11
            xcodwith        =   800
            NomTabla        =   "cp_proveedor"
            TituloAyuda     =   "Ayuda de Clientes"
            ListaCampos     =   $"frmAnularLetras.frx":08A0
            XcodCampo       =   "clientecodigo"
            XListCampo      =   "clienterazonsocial"
            ListaCamposDescrip=   "Codigo,Descripcion,Ruc,Direccion,Distrito,LimiteCred,Saldo,T,P,M,D"
            ListaCamposText =   $"frmAnularLetras.frx":0986
         End
         Begin MSMask.MaskEdBox MBox 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   7305
            TabIndex        =   9
            Top             =   1995
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   393216
            ClipMode        =   1
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   7305
            TabIndex        =   10
            Top             =   2340
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            ClipMode        =   1
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox 
            Height          =   285
            Index           =   4
            Left            =   1305
            TabIndex        =   11
            Top             =   2235
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            PromptChar      =   "_"
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_NumDoc 
            Height          =   330
            Left            =   1305
            TabIndex        =   2
            Top             =   1590
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   582
            XcodMaxLongitud =   0
            xcodwith        =   900
            NomTabla        =   "cp_cargo"
            ListaCampos     =   "documentocargo(1),cargonumdoc(1),clientecodigo(1)"
            XcodCampo       =   "cargonumdoc"
            XListCampo      =   "clientecodigo"
            ListaCamposDescrip=   "TD,NDoc,CodCli"
            ListaCamposText =   "documentocargo,cargonumdoc,clientecodigo"
            Requerido       =   0   'False
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_TipoDoc 
            Height          =   300
            Left            =   1320
            TabIndex        =   1
            Top             =   1245
            Width           =   3840
            _ExtentX        =   6773
            _ExtentY        =   529
            XcodMaxLongitud =   0
            xcodwith        =   500
            NomTabla        =   "cp_tipodocumento"
            ListaCampos     =   "tdocumentocodigo(1),tdocumentodescripcion(1)"
            XcodCampo       =   "tdocumentocodigo"
            XListCampo      =   "tdocumentodescripcion"
            ListaCamposDescrip=   "Código,Descripción"
            ListaCamposText =   "tdocumentocodigo,tdocumentodescripcion"
            Requerido       =   0   'False
         End
         Begin VB.Label lblMoneda 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   3840
            TabIndex        =   29
            Top             =   2220
            Width           =   1470
         End
         Begin VB.Label Label5 
            Caption         =   "Importe"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   7
            Left            =   180
            TabIndex        =   21
            Top             =   2265
            Width           =   915
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Registro"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   7320
            TabIndex        =   20
            Top             =   240
            Width           =   1125
         End
         Begin VB.Label Label5 
            Caption         =   "Tipo Doc."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   6
            Left            =   180
            TabIndex        =   19
            Top             =   1305
            Width           =   1395
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Planilla"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   480
            TabIndex        =   18
            Top             =   -300
            Width           =   1215
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000009&
            Index           =   0
            X1              =   30
            X2              =   9750
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   1
            X1              =   30
            X2              =   9720
            Y1              =   570
            Y2              =   570
         End
         Begin VB.Label Label3 
            Caption         =   "Documento Principal"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   210
            TabIndex        =   17
            Top             =   630
            Width           =   3795
         End
         Begin VB.Label Label4 
            Caption         =   "Proveedor"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   210
            TabIndex        =   16
            Top             =   930
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Fecha Emision"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   5655
            TabIndex        =   15
            Top             =   2025
            Width           =   1395
         End
         Begin VB.Label Label5 
            Caption         =   "Moneda"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   3105
            TabIndex        =   14
            Top             =   2265
            Width           =   600
         End
         Begin VB.Label Label5 
            Caption         =   "Fecha Vencimiento"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   5640
            TabIndex        =   13
            Top             =   2385
            Width           =   1665
         End
         Begin VB.Label Label5 
            Caption         =   "Documento"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   4
            Left            =   180
            TabIndex        =   12
            Top             =   1665
            Width           =   1305
         End
      End
   End
End
Attribute VB_Name = "frmAnularLetras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adll As New dllgeneral.dll_general

Dim nLongicampo(6) As Integer
Dim rsdeta As New ADODB.Recordset

Dim apedido As String
Dim aalmacen As String
Dim alista As String * 2


Private Sub Form_Load()
  MBox1(1) = Format(Date, "DD/MM/YYYY")
    
  Call Ctr_Cliente.conexion(VGCNx)
  Call Ctr_TipoDoc.conexion(VGCNx)
  Call Ctr_NumDoc.conexion(VGCNx)
  
  Ctr_TipoDoc.Filtro = "tdocumentotipo='C' and tdocumentodocrenovaletra='1'"
  Ctr_TipoDoc.Ejecutar
  
End Sub

Public Function GrabarData() As Integer
    Dim J As Integer
    Dim regi As Long
    Dim nsql As String
    Dim tcargo As String
    
    Dim acmd As New ADODB.Command
    Dim asql As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim SQL As String

    On Error GoTo vererror
    
    GrabarData = 0
    
    If adll.VerificaDatoExistente(VGCNx, "select * from cp_abono where documentoabono='" & Trim$(Ctr_TipoDoc.xclave) & "' and  abononumdoc='" & Trim$(Ctr_NumDoc.xclave) & "'") = 0 Then
         VGCNx.Execute " Update cp_cargo " & _
                    " Set cargoapeflgreg='1' " & _
                    " where documentocargo='" & Trim$(Ctr_TipoDoc.xclave) & "' and cargonumdoc='" & Trim$(Ctr_NumDoc.xclave) & "'"
         
         SQL = "UPDATE cp_cargo set "
         SQL = SQL & "cargoapeimppag=cargoapeimppag-YY.abonocanimpsol,"
         SQL = SQL & "cargoapeflgcan = 0 "
         SQL = SQL & "FROM cp_cargo AA,"
         SQL = SQL & "(select A.documentoabono,A.abononumdoc,A.abonocancli,A.abonotipoplanilla,"
         SQL = SQL & "A.abonocanmoneda,A.abonocanimcan ,A.abonocanforcan,A.abonocanimpsol "
         SQL = SQL & "from  dbo.cp_abono A,"
         SQL = SQL & "(select  A.* from cp_cargo A,cp_tipoplanilla d "
         SQL = SQL & "where A.abonotipoplanilla=d.tplanillacodigo and  d.tplanillacanjes='1' and "
         SQL = SQL & "A.clientecodigo like '" & Trim$(Ctr_Cliente.xclave) & "' and "
         SQL = SQL & "A.documentocargo='" & Trim$(Ctr_TipoDoc.xclave) & "' and "
         SQL = SQL & "A.cargonumdoc like '" & Trim$(Ctr_NumDoc.xclave) & "') as ZZ "
         SQL = SQL & "Where A.abononumplanilla=ZZ.abononumplanilla and "
         SQL = SQL & "A.abonotipoplanilla=ZZ.abonotipoplanilla and "
         SQL = SQL & "A.abononumplanilla like '%' ) as YY "
         SQL = SQL & "where AA.documentocargo=YY.documentoabono and "
         SQL = SQL & "AA.cargonumdoc=YY.abononumdoc and "
         SQL = SQL & "AA.clientecodigo=YY.abonocancli and "
         SQL = SQL & "AA.cargoapeimppag<>0"
         VGCNx.Execute (SQL)
                    
         rsdeta.MoveFirst
         Do Until rsdeta.EOF
            'SQL = "UPDATE cp_abono SET abonocanflreg='1' where "
            SQL = "DELETE cp_abono WHERE "
            SQL = SQL & "documentoabono='" & rsdeta("documentoabono").Value & "' and "
            SQL = SQL & "abononumdoc='" & rsdeta("abononumdoc").Value & "' and "
            SQL = SQL & "abonocancli='" & Trim$(Ctr_Cliente.xclave) & "'"
            VGCNx.Execute (SQL)
            rsdeta.MoveNext
         Loop
                    
        MsgBox "Se Anulo Satisfactoriamente el Documento: " & Chr(13) & Chr(10) & Ctr_NumDoc.xclave, vbInformation, MsgTitle
        GrabarData = 1
    Else
      MsgBox "No se puede Anular el Documento tiene Abonos " & Chr(13) & Chr(10) & Ctr_NumDoc.xclave, vbInformation, MsgTitle
      GrabarData = 0
    End If
    
vererror:
   If Err Then
      MsgBox Err.Number & "-" & Err.Description
      MsgBox "Comunicarse con Sistemas ...!!" & Chr(13) & Chr(10) & VGCNx.Errors(0).Number & "-" & VGCNx.Errors(0).Description
      Exit Function
   End If
End Function

Private Sub Ctr_TipoDoc_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
  Ctr_NumDoc.Filtro = "documentocargo='" & ColecCampos(0).Value & "' and cargoapeflgcan='0' and isnull(cargoapeflgreg,0)<>1"
  'Ctr_NumDoc.Filtro = "documentocargo='" & ColecCampos(0).Value & "' and cargoapeflgcan='0' and clientecodigo='82'"
  Ctr_NumDoc.Ejecutar
End Sub

Private Sub Ctr_NumDoc_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
  Call MuestraData
End Sub

Private Sub cmdBotones_Click(Index As Integer)
   Dim asql As String
   Dim acmd As New ADODB.Command
   Dim J, nl As Integer
   
   On Error GoTo vererror
   
   Select Case Index
    Case 11:
      If ValidaData() = True Then
         VGCNx.BeginTrans
         If GrabarData() = 1 Then
            VGCNx.CommitTrans
            g_TipoMovi = 0
            Exit Sub
         Else
            VGCNx.RollbackTrans
            g_TipoMovi = 0
            Exit Sub
         End If
         g_TipoMovi = 0
      End If
    Case 12:
       g_TipoMovi = 0
       Unload Me
   End Select
   
vererror:
  If Err Then
     MsgBox "Comunicarse con Sistemas ...!!" & Chr(13) & Chr(10) & Err.Number & "-" & Err.Description
     Err = 0
     Exit Sub
  End If

End Sub

Sub LimpiarData()
 Dim i  As Integer
  Ctr_Cliente.xclave = Empty: Ctr_Cliente.Ejecutar
  Ctr_TipoDoc.xclave = Empty: Ctr_TipoDoc.Ejecutar
  Ctr_NumDoc.xclave = Empty: Ctr_NumDoc.Ejecutar
  For i = 3 To 5
    MBox(i).Text = Empty
  Next
  lblMoneda.Caption = Empty

End Sub

Function ValidaData() As Boolean
   
   If rsdeta.RecordCount = 0 Then
     MsgBox "Esta Letra no tiene Detalle de Facturas Relacionadas", vbInformation, Caption
     ValidaData = False
     Exit Function
   End If
   
   If adll.VerificaDatoExistente(VGCNx, "select * from cp_cargo where documentocargo='" & Trim$(Ctr_TipoDoc.xclave) & "' and cargonumdoc='" & Trim$(Ctr_NumDoc.xclave) & "' and clientecodigo='" & Ctr_Cliente.xclave & "' and cargoapeflgreg='1'") = 1 Then
        MsgBox "El Documento esta anulado...!!!", vbInformation, Caption
        ValidaData = False
        Exit Function
    End If

    If MsgBox("Desea Anular el Documento?", vbYesNo, MsgTitle) = vbNo Then
        ValidaData = False
        Exit Function
    End If

    If IsNull(Ctr_Cliente.xclave) Or Trim$(Ctr_Cliente.xclave) = Empty Then
       MsgBox "Cliente no existe...Verifique!!!", vbInformation, Caption
       Ctr_Cliente.SetFocus
       ValidaData = False
       Exit Function
    End If

    ValidaData = True

End Function

Sub MuestraData()
  Dim SQL As String
  
  SQL = " select  A.documentoabono,A.abononumdoc,A.zonacodigo,A.abonotipoplanilla,A.vendedorcodigo,"
  SQL = SQL & "A.abononumplanilla,A.abonocanfecpla,A.abonocanfecpro,A.abonocanmoneda,A.abonocanimcan,"
  SQL = SQL & "A.abonocanforcan,A.abonocanfecan,A.abonocanmoncan,A.abonocanimpcan,A.abonocanimpsol,ZZ.documentocargo,"
  SQL = SQL & "ZZ.cargonumdoc,ZZ.abonotipoplanilla,ZZ.abononumplanilla,ZZ.cargoapefecemi,ZZ.cargoapefecvct,ZZ.cargoapeimpape,e.clienterazonsocial,"
  SQL = SQL & "f.tdocumentodesccorta as DescDocAbono,g.monedasimbolo as MonAbono,h.tdocumentodesccorta as DescDocCargo,"
  SQL = SQL & "i.monedasimbolo as MonCargo from  cp_abono A,"
  SQL = SQL & "(select  A.* from cp_cargo A,cp_tipoplanilla d "
  SQL = SQL & "where isnull(cargoapeflgreg,0)<>1 and A.abonotipoplanilla=d.tplanillacodigo and  (d.tplanillacanjes='1' or d.tplanillarenovar='1') and "
  SQL = SQL & "A.clientecodigo like '" & Trim$(Ctr_Cliente.xclave) & "' and A.documentocargo='" & Trim$(Ctr_TipoDoc.xclave) & "' and "
  SQL = SQL & "A.cargonumdoc='" & Trim$(Ctr_NumDoc.xclave) & "') as ZZ,"
  SQL = SQL & "cp_proveedor e,dbo.cp_tipodocumento f,"
  SQL = SQL & "gr_moneda g,dbo.cp_tipodocumento h,"
  SQL = SQL & "gr_moneda I "
  SQL = SQL & "where A.abononumplanilla=ZZ.abononumplanilla and A.abonotipoplanilla=ZZ.abonotipoplanilla and "
  SQL = SQL & "A.abononumplanilla like '%' and isnull(A.abonocanflreg,0)<>1 and A.abonocancli = e.clientecodigo and "
  SQL = SQL & "A.documentoabono = f.tdocumentocodigo and A.abonocanmoneda = g.monedacodigo and "
  SQL = SQL & "ZZ.documentocargo = h.tdocumentocodigo and ZZ.monedacodigo = I.monedacodigo "
  SQL = SQL & "ORDER BY A.documentoabono,A.abononumdoc "
  
  Set rsdeta = New ADODB.Recordset
  Set rsdeta = VGCNx.Execute(SQL)
  Dim Suma As Double
  Suma = 0
  Set TDBGDetalleDoc.DataSource = rsdeta
  If Not rsdeta.BOF Or Not rsdeta.EOF Then
    MBox(4).Text = rsdeta("cargoapeimpape").Value
    MBox(3).Text = rsdeta("cargoapefecemi").Value
    MBox(5).Text = rsdeta("cargoapefecvct").Value
    lblMoneda.Caption = rsdeta("moncargo").Value
    rsdeta.MoveFirst
    Do Until rsdeta.EOF
       Suma = Suma + rsdeta("abonocanimpsol").Value
       rsdeta.MoveNext
    Loop
  Else
    Set rsdeta = New ADODB.Recordset
  End If
  Text1(3).Text = Format(Suma, "###,###,##0.#0")

End Sub
