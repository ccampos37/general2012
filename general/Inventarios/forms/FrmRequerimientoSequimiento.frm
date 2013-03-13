VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmRequerimientoSeguimiento 
   Caption         =   "Seguimiento de Requerimientos"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10845
   LinkTopic       =   "Form2"
   ScaleHeight     =   8325
   ScaleWidth      =   10845
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FrameEliminar 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Anulacion"
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   5760
      TabIndex        =   22
      Top             =   720
      Width           =   4695
      Begin VB.CommandButton Cmdgrabaanulacion 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   2520
         Picture         =   "FrmRequerimientoSequimiento.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton CmdSalirAnulacion 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   675
         Left            =   3600
         Picture         =   "FrmRequerimientoSequimiento.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   240
         Width           =   775
      End
      Begin MSComCtl2.DTPicker DTPAnulacion 
         Height          =   300
         Left            =   360
         TabIndex        =   25
         Top             =   555
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   48037889
         CurrentDate     =   37623.1285069444
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Anulacion"
         Height          =   210
         Index           =   8
         Left            =   600
         TabIndex        =   26
         Top             =   240
         Width           =   1110
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8535
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   10650
      _ExtentX        =   18785
      _ExtentY        =   15055
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   420
      TabMaxWidth     =   2
      TabCaption(0)   =   "Requerimientos de Bienes"
      TabPicture(0)   =   "FrmRequerimientoSequimiento.frx":0884
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Fr1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Fr2(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame Fr2 
         Height          =   1365
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   10290
         Begin VB.Frame Frame3 
            Height          =   1005
            Left            =   6480
            TabIndex        =   11
            Top             =   240
            Width           =   3600
            Begin VB.CheckBox ChkFech 
               Caption         =   "Rango de Fechas"
               Height          =   285
               Left            =   75
               TabIndex        =   12
               Top             =   -45
               Width           =   1620
            End
            Begin MSComCtl2.DTPicker DTPFechaIni 
               Height          =   300
               Left            =   1620
               TabIndex        =   13
               Top             =   195
               Width           =   1785
               _ExtentX        =   3149
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   0   'False
               Format          =   48037889
               CurrentDate     =   37623.1285069444
            End
            Begin MSComCtl2.DTPicker DTPFechaFin 
               Height          =   300
               Left            =   1620
               TabIndex        =   14
               Top             =   555
               Width           =   1785
               _ExtentX        =   3149
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   0   'False
               Format          =   48037889
               CurrentDate     =   37623.1264351852
            End
            Begin VB.Label Label2 
               Caption         =   "Fecha Fin :"
               Height          =   210
               Left            =   435
               TabIndex        =   16
               Top             =   615
               Width           =   810
            End
            Begin VB.Label Label1 
               Caption         =   "Fecha Inicio :"
               Height          =   210
               Index           =   7
               Left            =   390
               TabIndex        =   15
               Top             =   240
               Width           =   1110
            End
         End
         Begin VB.CheckBox ChkTodos 
            Caption         =   "Incluir Todos"
            Height          =   375
            Left            =   5400
            TabIndex        =   10
            Top             =   360
            Width           =   1095
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayusolicitante 
            Height          =   315
            Left            =   1080
            TabIndex        =   8
            Top             =   720
            Width           =   3870
            _ExtentX        =   6826
            _ExtentY        =   556
            XcodMaxLongitud =   3
            xcodwith        =   500
            NomTabla        =   "co_solicitantes"
            TituloAyuda     =   "Busqueda de Solicitante"
            ListaCampos     =   "solicitantecodigo(1),solicitantenombre(1)"
            XcodCampo       =   "solicitantecodigo"
            XListCampo      =   "solicitantenombre"
            ListaCamposDescrip=   "Codigo,Descripcion"
            ListaCamposText =   "solicitantecodigo,solicitantenombre"
            Requerido       =   0   'False
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyutipoOrden 
            Height          =   270
            Left            =   1080
            TabIndex        =   28
            Top             =   240
            Width           =   3870
            _ExtentX        =   6826
            _ExtentY        =   476
            XcodMaxLongitud =   6
            xcodwith        =   500
            NomTabla        =   "co_tipodeorden"
            TituloAyuda     =   "Busqueda de Tipo de Orden"
            ListaCampos     =   "tipoordencodigo(1),tipoordendescripcion(1),tipoordennumeracion(2),ordendebienes(2)"
            XcodCampo       =   "tipoordencodigo"
            XListCampo      =   "tipoordendescripcion"
            ListaCamposDescrip=   "Codigo,Descripcion,numeracion,tipo de orden"
            ListaCamposText =   "tipoordencodigo,tipoordendescripcion,tipoordennumeracion,ordendebienes"
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "Tipo "
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   270
            Width           =   360
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Solicitante     :"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   765
            Width           =   885
         End
      End
      Begin VB.Frame Fr1 
         Height          =   7785
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   10416
         Begin VB.CommandButton CmdSalir 
            Caption         =   "&Salir"
            Height          =   675
            Left            =   5040
            Picture         =   "FrmRequerimientoSequimiento.frx":08A0
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   6720
            Width           =   775
         End
         Begin VB.CommandButton CmdEli 
            Caption         =   "&Anulacion"
            Height          =   675
            Left            =   2400
            Picture         =   "FrmRequerimientoSequimiento.frx":0CE2
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   6720
            Width           =   900
         End
         Begin VB.CommandButton command5 
            Caption         =   "&Reporte"
            Height          =   675
            Left            =   3840
            Picture         =   "FrmRequerimientoSequimiento.frx":1124
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   6720
            Width           =   775
         End
         Begin VB.Frame Frame5 
            Height          =   585
            Index           =   0
            Left            =   7644
            TabIndex        =   2
            Top             =   3720
            Width           =   2265
            Begin MSMask.MaskEdBox totreg 
               Height          =   372
               Index           =   0
               Left            =   1104
               TabIndex        =   3
               Top             =   144
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   635
               _Version        =   393216
               Appearance      =   0
               BackColor       =   12648447
               ForeColor       =   16777215
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PromptChar      =   "_"
            End
            Begin VB.Label Label3 
               Caption         =   "Total Reg."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   228
               Index           =   0
               Left            =   156
               TabIndex        =   4
               Top             =   192
               Width           =   1032
            End
         End
         Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
            Height          =   2415
            Left            =   120
            TabIndex        =   17
            Top             =   1320
            Width           =   10095
            _ExtentX        =   17806
            _ExtentY        =   4260
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
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).DataField=   ""
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).DataField=   ""
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).DataField=   ""
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   5
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   13160660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=5"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(9)=   "Column(2).Width=2725"
            Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2646"
            Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(13)=   "Column(3).Width=2725"
            Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2646"
            Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(17)=   "Column(4).Width=2725"
            Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
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
            DeadAreaBackColor=   13160660
            RowDividerColor =   13160660
            RowSubDividerColor=   13160660
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=54,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=58,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
            _StyleDefs(56)  =   "Named:id=33:Normal"
            _StyleDefs(57)  =   ":id=33,.parent=0"
            _StyleDefs(58)  =   "Named:id=34:Heading"
            _StyleDefs(59)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(60)  =   ":id=34,.wraptext=-1"
            _StyleDefs(61)  =   "Named:id=35:Footing"
            _StyleDefs(62)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(63)  =   "Named:id=36:Selected"
            _StyleDefs(64)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(65)  =   "Named:id=37:Caption"
            _StyleDefs(66)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(67)  =   "Named:id=38:HighlightRow"
            _StyleDefs(68)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(69)  =   "Named:id=39:EvenRow"
            _StyleDefs(70)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(71)  =   "Named:id=40:OddRow"
            _StyleDefs(72)  =   ":id=40,.parent=33"
            _StyleDefs(73)  =   "Named:id=41:RecordSelector"
            _StyleDefs(74)  =   ":id=41,.parent=34"
            _StyleDefs(75)  =   "Named:id=42:FilterBar"
            _StyleDefs(76)  =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
            Height          =   2280
            Left            =   240
            TabIndex        =   21
            Top             =   4320
            Width           =   12945
            _ExtentX        =   22834
            _ExtentY        =   4022
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Item"
            Columns(0).DataField=   "oc_citem"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Codigo"
            Columns(1).DataField=   "oc_ccodigo"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Descripcion"
            Columns(2).DataField=   "descripcion"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "C.Requerida"
            Columns(3).DataField=   "oc_ncantid"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "C.Atendida"
            Columns(4).DataField=   "oc_ncanaten"
            Columns(4).NumberFormat=   "#####.##"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Saldo x Atender"
            Columns(5).DataField=   "oc_nsaldo"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   4
            Columns(6)._MaxComboItems=   5
            Columns(6).ValueItems(0)._DefaultItem=   0
            Columns(6).ValueItems(0).Value=   "1"
            Columns(6).ValueItems(0).Value.vt=   8
            Columns(6).ValueItems(0).DisplayValue=   "1"
            Columns(6).ValueItems(0).DisplayValue.vt=   8
            Columns(6).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
            Columns(6).ValueItems(1)._DefaultItem=   0
            Columns(6).ValueItems(1).Value=   "0"
            Columns(6).ValueItems(1).Value.vt=   8
            Columns(6).ValueItems(1).DisplayValue=   "0"
            Columns(6).ValueItems(1).DisplayValue.vt=   8
            Columns(6).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
            Columns(6).ValueItems.Count=   2
            Columns(6).Caption=   "Estado"
            Columns(6).DataField=   "oc_estadoorden"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   7
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   13160660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=7"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=979"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=900"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=2170"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2090"
            Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8196"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=6694"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=6615"
            Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=8196"
            Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(16)=   "Column(3).Width=1852"
            Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1773"
            Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=8196"
            Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(21)=   "Column(4).Width=2196"
            Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2117"
            Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=8196"
            Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(26)=   "Column(5).Width=2672"
            Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2593"
            Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=8196"
            Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(31)=   "Column(6).Width=1058"
            Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=979"
            Splits(0)._ColumnProps(34)=   "Column(6).Order=7"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=System"
            PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=System"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            Appearance      =   2
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            MultipleLines   =   0
            CellTipsWidth   =   0
            MultiSelect     =   2
            AnimateWindow   =   2
            AnimateWindowClose=   2
            DeadAreaBackColor=   13160660
            RowDividerColor =   13160660
            RowSubDividerColor=   13160660
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
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&H344A87&"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.locked=-1"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=54,.parent=13,.locked=-1"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=58,.parent=13,.locked=-1"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=62,.parent=13,.locked=-1"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=66,.parent=13,.locked=-1"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=70,.parent=13,.locked=-1"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=50,.parent=13,.bgcolor=&HBFFFAA&"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=17"
            _StyleDefs(64)  =   "Named:id=33:Normal"
            _StyleDefs(65)  =   ":id=33,.parent=0"
            _StyleDefs(66)  =   "Named:id=34:Heading"
            _StyleDefs(67)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(68)  =   ":id=34,.wraptext=-1"
            _StyleDefs(69)  =   "Named:id=35:Footing"
            _StyleDefs(70)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(71)  =   "Named:id=36:Selected"
            _StyleDefs(72)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(73)  =   "Named:id=37:Caption"
            _StyleDefs(74)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(75)  =   "Named:id=38:HighlightRow"
            _StyleDefs(76)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(77)  =   "Named:id=39:EvenRow"
            _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(79)  =   "Named:id=40:OddRow"
            _StyleDefs(80)  =   ":id=40,.parent=33"
            _StyleDefs(81)  =   "Named:id=41:RecordSelector"
            _StyleDefs(82)  =   ":id=41,.parent=34"
            _StyleDefs(83)  =   "Named:id=42:FilterBar"
            _StyleDefs(84)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Detalle del Requerimiento"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   345
            Index           =   0
            Left            =   3000
            TabIndex        =   6
            Top             =   3990
            Width           =   3765
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Requerimientos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   348
            Left            =   240
            TabIndex        =   5
            Top             =   96
            Width           =   9528
         End
      End
   End
End
Attribute VB_Name = "FrmRequerimientoSeguimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsacumula As New ADODB.Recordset
Dim rsdeta As New ADODB.Recordset
Dim rsdeta2 As New ADODB.Recordset
Dim csql As New ADODB.Recordset
Dim rsql As New ADODB.Recordset
Dim dllgeneral As New dllgeneral.dll_general
Dim al_tempo As String, al_tempo1 As String
Dim xsql, xAlma, xtipo, xnumero As String
Dim rsql_ok As Integer
Dim g_tipoped As String
Dim g_pedserie As String
Dim nivelatendido As String
Dim acepta As Integer
Dim nLongicampo(1) As Integer

Private Sub ChkFech_Click()
If ChkFech.Value = 1 Then
    DTPFechaIni.Enabled = True
    DTPFechaFin.Enabled = True
  Else
    DTPFechaIni.Enabled = False
    DTPFechaFin.Enabled = False
End If

End Sub

Private Sub ChkTodos_Click()
If ChkTodos.Value = 0 Then
   inicializaarchivo (3)
 Else
   inicializaarchivo (0)
 End If
End Sub

Private Sub CmdEli_Click()
FrameEliminar.Visible = True
Call CargaGrilla
End Sub

Private Sub Cmdgrabaanulacion_Click()
If rsql.BOF Then Exit Sub
rsql.Update
Dim actualiza As New ADODB.Recordset
Dim saldo As Double
Dim xsql As String
rsql.MoveFirst
Do Until rsql.EOF
   If rsql!oc_estadoorden Then
      xsql = " update co_detordcompra set oc_estadoorden=1, oc_nsaldo=0  where tipoordencodigo='" & csql!tipoordencodigo & "'"
      xsql = xsql & " and oc_cnumord='" & csql!oc_cnumord & "' And oc_citem ='" & rsql!oc_citem & " '"
      rsql!oc_nsaldo = 0
      Set actualiza = VGCNx.Execute(xsql)
   End If
   saldo = saldo + rsql!oc_nsaldo
   rsql.MoveNext
Loop
If saldo = 0 Then
   xsql = " update co_cabordcompra set estadooccodigo='" & nivelatendido & "'"
   xsql = xsql & " where tipoordencodigo='" & csql!tipoordencodigo & "'"
   xsql = xsql & " and oc_cnumord='" & csql!oc_cnumord & "'"
   Set actualiza = VGCNx.Execute(xsql)
End If
FrameEliminar.Visible = False
inicializaarchivo (0)
csql.MoveFirst
TDBGrid1.Refresh
End Sub

Private Sub cmdNuevo_Click()
   inicializaarchivo
   End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub cmdSalirFinal_Click(Index As Integer)
   Call dllgeneral.ActivaTab(0, 1, SSTab1)
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub



Private Sub Command2_Click()
  Unload Me

End Sub

Private Sub Command3_Click()
FrameEliminar.Visible = False
End Sub

Private Sub CmdSalirAnulacion_Click()
FrameEliminar.Visible = False
End Sub

Private Sub Ctr_Ayusolicitante_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
If Ctr_Ayusolicitante.xclave <> "" Then inicializaarchivo (1)
End Sub

Private Sub Ctrayu_tipoorden_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)

End Sub

Private Sub Ctr_AyutipoOrden1_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)

End Sub

Private Sub Ctr_Ayusolicitante_AlNoDevolverNada()
 inicializaarchivo (0)
End Sub

Private Sub Ctr_AyutipoOrden_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
 inicializaarchivo (1)
End Sub

Private Sub Ctr_AyutipoOrden_AlNoDevolverNada()
 inicializaarchivo (0)
End Sub

Private Sub DTPFechaFin_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
 inicializaarchivo (2)
End Sub

Private Sub DTPFechaIni_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
 inicializaarchivo (2)
End Sub

Private Sub Form_Load()
   al_tempo = "##al_" & ComputerName
   al_tempo1 = "##al1_" & ComputerName
  nLongicampo(1) = 0
  Call Ctr_Ayusolicitante.Conexion(VGCNx)
  Call Ctr_AyutipoOrden.Conexion(VGCNx): Ctr_AyutipoOrden.filtro = ("flagrequerimientos=1")
  DTPAnulacion.Value = Date
  DTPFechaIni = Date - 30
  DTPFechaFin = Date
  FrameEliminar.Visible = False
  TDBGrid2.FetchRowStyle = True
  Set csql = VGCNx.Execute("select top 1 estadooccodigo from co_estadorequerimiento where estadoocatendido=1 ")
  nivelatendido = csql!estadooccodigo
  ConfiguraGrid
End Sub
Private Sub inicializaarchivo(Optional dato As Integer)

'If ExisteElem(0, VGCNx, al_tempo) Then VGCNx.Execute ("drop table " & al_tempo)
xsql = " select  e.solicitantenombre,a.tipoordencodigo,a.OC_CNUMORD,a.OC_DFECDOC,a.OC_CRAZSOC,a.estadooccodigo,a.OC_CSOLICT,a.OC_CCOTIZA  from co_cabordcompra a "
xsql = xsql & " inner join co_estadorequerimiento c on a.estadooccodigo=c.estadooccodigo"
xsql = xsql & " inner join co_nivelrequerimiento d on C.nivelrequerimientocodigo=d.nivelrequerimientocodigo "
xsql = xsql & " inner join co_solicitantes e on a.oc_csolict=e.solicitantecodigo "
xsql = xsql & " inner join co_tipodeorden f on a.tipoordencodigo=f.tipoordencodigo "
If dato = 1 Then
   If Not Ctr_Ayusolicitante.xclave = "" Then
      xsql = xsql & " and a.oc_csolict='" & Ctr_Ayusolicitante.xclave & "'"
   End If
   If Not Ctr_AyutipoOrden.xclave = "" Then
      xsql = xsql & " and a.tipoordencodigo='" & Ctr_AyutipoOrden.xclave & "'"
   End If
End If
If dato = 2 Then xsql = xsql & " and a.oc_dfecdoc>=DTPFechaIni and a.oc_dfecdoc <= DTPFechaFin "
If dato = 3 Then xsql = xsql & " AND d.nivelaprobaciongerencia=1 and c.estadoocatendido<>1 "
Set csql = VGCNx.Execute(xsql)
Set TDBGrid1.DataSource = Nothing
TDBGrid1.ClearFields
Set TDBGrid1.DataSource = csql
TDBGrid1.Refresh
If csql.RecordCount > 0 Then
  csql.MoveFirst
End If
Call TDBGrid1_Click
End Sub


Public Function ConfiguraGrid()

   With TDBGrid1
       .Columns(0).Caption = "solicitante"
       .Columns(0).Width = 2200
       .Columns(0).HeadAlignment = dbgCenter
       .Columns(1).Caption = "tipo"
       .Columns(1).Width = 1000
       .Columns(1).HeadAlignment = dbgCenter
       .Columns(2).Caption = "Numero"
       .Columns(2).Width = 1600
       .Columns(2).HeadAlignment = dbgCenter
       .Columns(3).Caption = "fecha"
       .Columns(3).Width = 1600
       .Columns(3).HeadAlignment = dbgCenter
       .Columns(4).Caption = "Nro.Sistema."
       .Columns(4).Width = 1000
       .Columns(4).HeadAlignment = dbgCenter
       .Refresh
   End With
   
   
End Function
Private Sub TDBGrid1_Click()
Dim sql As String
If csql.RecordCount > 0 Then
   If rsql_ok = 1 Then rsql.Close
   sql = "select descripcion=case when oc_ccodigo='00' then rtrim(OC_CCOMEN1) "
   sql = sql & " else rtrim(OC_CDESREF)+'-'+rtrim(OC_CCOMEN1) end, a.* from co_detordcompra a "
   sql = sql & " where TIPOORDENCODIGO='" & csql!tipoordencodigo & "' and oc_cnumord='" & csql!oc_cnumord & "'"
   sql = sql & " and isnull(oc_estadoorden,0)<>1"
   rsql.Open (sql), VGCNx, adOpenDynamic, adLockBatchOptimistic
   TDBGrid2.DataSource = rsql
   TDBGrid2.Refresh
   rsql_ok = 1
End If
End Sub
Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
 With csql
    If .Sort = Empty Then
        .Sort = TDBGrid1.Columns.item(ColIndex).DataField & " asc"
    ElseIf Right(.Sort, 3) = "asc" Then
        .Sort = TDBGrid1.Columns.item(ColIndex).DataField & " desc"
    ElseIf Right(.Sort, 4) = "desc" Then
        .Sort = TDBGrid1.Columns.item(ColIndex).DataField & " asc"
    End If
    TDBGrid1.Refresh
 End With
End Sub


Private Sub TDBGrid1_DblClick()
Call TDBGrid1_Click
End Sub

Private Sub tdbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
   Call TDBGrid1_Click
End If
End Sub
Private Sub TDBGrid2_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
Dim rsclone As New ADODB.Recordset
    On Error Resume Next
    Set rsclone = rsql.Clone(adLockReadOnly)
    If rsclone.RecordCount = 0 Then Exit Sub
    rsclone.Bookmark = Bookmark
    If rsclone!oc_nsaldo > 0 Then
       RowStyle.BackColor = RGB(254, 251, 218)
       '185,251,210
    End If
    If rsclone!oc_estado = 1 Then
       RowStyle.BackColor = RGB(200, 250, 100)
    End If
    If TDBGrid2.ApproxCount > 0 And TDBGrid2.Columns(6).text = 1 Then
        xtipo = csql!tipoordencodigo
        rsdeta.AddNew
        rsdeta!tipo = xtipo
        rsdeta!numero = csql!oc_cnumord
        rsdeta!item = TDBGrid2.Columns(0).text
        rsdeta!cant = TDBGrid2.Columns(6).text
        rsdeta!Estado = TDBGrid2.Columns(7).text
    End If
End Sub

Public Function CargaGrilla()

   Set rsdeta = Nothing
   Call rsdeta.Fields.Append("Item", adChar, 3)
   Call rsdeta.Fields.Append("tipo", adChar, 2)
   Call rsdeta.Fields.Append("numero", adChar, 11)
   Call rsdeta.Fields.Append("Cant", adDouble)
   Call rsdeta.Fields.Append("estado", adBoolean)
   rsdeta.Open
   End Function

Private Sub cmdBotones_Click(Index As Integer)
  On Error GoTo nerror
  Select Case Index
  Case 11
    Call dllgeneral.ActivaTab(0, 1, SSTab1)
    TDBGrid2.Refresh
   
  Case 12
    Call dllgeneral.ActivaTab(0, 1, SSTab1)
  
  End Select
  
nerror:
   If err Then
       MsgBox err.Description & "-" & err.Description, vbInformation, MsgTitle
       err = 0
       Resume Next
       Exit Sub
   End If
End Sub

