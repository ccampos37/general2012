VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form frmingresoOC 
   AutoRedraw      =   -1  'True
   Caption         =   "Movimientos por Ordenes/Requerimeintos de articulos"
   ClientHeight    =   6195
   ClientLeft      =   1725
   ClientTop       =   1725
   ClientWidth     =   10170
   Icon            =   "frmingresoOC.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   10170
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   4320
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   10610
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Doc. Pendientes de Procesar"
      TabPicture(0)   =   "frmingresoOC.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "TDBGridordenes"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(2)=   "Command1"
      Tab(0).Control(3)=   "Command2"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Ingresos/salidas de Almacen"
      TabPicture(1)   =   "frmingresoOC.frx":08E6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "DBGrid1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "CmdSalir"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdGra"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "fraDatos"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "fraCabec"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Codificacion de articulos"
      TabPicture(2)   =   "frmingresoOC.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "CmdSalirproducto"
      Tab(2).Control(1)=   "CmdAceptar"
      Tab(2).Control(2)=   "TDBGridarticulos"
      Tab(2).Control(3)=   "Ctr_Ayuarticulo"
      Tab(2).Control(4)=   "Label11"
      Tab(2).ControlCount=   5
      Begin VB.CommandButton Command2 
         Caption         =   "&Adicionar"
         Height          =   650
         Left            =   -70320
         Picture         =   "frmingresoOC.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   5040
         Width           =   1000
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Salir"
         Height          =   650
         Left            =   -69000
         Picture         =   "frmingresoOC.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   5040
         Width           =   1000
      End
      Begin VB.Frame fraCabec 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   240
         TabIndex        =   35
         Top             =   360
         Width           =   9516
         Begin VB.TextBox txtNum 
            Height          =   285
            Left            =   7710
            MaxLength       =   13
            TabIndex        =   36
            Top             =   240
            Width           =   1335
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctrayu_tipoorden 
            Height          =   270
            Left            =   1470
            TabIndex        =   37
            Top             =   240
            Width           =   3390
            _ExtentX        =   5980
            _ExtentY        =   476
            XcodMaxLongitud =   11
            xcodwith        =   1100
            NomTabla        =   "co_tipodeorden"
            TituloAyuda     =   "Busqueda de Tipo de Orden"
            ListaCampos     =   "tipoordencodigo(1),tipoordendescripcion(1),tipoordennumeracion(2),flagrequerimientosordenes(1)"
            XcodCampo       =   "tipoordencodigo"
            XListCampo      =   "tipoordendescripcion"
            ListaCamposDescrip=   "Codigo,Descripcion"
            ListaCamposText =   "tipoordencodigo,tipoordendescripcion,tipoordennumeracion,flagrequerimientosordenes"
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Número  :"
            Height          =   195
            Left            =   6960
            TabIndex        =   39
            Top             =   240
            Width           =   690
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Orden     :"
            Height          =   195
            Left            =   360
            TabIndex        =   38
            Top             =   270
            Width           =   1035
         End
      End
      Begin VB.Frame fraDatos 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2145
         Left            =   270
         TabIndex        =   10
         Top             =   930
         Width           =   9516
         Begin VB.TextBox txtAlm 
            Height          =   285
            Left            =   1440
            MaxLength       =   3
            TabIndex        =   15
            Top             =   1560
            Width           =   315
         End
         Begin VB.TextBox txtTM 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
            Height          =   285
            Left            =   5520
            MaxLength       =   2
            TabIndex        =   14
            Top             =   1560
            Width           =   375
         End
         Begin VB.TextBox txtNTF 
            Height          =   285
            Left            =   5976
            MaxLength       =   10
            TabIndex        =   13
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox txtTF 
            Height          =   285
            Left            =   4770
            MaxLength       =   3
            TabIndex        =   12
            Top             =   1200
            Width           =   375
         End
         Begin VB.TextBox txtserie 
            Height          =   285
            Left            =   5280
            MaxLength       =   4
            TabIndex        =   11
            Top             =   1200
            Width           =   615
         End
         Begin MSComCtl2.DTPicker txtfec 
            Height          =   285
            Left            =   1440
            TabIndex        =   40
            Top             =   1080
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Format          =   40239105
            CurrentDate     =   37015
         End
         Begin VB.Label lblCen 
            AutoSize        =   -1  'True
            Caption         =   "Tip./Fact.    :"
            Height          =   195
            Left            =   3690
            TabIndex        =   34
            Top             =   1215
            Width           =   930
         End
         Begin VB.Label lblProv 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2640
            TabIndex        =   33
            Top             =   360
            Width           =   2970
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Almacén       :"
            Height          =   195
            Left            =   375
            TabIndex        =   32
            Top             =   1575
            Width           =   975
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fecha           :"
            Height          =   195
            Left            =   375
            TabIndex        =   31
            Top             =   1215
            Width           =   990
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Emisión         :"
            Height          =   195
            Left            =   375
            TabIndex        =   30
            Top             =   740
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor     :"
            Height          =   195
            Left            =   375
            TabIndex        =   29
            Top             =   380
            Width           =   1005
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "R.U.C.  :"
            Height          =   192
            Left            =   7056
            TabIndex        =   28
            Top             =   372
            Width           =   612
         End
         Begin VB.Label lblRuc 
            BorderStyle     =   1  'Fixed Single
            Height          =   288
            Left            =   7728
            TabIndex        =   27
            Top             =   360
            Width           =   1596
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Entrega   :"
            Height          =   192
            Left            =   3720
            TabIndex        =   26
            Top             =   732
            Width           =   732
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Estado  :"
            Height          =   192
            Left            =   6216
            TabIndex        =   25
            Top             =   732
            Width           =   636
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Movim. :"
            Height          =   192
            Left            =   4536
            TabIndex        =   24
            Top             =   1572
            Width           =   960
         End
         Begin VB.Label lblAlm 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1920
            TabIndex        =   23
            Top             =   1560
            Width           =   2535
         End
         Begin VB.Label lblPro 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1440
            TabIndex        =   22
            Top             =   360
            Width           =   1110
         End
         Begin VB.Label lblEmi 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1440
            TabIndex        =   21
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label lblEnt 
            BorderStyle     =   1  'Fixed Single
            Height          =   288
            Left            =   4560
            TabIndex        =   20
            Top             =   720
            Width           =   1092
         End
         Begin VB.Label lblEsta 
            BorderStyle     =   1  'Fixed Single
            Height          =   288
            Left            =   7536
            TabIndex        =   19
            Top             =   720
            Width           =   1788
         End
         Begin VB.Label lblEst 
            BorderStyle     =   1  'Fixed Single
            Height          =   288
            Left            =   7056
            TabIndex        =   18
            Top             =   720
            Width           =   396
         End
         Begin VB.Label lblTF 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7320
            TabIndex        =   17
            Top             =   1200
            Width           =   2070
         End
         Begin VB.Label lblTM 
            BorderStyle     =   1  'Fixed Single
            Height          =   288
            Left            =   6000
            TabIndex        =   16
            Top             =   1560
            Width           =   3372
         End
      End
      Begin VB.CommandButton cmdGra 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   675
         Left            =   3975
         Picture         =   "frmingresoOC.frx":106A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   5250
         Width           =   775
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5310
         Picture         =   "frmingresoOC.frx":14AC
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   5250
         Width           =   775
      End
      Begin VB.CommandButton CmdSalirproducto 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   675
         Left            =   -66240
         Picture         =   "frmingresoOC.frx":18EE
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   4440
         Width           =   775
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   675
         Left            =   -67320
         Picture         =   "frmingresoOC.frx":1D30
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   4440
         Width           =   825
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   15
         Left            =   -72840
         TabIndex        =   1
         Top             =   5760
         Width           =   135
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGridarticulos 
         Height          =   3375
         Left            =   -74760
         TabIndex        =   4
         Top             =   600
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   5953
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "item"
         Columns(0).DataField=   "oc_citem"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Codigo"
         Columns(1).DataField=   "oc_ccodigo"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Desripcion"
         Columns(2).DataField=   "oc_cdesref"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Comentario Adicional"
         Columns(3).DataField=   "oc_ccomen2"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   4
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=4"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=794"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=714"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=1773"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1693"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=7091"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=7011"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=6376"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=6297"
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
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuarticulo 
         Height          =   345
         Left            =   -73635
         TabIndex        =   5
         Top             =   4440
         Width           =   5340
         _ExtentX        =   9419
         _ExtentY        =   609
         XcodMaxLongitud =   20
         xcodwith        =   1000
         NomTabla        =   "maeart"
         ListaCampos     =   "acodigo(1),adescri(1),acodigo2(2),aunidad(2)"
         XcodCampo       =   "acodigo"
         XListCampo      =   "adescri"
         ListaCamposDescrip=   "Vodigo,Descripcion"
         ListaCamposText =   "acodigo,adescri,acodigo2,aunidad"
      End
      Begin TrueOleDBGrid70.TDBGrid DBGrid1 
         Height          =   1905
         Left            =   240
         TabIndex        =   9
         Top             =   3120
         Width           =   9570
         _ExtentX        =   16880
         _ExtentY        =   3360
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=15,.bold=0,.fontsize=825,.italic=0"
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
      Begin TrueOleDBGrid70.TDBGrid TDBGridordenes 
         Height          =   4215
         Left            =   -74760
         TabIndex        =   41
         Top             =   600
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   7435
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Tipo"
         Columns(0).DataField=   "tipoordencodigo"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Numero"
         Columns(1).DataField=   "oc_cnumord"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Proveedor"
         Columns(2).DataField=   "oc_crazsoc"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Descripcion del estado"
         Columns(3).DataField=   "estadodescripcion"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   4
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=4"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=900"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=820"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2011"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1931"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=5715"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=5636"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=6932"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=6853"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=15,.bold=0,.fontsize=825,.italic=0"
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
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   -74640
         TabIndex        =   6
         Top             =   4470
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmingresoOC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rbusca1 As New ADODB.Recordset
Dim rsdeta2 As New ADODB.Recordset
Dim adodc1 As New ADODB.Recordset
Dim Adodc2 As New ADODB.Recordset
Dim Adodc3 As New ADODB.Recordset
Dim ok As Integer
Dim conexion As String
Dim dllgeneral As New dllgeneral.dll_general
Public error As Integer
Dim tipo As String
Dim SQL As String
Dim nTra As Integer
Dim Mensaje As String

Private Sub Cmd_Revalorizar_Click()

 rsdeta2.Fields(1) = Ctr_Ayuarticulo.xclave
 rsdeta2.Fields(2) = Left(Ctr_Ayuarticulo.xnombre, 30)
 dbGrid1.SetFocus
 End Sub

Public Sub cmdAceptar_Click()
Dim SQL As String
Dim rs1 As New ADODB.Recordset
If Ctr_Ayuarticulo.xclave <> "" Then
   SQL = " Update co_detordcompra set oc_ccodigo='" & Ctr_Ayuarticulo.xclave & "'"
   SQL = SQL & " , oc_cdesref='" & Left(Ctr_Ayuarticulo.xnombre, 65) & "' where  tipoordencodigo='" & rbusca1!tipoordencodigo & "' and oc_cnumord='"
   SQL = SQL & rbusca1!oc_cnumord & "' and oc_citem='" & rbusca1!oc_citem & "'"
   Set rs1 = VGCNx.Execute(SQL)
End If
Cargaproducto
TDBGridarticulos.Refresh
End Sub

Private Sub cmdGra_Click()
    Dim SQLc As String
    Dim SQLd As String
    Dim I As Integer, TipoIng As Integer
    Dim vNI As Integer
    Dim vNF As Single, vNC As Single
    Dim vNP As Single, vNP1 As Single
    Dim TIPOMOV As String
    Dim tipodoc As String
    Dim CANTIDAD As String
    Dim txtNTF1 As String
    On Error GoTo GrabErr
    tipodoc = "NI"
        If CDate(txtfec.Value) < CDate(lblEmi) Then
            Mensaje = "La Fecha debe ser igual o posterior que la Fecha de emisión"
            MsgBox Mensaje, vbExclamation, "Mensaje"
            txtfec.SetFocus
            Exit Sub
        End If
    
    txtTF = Trim(txtTF)
    If txtTF = "" Then
        Mensaje = "Debe especificar Tipo de Documento"
        MsgBox Mensaje, vbExclamation, "Mensaje"
        txtTF.SetFocus
        Exit Sub
    
    Else
        If Not Existe(1, txtTF, "tipo_docu", "tdo_tipdoc", False) Then
            Mensaje = "No existe el Tipo de Documento ingresado"
            MsgBox Mensaje, vbExclamation, "Error"
            txtTF.SetFocus
            Exit Sub
        Else
            If lblTF = "" Then

                lblTF = Devolver_Dato(1, txtTF, "tipo_docu", "tdo_tipdoc", False, _
                    "tdo_descri")
                If TIPOMOV = "S" Then tipodoc = "NS"
            End If
        End If
    End If
    txtNTF.Enabled = True
    TIPOMOV = Devolver_Dato(1, txtTM, "tabtransa", "tt_codmov", False, "tt_tipmov")
    txtserie = Trim(txtserie)
    If txtserie = "" Then
        Mensaje = "Debe especificar Serie del Documento"
        MsgBox Mensaje, vbExclamation, "Mensaje"
        txtserie.SetFocus
        Exit Sub
    End If
    txtNTF = Trim(txtNTF)
    If txtNTF = "" Then
        Mensaje = "Debe especificar el Número de Documento"
        MsgBox Mensaje, vbExclamation, "Mensaje"
        txtNTF.SetFocus
        Exit Sub
    End If
    txtAlm = Trim(txtAlm)
    If txtAlm = "" Then
        Mensaje = "Debe especificar Código de Almacen"
        MsgBox Mensaje, vbExclamation, "Mensaje"
        txtAlm.SetFocus
        Exit Sub
    Else
        If Not Existe(1, txtAlm, "tabalm", "taalma", False) Then
            Mensaje = "El Código de Almacén ingresado no existe"
            MsgBox Mensaje, vbExclamation, "Error"
            txtAlm.SetFocus
            Exit Sub
        Else
            If lblAlm = "" Then lblAlm = Devolver_Dato(1, txtAlm, "tabalm", "taalma", _
                False, "tadescri")
        End If
    End If

    txtTM = Trim(txtTM)
    If txtTM = "" Then
        Mensaje = "Debe especificar Tipo de Movimiento"
        MsgBox Mensaje, vbExclamation, "Mensaje"
        txtTM.SetFocus
        Exit Sub
    Else
        If Not Existe(1, txtTM, "TABTRANSA", "TT_CODMOV", False) Then
            Mensaje = "El Tipo de Movimiento ingresado no existe"
            MsgBox Mensaje, vbExclamation, "Error"
            txtTM.SetFocus
        Else
            If lblTM = "" Then lblTM = Devolver_Dato(1, txtTM, "TABTRANSA", "TT_CODMOV", _
                False, "TT_DESCRI")
        End If
    End If
    TipoIng = Ingreso_Realizado
    If TipoIng = 0 Then
        Mensaje = "No se puede grabar." & vbCrLf & "No se ha recepcionado ningún artículo"
        MsgBox Mensaje, vbExclamation, "Mensaje"
        Exit Sub
    End If
If error = 1 Then
        Mensaje = "Existen Item que no tienen codigo de articulo, verifique " & vbCrLf & "codigo de artículo"
        MsgBox Mensaje, vbExclamation, "Mensaje"
        Exit Sub
End If
Mensaje = "¿Desea guardar los cambios realizados?"
If MsgBox(Mensaje, vbQuestion + vbYesNo, "Mensaje") = vbYes Then
   txtNTF1 = Format(txtserie, "0000") + Format(txtNTF, "0000000000")
   nTra = 1
   VGCNx.BeginTrans
   If TIPOMOV = "I" Then
      vNI = Devolver_Dato(1, txtAlm, "tabalm", "taalma", False, "tanument")
      vNI = vNI + 1
      SQLc = "UPDATE tabalm SET tanument=" & vNI & " WHERE taalma='" & txtAlm & "'"
    Else
      vNI = Devolver_Dato(1, txtAlm, "tabalm", "taalma", False, "tanumsal")
      vNI = vNI + 1
      SQLc = "UPDATE tabalm SET tanumsal=" & vNI & " WHERE taalma='" & txtAlm & "'"
   End If
   VGCNx.Execute SQLc
   VGCNx.CommitTrans
   VGCNx.BeginTrans
   SQLc = "UPDATE co_cabordcompra SET estadooccodigo='" & Format(TipoIng, "0") & _
        "' WHERE tipoordencodigo='" & tipo & "' and oc_cnumord='" & txtNum & "'"
   VGCNx.Execute SQLc
   If rsdeta2.RecordCount > 0 Then
      rsdeta2.MoveFirst
      Do Until rsdeta2.EOF
         If rsdeta2.Fields(6) > 0 Then
            SQLd = "UPDATE co_detordcompra SET oc_ncanten=oc_ncanten +" & rsdeta2.Fields(6) & ","
            SQLd = SQLd & "oc_nsaldo =" & rsdeta2.Fields(5) - rsdeta2.Fields(6) & ","
            SQLd = SQLd & "estadooccodigo='" & Format(TipoIng, "0") & "'"
            SQLd = SQLd & " WHERE tipoordencodigo='" & tipo & "' and oc_cnumord='" & txtNum & "' AND oc_citem ='" & rsdeta2.Fields(0) & "'"
          VGCNx.Execute SQLd
        End If
          rsdeta2.MoveNext
      Loop
   End If
   VGCNx.CommitTrans
   SQLd = "INSERT INTO movalmcab (caalma,catd,canumdoc,cafecdoc,catipmov,cacodmov," & _
        "carftdoc,carfndoc,cacodpro,cafecact,cahora,causuari,cacodmon," & _
        "canumord) VALUES ('" & txtAlm & "','" & tipodoc & "','" & Format(vNI, "00000000000") & _
        "','" & txtfec & "','I','" & txtTM & "','" & _
        txtTF & "','" & txtNTF1 & "','" & LblPro & "',getdate(),'" & Format(Time, "hh.mm.ss") & "','" & VGUsuario & "','" & " " & "','" & _
        tipo & txtNum & "')"
   VGCNx.BeginTrans
   VGCNx.Execute SQLd
   VGCNx.CommitTrans
   VGCNx.BeginTrans
   If rsdeta2.RecordCount > 0 Then
      rsdeta2.MoveFirst
      vNP1 = 0
      I = 0
      Do Until rsdeta2.EOF
         If Val(rsdeta2.Fields(6)) > 0 Then
         ' se a colocado a 0 por esta unica vez
           If VGtipoAprobacion = 1 Then
              vNP = Devolver_Dato(1, txtNum, "co_detordcompra", "oc_cnumord", False, "oc_nprenet", _
                    rsdeta2.Fields(1), "oc_ccodigo")
             Else
               vNP = 0
             End If
            vNP1 = vNP * Val(rsdeta2.Fields(6))
            I = I + 1
            SQLd = "INSERT INTO movalmdet (dealma,detd,denumdoc,deitem,decodigo," & _
                 "decantid,deprecio,defecdoc,deestado,decodmov,devaltot,decodmon," & _
                 "desoli) VALUES ('" & txtAlm & "','" & tipodoc & "','" & Format(vNI, "00000000000") & _
                 "'," & I & ",'" & rsdeta2.Fields(1) & "'," & _
                 Val(rsdeta2.Fields(6)) & "," & vNP & ",'" & _
                 txtfec & "','V','" & txtTM & "'," & _
                 vNP1 & ",'" & " " & _
                 "','" & " " & "')"
             VGCNx.Execute SQLd
            CANTIDAD = Val(rsdeta2.Fields(6))
            If TIPOMOV = "S" Then CANTIDAD = CANTIDAD * -1
         
         If Existe(1, rsdeta2.Fields(1), "stkart", "stcodigo", False, txtAlm, _
            "stalma") Then
            SQLd = "UPDATE stkart SET stskdis=stskdis+ (" & _
                 CANTIDAD & ") WHERE stalma='" & txtAlm & _
                 "'AND stcodigo='" & rsdeta2.Fields(1) & "'"
           Else
            SQLd = "INSERT INTO stkart (stalma,stcodigo,stskdis) VALUES ('" & txtAlm & _
                 "','" & rsdeta2.Fields(1) & "'," & CANTIDAD & ")"
         End If
         VGCNx.Execute SQLd
         End If
         rsdeta2.MoveNext
      Loop
   End If
   VGCNx.CommitTrans
   nTra = 0
  'adodc1.Requery
   tipodoc = "Nota de Ingreso "
   If TIPOMOV = "S" Then tipodoc = "Nota de Salida "
   Mensaje = "Se Proceso el documento " & vbCrLf & tipodoc & _
           Format(vNI, "00000000000") & vbCrLf & " Almacén : " & txtAlm & vbCrLf & _
          "Tipo de Movimiento :  " & txtTM
   MsgBox Mensaje, vbInformation, "Ingreso"
   Dim rpta As String
   rpta = MsgBox("¿Desea Imprimir?", vbYesNo + vbQuestion, "Aviso")
   If rpta = vbYes Then
    Call imprimir(vNI, txtNTF1)
  End If
  txtNum = ""
  txtNum.SetFocus
End If
Exit Sub
GrabErr:
    MsgBox Err.Description
 '  Resume
    If nTra = 1 Then VGCNx.RollbackTrans
        Mensaje = "No se puede grabar." & vbCrLf & "No se ha recepcionado ningún artículo"
        MsgBox Mensaje, vbExclamation, "Mensaje"
        Exit Sub
        Resume
End Sub
        
        
Private Sub CmdSalir_Click()
    Call dllgeneral.ActivaTab(0, 1, SSTab1)
    grillainicial
End Sub
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Ctrayu_tipoorden.xclave = TDBGridordenes.Columns(0)
Ctrayu_tipoorden.Ejecutar
txtNum.text = TDBGridordenes.Columns(1)
txtnum_KeyPress (13)
Call dllgeneral.ActivaTab(1, 1, SSTab1)
End Sub

Private Sub Ctrayu_tipoorden_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
tipo = Ctrayu_tipoorden.xclave
If ColecCampos("flagrequerimientosordenes") = "1" Then
   VGtipoAprobacion = 0
 Else
   VGtipoAprobacion = 1
End If
End Sub


Private Sub Form_Load()
    central Me
    Call Ctrayu_tipoorden.conexion(VGCNx)
    If Not VGparametros.PermiteIngresosconRequerimientos Then
       Ctrayu_tipoorden.filtro = " isnull(flagrequerimientos,0)= 0 and ordendebienes='B'"
     Else
       Ctrayu_tipoorden.filtro = " ordendebienes='B'"
    End If
    Call Ctr_Ayuarticulo.conexion(VGCNx)
   Call dllgeneral.ActivaTab(0, 1, SSTab1)
   grillainicial
  End Sub

Sub Limpiar()
    LblPro = "": lblProv = "": lblRuc = ""
    lblEmi = "": lblEnt = "": lblEst = ""
    lblEsta = "": txtTF = ""
    txtAlm = "": txtTM = ""

End Sub

Private Sub TDBGridarticulos_Click()
Ctr_Ayuarticulo.xclave = rbusca1!oc_ccodigo
End Sub

Private Sub txtAlm_Change()
    If lblAlm <> "" Then lblAlm = ""
End Sub

Private Sub txtAlm_DblClick()
    Static Adodc2 As ADODB.Recordset
    Dim strsql As String
    
    Set Adodc2 = New ADODB.Recordset
    
    strsql = "SELECT taalma,tadescri FROM tabalm"
    Adodc2.Open strsql, VGCNx, adOpenStatic, adLockReadOnly
    
    frmReferencia.Conectar Adodc2, strsql
    frmReferencia.Label1 = "Lista de Almacenes"
    frmReferencia.inicio
    frmReferencia.Show vbModal

    
    If vGUtil(1) <> "" Then
        txtAlm = vGUtil(1)
        lblAlm = vGUtil(2)
        txtTM.SetFocus
    End If
End Sub

Private Sub txtAlm_GotFocus()
    Enfoque txtAlm
End Sub

Private Sub txtAlm_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then txtAlm_DblClick
End Sub

Private Sub txtAlm_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtAlm = Trim(txtAlm)
        If txtAlm <> "" Then
            If Not Existe(1, txtAlm, "tabalm", "taalma", False) Then
                Mensaje = "El Código de Almacén ingresado no existe"
                MsgBox Mensaje, vbExclamation, "Error"
                txtAlm.SetFocus
            Else
                lblAlm = Devolver_Dato(1, txtAlm, "tabalm", "taalma", False, "tadescri")
                txtTM.SetFocus
            End If
        Else
            txtTM.SetFocus
        End If
    End If
    Enteros_Positivos KeyAscii, txtAlm
End Sub

Private Sub txtFec_GotFocus()
   ' txtfec.SelStart = 0
   ' txtfec.SelLength = 12
End Sub

Private Sub txtFec_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not ValidFecha(txtfec) Then
            Mensaje = "Fecha No Válida"
            MsgBox Mensaje, vbExclamation, "Error"
            txtfec.SetFocus
        Else
            If CDate(txtfec.Value) < CDate(lblEmi) Then
                Mensaje = "La Fecha debe ser igual o posterior que la Fecha de emisión"
                MsgBox Mensaje, vbExclamation, "Mensaje"
                txtfec.SetFocus
            Else
                txtTF.SetFocus
            End If
        End If
    End If
End Sub

Private Sub txtNTF_GotFocus()
    Enfoque txtNTF
End Sub

Private Sub txtNTF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtNTF = Trim(txtNTF)
        txtAlm.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub Txtnum_Change()
    If LblPro <> "" Then
        Limpiar
        Fradatos.Enabled = False
        cmdGra.Enabled = False
    End If
End Sub

Private Sub txtNum_DblClick()
    Set Adodc2 = New ADODB.Recordset
If VGtipoAprobacion = 1 Then
    SQL = "SELECT a.oc_cnumord, a.oc_dfecdoc, b.estadoocdescripcion " & _
             " FROM co_cabordcompra a inner join co_estadoorden b " & _
             " on a.estadooccodigo=b.estadooccodigo inner join co_tipodeorden c " & _
             " on a.tipoordencodigo = c.tipoordencodigo where  b.estadoocatendido=0 " & _
             " and a.oc_estadoorden<>1 and a.tipoordencodigo='" & tipo & "'"
Else
    SQL = "SELECT a.oc_cnumord, a.oc_dfecdoc, b.estadoocdescripcion " & _
             " FROM co_cabordcompra a inner join co_estadorequerimiento b " & _
             " on a.estadooccodigo=b.estadooccodigo inner join co_tipodeorden c " & _
             " on a.tipoordencodigo = c.tipoordencodigo where b.estadooccodigo=1 " & _
             " and a.oc_estadoorden<>1 and a.tipoordencodigo='" & tipo & "'"
End If
Adodc2.Open SQL, VGCNx, adOpenStatic, adLockReadOnly
    frmReferencia1.Conectar Adodc2, SQL
    frmReferencia1.Caption = "Ingresos de Ordenes "
    frmReferencia1.inicio
    frmReferencia1.Show vbModal
    Adodc2.Close
    frmReferencia.Conectar Adodc2, "Select ACODIGO, ADESCRI,AUNIDAD from MaeArt"
     
    If vGUtil(1) <> "" Then
        txtNum = vGUtil(1)
        ok = 0
      '  lblSol = vGUtil(2)
      '  txtcen.SetFocus
    End If
End Sub

Private Sub txtNum_GotFocus()
    Set dbGrid1.DataSource = Nothing
    Enfoque txtNum
End Sub

Private Sub txtNum_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then txtNum_DblClick
End Sub

Private Sub txtnum_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtNum <> "" Then
            txtNum = Format(txtNum, "00000000000")
            If Not Existe(1, txtNum, "co_cabordcompra", "oc_cnumord", False) Then
                Mensaje = "El Número de Orden de Compra ingresado no existe"
                MsgBox Mensaje, vbExclamation, "Mensaje"
                Enfoque txtNum
                txtNum.SetFocus
                Exit Sub
            Else
                If Not Estado_Valido Then
                    Mensaje = "Estado de Orden de compra no válido"
                    MsgBox Mensaje, vbExclamation, "Mensaje"
                    Enfoque txtNum
                    txtNum.SetFocus
                    Exit Sub
                Else
                    Muestra_datos_de_co_cabordcompra
                    CargaGrilla
                    Fradatos.Enabled = True
                    cmdGra.Enabled = True
                    txtfec = VGParamSistem.FechaTrabajo
                    txtfec.SetFocus
                End If
            End If
        End If
    End If
    If ok = 0 Then
       Cargaproducto
    End If
    Enteros_Positivos KeyAscii, txtNum
End Sub

Public Function CargaGrilla()
Call cargar_grilla2
Set rbusca1 = Nothing

SQL = "SELECT oc_citem,oc_ccodigo,oc_cdesref=case when oc_ccodigo='00' then oc_ccomen1 else oc_cdesref end,"
SQL = SQL & "ord_fabnum,oc_ncantid,oc_nsaldo,oc_nsaldo as oc_ncanten "
SQL = SQL & " FROM co_detordcompra WHERE tipoordencodigo='" & Ctrayu_tipoorden.xclave
SQL = SQL & "' and oc_cnumord='" & txtNum & "' and isnull(estadooccodigo,0)<>'2'  ORDER BY oc_citem"
Set rbusca1 = VGCNx.Execute(SQL)
If rbusca1.RecordCount > 0 Then
   rbusca1.MoveFirst
   Do Until rbusca1.EOF
      rsdeta2.AddNew
      rsdeta2.Fields(0) = rbusca1!oc_citem
      rsdeta2.Fields(1) = rbusca1!oc_ccodigo
      rsdeta2.Fields(2) = Escadena(Left(rbusca1!OC_CDESREF, 40))
      rsdeta2.Fields(3) = rbusca1!ord_fabnum
      rsdeta2.Fields(4) = rbusca1!oc_ncantid
      rsdeta2.Fields(5) = rbusca1!oc_nsaldo
      rsdeta2.Fields(6) = rbusca1!oc_ncanten
      rbusca1.MoveNext
   Loop
End If
rbusca1.Close
Set rbusca1 = Nothing
End Function

Public Function cargar_grilla2()

Set rsdeta2 = Nothing
Call rsdeta2.Fields.Append("Item", adVarChar, 20)
Call rsdeta2.Fields.Append("Codigo", adVarChar, 20)
Call rsdeta2.Fields.Append("Descripcion", adVarChar, 40)
Call rsdeta2.Fields.Append("Ord.Fab.", adVarChar, 20)
Call rsdeta2.Fields.Append("Cant.Pedida", adVarChar, 20)
Call rsdeta2.Fields.Append("Saldo", adVarChar, 20)
Call rsdeta2.Fields.Append("Cant.Alm.", adVarChar, 20)
rsdeta2.Open

Set dbGrid1.DataSource = Nothing
Set dbGrid1.DataSource = rsdeta2

dbGrid1.Columns(0).AllowFocus = False
dbGrid1.Columns(1).AllowFocus = False
dbGrid1.Columns(2).AllowFocus = False
dbGrid1.Columns(3).AllowFocus = False
dbGrid1.Columns(4).AllowFocus = False
dbGrid1.Columns(5).AllowFocus = False
dbGrid1.Columns(0).Width = 400
dbGrid1.Columns(1).Width = 1200
dbGrid1.Columns(2).Width = 4000
dbGrid1.Columns(3).Width = 1000
dbGrid1.Columns(4).Width = 800
dbGrid1.Columns(5).Width = 800
dbGrid1.Columns(6).Width = 800
dbGrid1.Columns(4).NumberFormat = "###,##0.00"
dbGrid1.Columns(5).NumberFormat = "###,##0.00"
dbGrid1.Columns(6).NumberFormat = "###,##0.00"

dbGrid1.Refresh

End Function

Function Estado_Valido() As Boolean
    Dim vest As String
    
    vest = Devolver_Dato(1, txtNum, "co_cabordcompra", "oc_cnumord", False, "estadooccodigo")
    Estado_Valido = False
    If vest <> "2" Then Estado_Valido = True
End Function

Sub Muestra_datos_de_co_cabordcompra()
    Static adodc1 As New ADODB.Recordset
    Set adodc1 = New ADODB.Recordset
    SQL = "SELECT oc_ccodpro,oc_crazsoc=clienterazonsocial,oc_dfecdoc,oc_dfecent,estadooccodigo,oc_ccodmon,"
    SQL = SQL & "oc_csolict FROM co_cabordcompra A left join cp_proveedor b "
    SQL = SQL & " on a.oc_ccodpro=b.clientecodigo WHERE tipoordencodigo='" & Ctrayu_tipoorden.xclave
    SQL = SQL & "' and oc_cnumord='" & txtNum & "'"
    adodc1.Open SQL, VGCNx, adOpenDynamic, adLockOptimistic
    
    LblPro = Escadena(adodc1("oc_ccodpro"))
    lblProv = Escadena(adodc1("oc_crazsoc"))
    lblRuc = Devolver_Dato(1, LblPro, "cp_proveedor", "clientecodigo", False, "clienteruc")
    lblEmi = adodc1("oc_dfecdoc")
    lblEnt = adodc1("oc_dfecent")
    lblEst = adodc1("estadooccodigo")
    If VGtipoAprobacion = 1 Then
       lblEsta = Devolver_Dato(1, lblEst, "co_estadoorden", "estadooccodigo", False, "estadoocdescripcion")
    Else
       lblEsta = Devolver_Dato(1, lblEst, "co_estadorequerimiento", "estadooccodigo", False, "estadoocdescripcion")
       txtTF.text = Ctrayu_tipoorden.xclave
       txtNTF.text = txtNum.text
    End If
    
End Sub

Private Sub txtTF_Change()
    If lblTF <> "" Then
        lblTF = ""
        txtNTF = ""
        txtNTF.Enabled = True
    End If
End Sub

Private Sub txtTF_DblClick()
    Static Adodc2 As ADODB.Recordset
    Dim strsql As String
    
    Set Adodc2 = New ADODB.Recordset
    
    strsql = "SELECT tdo_tipdoc,tdo_descri FROM tipo_docu"
    Adodc2.Open strsql, VGCNx, adOpenStatic, adLockReadOnly
    
    frmReferencia.Conectar Adodc2, strsql
    frmReferencia.Label1 = "Tipo de Documentos"
    frmReferencia.inicio
    frmReferencia.Show vbModal
    
    If vGUtil(1) <> "" Then
        txtTF = vGUtil(1)
    End If
End Sub

Private Sub txtTF_GotFocus()
    Enfoque txtTF
End Sub

Private Sub txtTF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then txtTF_DblClick
End Sub

Private Sub txtTF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtTF = Trim(txtTF)
        If txtTF <> "" Then
            If Not Existe(1, txtTF, "tipo_docu", "tdo_tipdoc", False) Then
                Mensaje = "No existe el Tipo de documento ingresado"
                MsgBox Mensaje, vbExclamation, "Error"
                txtTF.SetFocus
            Else
                lblTF = Devolver_Dato(1, txtTF, "tipo_docu", "tdo_tipdoc", False, _
                    "tdo_descri")
                txtNTF.Enabled = True
                txtNTF.SetFocus
            End If
        Else
            txtAlm.SetFocus
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtTM_Change()
    If lblTM <> "" Then lblTM = ""
End Sub

Private Sub txtTM_DblClick()
    Static Adodc2 As ADODB.Recordset
    Set Adodc2 = New ADODB.Recordset
    SQL = "SELECT tt_codmov,tt_descri FROM tabtransa"
    Adodc2.Open SQL, VGCNx, adOpenStatic, adLockReadOnly
    frmReferencia.Conectar Adodc2, SQL
    frmReferencia.Label1 = "Tipo de Transaccion"
    frmReferencia.inicio
    frmReferencia.Show vbModal

    If vGUtil(1) <> "" Then
        txtTM = vGUtil(1)
        txtTM_KeyPress 13
        dbGrid1.SetFocus
    End If
End Sub

Private Sub txtTM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then txtTM_DblClick
End Sub

Private Sub txtTM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtTM = Trim(txtTM)
        If txtTM <> "" Then
            If Not Existe(1, txtTM, "tabtransa", "tt_codmov", False) Then
                Mensaje = "El Código de transaccion no existe"
                MsgBox Mensaje, vbExclamation, "Error"
                txtTM.SetFocus
            Else
                lblTM = Devolver_Dato(1, txtTM, "tabtransa", "tt_codmov", False, "tt_descri")
            End If
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Function Ingreso_Realizado() As Integer
    Dim I As Integer
    Dim tSal As Single
    Dim tRec As Single
If rsdeta2.RecordCount > 0 Then
   rsdeta2.MoveFirst
   Do Until rsdeta2.EOF
      If (rsdeta2.Fields(1) = "00" Or rsdeta2.Fields(1) = "") And rsdeta2.Fields(6) > 0 Then
         error = 1
      End If
      tSal = tSal + rsdeta2.Fields(5)
      tRec = tRec + rsdeta2.Fields(6)
      rsdeta2.MoveNext
   Loop
End If
If tRec = 0 Then
   Ingreso_Realizado = 0
    ElseIf tRec < tSal Then
        Ingreso_Realizado = 1
    Else
        Ingreso_Realizado = 2
    End If
End Function
Public Sub Cargaproducto()
Dim rsproducto As New ADODB.Recordset
SQL = "SELECT tipoordencodigo,oc_cnumord,oc_citem,oc_ccodigo,oc_cdesref,OC_ccomen2 from co_detordcompra "
SQL = SQL & " WHERE tipoordencodigo='" & Ctrayu_tipoorden.xclave & "'"
SQL = SQL & " and  oc_cnumord='" & txtNum & "'and oc_ccodigo='00'"

Set rbusca1 = VGCNx.Execute(SQL)
If rbusca1.RecordCount > 0 Then
    Call dllgeneral.ActivaTab(2, 1, SSTab1)
 Else
   ok = 1
   Call dllgeneral.ActivaTab(1, 1, SSTab1)
   txtnum_KeyPress (13)
End If
TDBGridarticulos.DataSource = rbusca1
TDBGridarticulos.Refresh
End Sub
Private Sub grillainicial()
SQL = "SELECT a.tipoordencodigo,a.oc_cnumord,oc_ccodpro, oc_crazsoc,a.oc_dfecdoc, b.estadoocdescripcion " & _
     " FROM co_cabordcompra a inner join co_estadoorden b " & _
             " on a.estadooccodigo=b.estadooccodigo inner join co_tipodeorden c " & _
             " on a.tipoordencodigo = c.tipoordencodigo where  b.estadoocatendido=0 " & _
             " and a.oc_estadoorden<>1 and ordeningresoalmacen=1 "
'SQL = SQL & " union all "
' SQL = SQL & " SELECT a.tipoordencodigo,a.oc_cnumord, oc_ccodpro,oc_crazsoc,a.oc_dfecdoc, b.estadoocdescripcion " & _
'             " FROM co_cabordcompra a inner join co_estadorequerimiento b " & _
             " on a.estadooccodigo=b.estadooccodigo inner join co_tipodeorden c " & _
             " on a.tipoordencodigo = c.tipoordencodigo where b.estadooccodigo=1 " & _
             " and a.oc_estadoorden<>1 "
Set rsdeta2 = New ADODB.Recordset
Set rsdeta2 = VGCNx.Execute(SQL)
TDBGridordenes.DataSource = rsdeta2
TDBGridordenes.Refresh
End Sub

Private Sub imprimir(ByRef NumDoc As Integer, ByRef numrefere As String)
    Dim CADENA As String
    Dim cFormato As String
    Dim cDireccion As String
    Dim cRuc As String
    Dim cNomRepor  As String
    Dim aBusca As New ADODB.Recordset
    Dim numdoc1 As String
    numdoc1 = Format(NumDoc, "00000000000")
                           CrystalReport1.Reset
                            cNomRepor = "REPNOTAING.rpt"
                            CrystalReport1.ReportFileName = VGParamSistem.RutaReport & cNomRepor
               
                            CrystalReport1.Connect = VGcadenareport2
                            CrystalReport1.StoredProcParam(0) = VGParamSistem.BDEmpresa
                            CrystalReport1.StoredProcParam(1) = txtAlm
                            CrystalReport1.StoredProcParam(2) = "NI"
                            CrystalReport1.StoredProcParam(3) = numdoc1
                            CrystalReport1.DiscardSavedData = True
                            CrystalReport1.Destination = crptToWindow
                            CrystalReport1.formulas(0) = "fecha='" & txtfec.Value & "'"
                            CrystalReport1.formulas(1) = "xtrans = '" & lblTM.Caption & "' "
                            CrystalReport1.formulas(2) = "xtd = 'NI'"
                            CrystalReport1.formulas(3) = "xndoc = '" & numdoc1 & "' "
                            CrystalReport1.WindowTitle = "RepNotaIng -- Impresion de Notas de Ingreso"
                                CrystalReport1.formulas(4) = "Xnalma = ' '"
                                CrystalReport1.formulas(5) = "Dalma = ' ' "
                                CrystalReport1.formulas(6) = "AlmaDes ='" & txtAlm & "' "
                                CrystalReport1.formulas(7) = "Dalmades = '" & lblAlm.Caption & "' "
                            CrystalReport1.formulas(8) = "NRef = '" & numrefere & "' "
                            CrystalReport1.formulas(9) = "DocRef = '" & txtTF & "' "
                            CrystalReport1.formulas(10) = "TTrans = '" & txtTM & "' "
                            CrystalReport1.formulas(11) = "emp = '" & VGparametros.RucEmpresa & "'"
                            CrystalReport1.WindowShowPrintBtn = True
                            CrystalReport1.WindowShowRefreshBtn = True
                            CrystalReport1.WindowShowSearchBtn = True
                            CrystalReport1.WindowShowPrintSetupBtn = True
                            CrystalReport1.WindowState = crptMaximized
                            If CrystalReport1.Status <> 2 Then
                                CrystalReport1.Action = 1
                            End If
        Exit Sub
ErrImp:
     MsgBox Err.Description
     Resume Next
End Sub

