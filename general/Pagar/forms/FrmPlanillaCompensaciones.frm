VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmPlanillaCompensaciones 
   Caption         =   "Compensaciones"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12240
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   12240
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   9045
      Left            =   150
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15954
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "FrmPlanillaCompensaciones.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(1)=   "Frame2"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "FrmPlanillaCompensaciones.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame7"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   4305
         Left            =   330
         TabIndex        =   35
         Top             =   4680
         Width           =   11175
         Begin VB.Frame Frame6 
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            Height          =   840
            Left            =   3750
            TabIndex        =   50
            Top             =   3180
            Width           =   4230
            Begin VB.CommandButton cmdBotones 
               Caption         =   "&Salir"
               Height          =   690
               Index           =   7
               Left            =   3255
               Picture         =   "FrmPlanillaCompensaciones.frx":0038
               Style           =   1  'Graphical
               TabIndex        =   54
               Top             =   90
               Width           =   870
            End
            Begin VB.CommandButton cmdBotones 
               Caption         =   "&Eliminar"
               Height          =   690
               Index           =   6
               Left            =   2280
               Picture         =   "FrmPlanillaCompensaciones.frx":047A
               Style           =   1  'Graphical
               TabIndex        =   53
               Top             =   90
               Width           =   825
            End
            Begin VB.CommandButton cmdBotones 
               Caption         =   "&Grabar"
               Height          =   690
               Index           =   5
               Left            =   1320
               Picture         =   "FrmPlanillaCompensaciones.frx":08BC
               Style           =   1  'Graphical
               TabIndex        =   52
               Top             =   120
               Width           =   870
            End
            Begin VB.CommandButton cmdBotones 
               Caption         =   "&Nuevo"
               Height          =   690
               Index           =   4
               Left            =   180
               Picture         =   "FrmPlanillaCompensaciones.frx":0CFE
               Style           =   1  'Graphical
               TabIndex        =   51
               Top             =   90
               Width           =   870
            End
         End
         Begin VB.Frame Frame8 
            Height          =   795
            Left            =   150
            TabIndex        =   37
            Top             =   2310
            Width           =   10845
            Begin VB.CommandButton cAyuda 
               Caption         =   "..."
               Height          =   315
               Index           =   2
               Left            =   8160
               TabIndex        =   57
               Top             =   360
               Width           =   195
            End
            Begin VB.TextBox Text2 
               Height          =   285
               Index           =   4
               Left            =   9270
               MaxLength       =   12
               TabIndex        =   42
               Top             =   420
               Width           =   1125
            End
            Begin VB.TextBox Text2 
               Height          =   285
               Index           =   3
               Left            =   8565
               MaxLength       =   2
               TabIndex        =   41
               Top             =   420
               Width           =   480
            End
            Begin VB.TextBox Text2 
               Height          =   285
               Index           =   2
               Left            =   6915
               MaxLength       =   10
               TabIndex        =   40
               Top             =   420
               Width           =   1110
            End
            Begin VB.TextBox Text2 
               Height          =   285
               Index           =   1
               Left            =   6240
               MaxLength       =   4
               TabIndex        =   39
               Top             =   420
               Width           =   555
            End
            Begin VB.TextBox Text2 
               Height          =   285
               Index           =   0
               Left            =   5820
               MaxLength       =   2
               TabIndex        =   38
               Top             =   420
               Width           =   375
            End
            Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuProv 
               Height          =   300
               Left            =   45
               TabIndex        =   43
               Top             =   405
               Width           =   5670
               _ExtentX        =   10001
               _ExtentY        =   529
               XcodMaxLongitud =   11
               xcodwith        =   1200
               NomTabla        =   "cp_proveedor"
               ListaCampos     =   "clientecodigo(1),clienterazonsocial(1)"
               XcodCampo       =   "clientecodigo"
               XListCampo      =   "clienterazonsocial"
               ListaCamposDescrip=   "Codigo,Descripcion"
               ListaCamposText =   "clientecodigo,clienterazonsocial"
               Requerido       =   0   'False
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Caption         =   "Importe"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   13
               Left            =   9315
               TabIndex        =   49
               Top             =   180
               Width           =   1035
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Caption         =   "Mon."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   14
               Left            =   8580
               TabIndex        =   48
               Top             =   240
               Width           =   525
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Caption         =   "Numero"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   17
               Left            =   6930
               TabIndex        =   47
               Top             =   180
               Width           =   765
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Caption         =   "Serie"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   18
               Left            =   6315
               TabIndex        =   46
               Top             =   180
               Width           =   585
            End
            Begin VB.Label Label2 
               Caption         =   "TD"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   19
               Left            =   6000
               TabIndex        =   45
               Top             =   180
               Width           =   315
            End
            Begin VB.Label Label8 
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
               Height          =   165
               Left            =   1680
               TabIndex        =   44
               Top             =   150
               Width           =   1380
            End
         End
         Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
            Height          =   1965
            Left            =   135
            TabIndex        =   36
            Top             =   390
            Width           =   10845
            _ExtentX        =   19129
            _ExtentY        =   3466
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
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=System"
            PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=System"
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=43,.bold=0,.fontsize=825,.italic=0"
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
         Begin VB.Label Label6 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Index           =   1
            Left            =   9660
            TabIndex        =   63
            Top             =   3240
            Width           =   1275
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
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
            Height          =   225
            Index           =   1
            Left            =   8880
            TabIndex        =   62
            Top             =   3360
            Width           =   645
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            Caption         =   "Importe"
            DataField       =   "Importe"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   1
            Left            =   240
            TabIndex        =   59
            Top             =   3720
            Visible         =   0   'False
            Width           =   2595
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            Caption         =   "Importe"
            DataField       =   "Importe"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   58
            Top             =   3240
            Visible         =   0   'False
            Width           =   2595
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Documentos por Pagar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   55
            Top             =   0
            Width           =   10785
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H0000C0C0&
            BorderWidth     =   2
            FillColor       =   &H00FFC0C0&
            FillStyle       =   0  'Solid
            Height          =   3915
            Index           =   1
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   120
            Width           =   11175
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Height          =   3480
         Left            =   -72818
         TabIndex        =   27
         Top             =   2220
         Width           =   7680
         Begin VB.Frame Frame1 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            Height          =   3120
            Left            =   150
            TabIndex        =   28
            Top             =   240
            Width           =   7425
            Begin MSMask.MaskEdBox MBox1 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "MM/dd/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   3
               EndProperty
               Height          =   300
               Left            =   3450
               TabIndex        =   29
               Top             =   1110
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuTipoPla 
               Height          =   285
               Left            =   3480
               TabIndex        =   30
               Top             =   720
               Width           =   3645
               _ExtentX        =   6429
               _ExtentY        =   503
               XcodMaxLongitud =   2
               xcodwith        =   150
               NomTabla        =   "cp_tipoplanilla"
               TituloAyuda     =   "Ayuda de Tipo de Planilla"
               ListaCampos     =   "tplanillacodigo(1),tplanilladesccorta(1)"
               XcodCampo       =   "tplanillacodigo"
               XListCampo      =   "tplanilladesccorta"
               ListaCamposDescrip=   "Codigo,Descripcion"
               ListaCamposText =   "tplanillacodigo,tplanilladesccorta"
               Requerido       =   0   'False
            End
            Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuoficina 
               Height          =   285
               Left            =   3450
               TabIndex        =   31
               Top             =   1680
               Width           =   3645
               _ExtentX        =   6429
               _ExtentY        =   503
               XcodMaxLongitud =   3
               xcodwith        =   200
               NomTabla        =   "cp_oficina"
               ListaCampos     =   "vendedorcodigo(1),vendedornombres(1)"
               XcodCampo       =   "vendedorcodigo"
               XListCampo      =   "vendedornombres"
               ListaCamposDescrip=   "Codigo,Descripcion"
               ListaCamposText =   "vendedorcodigo,vendedornombres"
               Requerido       =   0   'False
            End
            Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuempresa 
               Height          =   375
               Left            =   3480
               TabIndex        =   64
               Top             =   2280
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   661
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
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "EMPRESA"
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
               Height          =   225
               Index           =   2
               Left            =   960
               TabIndex        =   65
               Top             =   2400
               Width           =   2085
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "FECHA DE PLANILLA"
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
               Height          =   225
               Index           =   1
               Left            =   960
               TabIndex        =   34
               Top             =   1200
               Width           =   2085
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "TIPO DE PLANILLA"
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
               Height          =   225
               Index           =   0
               Left            =   960
               TabIndex        =   33
               Top             =   600
               Width           =   2085
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "OFICINA"
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
               Height          =   225
               Index           =   3
               Left            =   960
               TabIndex        =   32
               Top             =   1755
               Width           =   2085
            End
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   4275
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   11175
         Begin VB.Frame frmbotones 
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            Height          =   810
            Left            =   4140
            TabIndex        =   22
            Top             =   3330
            Width           =   3630
            Begin VB.CommandButton cmdBotones 
               Caption         =   "&Nuevo"
               Height          =   660
               Index           =   0
               Left            =   180
               Picture         =   "FrmPlanillaCompensaciones.frx":1140
               Style           =   1  'Graphical
               TabIndex        =   25
               Top             =   90
               Width           =   870
            End
            Begin VB.CommandButton cmdBotones 
               Caption         =   "&Compensar"
               Height          =   660
               Index           =   2
               Left            =   2490
               Picture         =   "FrmPlanillaCompensaciones.frx":1582
               Style           =   1  'Graphical
               TabIndex        =   24
               Top             =   90
               Width           =   990
            End
            Begin VB.CommandButton cmdBotones 
               Caption         =   "&Eliminar"
               Height          =   660
               Index           =   1
               Left            =   1230
               Picture         =   "FrmPlanillaCompensaciones.frx":19C4
               Style           =   1  'Graphical
               TabIndex        =   23
               Top             =   90
               Width           =   1065
            End
         End
         Begin VB.Frame Frame5 
            Height          =   825
            Left            =   135
            TabIndex        =   6
            Top             =   2430
            Width           =   10845
            Begin VB.TextBox Text1 
               Height          =   285
               Index           =   4
               Left            =   9600
               MaxLength       =   12
               TabIndex        =   12
               Top             =   450
               Width           =   1065
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Index           =   3
               Left            =   9000
               MaxLength       =   2
               TabIndex        =   11
               Top             =   450
               Width           =   435
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Index           =   0
               Left            =   6000
               MaxLength       =   2
               TabIndex        =   10
               Top             =   450
               Width           =   435
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Index           =   1
               Left            =   6480
               MaxLength       =   4
               TabIndex        =   9
               Top             =   450
               Width           =   585
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Index           =   2
               Left            =   7080
               MaxLength       =   10
               TabIndex        =   8
               Top             =   450
               Width           =   1305
            End
            Begin VB.CommandButton cAyuda 
               Caption         =   "..."
               Height          =   315
               Index           =   0
               Left            =   8400
               TabIndex        =   7
               Top             =   450
               Width           =   195
            End
            Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayucliente 
               Height          =   300
               Left            =   60
               TabIndex        =   13
               Top             =   420
               Width           =   5730
               _ExtentX        =   10107
               _ExtentY        =   529
               XcodMaxLongitud =   11
               xcodwith        =   1200
               NomTabla        =   "vt_cliente"
               ListaCampos     =   "clientecodigo(1),clienterazonsocial(1)"
               XcodCampo       =   "clientecodigo"
               XListCampo      =   "clienterazonsocial"
               ListaCamposDescrip=   "Codigo,Descripcion"
               ListaCamposText =   "clientecodigo,clienterazonsocial"
               Requerido       =   0   'False
            End
            Begin VB.Label Label2 
               Caption         =   "TD"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   1
               Left            =   6090
               TabIndex        =   21
               Top             =   180
               Width           =   315
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Caption         =   "Serie"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   2
               Left            =   6510
               TabIndex        =   20
               Top             =   180
               Width           =   525
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Caption         =   "Numero"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   3
               Left            =   7170
               TabIndex        =   19
               Top             =   180
               Width           =   1065
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Caption         =   "Mon."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   6
               Left            =   9000
               TabIndex        =   18
               Top             =   180
               Width           =   465
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Caption         =   "Importe"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   7
               Left            =   9690
               TabIndex        =   17
               Top             =   210
               Width           =   945
            End
            Begin VB.Label Label3 
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   2
               Left            =   0
               TabIndex        =   16
               Top             =   960
               Width           =   1185
            End
            Begin VB.Label Label3 
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   3
               Left            =   0
               TabIndex        =   15
               Top             =   900
               Width           =   1185
            End
            Begin VB.Label Label7 
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
               Height          =   240
               Index           =   1
               Left            =   1920
               TabIndex        =   14
               Top             =   210
               Width           =   1380
            End
         End
         Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
            Height          =   1935
            Left            =   150
            TabIndex        =   5
            Top             =   540
            Width           =   10845
            _ExtentX        =   19129
            _ExtentY        =   3413
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
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=43,.bold=0,.fontsize=825,.italic=0"
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
         Begin VB.Label Label6 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Index           =   0
            Left            =   9570
            TabIndex        =   61
            Top             =   3360
            Width           =   1275
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
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
            Height          =   225
            Index           =   0
            Left            =   9000
            TabIndex        =   60
            Top             =   3510
            Width           =   645
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Documentos de Clientes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   210
            TabIndex        =   26
            Top             =   120
            Width           =   10755
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H0000C0C0&
            BorderWidth     =   2
            FillColor       =   &H00FFC0C0&
            FillStyle       =   0  'Solid
            Height          =   3885
            Index           =   0
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   360
            Width           =   11205
         End
      End
      Begin VB.Frame Frame4 
         Height          =   930
         Left            =   -69690
         TabIndex        =   1
         Top             =   6120
         Width           =   1980
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Cancelar"
            Height          =   690
            Index           =   12
            Left            =   1080
            Picture         =   "FrmPlanillaCompensaciones.frx":1E06
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   180
            Width           =   855
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Acepta"
            Height          =   690
            Index           =   11
            Left            =   90
            Picture         =   "FrmPlanillaCompensaciones.frx":2248
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   180
            Width           =   870
         End
      End
   End
   Begin MSComctlLib.StatusBar Panel 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   56
      Top             =   9150
      Width           =   12240
      _ExtentX        =   21590
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmPlanillaCompensaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adll As New dllgeneral.dll_general
Dim rsdetac1 As New ADODB.Recordset
Dim rsdetac2 As New ADODB.Recordset

Public Function Cargar_grilla2()
   Set rsdetac2 = Nothing
   Call rsdetac2.Fields.Append("Cliente", adChar, 11)
   Call rsdetac2.Fields.Append("Descripcion", adChar, 30)
   Call rsdetac2.Fields.Append("TD", adChar, 2)
   Call rsdetac2.Fields.Append("Serie", adChar, 4)
   Call rsdetac2.Fields.Append("Numero", adChar, 10)
   Call rsdetac2.Fields.Append("FEmision", adDate)
   Call rsdetac2.Fields.Append("FVencimiento", adDate)
   Call rsdetac2.Fields.Append("Moneda", adChar, 2)
   Call rsdetac2.Fields.Append("Importe", adDouble)
   rsdetac2.Open
   Set TDBGrid2.DataSource = rsdetac2
   TDBGrid2.Refresh
   Call ConfigGrid2

End Function

Public Function ConfigGrid2()
    With TDBGrid2
       .Columns(0).Width = 1200
       .Columns(1).Width = 2800
       .Columns(2).Width = 500
       .Columns(3).Width = 700
       .Columns(4).Width = 1100
       .Columns(5).Width = 1100
       .Columns(6).Width = 1200
       .Columns(7).Width = 700
       .Columns(8).Width = 1200
       .Columns(8).NumberFormat = "###,###,##0.00"
       .Refresh
   End With
End Function

Private Sub cAyuda_Click(Index As Integer)
  Dim gfiltra(1, 2) As String
  nAyuda = "": nDetalle = ""
  If Index = 0 Then
     If adll.VerificaDatoExistente(VGCNx, "select * from vt_cargo where clientecodigo='" & Ctr_Ayucliente.xclave & "'") = 1 Then

        gfiltra(1, 1) = "Documento": gfiltra(1, 2) = "cargonumdoc"
        FrmAyuda.TipoForma = 5
        FrmAyuda.BConexion = VGCNx
        FrmAyuda.Bdata = "0"
        FrmAyuda.BTabla = "vt_cargo A inner join cc_tipodocumento B On a.documentocargo=b.tdocumentocodigo"
        FrmAyuda.BCampos = "documentocargo as TD,cargonumdoc as Numero,monedacodigo as Mnd,cargoapeimpape as Total,(isnull(cargoapeimpape,0)-isnull(cargoapeimppag,0)) as Saldo,cargoapefecvct as FechaVenc"
        FrmAyuda.BOrden = "Clientecodigo,cargoapefecemi"
        FrmAyuda.BCondi = "clientecodigo='" & Ctr_Ayucliente.xclave & "' and cargoapeflgcan<>'1' and b.tdocumentorenovarletras='1'"
        FrmAyuda.BFiltro = gfiltra
        FrmAyuda.Show 1
        Text1(0) = nAyuda
        Text1(1) = Left$(nDetalle, Text1(1).MaxLength)
        If Len(nDetalle) >= Text1(1).MaxLength + Text1(2).MaxLength Then
           Text1(2) = Right$(nDetalle, Text1(2).MaxLength)
         Else
           Text1(2) = Right$(nDetalle, Len(nDetalle) - Text1(1).MaxLength)
        End If
        Text1(4) = nSaldo
        Call Text1_KeyPress(2, 13)
        Text1(3).SetFocus
     Else
         nAyuda = "": nDetalle = ""
         MsgBox "No existen Documentos...", vbInformation, MsgTitle
         Exit Sub
     End If
  ElseIf Index = 1 Then
    If adll.VerificaDatoExistente(VGCNx, "select * from cp_tipodocumento where tdocumentodocrenovaletra='1'") = 1 Then
         Dim dfiltra(1, 2) As String
         dfiltra(1, 1) = "Documento": dfiltra(1, 2) = "cargonumdoc"
         FrmAyuda.TipoForma = 1
         FrmAyuda.BConexion = VGCNx
         FrmAyuda.Bdata = "0"
         FrmAyuda.BTabla = "cp_tipodocumento"
         FrmAyuda.BCampos = "tdocumentocodigo as Codigo,tdocumentodescripcion as Descripcion"
         FrmAyuda.BOrden = "tdocumentocodigo"
         FrmAyuda.BCondi = "tdocumentodocrenovaletra='1'"
         FrmAyuda.BFiltro = dfiltra
         FrmAyuda.Show 1
         Text2(0) = nAyuda
         Call Text2_KeyPress(0, 13)
         Text2(0).SetFocus
    Else
          nAyuda = "": nDetalle = ""
         MsgBox "No existen Documentos...", vbInformation, MsgTitle
         Exit Sub
     End If
     ElseIf Index = 2 Then
          If adll.VerificaDatoExistente(VGCNx, "select * from cp_cargo where clientecodigo='" & Ctr_AyuProv.xclave & "'") = 1 Then
             gfiltra(1, 1) = "Documento": gfiltra(1, 2) = "cargonumdoc"
             FrmAyuda.TipoForma = 5
             FrmAyuda.BConexion = VGCNx
             FrmAyuda.Bdata = "0"
             FrmAyuda.BTabla = "cp_cargo A inner join cp_tipodocumento B On a.documentocargo=b.tdocumentocodigo"
             FrmAyuda.BCampos = "documentocargo as TD,cargonumdoc as Numero,monedacodigo as Mnd,cargoapeimpape as Total,(isnull(cargoapeimpape,0)-isnull(cargoapeimppag,0)) as Saldo,cargoapefecvct as FechaVenc"
             FrmAyuda.BOrden = "Clientecodigo,cargoapefecemi"
             FrmAyuda.BCondi = "clientecodigo='" & Ctr_AyuProv.xclave & "' and cargoapeflgcan<>'1' and b.tdocumentorenovarletras='1'"
             FrmAyuda.BFiltro = gfiltra
             FrmAyuda.Show 1
             Text2(0) = nAyuda
             If Len(nDetalle) = 0 Then
                Exit Sub
              Else
                Text2(1).Text = Left$(nDetalle, Text1(1).MaxLength)
                If Len(nDetalle) >= Text1(1).MaxLength + Text2(2).MaxLength Then
                   Text2(2) = Right$(nDetalle, Text2(2).MaxLength)
                Else
                   Text2(2) = Right$(nDetalle, Len(nDetalle) - Text2(1).MaxLength)
               End If
               Call Text2_KeyPress(2, 13)
              Text2(3).SetFocus
            End If
          Else
             nAyuda = "": nDetalle = ""
             MsgBox "No existen Documentos...", vbInformation, MsgTitle
             Exit Sub
         End If
     Else: nAyuda = "": nDetalle = ""
         MsgBox "No existen Documentos...", vbInformation, MsgTitle
         Exit Sub
     End If
  
   nAyuda = "": nDetalle = ""
End Sub

Private Sub cmdBotones_Click(Index As Integer)
  Dim acmd As New ADODB.Command
  Dim rb As New ADODB.Recordset
  Dim xabono, xzona, xmone, xcuenta, xcargo, xcance As String
  Dim xparcial, xtipo As String
  Dim xnumplan, ximpsol, xtcam, xnumpag As Double
  
  On Error GoTo nerror
  
  Select Case Index
    Case 0
      Limpiartexto Text1, 0, 4
      Text1(0).SetFocus
    Case 1   'Eliminar Datos
      If TDBGrid1.ApproxCount > 0 Then
         TDBGrid1.Delete
         TDBGrid1.Update
         TDBGrid1.Refresh
         Call PlanillaTotales(rsdetac1, "importe", Label6(0))
      End If
    Case 2   'Grabar Datos de Documentos a Canjear
      If TDBGrid1.ApproxCount > 0 Then
         Frame7.Enabled = True
         'Text2(0).SetFocus
         Ctr_AyuProv.SetFocus
      Else
         Limpiartexto Text2, 0, 7
         Call adll.ActivaTab(0, 1, SSTab1)
      End If
    Case 4
        Limpiartexto Text2, 0, 7
        Text2(0).SetFocus
    Case 5  'Grabar Datos
       'Grabar datos a canjear
        If rsdetac1.RecordCount = 0 Or rsdetac2.RecordCount = 0 Then
            MsgBox "No existen documentos canjeados ó a canjear....Verifique!!!", vbInformation, MsgTitle
            Exit Sub
        End If
        If CDbl(Label6(0)) <> CDbl(Label6(1)) Then
            If Not MsgBox("Los PlanillaTotales no son iguales....Verifique!!! ,Desea Continuar ?", vbYesNo, MsgTitle) = vbYes Then
               Exit Sub
            End If
        End If
        If rsdetac1.RecordCount > 0 Then
            Set rb = VGCNx.Execute("select * from cp_tipoplanilla where tplanillacodigo='" & Escadena(Ctr_AyuTipoPla.xclave) & "'")
            If rb.RecordCount > 0 Then
               xnumplan = Val(Trim$(rb!tplanillanumerador)) + 1
            Else
               xnumplan = 1
            End If
            rb.Close
            Set rb = Nothing
            
            VGCNx.Execute "update cp_tipoplanilla " & _
                        " set tplanillanumerador='" & xnumplan & "' " & _
                        " where tplanillacodigo='" & Escadena(Ctr_AyuTipoPla.xclave) & "'"
        
            VGCNx.Execute "update CC_tipoplanilla " & _
                        " set tplanillanumerador='" & xnumplan & "' " & _
                        " where tplanillacodigo='" & Escadena(Ctr_AyuTipoPla.xclave) & "'"
        End If
        grabacobrar (xnumplan)
        grabapagar (xnumplan)
        MsgBox "Los datos han sido grabados satisfactoriamente...!!!", vbInformation, MsgTitle
        
        Call adll.ActivaTab(0, 1, SSTab1)

    
    Case 6
       If TDBGrid2.ApproxCount > 0 Then
         TDBGrid2.Delete
         TDBGrid2.Update
         TDBGrid2.Refresh
         Call PlanillaTotales(rsdetac2, "importe", Label6(1))
       End If
    Case 7
        Call adll.ActivaTab(0, 1, SSTab1)
    Case 11
      If Len(Trim$(Ctr_AyuTipoPla.xclave)) = 0 Then
        MsgBox "Falta Ingresar Tipo de Planilla...Verifique!!", vbInformation, MsgTitle
        Exit Sub
      End If
      'If Len(trim$(ctr_ayucliente.xclave)) = 0 Then
      '  MsgBox "Falta Ingresar Oficina/Vendedor...Verifique!!", vbInformation, MsgTitle
      '  Exit Sub
      'End If
      If Len(Trim$(Ctr_Ayuoficina.xclave)) = 0 Then
        MsgBox "Falta Ingresar Oficina/Vendedor...Verifique!!", vbInformation, MsgTitle
        Exit Sub
      End If
      If adll.VerificaDatoExistente(VGCNx, "select * from cp_tipoplanilla where tplanillaCompensa='1' and tplanillacodigo='" & Escadena(Ctr_AyuTipoPla.xclave) & "' ") = 0 Then
            MsgBox "La planilla no es valida para realizar los canjes...Verifique!!!", vbInformation, MsgTitle
            Ctr_AyuTipoPla.SetFocus
            Exit Sub
      End If

      Set rsdetac1 = Nothing
      TDBGrid1.ClearFields
      Set TDBGrid1.DataSource = Nothing
      Call cargar_grilla
      
      Set rsdetac2 = Nothing
      TDBGrid2.ClearFields
      Set TDBGrid2.DataSource = Nothing
      Call Cargar_grilla2
       
      Label6(0) = "": Label6(1) = ""
      Limpiartexto Text1, 0, 4
  '    Limpiartexto Text2, 0, 7
      
      Call adll.ActivaTab(1, 1, SSTab1)
      'Text1(0).SetFocus
      Ctr_Ayucliente.SetFocus
    Case 12
      Unload Me
  End Select
  Exit Sub
nerror:
  If Err Then
    MsgBox "Error : " & Err.Number & "-" & Err.Description, vbInformation, MsgTitle
    Err = 0
    Exit Sub
  End If
End Sub

Private Sub Form_Load()
  MostrarForm Me, "C"
  
  MBox1 = Format(Date, "DD/MM/YYYY")
  Call Ctr_AyuTipoPla.conexion(VGCNx)
  Call Ctr_Ayucliente.conexion(VGCNx)
  Call Ctr_Ayuoficina.conexion(VGCNx)
  Call Ctr_AyuProv.conexion(VGCNx)
  Call Ctr_Ayuempresa.conexion(VGCNx)
  Ctr_AyuTipoPla.Filtro = "tplanillacompensa='1'"
  Call adll.ActivaTab(0, 1, SSTab1)
End Sub

Public Sub ConfigGrid()
   With TDBGrid1
       .Columns(0).Width = 1200
       .Columns(1).Width = 2800
       .Columns(2).Width = 500
       .Columns(3).Width = 700
       .Columns(4).Width = 1100
       .Columns(5).Width = 1100
       .Columns(6).Width = 1200
       .Columns(7).Width = 700
       .Columns(8).Width = 1200
       .Columns(8).NumberFormat = "###,###,##0.00"
       .Refresh
   End With
End Sub

Public Sub cargar_grilla()
   Set rsdetac1 = Nothing
   Call rsdetac1.Fields.Append("Cliente", adChar, 11)
   Call rsdetac1.Fields.Append("Descripcion", adChar, 30)
   Call rsdetac1.Fields.Append("TD", adChar, 2)
   Call rsdetac1.Fields.Append("Serie", adChar, 4)
   Call rsdetac1.Fields.Append("Numero", adChar, 10)
   Call rsdetac1.Fields.Append("FEmision", adDate)
   Call rsdetac1.Fields.Append("FVencimiento", adDate)
   Call rsdetac1.Fields.Append("Moneda", adChar, 2)
   Call rsdetac1.Fields.Append("Importe", adDouble)
   rsdetac1.Open
   Set TDBGrid1.DataSource = rsdetac1
   TDBGrid1.Refresh
   Call ConfigGrid
End Sub

Private Sub MBox1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
     SendKeys "{tab}"
  End If
  
End Sub



Private Sub Text1_GotFocus(Index As Integer)
  If Index = 4 Then
      Call adll.Enfoquetexto(Text1(Index))
  End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
  Dim rb As New ADODB.Recordset
  Dim rb2 As New ADODB.Recordset
  Dim xpago, xcam, J, flag As Double
  On Error Resume Next
  
  If KeyAscii = 13 Then
          
     If Index = 0 Then
       Set rb = VGCNx.Execute("select * from cp_tipodocumento where tdocumentocodigo='" & Text1(Index) & "' and tdocumentotipo='C'")
       If rb.RecordCount = 0 Then
           MsgBox "El documento no es valido....Verifique!!!", vbInformation, MsgTitle
           rb.Close
           Set rb = Nothing
           Exit Sub
       End If
       rb.Close
       Set rb = Nothing
    ElseIf Index = 1 Then
       Text1(Index) = Right$("000000000" & Trim$(Text1(Index)), Text1(Index).MaxLength)
    ElseIf Index = 2 Then
   '    Text1(Index) = right$("000000000" & trim$(Text1(Index)), Text1(Index).MaxLength)
       
       Set rb = VGCNx.Execute("select * from vt_cargo where documentocargo='" & Text1(0).Text & "' and cargonumdoc='" & Trim$(Text1(1).Text & Text1(2).Text) & "' and Clientecodigo='" & Trim$(Ctr_Ayucliente.xclave) & "'")
       If rb.RecordCount = 0 Then
           MsgBox "El documento no existe....Verifique!!!", vbInformation, MsgTitle
           rb.Close
           Set rb = Nothing
           Label3(0) = "": Label3(1) = "": Label3(2) = "": Label3(3) = "": Text1(3) = "": Text1(4) = ""
           Exit Sub
       Else
            Text1(3) = IIf(IsNull(rb!monedacodigo), "", rb!monedacodigo)
            If IsNull(rb!cargoapeimppag) Then
               Text1(4) = Numero(rb!cargoapeimpape)
            Else
               Text1(4) = Numero(rb!cargoapeimpape - rb!cargoapeimppag)

            End If
            
            Label3(2) = IIf(IsNull(rb!cargoapefecemi), "", rb!cargoapefecemi)
            Label3(3) = IIf(IsNull(rb!cargoapefecvct), "", rb!cargoapefecvct)
       
            Set rb2 = VGCNx.Execute("select * from vt_cliente  where clientecodigo='" & Trim$(Escadena(rb!clientecodigo)) & "'")
            If rb2.RecordCount = 0 Then
                MsgBox "Cliente No existe...Verifique!!!", vbInformation, MsgTitle
            Else
               Label3(0) = Escadena(rb2!clientecodigo)
               Label3(1) = Escadena(rb2!clienterazonsocial)
            End If
            rb2.Close
            Set rb2 = Nothing
       End If
       rb.Close
       Set rb = Nothing
     ElseIf Index = 3 Then
       Set rb = VGCNx.Execute("select * from gr_moneda where monedacodigo='" & Trim$(Escadena(Text1(Index))) & "'")
       If rb.RecordCount = 0 Then
           MsgBox "La moneda no existe....Verifique!!!", vbInformation, MsgTitle
           rb.Close
           Set rb = Nothing
           Exit Sub
       End If
       rb.Close
       Set rb = Nothing
 
    ElseIf Index = 4 Then
        Text1(Index) = Format(Trim$(Text1(Index)), "##,###,##0.00")
        flag = 0
        For J = 0 To 4
          If Len(Trim$(Text1(J))) = 0 Then
             flag = 1
             Exit For
          End If
        Next J
        If flag = 1 Then
           MsgBox "Falta completar datos...Verifique!!!", vbInformation, MsgTitle
           Exit Sub
        End If
        rsdetac1.AddNew
        rsdetac1.Fields(0) = Ctr_Ayucliente.xclave
        rsdetac1.Fields(1) = Left$(Ctr_Ayucliente.xnombre, 30)
        rsdetac1.Fields(2) = Escadena(Text1(0))
        rsdetac1.Fields(3) = Escadena(Text1(1))
        rsdetac1.Fields(4) = Escadena(Text1(2))
        
        rsdetac1.Fields(5) = Escadena(Label3(2))    'Fecha de Emision
        rsdetac1.Fields(6) = Escadena(Label3(3))    'Fecha de Vencimiento
        
        rsdetac1.Fields(7) = Escadena(Text1(3))
        rsdetac1.Fields(8) = Escadena(Text1(4))
        rsdetac1.Update
                        
        Call PlanillaTotales(rsdetac1, "importe", Label6(0))
        
        Limpiartexto Text1, 0, 4
        'Text1(0).SetFocus
        Ctr_AyuProv.SetFocus
        
        Exit Sub
     End If
     SendKeys "{tab}"
  End If
  Set rb = Nothing
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
Dim rb As New ADODB.Recordset
Dim rb2 As New ADODB.Recordset
Dim xpago, xcam, J, flag As Double
On Error Resume Next

  If KeyAscii = 13 Then
          
     If Index = 0 Then
       Set rb = VGCNx.Execute("select * from cp_tipodocumento where tdocumentocodigo='" & Text2(Index) & "' and tdocumentodocrenovaletra='1'")
       If rb.RecordCount = 0 Then
           MsgBox "El documento no es valido....Verifique!!!", vbInformation, MsgTitle
           rb.Close
           Set rb = Nothing
           Exit Sub
       Else
          If Not IsNull(rb.Fields("tdocumentonumeauto")) Then
             Text2(1).Text = IIf(rb.Fields("tdocumentonumeauto") = "0", "", Right$("0000000000" & g_pedserie, 3))
             Text2(2).Text = IIf(rb.Fields("tdocumentonumeauto") = "0", "", Right$("0000000000" & Trim$(rb.Fields("tdocumentonumerador")), 8))
             Text2(0).SetFocus
          End If
          
       End If
       rb.Close
       Set rb = Nothing
    ElseIf Index = 1 Then
       Text2(Index) = Right$("000000000" & Trim$(Text2(Index)), Text2(Index).MaxLength)
    ElseIf Index = 2 Then
       SQL = "select * from cp_cargo where documentocargo='" & Text2(0) & "' and cargonumdoc='" & Trim$(Text2(1) & Text2(2)) & "' and Clientecodigo='" & Ctr_AyuProv.xclave & "'"
       Set rb = VGCNx.Execute(SQL)
       If rb.RecordCount = 0 Then
           MsgBox "El documento no existe....Verifique!!!", vbInformation, MsgTitle
           rb.Close
           Set rb = Nothing
           Label3(0) = "": Label3(1) = "": Label3(2) = "": Label3(3) = "": Text1(3) = "": Text1(4) = ""
           Exit Sub
       Else
            Text2(3) = IIf(IsNull(rb!monedacodigo), "", rb!monedacodigo)
            If IsNull(rb!cargoapeimppag) Then
               Text2(4) = Numero(rb!cargoapeimpape)
            Else
               Text2(4) = Numero(rb!cargoapeimpape - rb!cargoapeimppag)

            End If
            
            Label9(0) = IIf(IsNull(rb!cargoapefecemi), "", rb!cargoapefecemi)
            Label9(1) = IIf(IsNull(rb!cargoapefecvct), "", rb!cargoapefecvct)
       
 
            Set rb = Nothing
       End If
       rb.Close
       Set rb = Nothing
    ElseIf Index = 3 Then
       Set rb = VGCNx.Execute("select * from gr_moneda where monedacodigo='" & Trim$(Escadena(Text2(Index))) & "'")
       If rb.RecordCount = 0 Then
           MsgBox "La moneda no existe....Verifique!!!", vbInformation, MsgTitle
           rb.Close
           Set rb = Nothing
           Exit Sub
       End If
       rb.Close
       Set rb = Nothing
    ElseIf Index = 4 Then
        flag = 0
        For J = 0 To 4
          If Len(Trim$(Text2(J))) = 0 Then
             flag = 1
             Exit For
          End If
        Next J
        If flag = 1 Then
           MsgBox "Falta completar datos...Verifique!!!", vbInformation, MsgTitle
           Exit Sub
        End If
        
        rsdetac2.AddNew
        rsdetac2.Fields(0) = Ctr_AyuProv.xclave
        rsdetac2.Fields(1) = Left$(Ctr_AyuProv.xnombre, 30)
        rsdetac2.Fields(2) = Escadena(Text2(0))
        rsdetac2.Fields(3) = Escadena(Text2(1))
        rsdetac2.Fields(4) = Escadena(Text2(2))
        
        rsdetac2.Fields(5) = Escadena(Label9(0))    'Fecha de Emision
        rsdetac2.Fields(6) = Escadena(Label9(1))    'Fecha de Vencimiento
        
        rsdetac2.Fields(7) = Escadena(Text2(3))
        rsdetac2.Fields(8) = Escadena(Text2(4))
        
        rsdetac2.Update
                        
                      
        Call PlanillaTotales(rsdetac2, "importe", Label6(1))
        
        Limpiartexto Text2, 0, 7
        Text2(0).SetFocus
        Exit Sub
    End If
    SendKeys "{tab}"
  End If
End Sub
Private Sub grabacobrar(numpla As Double)
Dim acmd As New ADODB.Command
Dim rb As New ADODB.Recordset
Dim xabono, xcance, xtipo, xcuenta, xparcial, xzona, xmone As String
Dim xnumpag, ximpsol, xtcam As Double
rsdetac1.MoveFirst
Do Until rsdetac1.EOF
                Set rb = VGCNx.Execute("select * from cc_tipodocumento where tdocumentocodigo ='" & rsdetac1.Fields(2) & "'")
                If rb.RecordCount > 0 Then
                   xabono = rb!tdocumentotipo
                   xcance = rb!tdocumentocodigo
                   xtipo = IIf(IsNull(rb!tdocumentopermiteaplica), Null, rb!tdocumentopermiteaplica)
                   If rsdetac1.Fields("moneda") = g_TipoSol Then
                      xcuenta = "" & Trim$(rb!tdocumentocuentasoles)
                   Else
                      xcuenta = "" & Trim$(rb!tdocumentocuentadolares)
                   End If
                Else
                   xabono = "": xcuenta = "": xtipo = "": xcance = ""
                End If
                rb.Close
                Set rb = Nothing
                
                xparcial = ""
                SQL = "select * from vt_cargo where documentocargo='" & rsdetac1.Fields(2) & "'"
                SQL = SQL & " and cargonumdoc='" & rsdetac1.Fields(3) & rsdetac1.Fields(4) & "'"
                SQL = SQL & " and clientecodigo='" & Trim$(Ctr_Ayucliente.xclave) & "'"
                Set rb = VGCNx.Execute(SQL)
                If rb.RecordCount > 0 Then
                   xzona = rb!zonacodigo
                   xmone = rb!monedacodigo
                   If IsNull(rb.Fields("cargoapeimppag")) Then
                     xparcial = IIf(rb.Fields("cargoapeimpape") - rsdetac1.Fields("importe") <= 0, "T", "P")
                   Else
                     xparcial = IIf(rb.Fields("cargoapeimpape") - rb.Fields("cargoapeimppag") - rsdetac1.Fields("importe") <= 0, "T", "P")
                   End If
                  
                   If IsNull(rb!cargoapenumpag) Then
                     xnumpag = 1
                   Else
                     xnumpag = Val(rb!cargoapenumpag)
                   End If
                Else
                   xzona = "01": xmone = g_TipoSol: xnumpag = 1: xparcial = ""
                End If
                rb.Close
                Set rb = Nothing

                ximpsol = CDbl(rsdetac1.Fields("importe"))
                xtcam = DatoTipoCambio(VGcnxCT, MBox1.Text)               'TraeTipoCambio(Date, VGcnx)
                If rsdetac1.Fields("moneda") <> xmone Then
                   If rsdetac1.Fields("moneda") = g_TipoSol Then
                      ximpsol = CDbl(rsdetac1.Fields("importe")) / CDbl(xtcam)
                   Else
                      ximpsol = CDbl(rsdetac1.Fields("importe")) * CDbl(xtcam)
                   End If
                End If
                
                Set acmd.ActiveConnection = VGgeneral
                acmd.CommandType = adCmdStoredProc
                acmd.CommandText = "cc_abonadocumento_pro"
                acmd.CommandTimeout = 0
                acmd.Prepared = True
                With acmd
                    .Parameters("@base") = VGCNx.DefaultDatabase
                    .Parameters("@tipo") = "1"
                    .Parameters("@empresa") = Ctr_Ayuempresa.xclave
                    .Parameters("@documentoabono") = rsdetac1.Fields(2)
                    .Parameters("@abononumdoc") = RTrim$(rsdetac1.Fields(3) & rsdetac1.Fields(4))
                    .Parameters("@abonocannumpag") = xnumpag
                    .Parameters("@zonacodigo") = xzona
                    .Parameters("@tipoplanilla") = Escadena(Ctr_AyuTipoPla.xclave)
                    .Parameters("@vendedor") = Escadena(Ctr_Ayuoficina.xclave)
                    .Parameters("@numplanilla") = Right$("00000000" & RTrim$(CStr(numpla)), 6)
                    .Parameters("@fechapla") = MBox1.Text
                    .Parameters("@fechapro") = MBox1.Text
                    .Parameters("@moneda") = xmone
                    .Parameters("@abonocancarabo") = "A"  'xabono
                    .Parameters("@cuenta") = xcuenta
                    .Parameters("@banco") = ""
                    .Parameters("@tipocam") = CDbl(xtcam)
                    .Parameters("@abonoflpres") = "1"
                    .Parameters("@abonocanimpcan") = CDbl(rsdetac1.Fields("importe"))
                    .Parameters("@abonocanimpsol") = ximpsol
                    .Parameters("@usuario") = VGparamsistem.Usuario
                    .Parameters("@fechaact") = Now
                    .Parameters("@forma") = xparcial
                    .Parameters("@monedacan") = rsdetac1.Fields("moneda")
                    .Parameters("@abonocantd") = "33"  ' rsdetac1.Fields("moneda")
                    .Parameters("@abonocannro") = Right$("00000000" & RTrim$(CStr(numpla)), 6)
                    .Parameters("@fechacan") = MBox1.Text
                    .Parameters("@cliente") = rsdetac1.Fields(0)
                End With
               acmd.Execute
                
                Set acmd = Nothing
                DoEvents
                                
                '**** Actualizamos Saldos de documento pendiente
                If rsdetac1.Fields("moneda") = g_TipoDolar Then
                   If xmone = g_TipoSol Then
                       VGCNx.Execute "Update  vt_cargo Set cargoapeimppag=isnull(cargoapeimppag,0)+" & CDbl(rsdetac1.Fields(10) * xtcam) & "," & _
                                   " cargoapenumpag='" & xnumpag + 1 & "'" & _
                                  " Where empresacodigo='" & Ctr_Ayuempresa.xclave & "' and  documentocargo='" & rsdetac1.Fields(1) & "' and cargonumdoc='" & RTrim$(rsdetac1.Fields(2) & rsdetac1.Fields(3)) & "'"
                   Else
                       VGCNx.Execute "Update  vt_cargo Set cargoapeimppag=isnull(cargoapeimppag,0)+" & rsdetac1.Fields("importe") & "," & _
                                  " cargoapenumpag='" & xnumpag + 1 & "'" & _
                                  " Where empresacodigo='" & Ctr_Ayuempresa.xclave & "' and documentocargo='" & rsdetac1.Fields(1) & "' and cargonumdoc='" & RTrim$(rsdetac1.Fields(2) & rsdetac1.Fields(3)) & "'"
                   End If
                ElseIf rsdetac1.Fields("moneda") = g_TipoSol Then
                   If xmone = g_TipoDolar Then
                       VGCNx.Execute "Update  vt_cargo Set cargoapeimppag=isnull(cargoapeimppag,0)+" & rsdetac1.Fields("importe") / xtcam & "," & _
                                  " cargoapenumpag='" & xnumpag + 1 & "'" & _
                                  " Where empresacodigo='" & Ctr_Ayuempresa.xclave & "' and documentocargo='" & rsdetac1.Fields(1) & "' and cargonumdoc='" & RTrim$(rsdetac1.Fields(2) & rsdetac1.Fields(3)) & "'"
                   Else
                       VGCNx.Execute "Update  vt_cargo Set cargoapeimppag=isnull(cargoapeimppag,0)+" & rsdetac1.Fields("importe") & "," & _
                                  " cargoapenumpag='" & xnumpag + 1 & "'" & _
                                  " Where empresacodigo='" & Ctr_Ayuempresa.xclave & "' and documentocargo='" & rsdetac1.Fields(1) & "' and cargonumdoc='" & RTrim$(rsdetac1.Fields(2) & rsdetac1.Fields(3)) & "'"
                   End If
                End If
                
                VGCNx.Execute "Update  vt_cargo " & _
                            " Set cargoapeflgcan= '0'" & _
                            " Where empresacodigo='" & Ctr_Ayuempresa.xclave & "' and documentocargo='" & rsdetac1.Fields(5) & "' and cargonumdoc='" & RTrim$(rsdetac1.Fields(6) & rsdetac1.Fields(7)) & "'"
                                        
                VGCNx.Execute "Update  vt_cargo " & _
                            " Set cargoapeflgcan= '1'," & _
                            "   cargoapefeccan='" & MBox1.Text & "'" & _
                            " Where empresacodigo='" & Ctr_Ayuempresa.xclave & "' and documentocargo='" & rsdetac1.Fields(1) & "' and cargonumdoc='" & RTrim$(rsdetac1.Fields(2) & rsdetac1.Fields(3)) & "'" & _
                            " and cargoapeimpape-isnull(cargoapeimppag,0)= 0 "
             rsdetac1.MoveNext
Loop
End Sub
Private Sub grabapagar(numpla As Double)

Dim acmd As New ADODB.Command
Dim rb As New ADODB.Recordset
Dim xabono, xcance, xtipo, xcuenta, xparcial, xzona, xmone As String
Dim xnumpag, ximpsol, xtcam As Double
            rsdetac2.MoveFirst
            Do Until rsdetac2.EOF
                
                Set rb = VGCNx.Execute("select * from cP_tipodocumento where tdocumentocodigo ='" & rsdetac2.Fields(2) & "'")
                If rb.RecordCount > 0 Then
                   xabono = rb!tdocumentotipo
                   xcance = rb!tdocumentocodigo
                   xtipo = IIf(IsNull(rb!tdocumentopermiteaplica), Null, rb!tdocumentopermiteaplica)
                   If rsdetac2.Fields("moneda") = g_TipoSol Then
                      xcuenta = "" & Trim$(rb!tdocumentocuentasoles)
                   Else
                      xcuenta = "" & Trim$(rb!tdocumentocuentadolares)
                   End If
                Else
                   xabono = "": xcuenta = "": xtipo = "": xcance = ""
                End If
                rb.Close
                Set rb = Nothing
                
                xparcial = ""
                SQL = "select * from cp_cargo where empresacodigo='" & Ctr_Ayuempresa.xclave & "' and documentocargo='" & rsdetac2.Fields(2) & "'"
                SQL = SQL & " and cargonumdoc='" & rsdetac2.Fields(3) & rsdetac2.Fields(4) & "'"
                SQL = SQL & " and clientecodigo='" & Trim$(Ctr_AyuProv.xclave) & "'"
                Set rb = VGCNx.Execute(SQL)
                If rb.RecordCount > 0 Then
                   xzona = rb!zonacodigo
                   xmone = rb!monedacodigo
                   If IsNull(rb.Fields("cargoapeimppag")) Then
                     xparcial = IIf(rb.Fields("cargoapeimpape") - rsdetac2.Fields("importe") <= 0, "T", "P")
                   Else
                     xparcial = IIf(rb.Fields("cargoapeimpape") - rb.Fields("cargoapeimppag") - rsdetac2.Fields("importe") <= 0, "T", "P")
                   End If
                  
                   If IsNull(rb!cargoapenumpag) Then
                     xnumpag = 1
                   Else
                     xnumpag = Val(rb!cargoapenumpag)
                   End If
                Else
                   xzona = "01": xmone = g_TipoSol: xnumpag = 1: xparcial = ""
                End If
                rb.Close
                Set rb = Nothing

                ximpsol = CDbl(rsdetac2.Fields("importe"))
                xtcam = DatoTipoCambio(VGcnxCT, MBox1.Text)               'TraeTipoCambio(Date, VGcnx)
                If rsdetac2.Fields("moneda") <> xmone Then
                   If rsdetac1.Fields("moneda") = g_TipoSol Then
                      ximpsol = CDbl(rsdetac2.Fields("importe")) / CDbl(xtcam)
                   Else
                      ximpsol = CDbl(rsdetac2.Fields("importe")) * CDbl(xtcam)
                   End If
                End If

                Set acmd.ActiveConnection = VGgeneral
                acmd.CommandType = adCmdStoredProc
                acmd.CommandText = "cp_abonadocumento_pro"
                acmd.CommandTimeout = 0
                acmd.Prepared = True
                With acmd
                    .Parameters("@base") = VGCNx.DefaultDatabase
                    .Parameters("@tipo") = "1"
                    .Parameters("@documentoabono") = rsdetac2.Fields(2)
                    .Parameters("@abononumdoc") = Trim$(rsdetac2.Fields(3) & rsdetac2.Fields(4))
                    .Parameters("@abonocannumpag") = xnumpag
                    .Parameters("@zonacodigo") = xzona
                    .Parameters("@tipoplanilla") = Escadena(Ctr_AyuTipoPla.xclave)
                    .Parameters("@vendedor") = Escadena(Ctr_Ayuoficina.xclave)
                    .Parameters("@numplanilla") = Right$("00000000" & Trim$(CStr(numpla)), 6)
                    .Parameters("@fechapla") = MBox1.Text
                    .Parameters("@fechapro") = MBox1.Text
                    .Parameters("@moneda") = xmone
                    .Parameters("@abonocancarabo") = xabono
                    .Parameters("@cuenta") = xcuenta
                    .Parameters("@banco") = ""
                    .Parameters("@tipocam") = CDbl(xtcam)
                    .Parameters("@ctabanco") = ""
                    .Parameters("@abonoflpres") = "1"
                    .Parameters("@abonocanimpcan") = CDbl(rsdetac2.Fields("importe"))
                    .Parameters("@abonocanimpsol") = ximpsol
                    .Parameters("@usuario") = VGusuario
                    .Parameters("@fechaact") = Now
                    .Parameters("@forma") = xparcial
                    .Parameters("@monedacan") = rsdetac2.Fields("moneda")
                    .Parameters("@abonocantd") = xcance
                    .Parameters("@abonocannro") = ""
                    .Parameters("@fechacan") = MBox1.Text
                    '.Parameters("@cliente") = Escadena(ctr_ayucliente.xclave)
                    .Parameters("@cliente") = Trim$(rsdetac2.Fields("Cliente"))
                    .Parameters("@empresa") = Ctr_Ayuempresa.xclave
                End With
                acmd.Execute

                Set acmd = Nothing
                DoEvents

                '**** Actualizamos Saldos de documento pendiente
                If rsdetac2.Fields("moneda") = g_TipoDolar Then
                   If xmone = g_TipoSol Then
                       VGCNx.Execute "Update  cp_cargo Set cargoapeimppag=isnull(cargoapeimppag,0)+" & CDbl(rsdetac2.Fields("importe") / xtcam) & "," & _
                                   " cargoapenumpag='" & xnumpag + 1 & "'" & _
                                  " Where documentocargo='" & rsdetac2.Fields(2) & "' and cargonumdoc='" & Trim$(rsdetac2.Fields(3) & rsdetac2.Fields(4)) & "' and " & _
                                  " clientecodigo='" & Trim$(rsdetac2.Fields("Cliente")) & "' and empresacodigo='" & Ctr_Ayuempresa.xclave & "'"
                   Else
                       VGCNx.Execute "Update  cp_cargo Set cargoapeimppag=isnull(cargoapeimppag,0)+" & CDbl(rsdetac2.Fields("importe")) & "," & _
                                  " cargoapenumpag='" & xnumpag + 1 & "'" & _
                                  " Where documentocargo='" & rsdetac2.Fields(2) & "' and cargonumdoc='" & Trim$(rsdetac2.Fields(3) & rsdetac2.Fields(4)) & "' and " & _
                                  " clientecodigo='" & Trim$(rsdetac2.Fields("Cliente")) & "' and empresacodigo='" & Ctr_Ayuempresa.xclave & "'"
                   End If
                ElseIf rsdetac2.Fields("moneda") = g_TipoSol Then
                   If xmone = g_TipoDolar Then
                       VGCNx.Execute "Update  cp_cargo Set cargoapeimppag=isnull(cargoapeimppag,0)+" & CDbl(rsdetac2.Fields("importe") * xtcam) & "," & _
                                  " cargoapenumpag='" & xnumpag + 1 & "'" & _
                                  " Where documentocargo='" & rsdetac2.Fields(2) & "' and cargonumdoc='" & Trim$(rsdetac2.Fields(3) & rsdetac2.Fields(4)) & "' and " & _
                                  " clientecodigo='" & Trim$(rsdetac2.Fields("Cliente")) & "' and empresacodigo='" & Ctr_Ayuempresa.xclave & "'"
                   Else
                       VGCNx.Execute "Update  cp_cargo Set cargoapeimppag=isnull(cargoapeimppag,0)+" & CDbl(rsdetac2.Fields("importe")) & "," & _
                                  " cargoapenumpag='" & xnumpag + 1 & "'" & _
                                  " Where documentocargo='" & rsdetac2.Fields(2) & "' and cargonumdoc='" & Trim$(rsdetac2.Fields(3) & rsdetac2.Fields(4)) & "' and " & _
                                  " clientecodigo='" & Trim$(rsdetac2.Fields("Cliente")) & "' and empresacodigo='" & Ctr_Ayuempresa.xclave & "'"
                   End If
                End If

                VGCNx.Execute "Update  cp_cargo " & _
                            " Set cargoapeflgcan= CASE isnull(cargoapeimpape,0)-isnull(cargoapeimppag,0) WHEN 0 THEN '1' ELSE '0' END ," & _
                            "   cargoapefeccan='" & Date & "'" & _
                            " Where documentocargo='" & rsdetac2.Fields(2) & "' and cargonumdoc='" & Trim$(rsdetac2.Fields(3) & rsdetac2.Fields(4)) & "' and " & _
                            " clientecodigo='" & Trim$(rsdetac2.Fields("Cliente")) & "' and empresacodigo='" & Ctr_Ayuempresa.xclave & "'"

                rsdetac2.MoveNext
           Loop
End Sub

