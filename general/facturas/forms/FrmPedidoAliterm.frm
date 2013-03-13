VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Begin VB.Form FrmPedidoAliterm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pedido"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12255
   Icon            =   "FrmPedidoAliterm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   12255
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   8205
      Left            =   90
      TabIndex        =   35
      Top             =   90
      Width           =   12105
      _ExtentX        =   21352
      _ExtentY        =   14473
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Consulta"
      TabPicture(0)   =   "FrmPedidoAliterm.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label10"
      Tab(0).Control(1)=   "Label11"
      Tab(0).Control(2)=   "Label12"
      Tab(0).Control(3)=   "LblReg"
      Tab(0).Control(4)=   "DtFechaHasta"
      Tab(0).Control(5)=   "DtFechaDesde"
      Tab(0).Control(6)=   "Frame1"
      Tab(0).Control(7)=   "Fr1(1)"
      Tab(0).Control(8)=   "cmdBotones(4)"
      Tab(0).Control(9)=   "cmdBotones(2)"
      Tab(0).Control(10)=   "cmdBotones(1)"
      Tab(0).Control(11)=   "cmdBotones(0)"
      Tab(0).Control(12)=   "TxtCliente"
      Tab(0).Control(13)=   "TxtNro"
      Tab(0).Control(14)=   "CmdBuscar"
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Mantenimiento"
      TabPicture(1)   =   "FrmPedidoAliterm.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Fr2(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame6"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "TDBGrid1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Fr2(2)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame3"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdBotones(12)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdBotones(11)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "SSTab2"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Fr4"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Ingreso Masivo"
      TabPicture(2)   =   "FrmPedidoAliterm.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Cmdsalirmasivo"
      Tab(2).Control(1)=   "Cmdgrabamasivo"
      Tab(2).Control(2)=   "Text4(3)"
      Tab(2).Control(3)=   "Text7"
      Tab(2).Control(4)=   "Text10"
      Tab(2).Control(5)=   "TDBGrid3"
      Tab(2).Control(6)=   "Label4(2)"
      Tab(2).Control(7)=   "Label3(6)"
      Tab(2).Control(8)=   "Label1(26)"
      Tab(2).ControlCount=   9
      Begin VB.Frame Fr4 
         BackColor       =   &H00C9955A&
         BorderStyle     =   0  'None
         Height          =   3705
         Left            =   1440
         TabIndex        =   91
         Top             =   2250
         Visible         =   0   'False
         Width           =   8790
         Begin VB.Frame Frame2 
            BackColor       =   &H00C9955A&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   165
            TabIndex        =   93
            Top             =   255
            Width           =   5355
            Begin VB.OptionButton cOpc2 
               BackColor       =   &H00C9955A&
               Caption         =   "TICKET"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Index           =   3
               Left            =   120
               TabIndex        =   133
               Top             =   -30
               Value           =   -1  'True
               Width           =   975
            End
            Begin VB.OptionButton cOpc2 
               BackColor       =   &H00C9955A&
               Caption         =   "BO"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Index           =   2
               Left            =   4245
               TabIndex        =   96
               Top             =   -30
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.OptionButton cOpc2 
               BackColor       =   &H00C9955A&
               Caption         =   "BOLETA"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Index           =   1
               Left            =   2850
               TabIndex        =   95
               Top             =   -30
               Width           =   1050
            End
            Begin VB.OptionButton cOpc2 
               BackColor       =   &H00C9955A&
               Caption         =   "FACTURA"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Index           =   0
               Left            =   1335
               TabIndex        =   94
               Top             =   -30
               Width           =   1155
            End
         End
         Begin VB.CommandButton cSeleccion 
            BackColor       =   &H80000009&
            Caption         =   "Canc&ela"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1005
            Index           =   1
            Left            =   7290
            Style           =   1  'Graphical
            TabIndex        =   92
            Top             =   2460
            Width           =   1185
         End
         Begin VB.CommandButton cSeleccion 
            BackColor       =   &H80000009&
            Caption         =   "Ace&pta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1005
            Index           =   0
            Left            =   7290
            Style           =   1  'Graphical
            TabIndex        =   130
            Top             =   1380
            Width           =   1185
         End
         Begin TextFer.TxFer TxFernumero 
            Height          =   315
            Left            =   4320
            TabIndex        =   124
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            Appearance      =   0
            Object.CausesValidation=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   20
            Text            =   ""
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            TipoDato        =   1
            SignoNegativo   =   0   'False
            MarcarTextoAlEnfoque=   -1  'True
         End
         Begin TextFer.TxFer TxFerimporte 
            Height          =   315
            Left            =   7140
            TabIndex        =   128
            Top             =   960
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            Appearance      =   0
            Object.CausesValidation=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            TipoDato        =   1
            NumeroDecimales =   2
            Formato         =   "###,###.##"
            MarcarTextoAlEnfoque=   -1  'True
         End
         Begin TextFer.TxFer TxFermoneda 
            Height          =   315
            Left            =   6270
            TabIndex        =   126
            Top             =   960
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            Appearance      =   0
            Object.CausesValidation=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            TipoDato        =   1
            NoCaracteres    =   "a-z,A-Z"
            MarcarTextoAlEnfoque=   -1  'True
            NoRangoCadena   =   -1  'True
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuoperacion 
            Height          =   315
            Left            =   240
            TabIndex        =   121
            Top             =   960
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   556
            XcodMaxLongitud =   11
            xcodwith        =   200
            NomTabla        =   "vt_conceptosdepago"
            TituloAyuda     =   "Busqueda de Concepto de Pagos"
            ListaCampos     =   "pagocodigo(1),pagodescripcion(1),pagoefectivo(1)"
            XcodCampo       =   "pagocodigo"
            XListCampo      =   "pagodescripcion"
            ListaCamposDescrip=   "Codigo,Descripcion,efectivo"
            ListaCamposText =   "pagocodigo,pagodescripcion,pagoefectivo"
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayutipo 
            Height          =   315
            Left            =   2040
            TabIndex        =   122
            Top             =   960
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   556
            XcodMaxLongitud =   11
            xcodwith        =   200
            NomTabla        =   "vt_conceptostipodepago"
            TituloAyuda     =   "Busqueda de Concepto de Pagos"
            ListaCampos     =   "pagotipocodigo(1),pagotipodescripcion(1)"
            XcodCampo       =   "pagotipocodigo"
            XListCampo      =   "pagotipodescripcion"
            ListaCamposDescrip=   "Codigo,Descripcion"
            ListaCamposText =   "pagotipocodigo,pagotipodescripcion"
         End
         Begin TrueOleDBGrid70.TDBGrid TDBpagos 
            Height          =   2055
            Left            =   240
            TabIndex        =   132
            Top             =   1410
            Width           =   6735
            _ExtentX        =   11880
            _ExtentY        =   3625
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Operacion"
            Columns(0).DataField=   "pagocodigo"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Tipo Doc."
            Columns(1).DataField=   "pagotipocodigo"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Nro.Doc."
            Columns(2).DataField=   "pagonumdoc"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Importe"
            Columns(3).DataField=   "Pagoimporte"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   4
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   14215660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=4"
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
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            MultipleLines   =   0
            CellTipsWidth   =   0
            InsertMode      =   0   'False
            DeadAreaBackColor=   16777215
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
         Begin VB.Shape Shape3 
            BorderColor     =   &H80000009&
            BorderWidth     =   3
            Height          =   3615
            Left            =   45
            Top             =   45
            Width           =   8700
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Operacion"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   195
            Index           =   28
            Left            =   270
            TabIndex        =   131
            Top             =   690
            Width           =   855
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de tarjeta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   195
            Index           =   10
            Left            =   2070
            TabIndex        =   129
            Top             =   690
            Width           =   1260
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Moneda"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   195
            Index           =   9
            Left            =   6330
            TabIndex        =   127
            Top             =   690
            Width           =   675
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Numero"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   195
            Index           =   8
            Left            =   4380
            TabIndex        =   125
            Top             =   690
            Width           =   660
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Importe"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   195
            Index           =   7
            Left            =   7200
            TabIndex        =   123
            Top             =   690
            Width           =   705
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   2460
         Left            =   60
         TabIndex        =   59
         Top             =   690
         Width           =   11940
         _ExtentX        =   21061
         _ExtentY        =   4339
         _Version        =   393216
         Style           =   1
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Datos Generales"
         TabPicture(0)   =   "FrmPedidoAliterm.frx":0060
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Fr2(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Datos Detalle"
         TabPicture(1)   =   "FrmPedidoAliterm.frx":007C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Fr1(0)"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Datos Complementarios"
         TabPicture(2)   =   "FrmPedidoAliterm.frx":0098
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Fr3(0)"
         Tab(2).ControlCount=   1
         Begin VB.Frame Fr1 
            Height          =   2055
            Index           =   0
            Left            =   -74955
            TabIndex        =   151
            Top             =   270
            Width           =   11655
            Begin VB.CheckBox Chkmasivo 
               Caption         =   "Ing.Masivo"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   192
               Left            =   7875
               TabIndex        =   191
               Top             =   765
               Width           =   1290
            End
            Begin VB.CheckBox TClie 
               Caption         =   "Cliente Eventual"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   9630
               TabIndex        =   190
               Top             =   765
               Width           =   1920
            End
            Begin VB.TextBox Text4 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   2
               Left            =   2430
               MaxLength       =   8
               TabIndex        =   189
               Top             =   630
               Width           =   1215
            End
            Begin VB.TextBox Text4 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   1
               Left            =   1890
               MaxLength       =   3
               TabIndex        =   188
               Top             =   630
               Width           =   495
            End
            Begin VB.TextBox Text4 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   1470
               MaxLength       =   2
               TabIndex        =   187
               Top             =   630
               Width           =   375
            End
            Begin VB.ComboBox Combo2 
               Height          =   315
               Left            =   1470
               Style           =   2  'Dropdown List
               TabIndex        =   23
               Top             =   1020
               Width           =   1065
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   270
               Index           =   0
               Left            =   1470
               TabIndex        =   152
               Top             =   285
               Width           =   510
               _ExtentX        =   900
               _ExtentY        =   476
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   2
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   270
               Index           =   1
               Left            =   3150
               TabIndex        =   153
               Top             =   285
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   476
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   11
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   270
               Index           =   2
               Left            =   5625
               TabIndex        =   154
               Top             =   285
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   476
               _Version        =   393216
               ClipMode        =   1
               Appearance      =   0
               Enabled         =   0   'False
               MaxLength       =   11
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   270
               Index           =   3
               Left            =   8010
               TabIndex        =   155
               Top             =   285
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   476
               _Version        =   393216
               ClipMode        =   1
               Appearance      =   0
               Enabled         =   0   'False
               MaxLength       =   11
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   270
               Index           =   4
               Left            =   10395
               TabIndex        =   156
               Top             =   285
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   476
               _Version        =   393216
               ClipMode        =   1
               Appearance      =   0
               Enabled         =   0   'False
               MaxLength       =   11
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   315
               Index           =   9
               Left            =   1485
               TabIndex        =   24
               Top             =   1395
               Width           =   7965
               _ExtentX        =   14049
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   300
               Index           =   18
               Left            =   10935
               TabIndex        =   198
               Top             =   1395
               Visible         =   0   'False
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   529
               _Version        =   393216
               ClipMode        =   1
               Appearance      =   0
               MaxLength       =   3
               PromptChar      =   "_"
            End
            Begin VB.Label Label1 
               Caption         =   "Dias Pago :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   18
               Left            =   9720
               TabIndex        =   199
               Top             =   1440
               Visible         =   0   'False
               Width           =   990
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Referen :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   180
               TabIndex        =   192
               Top             =   630
               Width           =   765
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Observacion :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   11
               Left            =   180
               TabIndex        =   163
               Top             =   1440
               Width           =   1140
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Lista Precios :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   10
               Left            =   180
               TabIndex        =   162
               Top             =   1080
               Width           =   1155
            End
            Begin VB.Label Label1 
               Caption         =   "Nro. Guia :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   8
               Left            =   9450
               TabIndex        =   161
               Top             =   330
               Width           =   855
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Nro. Pedido :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   7
               Left            =   2070
               TabIndex        =   160
               Top             =   330
               Width           =   1035
            End
            Begin VB.Label Label1 
               Caption         =   "Nro. Boleta :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   4
               Left            =   6885
               TabIndex        =   159
               Top             =   330
               Width           =   1215
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Nro. Factura :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   4410
               TabIndex        =   158
               Top             =   330
               Width           =   1110
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Punto Venta :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   180
               TabIndex        =   157
               Top             =   330
               Width           =   1125
            End
         End
         Begin VB.Frame Fr2 
            Height          =   2130
            Index           =   0
            Left            =   45
            TabIndex        =   138
            Top             =   270
            Width           =   11820
            Begin VB.ComboBox Combo1 
               Height          =   315
               Left            =   7740
               Style           =   2  'Dropdown List
               TabIndex        =   11
               Top             =   1710
               Width           =   1308
            End
            Begin VB.ComboBox Combo5 
               Height          =   315
               Left            =   10665
               Style           =   2  'Dropdown List
               TabIndex        =   7
               Top             =   1290
               Width           =   825
            End
            Begin VB.ComboBox Combo4 
               Height          =   315
               Left            =   5580
               Style           =   2  'Dropdown List
               TabIndex        =   1
               Top             =   180
               Width           =   2085
            End
            Begin VB.ComboBox Combo3 
               Height          =   315
               Left            =   1215
               Style           =   2  'Dropdown List
               TabIndex        =   0
               Top             =   180
               Width           =   2805
            End
            Begin VB.CommandButton cAyuda 
               Caption         =   "..."
               Height          =   300
               Index           =   0
               Left            =   8415
               TabIndex        =   140
               Top             =   945
               Width           =   375
            End
            Begin VB.TextBox Text5 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   288
               Left            =   9585
               TabIndex        =   139
               Top             =   945
               Width           =   1500
            End
            Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda3 
               Height          =   315
               Left            =   6090
               TabIndex        =   6
               Top             =   1290
               Width           =   3165
               _ExtentX        =   5583
               _ExtentY        =   556
               XcodMaxLongitud =   2
               xcodwith        =   100
               NomTabla        =   "tabalm"
               TituloAyuda     =   "Ayuda de Almacenes"
               ListaCampos     =   "taalma(1),tadescri(1),tipoalmacencodigo(1)"
               XcodCampo       =   "taalma"
               XListCampo      =   "tadescri"
               ListaCamposDescrip=   "Codigo,Descripcion,Tipo"
               ListaCamposText =   "taalma,tadescri,tipoalmacencodigo"
            End
            Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda2 
               Height          =   315
               Left            =   1215
               TabIndex        =   5
               Top             =   1320
               Width           =   3795
               _ExtentX        =   6694
               _ExtentY        =   556
               XcodMaxLongitud =   3
               xcodwith        =   200
               NomTabla        =   "vt_vendedor"
               TituloAyuda     =   "Ayuda de Vendedores"
               ListaCampos     =   "vendedorcodigo(1),vendedornombres(1)"
               XcodCampo       =   "vendedorcodigo"
               XListCampo      =   "vendedornombres"
               ListaCamposDescrip=   "Codigo,Descripcion"
               ListaCamposText =   "vendedorcodigo,vendedornombres"
            End
            Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda1 
               Height          =   315
               Left            =   1215
               TabIndex        =   2
               Top             =   555
               Width           =   6570
               _ExtentX        =   11589
               _ExtentY        =   556
               XcodMaxLongitud =   11
               xcodwith        =   800
               NomTabla        =   "vt_Cliente"
               TituloAyuda     =   "Ayuda de Clientes"
               ListaCampos     =   $"FrmPedidoAliterm.frx":00B4
               XcodCampo       =   "clientecodigo"
               XListCampo      =   "clienterazonsocial"
               ListaCamposDescrip=   "Codigo,Descripcion,Ruc,Direccion,Distrito,LimiteCred,Saldo,T,P,M,D"
               ListaCamposText =   $"FrmPedidoAliterm.frx":019A
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
               Index           =   10
               Left            =   8505
               TabIndex        =   25
               Top             =   180
               Width           =   1155
               _ExtentX        =   2037
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
               Height          =   315
               Index           =   17
               Left            =   9765
               TabIndex        =   201
               Top             =   180
               Visible         =   0   'False
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               ClipMode        =   1
               Appearance      =   0
               MaxLength       =   10
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   300
               Index           =   19
               Left            =   1215
               TabIndex        =   4
               Top             =   945
               Width           =   7110
               _ExtentX        =   12541
               _ExtentY        =   529
               _Version        =   393216
               Appearance      =   0
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   270
               Index           =   5
               Left            =   1215
               TabIndex        =   8
               Top             =   1710
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   476
               _Version        =   393216
               ClipMode        =   1
               Appearance      =   0
               MaxLength       =   9
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   270
               Index           =   6
               Left            =   3540
               TabIndex        =   9
               Top             =   1710
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   476
               _Version        =   393216
               ClipMode        =   1
               Appearance      =   0
               MaxLength       =   9
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   270
               Index           =   7
               Left            =   5760
               TabIndex        =   10
               Top             =   1710
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   476
               _Version        =   393216
               ClipMode        =   1
               Appearance      =   0
               MaxLength       =   9
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   270
               Index           =   8
               Left            =   10350
               TabIndex        =   12
               Top             =   1710
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   476
               _Version        =   393216
               ClipMode        =   1
               Appearance      =   0
               MaxLength       =   9
               PromptChar      =   "_"
            End
            Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuRef 
               Height          =   315
               Left            =   8505
               TabIndex        =   3
               Top             =   585
               Width           =   3240
               _ExtentX        =   5715
               _ExtentY        =   556
               XcodMaxLongitud =   11
               xcodwith        =   900
               NomTabla        =   "CT_ENTIDAD"
               TituloAyuda     =   "Busqueda de Centro de Costos"
               ListaCampos     =   "entidadcodigo(1),entidadrazonsocial(1)"
               XcodCampo       =   "entidadcodigo"
               XListCampo      =   "entidadrazonsocial"
               ListaCamposDescrip=   "Código,Descripción"
               ListaCamposText =   "entidadcodigo,entidadrazonsocial"
               Requerido       =   0   'False
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Dscto. Gral :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   2
               Left            =   135
               TabIndex        =   197
               Top             =   1755
               Width           =   1005
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Camb :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   3
               Left            =   9270
               TabIndex        =   196
               Top             =   1755
               Width           =   975
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Dscto. Prom :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   5
               Left            =   2340
               TabIndex        =   195
               Top             =   1755
               Width           =   1110
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Moneda :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   6
               Left            =   6885
               TabIndex        =   194
               Top             =   1755
               Width           =   765
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Dscto. Esp :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   9
               Left            =   4710
               TabIndex        =   193
               Top             =   1755
               Width           =   945
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Autorizacion :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   22
               Left            =   9450
               TabIndex        =   150
               Top             =   1350
               Width           =   1155
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   " Ref :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   21
               Left            =   7965
               TabIndex        =   149
               Top             =   630
               Width           =   420
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Almacen :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   17
               Left            =   5160
               TabIndex        =   148
               Top             =   1350
               Width           =   825
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Vendedor :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   16
               Left            =   135
               TabIndex        =   147
               Top             =   1350
               Width           =   900
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Cliente :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   15
               Left            =   135
               TabIndex        =   146
               Top             =   615
               Width           =   675
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Forma de  Pago :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   14
               Left            =   4095
               TabIndex        =   145
               Top             =   240
               Width           =   1395
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Fecha :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   13
               Left            =   7830
               TabIndex        =   144
               Top             =   225
               Width           =   585
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Modo Vta :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   12
               Left            =   135
               TabIndex        =   143
               Top             =   240
               Width           =   885
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Destino :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   24
               Left            =   135
               TabIndex        =   142
               Top             =   990
               Width           =   735
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "RUC :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   25
               Left            =   8910
               TabIndex        =   141
               Top             =   990
               Width           =   435
            End
         End
         Begin VB.Frame Fr3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1755
            Index           =   0
            Left            =   -74880
            TabIndex        =   62
            Top             =   450
            Width           =   11565
            Begin VB.ComboBox Combo8 
               Height          =   315
               Left            =   1290
               Style           =   2  'Dropdown List
               TabIndex        =   86
               Top             =   1290
               Width           =   1185
            End
            Begin VB.ComboBox Combo7 
               Height          =   315
               Left            =   9540
               Style           =   2  'Dropdown List
               TabIndex        =   84
               Top             =   930
               Width           =   1410
            End
            Begin VB.ComboBox Combo6 
               Height          =   315
               Left            =   7320
               Style           =   2  'Dropdown List
               TabIndex        =   83
               Top             =   930
               Width           =   1125
            End
            Begin MSMask.MaskEdBox MBox3 
               Height          =   315
               Index           =   0
               Left            =   1290
               TabIndex        =   73
               Top             =   210
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   11
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox3 
               Height          =   315
               Index           =   1
               Left            =   2745
               TabIndex        =   74
               Top             =   210
               Width           =   6495
               _ExtentX        =   11456
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox3 
               Height          =   315
               Index           =   2
               Left            =   9840
               TabIndex        =   75
               Top             =   210
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   11
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox3 
               Height          =   315
               Index           =   3
               Left            =   1290
               TabIndex        =   76
               Top             =   570
               Width           =   10185
               _ExtentX        =   17965
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox3 
               Height          =   315
               Index           =   4
               Left            =   1290
               TabIndex        =   77
               Top             =   930
               Width           =   4545
               _ExtentX        =   8017
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   20
               PromptChar      =   "_"
            End
            Begin VB.Label Dclie 
               BackStyle       =   0  'Transparent
               Caption         =   "Multidireccion"
               ForeColor       =   &H80000006&
               Height          =   255
               Index           =   8
               Left            =   240
               TabIndex        =   85
               Top             =   1380
               Width           =   1005
            End
            Begin VB.Label Dclie 
               BackStyle       =   0  'Transparent
               Caption         =   "Pais"
               ForeColor       =   &H80000006&
               Height          =   255
               Index           =   6
               Left            =   9030
               TabIndex        =   82
               Top             =   990
               Width           =   465
            End
            Begin VB.Label Dclie 
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo Persona"
               ForeColor       =   &H80000006&
               Height          =   255
               Index           =   3
               Left            =   6120
               TabIndex        =   81
               Top             =   960
               Width           =   975
            End
            Begin VB.Label Dclie 
               BackStyle       =   0  'Transparent
               Caption         =   "Ruc"
               ForeColor       =   &H80000006&
               Height          =   255
               Index           =   7
               Left            =   9420
               TabIndex        =   80
               Top             =   270
               Width           =   675
            End
            Begin VB.Label lcred 
               Alignment       =   1  'Right Justify
               ForeColor       =   &H8000000E&
               Height          =   285
               Index           =   1
               Left            =   9870
               TabIndex        =   79
               Top             =   1320
               Width           =   1605
            End
            Begin VB.Label lcred 
               Alignment       =   1  'Right Justify
               ForeColor       =   &H8000000C&
               Height          =   285
               Index           =   0
               Left            =   6780
               TabIndex        =   78
               Top             =   1350
               Width           =   1575
            End
            Begin VB.Label Dclie 
               BackStyle       =   0  'Transparent
               Caption         =   "Cliente"
               ForeColor       =   &H80000006&
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   67
               Top             =   240
               Width           =   1815
            End
            Begin VB.Label Dclie 
               BackStyle       =   0  'Transparent
               Caption         =   "Direccion"
               ForeColor       =   &H80000006&
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   66
               Top             =   600
               Width           =   1815
            End
            Begin VB.Label Dclie 
               BackStyle       =   0  'Transparent
               Caption         =   "Distrito"
               ForeColor       =   &H80000006&
               Height          =   255
               Index           =   2
               Left            =   240
               TabIndex        =   65
               Top             =   990
               Width           =   1815
            End
            Begin VB.Label Dclie 
               BackStyle       =   0  'Transparent
               Caption         =   "Saldo US$"
               ForeColor       =   &H80000006&
               Height          =   255
               Index           =   4
               Left            =   5790
               TabIndex        =   64
               Top             =   1380
               Width           =   1335
            End
            Begin VB.Label Dclie 
               BackStyle       =   0  'Transparent
               Caption         =   "Limite Cred US$"
               ForeColor       =   &H80000006&
               Height          =   255
               Index           =   5
               Left            =   8520
               TabIndex        =   63
               Top             =   1380
               Width           =   1815
            End
         End
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   -69510
         Picture         =   "FrmPedidoAliterm.frx":025F
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   585
         Width           =   1140
      End
      Begin VB.TextBox TxtNro 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73425
         MaxLength       =   11
         TabIndex        =   26
         Top             =   675
         Width           =   1365
      End
      Begin VB.TextBox TxtCliente 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73425
         MaxLength       =   50
         TabIndex        =   27
         Top             =   1035
         Width           =   3660
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Nuevo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   0
         Left            =   -68130
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   585
         Width           =   1140
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "E&ditar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   1
         Left            =   -66870
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   585
         Width           =   1140
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Eliminar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   2
         Left            =   -65640
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   585
         Width           =   1095
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   4
         Left            =   -64440
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   585
         Width           =   1140
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Procesar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   11
         Left            =   8550
         TabIndex        =   21
         Top             =   7560
         Width           =   1740
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   12
         Left            =   10395
         TabIndex        =   22
         Top             =   7560
         Width           =   1545
      End
      Begin VB.CommandButton Cmdsalirmasivo 
         Caption         =   "Cancelar"
         Height          =   540
         Left            =   -64695
         TabIndex        =   119
         Top             =   6525
         Width           =   972
      End
      Begin VB.CommandButton Cmdgrabamasivo 
         Caption         =   "Grabar"
         Height          =   540
         Left            =   -66045
         TabIndex        =   118
         Top             =   6570
         Width           =   972
      End
      Begin VB.TextBox Text4 
         Height          =   396
         Index           =   3
         Left            =   -69096
         TabIndex        =   117
         Text            =   "Text4"
         Top             =   6645
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.TextBox Text7 
         Height          =   396
         Left            =   -71064
         TabIndex        =   116
         Text            =   "0"
         Top             =   6705
         Width           =   972
      End
      Begin VB.TextBox Text10 
         Height          =   396
         Left            =   -74712
         TabIndex        =   115
         Text            =   "Text1"
         Top             =   6705
         Width           =   972
      End
      Begin VB.Frame Frame3 
         Caption         =   "Nota"
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
         Height          =   555
         Left            =   180
         TabIndex        =   102
         Top             =   7425
         Width           =   4065
         Begin VB.Label Label5 
            Caption         =   "[ENTER]"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   2130
            TabIndex        =   106
            Top             =   225
            Width           =   675
         End
         Begin VB.Label Label5 
            Caption         =   "[DEL]"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   150
            TabIndex        =   105
            Top             =   225
            Width           =   495
         End
         Begin VB.Label Label4 
            Caption         =   "Editar Item"
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
            Index           =   1
            Left            =   2820
            TabIndex        =   104
            Top             =   225
            Width           =   1125
         End
         Begin VB.Label Label4 
            Caption         =   "Eliminar Item"
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
            Index           =   0
            Left            =   840
            TabIndex        =   103
            Top             =   225
            Width           =   1095
         End
      End
      Begin VB.Frame Fr1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C9955A&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1545
         Index           =   1
         Left            =   -70995
         TabIndex        =   68
         Top             =   3015
         Visible         =   0   'False
         Width           =   4875
         Begin VB.CommandButton cBoton 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Cancela"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   2565
            Style           =   1  'Graphical
            TabIndex        =   72
            Top             =   990
            Width           =   1275
         End
         Begin VB.CommandButton cBoton 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Acepta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   1125
            Style           =   1  'Graphical
            TabIndex        =   71
            Top             =   990
            Width           =   1275
         End
         Begin VB.OptionButton cOpc 
            BackColor       =   &H00C9955A&
            Caption         =   "&FACTURACION"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   1
            Left            =   3060
            MaskColor       =   &H00FFFFFF&
            TabIndex        =   70
            Top             =   405
            Width           =   1665
         End
         Begin VB.OptionButton cOpc 
            BackColor       =   &H00C9955A&
            Caption         =   "&PEDIDO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   0
            Left            =   855
            TabIndex        =   69
            Top             =   405
            Value           =   -1  'True
            Width           =   1125
         End
         Begin VB.Image Image2 
            Height          =   540
            Left            =   270
            Picture         =   "FrmPedidoAliterm.frx":06A1
            Stretch         =   -1  'True
            Top             =   225
            Width           =   540
         End
         Begin VB.Image Image3 
            Height          =   540
            Left            =   2430
            Picture         =   "FrmPedidoAliterm.frx":0EB9
            Stretch         =   -1  'True
            Top             =   225
            Width           =   540
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H80000009&
            BorderWidth     =   4
            Height          =   1455
            Left            =   45
            Top             =   45
            Width           =   4785
         End
      End
      Begin VB.Frame Fr2 
         BackColor       =   &H00C9955A&
         BorderStyle     =   0  'None
         Height          =   675
         Index           =   2
         Left            =   120
         TabIndex        =   48
         Top             =   6660
         Width           =   11805
         Begin MSMask.MaskEdBox MBox2 
            Height          =   330
            Index           =   6
            Left            =   300
            TabIndex        =   49
            Top             =   75
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   582
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
         Begin MSMask.MaskEdBox MBox2 
            Height          =   330
            Index           =   7
            Left            =   2400
            TabIndex        =   50
            Top             =   75
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   582
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
         Begin MSMask.MaskEdBox MBox2 
            Height          =   330
            Index           =   8
            Left            =   4800
            TabIndex        =   51
            Top             =   75
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   582
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
         Begin MSMask.MaskEdBox MBox2 
            Height          =   330
            Index           =   9
            Left            =   7290
            TabIndex        =   52
            Top             =   75
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   582
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
         Begin MSMask.MaskEdBox MBox2 
            Height          =   330
            Index           =   10
            Left            =   9540
            TabIndex        =   53
            Top             =   75
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   582
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
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   58
            Top             =   435
            Width           =   1335
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Total Bruto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   255
            Index           =   1
            Left            =   2760
            TabIndex        =   57
            Top             =   435
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Total Dctos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   255
            Index           =   2
            Left            =   5160
            TabIndex        =   56
            Top             =   435
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Total I.G.V."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   255
            Index           =   3
            Left            =   7680
            TabIndex        =   55
            Top             =   435
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Neto Factura"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   255
            Index           =   4
            Left            =   9840
            TabIndex        =   54
            Top             =   435
            Width           =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            Index           =   0
            X1              =   2175
            X2              =   2175
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            Index           =   1
            X1              =   4440
            X2              =   4440
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            Index           =   2
            X1              =   6960
            X2              =   6960
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            Index           =   3
            X1              =   9360
            X2              =   9360
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            Index           =   4
            X1              =   2160
            X2              =   2160
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            Index           =   5
            X1              =   4420
            X2              =   4420
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            Index           =   6
            X1              =   6940
            X2              =   6940
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            Index           =   7
            X1              =   9340
            X2              =   9340
            Y1              =   120
            Y2              =   1215
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   1635
         Left            =   105
         TabIndex        =   36
         Top             =   4995
         Width           =   11820
         _ExtentX        =   20849
         _ExtentY        =   2884
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
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).DataField=   ""
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
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
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(19)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(24)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(29)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(34)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=System"
         PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=System"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         InsertMode      =   0   'False
         DeadAreaBackColor=   14417405
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=18,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&HC0C0C0&"
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
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
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
      Begin VB.Frame Frame1 
         Height          =   5955
         Left            =   -74790
         TabIndex        =   87
         Top             =   1710
         Width           =   11535
         Begin VB.Frame Fr5 
            BackColor       =   &H00C9955A&
            BorderStyle     =   0  'None
            Caption         =   "TIPO TRANSACCION"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C9955A&
            Height          =   1545
            Left            =   3800
            TabIndex        =   97
            Top             =   1305
            Visible         =   0   'False
            Width           =   4875
            Begin VB.CommandButton cBoton2 
               BackColor       =   &H00FFFFFF&
               Caption         =   "&Cancela"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Index           =   1
               Left            =   2610
               Style           =   1  'Graphical
               TabIndex        =   101
               Top             =   900
               Width           =   1215
            End
            Begin VB.CommandButton cBoton2 
               BackColor       =   &H00FFFFFF&
               Caption         =   "&Acepta"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Index           =   0
               Left            =   1215
               MaskColor       =   &H0000C0C0&
               Style           =   1  'Graphical
               TabIndex        =   100
               Top             =   900
               Width           =   1215
            End
            Begin VB.OptionButton cOpc3 
               BackColor       =   &H00C9955A&
               Caption         =   "FACTURACION"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   1
               Left            =   2880
               TabIndex        =   99
               Top             =   405
               Width           =   1695
            End
            Begin VB.OptionButton cOpc3 
               BackColor       =   &H00C9955A&
               Caption         =   "MODIFICA"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   0
               Left            =   855
               TabIndex        =   98
               Top             =   405
               Width           =   1275
            End
            Begin VB.Image Image5 
               Height          =   540
               Left            =   2295
               Picture         =   "FrmPedidoAliterm.frx":12E9
               Stretch         =   -1  'True
               Top             =   180
               Width           =   540
            End
            Begin VB.Image Image4 
               Height          =   540
               Left            =   270
               Picture         =   "FrmPedidoAliterm.frx":1719
               Stretch         =   -1  'True
               Top             =   180
               Width           =   540
            End
            Begin VB.Shape Shape1 
               BorderColor     =   &H80000009&
               BorderWidth     =   4
               Height          =   1455
               Left            =   45
               Top             =   45
               Width           =   4785
            End
         End
         Begin VB.Frame Frame5 
            Height          =   585
            Index           =   0
            Left            =   9540
            TabIndex        =   88
            Top             =   6570
            Width           =   2265
            Begin VB.TextBox TReg 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   285
               Left            =   1350
               TabIndex        =   90
               Top             =   210
               Width           =   765
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
               Height          =   225
               Index           =   5
               Left            =   150
               TabIndex        =   89
               Top             =   270
               Width           =   1035
            End
         End
         Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
            Height          =   5715
            Left            =   30
            TabIndex        =   107
            Top             =   150
            Width           =   11475
            _ExtentX        =   20241
            _ExtentY        =   10081
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=825"
            _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&HC9955A&"
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
      End
      Begin VB.Frame Frame6 
         Height          =   195
         Left            =   90
         TabIndex        =   111
         Top             =   3060
         Visible         =   0   'False
         Width           =   11820
      End
      Begin VB.Frame Fr2 
         Height          =   1740
         Index           =   1
         Left            =   90
         TabIndex        =   37
         Top             =   3150
         Width           =   11835
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000009&
            Height          =   825
            Left            =   900
            MaxLength       =   240
            MultiLine       =   -1  'True
            TabIndex        =   13
            Top             =   180
            Width           =   8265
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Cancelar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   10260
            Picture         =   "FrmPedidoAliterm.frx":1EF9
            Style           =   1  'Graphical
            TabIndex        =   186
            Top             =   225
            Width           =   870
         End
         Begin VB.CommandButton CmdAdd 
            Caption         =   "Añadir"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   9270
            Picture         =   "FrmPedidoAliterm.frx":233B
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   225
            Width           =   870
         End
         Begin VB.CommandButton cAyuda 
            Caption         =   "..."
            Height          =   285
            Index           =   3
            Left            =   2940
            TabIndex        =   16
            Top             =   1365
            Width           =   285
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   285
            Index           =   0
            Left            =   810
            TabIndex        =   14
            Top             =   1365
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            _Version        =   393216
            ClipMode        =   1
            Appearance      =   0
            MaxLength       =   6
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   285
            Index           =   1
            Left            =   1620
            TabIndex        =   15
            Top             =   1365
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   285
            Index           =   2
            Left            =   8700
            TabIndex        =   61
            Top             =   1365
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   -2147483648
            Enabled         =   0   'False
            MaxLength       =   8
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   285
            Index           =   3
            Left            =   9645
            TabIndex        =   17
            Top             =   1365
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            _Version        =   393216
            ClipMode        =   1
            Appearance      =   0
            MaxLength       =   10
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   285
            Index           =   4
            Left            =   10845
            TabIndex        =   18
            Top             =   1365
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            _Version        =   393216
            ClipMode        =   1
            Appearance      =   0
            MaxLength       =   10
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   285
            Index           =   5
            Left            =   11325
            TabIndex        =   19
            Top             =   555
            Visible         =   0   'False
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   503
            _Version        =   393216
            ClipMode        =   1
            Appearance      =   0
            MaxLength       =   10
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   285
            Index           =   11
            Left            =   150
            TabIndex        =   38
            Top             =   1365
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   503
            _Version        =   393216
            ClipMode        =   1
            BackColor       =   -2147483644
            Enabled         =   0   'False
            MaxLength       =   6
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   285
            Index           =   12
            Left            =   180
            TabIndex        =   60
            Top             =   735
            Visible         =   0   'False
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   -2147483633
            Enabled         =   0   'False
            MaxLength       =   6
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   270
            Index           =   13
            Left            =   240
            TabIndex        =   109
            Top             =   765
            Visible         =   0   'False
            Width           =   180
            _ExtentX        =   318
            _ExtentY        =   476
            _Version        =   393216
            BackColor       =   -2147483648
            Enabled         =   0   'False
            MaxLength       =   6
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   255
            Index           =   14
            Left            =   315
            TabIndex        =   110
            Top             =   765
            Visible         =   0   'False
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            _Version        =   393216
            ClipMode        =   1
            BackColor       =   -2147483644
            Enabled         =   0   'False
            MaxLength       =   6
            PromptChar      =   "_"
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Glosa :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   225
            TabIndex        =   200
            Top             =   225
            Width           =   555
         End
         Begin VB.Label deta 
            Alignment       =   2  'Center
            Caption         =   "Cnt. Ref"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   90
            TabIndex        =   108
            Top             =   450
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Label deta 
            Alignment       =   2  'Center
            Caption         =   "Codigo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   1710
            TabIndex        =   47
            Top             =   1125
            Width           =   1215
         End
         Begin VB.Label deta 
            Alignment       =   2  'Center
            Caption         =   "Descripción"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   3150
            TabIndex        =   46
            Top             =   1125
            Width           =   1365
         End
         Begin VB.Label deta 
            Alignment       =   2  'Center
            Caption         =   "U.M."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   8745
            TabIndex        =   45
            Top             =   1125
            Width           =   675
         End
         Begin VB.Label deta 
            Alignment       =   2  'Center
            Caption         =   "Precio Vta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   9660
            TabIndex        =   44
            Top             =   1125
            Width           =   1005
         End
         Begin VB.Label deta 
            Alignment       =   2  'Center
            Caption         =   "Dscto"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   10890
            TabIndex        =   43
            Top             =   1125
            Width           =   735
         End
         Begin VB.Label deta 
            Alignment       =   2  'Center
            Caption         =   "%Com"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   11325
            TabIndex        =   42
            Top             =   315
            Visible         =   0   'False
            Width           =   390
         End
         Begin VB.Label deta 
            Alignment       =   2  'Center
            Caption         =   "Cant.UM"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   870
            TabIndex        =   41
            Top             =   1125
            Width           =   795
         End
         Begin VB.Label Label2 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3285
            TabIndex        =   40
            Top             =   1365
            Width           =   5250
         End
         Begin VB.Label deta 
            Alignment       =   2  'Center
            Caption         =   "Item"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   7
            Left            =   135
            TabIndex        =   39
            Top             =   1125
            Width           =   555
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid3 
         Height          =   5220
         Left            =   -74715
         TabIndex        =   120
         Top             =   720
         Width           =   11070
         _ExtentX        =   19526
         _ExtentY        =   9208
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "item"
         Columns(0).DataField=   "item"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Articulo"
         Columns(1).DataField=   "articulo"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "descripcion"
         Columns(2).DataField=   "descripcion"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "unidad"
         Columns(3).DataField=   "unidad"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "saldo"
         Columns(4).DataField=   "saldo"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "IGV"
         Columns(5).DataField=   "tieneigv"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "cantidad"
         Columns(6).DataField=   "cantidad"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=900"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=820"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2514"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2434"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=9128"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=9049"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=1640"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=1561"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=1296"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=1217"
         Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(21)=   "Column(5).Width=741"
         Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=661"
         Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(25)=   "Column(6).Width=2302"
         Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=2223"
         Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
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
         CollapseColor   =   65535
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=12,.bold=0,.fontsize=780,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&HFFFFC0&"
         _StyleDefs(18)  =   ":id=6,.fgcolor=&HFFFF80&"
         _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(28)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(29)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(31)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(32)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(33)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(34)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(35)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(36)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=58,.parent=13"
         _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
         _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
         _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
         _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=46,.parent=13"
         _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
         _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
         _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=62,.parent=13"
         _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
         _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
         _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
         _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=50,.parent=13"
         _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=14"
         _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=15"
         _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=17"
         _StyleDefs(65)  =   "Named:id=33:Normal"
         _StyleDefs(66)  =   ":id=33,.parent=0"
         _StyleDefs(67)  =   "Named:id=34:Heading"
         _StyleDefs(68)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(69)  =   ":id=34,.wraptext=-1"
         _StyleDefs(70)  =   "Named:id=35:Footing"
         _StyleDefs(71)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(72)  =   "Named:id=36:Selected"
         _StyleDefs(73)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(74)  =   "Named:id=37:Caption"
         _StyleDefs(75)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(76)  =   "Named:id=38:HighlightRow"
         _StyleDefs(77)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(78)  =   "Named:id=39:EvenRow"
         _StyleDefs(79)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(80)  =   "Named:id=40:OddRow"
         _StyleDefs(81)  =   ":id=40,.parent=33"
         _StyleDefs(82)  =   "Named:id=41:RecordSelector"
         _StyleDefs(83)  =   ":id=41,.parent=34"
         _StyleDefs(84)  =   "Named:id=42:FilterBar"
         _StyleDefs(85)  =   ":id=42,.parent=33"
      End
      Begin MSComCtl2.DTPicker DtFechaDesde 
         Height          =   285
         Left            =   -73425
         TabIndex        =   28
         Top             =   1395
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         _Version        =   393216
         Format          =   87556097
         CurrentDate     =   39763
         MaxDate         =   44196
         MinDate         =   36526
      End
      Begin MSComCtl2.DTPicker DtFechaHasta 
         Height          =   285
         Left            =   -71760
         TabIndex        =   29
         Top             =   1395
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         _Version        =   393216
         Format          =   87556097
         CurrentDate     =   39763
         MaxDate         =   44196
         MinDate         =   36526
      End
      Begin VB.Label LblReg 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "(0) Pedidos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -64290
         TabIndex        =   137
         Top             =   7740
         Width           =   960
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Nro Pedido:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74685
         TabIndex        =   136
         Top             =   720
         Width           =   945
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74685
         TabIndex        =   135
         Top             =   1440
         Width           =   585
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Razon Social :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74685
         TabIndex        =   134
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   390
         Index           =   2
         Left            =   -69045
         TabIndex        =   114
         Top             =   6165
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Precio"
         Height          =   390
         Index           =   6
         Left            =   -71010
         TabIndex        =   113
         Top             =   6165
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Cantidad"
         Height          =   390
         Index           =   26
         Left            =   -74610
         TabIndex        =   112
         Top             =   6165
         Width           =   975
      End
   End
   Begin VB.CheckBox Chkentrega 
      Caption         =   "Ent diferida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3285
      TabIndex        =   174
      Top             =   1770
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7320
      TabIndex        =   169
      Top             =   2580
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7470
      TabIndex        =   168
      Top             =   2430
      Visible         =   0   'False
      Width           =   1005
   End
   Begin Crystal.CrystalReport oCrystalReport 
      Left            =   300
      Top             =   8070
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileUseRptNumberFmt=   -1  'True
      PrintFileUseRptDateFmt=   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSMask.MaskEdBox MBox 
      Height          =   255
      Index           =   15
      Left            =   810
      TabIndex        =   164
      Top             =   2385
      Visible         =   0   'False
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   450
      _Version        =   393216
      ClipMode        =   1
      MaxLength       =   10
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MBox 
      Height          =   255
      Index           =   16
      Left            =   2715
      TabIndex        =   166
      Top             =   2430
      Visible         =   0   'False
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   450
      _Version        =   393216
      ClipMode        =   1
      MaxLength       =   10
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MBox 
      Height          =   255
      Index           =   13
      Left            =   5625
      TabIndex        =   170
      Top             =   2460
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      _Version        =   393216
      ClipMode        =   1
      MaxLength       =   10
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MBox 
      Height          =   255
      Index           =   11
      Left            =   450
      TabIndex        =   173
      Top             =   2790
      Visible         =   0   'False
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   45
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TxtHor 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "hh:mm AMPM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   4
      EndProperty
      Height          =   255
      Left            =   4830
      TabIndex        =   175
      Top             =   1770
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   5
      Format          =   "HH:mm"
      Mask            =   "##:##"
      PromptChar      =   "_"
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaTc 
      Height          =   315
      Left            =   1560
      TabIndex        =   179
      Top             =   3870
      Visible         =   0   'False
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   556
      XcodMaxLongitud =   3
      xcodwith        =   200
      NomTabla        =   "vt_tipodecontacto"
      TituloAyuda     =   "Tipos de Contacto"
      ListaCampos     =   "tipocontactocodigo(1),tipocontactodescripcion(1)"
      XcodCampo       =   "tipocontactocodigo"
      XListCampo      =   "tipocontactodescripcion"
      ListaCamposDescrip=   "Codigo,Descripcion"
      ListaCamposText =   "tipocontactocodigo,tipocontactodescripcion"
      Requerido       =   0   'False
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaPro 
      Height          =   315
      Left            =   6360
      TabIndex        =   180
      Top             =   3870
      Visible         =   0   'False
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   556
      XcodMaxLongitud =   3
      xcodwith        =   200
      NomTabla        =   "Profesionales"
      TituloAyuda     =   "Ayuda de Profesionales"
      ListaCampos     =   "profesionalcodigo(1),profesionalnombres(1)"
      XcodCampo       =   "profesionalcodigo"
      XListCampo      =   "profesionalnombres"
      ListaCamposDescrip=   "Codigo,Descripcion"
      ListaCamposText =   "profesionalcodigo,profesionalnombres"
      Requerido       =   0   'False
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuTransporte 
      Height          =   315
      Left            =   3375
      TabIndex        =   184
      Top             =   3465
      Visible         =   0   'False
      Width           =   5250
      _ExtentX        =   9260
      _ExtentY        =   556
      XcodMaxLongitud =   11
      xcodwith        =   800
      NomTabla        =   "al_transporte"
      TituloAyuda     =   "Ayuda de Transporte"
      ListaCampos     =   "TRACODIGO(1),TRANOMBRE(1)"
      XcodCampo       =   "TRACODIGO"
      XListCampo      =   "TRANOMBRE"
      ListaCamposDescrip=   "Codigo,Descripcion"
      ListaCamposText =   "TRACODIGO,TRANOMBRE"
   End
   Begin VB.Label Label1 
      Caption         =   "Transportista"
      Height          =   195
      Index           =   27
      Left            =   2295
      TabIndex        =   185
      Top             =   3525
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre :"
      Height          =   255
      Index           =   30
      Left            =   5580
      TabIndex        =   183
      Top             =   3900
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo de Contacto :"
      Height          =   255
      Index           =   29
      Left            =   180
      TabIndex        =   182
      Top             =   3930
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label LblTicSer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5580
      TabIndex        =   181
      Top             =   4275
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label Label14 
      Caption         =   "datos detalle"
      Height          =   375
      Left            =   585
      TabIndex        =   178
      Top             =   3330
      Width           =   1590
   End
   Begin VB.Label Label13 
      Caption         =   "Datos generales"
      Height          =   375
      Left            =   585
      TabIndex        =   177
      Top             =   1800
      Width           =   1590
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Hora :"
      Height          =   195
      Left            =   4335
      TabIndex        =   176
      Top             =   1755
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label Label6 
      Caption         =   "Dscto Cliente"
      Height          =   255
      Left            =   6285
      TabIndex        =   172
      Top             =   2505
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "% Comision"
      Height          =   255
      Index           =   23
      Left            =   4725
      TabIndex        =   171
      Top             =   2460
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Nota de Pedido"
      Height          =   255
      Index           =   20
      Left            =   2160
      TabIndex        =   167
      Top             =   2430
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label Label1 
      Caption         =   "Otros Gastos"
      Height          =   255
      Index           =   19
      Left            =   405
      TabIndex        =   165
      Top             =   2400
      Visible         =   0   'False
      Width           =   945
   End
End
Attribute VB_Name = "FrmPedidoAliterm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
                    
Option Explicit
Dim Detraccion As Byte
Dim nLongicampo(6) As Integer
Dim rsdeta As New ADODB.Recordset
Dim rsdetax As New ADODB.Recordset
Dim rsmasivo As New ADODB.Recordset
Dim rspagos As New ADODB.Recordset
Dim wCabe(46)
Dim almacentipo As Double
Public guias_num As String, xxtipo As String

'****** Totales de Pedidos***
Dim Tbruto As Double
Dim Tigv As Double
Dim Tdscto As Double
Dim TSub As Double
Dim TImporte As Double
Dim TNeto As Double
Dim TCant As Double
Dim flag As Integer
Dim Ctrlgrilla3 As Integer
Dim masivo As Integer
'***Total Descuentos  ***

Dim DTGlobal As Double
Dim DTCliente As Double
Dim DTPPago As Double
Dim DTOficina As Double
Dim DTItem As Double
Dim DTLinea As Double
Dim DTPromo As Double

'*****************

Dim dllgeneral As New dllgeneral.dll_general

'Mensajes de Pedidos

Const W1TXT1 = "El Cliente No Existe en el Maestro de Clientes"
Const W1TXT2 = "El Cliente No Tiene Número de R.U.C. en el Maestro"
Const W1TXT3 = "El Cliente Esta Suspendido No Atender"
Const W1TXT4 = "El Cliente Ya No Tiene Credito. No Atender"

Const W1TXT6 = "Codigo del Vendedor No Existe en Tabla de Vendedores"
Const W1TXT7 = "El Codigo del Almacen No Existe en Tabla de Almacenes"

Const W1TXT9 = "El Monto de Otros Gastos debe ser un Valor Positivo"

Const W1TXT12 = "El Descuento General debe ser un Valor Positivo"
Const W1TXT13 = "El Descuento de Promoci¢n debe ser un Valor Positivo"
Const W1TXT14 = "El Descuento Pronto Pago debe ser un Valor Positivo"
Const W1TXT17 = "Codigo de la Lista de Precios No Existe"
Const W1TXT18 = "Archivo Maestro de la Lista de Precios No Existe"
Const W1TXT19 = "Codigo del Artículo No Existe en Maestro de Artículos "
Const W1TXT20 = "El Codigo del Articulo No Existe en Maestro de Precios"
Const W1TXT21 = "El Codigo del Articulo Ya Existe en el Proceso de Ventas"
Const W1TXT22 = "La Cantidad a Vender debe ser un Valor Mayor que Cero"
Const W1TXT23 = "La Cantidad a Vender es Mayor que el Actual en Almacén"
Const W1TXT24 = "El Precio de Venta debe de ser un Valor Mayor que Cero"
Const W1TXT25 = "El Descuento por Item debe ser un Valor Positivo"
Const W1TXT28 = "Debe de Ingresar el Nro. de R.U.C. del Cliente"
Const W1TXT30 = "El Importe debe ser mayor a cero"
Const W1TXT31 = ""
Const W1TXT32 = ""
Const W1TXT33 = ""

Sub ImprimirBoleta()
Dim formulas(3) As Variant
Dim Param(5) As Variant
Dim reporte As String

Param(0) = VGParamSistem.BDEmpresa
Param(1) = MBox(1).Text
Param(2) = VGParametros.empresacodigo
Param(3) = VGParametros.puntovta
Param(4) = Left(Combo1.Text, 2)

formulas(0) = "LETRAS='" & dllgeneral.NUMLET(Round(CDbl(MBox2(10)), 2)) & IIf(dllgeneral.ComboDato(Combo1.Text) = g_TipoSol, "Nuevos Soles", "Dolares Americanos") & "'"
formulas(1) = "@ruc='" & VGParametros.RucEmpresa & "'"

If VGParametros.multifacturas Then
   reporte = "vt_bolimpresa_" & VGCNx.DefaultDatabase & ".rpt"
Else
   reporte = "vt_bolimpresa_" & VGCNx.DefaultDatabase & ".rpt"
End If

Call ImpresionRptProc(reporte, formulas, Param, , "Impresion de Boletas")
End Sub

Private Sub ingresosalmacen(rrsql As Recordset, valor As Integer, valor1 As Integer)
    Dim acmd As New ADODB.Command
    Dim n As Integer
    Dim nn As String
    Dim xrsql As New ADODB.Recordset
    rrsql.MoveFirst
    n = 0
Do While Not rrsql.EOF
   n = n + 1
  nn = Format(n, "000")
    If valor = 1 Then
       Set acmd.ActiveConnection = VGgeneral
                    acmd.CommandType = adCmdStoredProc
                    acmd.CommandTimeout = 0
                    acmd.CommandText = "vt_ingresodetallealma_pro"
                    acmd.Prepared = True
                    With acmd
                        .Parameters("@base") = VGCNx.DefaultDatabase
                        .Parameters("@tabla") = "movalmdet" ' nsql
                        .Parameters("@tipo") = "1"
                        .Parameters("@item") = rsdeta!Item
                        .Parameters("@numero") = wCabe(5)
                        .Parameters("@producto") = rrsql!codart
                        .Parameters("@unidad") = rsdeta.Fields(3)
                        .Parameters("@cantidad") = rsdeta.Fields(4) * rrsql!canart
                        .Parameters("@preciopacto") = rsdeta.Fields(5)
                        .Parameters("@dsctoxitem") = rsdeta.Fields(6)
                        .Parameters("@importebruto") = rsdeta.Fields(7)
                        .Parameters("@porcomision") = rsdeta.Fields(8)
                        .Parameters("@mdsctoitem") = Tdscto
                        .Parameters("@mdsctoxlinea") = 0
                        .Parameters("@mdsctoxprom") = 0
                        .Parameters("@mimpor") = rsdeta.Fields(7)       'Previo
                        .Parameters("@unidadref") = IIf(IsNull(rsdeta.Fields(9)) Or Len(Trim(rsdeta.Fields(9))) = 0, 0, CDbl(rsdeta.Fields(9)))
                        .Parameters("@almacen") = wCabe(19)
                    End With
                    acmd.Execute
                    Set acmd = Nothing
    End If
        Set acmd.ActiveConnection = VGgeneral
                      acmd.CommandType = adCmdStoredProc
                      acmd.CommandTimeout = 0
                      acmd.CommandText = "vt_actualizoalma_pro"
                      acmd.Prepared = True
                      With acmd
                        .Parameters("@basedes") = CStr(VGCNx.DefaultDatabase)
                        .Parameters("@almacen") = wCabe(19)
                        If VGParamSistem.stockcomp = 1 Then
                           If valor1 = 0 Then
                               .Parameters("@tipo") = "3"
                            Else
                               .Parameters("@tipo") = "4"
                            End If
                        Else
                               .Parameters("@tipo") = "1"
                        End If
                        .Parameters("@articulo") = rrsql!codart
                        .Parameters("@cantidad") = rsdeta.Fields(4) * rrsql!canart
                      End With
                      acmd.Execute
                      Set acmd = Nothing
 
                    
                    rrsql.MoveNext
Loop
SQL = " update movalmcab set casitgui='F' where caalma='" & wCabe(19) & "' and catd='GR' and canumdoc='" & wCabe(5) & "'"
Set xrsql = VGCNx.Execute(SQL)
End Sub

Private Sub procImprimirguia()
Dim nguia As String
Dim ntabla As String
Dim busca As New dll_apisgen.dll_apis
Dim rb1 As New ADODB.Recordset
Dim reporte As String
Dim numguias As Integer
Dim SQL As String
Dim formulas(11) As Variant, Param(1) As Variant
              
SQL = "UPDATE vt_pedido set pedidoobserva= rtrim(pedidoobserva)+'/'+ '" & guias_num & "'"
SQL = SQL & " Where pedidonumero='" & Right(MBox(1).ClipText, 7) & "'"
             
VGCNx.Execute SQL
formulas(0) = "nro='" & MBox(2) & "'"
formulas(1) = "cliente='" & MBox3(1) & "'"
' Formulas(2) = "fecha='" & CStr(Day(CDate(MBox(10)))) & "     " & dllgeneral.DESMES(Month(CDate(MBox(10)))) & "                       " & Right(CStr(Year(CDate(MBox(10)))), 4) & "'"
formulas(2) = "fecha='" & CStr(Day(CDate(MBox(10)))) & "  " & Format(Month(CDate(MBox(10))), "00") & "  " & Right(CStr(Year(CDate(MBox(10)))), 4) & "'"
formulas(3) = "direccion='" & MBox3(3) & "'"
formulas(4) = "dni='" & MBox3(2) & "'"
formulas(5) = "opedido='" & MBox(1) & "'"
formulas(6) = "ocompra='" & Ctr_AyuRef.xclave & "'" 'MBox(17)
formulas(7) = "guia='" & guias_num & "'"
formulas(8) = "distrito='" & MBox3(4).ClipText & "'"
formulas(9) = "destino='" & MBox(19).ClipText & "'"
Set rb1 = VGCNx.Execute("select * from gr_empresa ")
If rb1.RecordCount > 0 Then
   formulas(10) = "partida='" & Escadena(rb1!empresadireccion) & "'"
 Else
  formulas(10) = "partida=''"
End If
Param(0) = VGCNx.DefaultDatabase
reporte = "al_guiaremision_" & VGCNx.DefaultDatabase & ".rpt"
Call ImpresionRptProc(reporte, formulas, Param, , "Impresion de guias")
End Sub
Private Sub procImprimirguia2()
Dim nguia As String
Dim ntabla As String
Dim busca As New dll_apisgen.dll_apis
Dim rb As New ADODB.Recordset
Dim rb1 As New ADODB.Recordset
Dim contador As Double
Dim contador1 As Double
Dim numguias As Integer
Dim SQL As String
Dim inicio As Integer
Dim fin As Integer
Dim J As Integer
        
'        VGcnx.Execute "UPDATE vt_pedido set pedidonrogiarem='" & nguia & "'" & _
'                 " Where pedidonumero='" & MBox(1).ClipText & "'"
                   
'        Set rb = VGcnx.Execute("select pedidoobserva from vt_pedido Where pedidonumero='" & MBox(1).ClipText & "'")
            
            If cOpc2(0).Value Then
               ntabla = "vt_detallepedido"
             Else
               If cOpc2(1).Value Then
                  ntabla = "vt_detallepedido"
                Else
                  If cOpc2(2).Value Then
                     ntabla = "vt_detallepedido"
                   Else
                     ntabla = g_DetallePuntoVta
                  End If
               End If
            End If
        contador = 0
       ' VGcnx.Execute "delete from gtempfile2filas"
        Set rb = VGCNx.Execute("select * from gtempfile inner join maeart on productocodigo=acodigo order by alinea,agrupo,acodigo ")
        If rb.RecordCount > 0 Then
           If rb.RecordCount Mod 50 > 0 Then
              numguias = Int(rb.RecordCount / 50) + 1
            Else
              numguias = Int(rb.RecordCount / 50)
           End If
           rb.MoveFirst
           Do While contador < numguias
              contador = contador + 1
              inicio = (contador - 1) * 50 + 1
              If contador * 50 > rb.RecordCount Then
                 fin = rb.RecordCount
               Else
                 fin = contador * 50
              End If
              
              nguia = g_pedserie & Right("000000000000" & TraeDataSerie("select puntovtadoccorr from vt_puntovtadocumento where documentocodigo='GR' and puntovtadocserie='" & g_pedserie & "' and puntovtacodigo='" & g_ptoventa & "'", VGCNx), 8)
              
              VGCNx.Execute "Update vt_puntovtadocumento " & _
                  " set puntovtadoccorr='" & Right("00000000" & Trim(CStr(Val(nguia) + 1)), 8) & "'" & _
                  " Where documentocodigo='GR' and puntovtacodigo='" & g_ptoventa & "' and puntovtadocserie='" & g_pedserie & "'"

              SQL = "UPDATE vt_pedido set pedidoobserva= rtrim(pedidoobserva)+'/'+ '" & nguia & "'"
              SQL = SQL & " Where pedidonumero='" & Right(MBox(1).ClipText, 7) & "'"
             
             VGCNx.Execute SQL
             ' VGcnx.Execute "UPDATE vt_pedido set pedidoobserva= rtrim(pedidoobserva)+' / '+ '" & nguia & "'" & _
             '      " Where pedidonumero='" & MBox(1).ClipText & "'"
              guias_num = guias_num + nguia + " / "
              contador1 = 0
              If fin > rb.RecordCount Then
                 fin = rb.RecordCount - inicio
              End If
              VGCNx.Execute "delete from gtempfile2filas"
              For J = inicio To fin
                     contador1 = contador1 + 1
                     If contador1 <= 25 Then
                        SQL = "INSERT INTO gtempfile2filas(item,producto1,descripcion1,cantidad1,importe1)"
                        SQL = SQL & " VALUES ( '" & contador1 & "','" & RTrim(rb!productocodigo) & "','" & RTrim(rb!productodescripcion) & "','" & rb!detpedcantpedida & "','" & rb!detpedimpbruto & "')"
                      Else
                        TCant = contador1 - 25
                        SQL = "UPDATE gtempfile2filas set producto2 ='" & RTrim(rb!productocodigo) & "',"
                        SQL = SQL & " descripcion2='" & RTrim(rb!productodescripcion) & "',"
                        SQL = SQL & "cantidad2='" & rb!detpedcantpedida & "',"
                        SQL = SQL & "importe2= '" & rb!detpedimpbruto & "'"
                        SQL = SQL & " where item = " & TCant & ""
                     End If
                     VGCNx.Execute SQL
                     rb.MoveNext
               Next J
               oCrystalReport.Reset
               oCrystalReport.ReportFileName = VGParamSistem.Rutareport & "Repguiaimpresa.rpt"
               oCrystalReport.LogOnServer "pdssql.dll", _
                    busca.LeerIni(App.Path & "\MARFICE.INI", "conexion", "SERVIDOR", ""), _
                    busca.LeerIni(App.Path & "\MARFICE.INI", "conexion", "BDDATOS", ""), _
                    busca.LeerIni(App.Path & "\MARFICE.INI", "conexion", "USUARIO", ""), _
                    busca.LeerIni(App.Path & "\MARFICE.INI", "conexion", "PASSW", "")
               oCrystalReport.Connect = VGcadenareport2
                oCrystalReport.Destination = crptToWindow
                oCrystalReport.WindowState = crptMaximized
                oCrystalReport.DiscardSavedData = True
                With oCrystalReport
                       .formulas(0) = "nro='" & MBox(2) & "'"
                       .formulas(1) = "cliente='" & MBox3(1) & "'"
   '                    .Formulas(2) = "fecha='" & CStr(Day(CDate(MBox(10)))) & "     " & dllgeneral.DESMES(Month(CDate(MBox(10)))) & "                       " & Right(CStr(Year(CDate(MBox(10)))), 4) & "'"
                       .formulas(2) = "fecha='" & CStr(Day(CDate(MBox(10)))) & "     " & Format(Month(CDate(MBox(10))), "00") & "      " & Right(CStr(Year(CDate(MBox(10)))), 4) & "'"
                       .formulas(3) = "direccion='" & MBox3(3) & "'"
                       .formulas(4) = "dni='" & MBox3(2) & "'"
                       .formulas(5) = "opedido='" & MBox(1) & "'"
                       .formulas(6) = "ocompra='" & Ctr_AyuRef.xclave & "'" 'MBox(17)
                       .formulas(7) = "guia='" & nguia & "'"
                       .formulas(8) = "distrito='" & MBox3(4).ClipText & "'"
                       .formulas(9) = "destino='" & MBox(19).ClipText & "'"
                       Set rb1 = VGCNx.Execute("select * from gr_empresa where empresacodigo='" & VGParametros.empresacodigo & "'")
                       If rb1.RecordCount > 0 Then
                           .formulas(10) = "partida='" & Escadena(rb1!empresadireccion) & "'"
                       Else
                           .formulas(10) = "partida=''"
                       End If
                       If .Status <> 2 Then .Action = 1
                End With
                  SQL = nguia
                 MsgBox "Proceda a imprimir la GUIA DE REMISION .", vbInformation, SQL
            Loop
        End If
        rb.Close

End Sub

Private Sub cAyuda_Click(Index As Integer)
Dim xsql As New ADODB.Recordset
  nAyuda = "": nDetalle = ""
  If Index = 0 And Len(Trim(MBox(19))) = 0 Then    'Ayuda de Punto de LLegada
    If dllgeneral.VerificaDatoExistente(VGCNx, "select * from vt_cliente where clientecodigo='" & Ctr_Ayuda1.xclave & "'") = 1 Then
       Dim gfiltra(1, 2) As String
       gfiltra(1, 1) = "Descripcion": gfiltra(1, 2) = "clientedireccion"
       FrmAyuda.TipoForma = 1
       FrmAyuda.BConexion = VGCNx
       FrmAyuda.BTabla = "vt_clientedireccion"
       FrmAyuda.Bdata = "0"
       FrmAyuda.BCampos = "Cliedirnumero as Codigo,Cliedirdireccion as Descripcion"
       FrmAyuda.BOrden = "Cliedirnumero"
       FrmAyuda.BCondi = "clientecodigo='" & Ctr_Ayuda1.xclave & "'"
       FrmAyuda.BFiltro = gfiltra
    Else
        nAyuda = "": nDetalle = ""
        MsgBox "No existen Direcciones Anexas...", vbInformation, MsgTitle
        Exit Sub
    End If
  ElseIf Index = 3 Then                             ' Ayuda de Productos
'       If Len(Label2) > 0 Then
'         SendKeys "{tab}"
'         Exit Sub
'       End If
       Dim sfiltra(1 To 2, 1 To 2) As String
       sfiltra(1, 1) = "Codigo": sfiltra(1, 2) = "acodigo"
       sfiltra(2, 1) = "Descripcion": sfiltra(2, 2) = "adescri"
       FrmAyuda.TipoForma = 1
       FrmAyuda.BConexion = VGCNx
       SQL = " select stalma=isnull(stalma,0),acodigo,acodigo2,adescri,stskdis=isnull(stskdis,0),stskcom=isnull(stskcom,0) into ##XX_VENTAS from maeart left join stkart b on acodigo=stcodigo "
 '      If almacentipo = 1 Then
 '        SQL = SQL & " Union All select stalma,codkit,acodigo2,adescri,stskdis=min(stskdis),stskcom=min(stskcom) from (select stalma,codkit,acodigo2=acodigo2+' ** ',adescri,"
 '        SQL = SQL & " codart,stskdis=floor((stskdis)/canart),stskcom=floor((stskcom)/canart) from kits b inner join maeart on "
 '        SQL = SQL & " codkit=acodigo inner join stkart c on codart=stcodigo) z group by stalma,codkit,acodigo2,adescri"
 '      End If
       If Combo2.ListCount > 0 Then
          If VGParamSistem.kitvirtual = 1 Then
               SQL = " select stalma=0,acodigo,acodigo2,adescri,stskdis=0,stskcom=0 into ##XX_VENTAS from maeart "
               If ExisteElem(0, VGCNx, "##xx_ventas") Then Set xsql = VGCNx.Execute("drop table ##xx_ventas")
                Set xsql = VGCNx.Execute(SQL)
                FrmAyuda.BTabla = " ##xx_ventas"
             Else
                FrmAyuda.BTabla = " maeart left join stkart ON acodigo=stcodigo"
          End If
       Else
              FrmAyuda.BTabla = " maeart left join stkart ON acodigo=stcodigo "
       End If
       FrmAyuda.Bdata = "2"
       FrmAyuda.Bdato = Escadena(MBox2(1).Text)
       If modoventa.ctrlinventario = 0 Then
          FrmAyuda.BCampos = "acodigo as Codigo,adescri as Descripcion"
       Else
          If VGParamSistem.stockcomp = 1 Then
             If VGParamSistem.kitvirtual = 1 Then
                 FrmAyuda.BCampos = "acodigo as Codigo,acodigo2 as tipo,adescri as Descripcion,stskdis-stskcom as Stock"
               Else
                 FrmAyuda.BCampos = "acodigo as Codigo,adescri as Descripcion,stskdis-stskcom as Stock"
             End If
           Else
             FrmAyuda.BCampos = "acodigo as Codigo,adescri as Descripcion,stskdis as Stock"
          End If
       End If
       FrmAyuda.BOrden = "adescri"
       If modoventa.ctrlinventario = "1" Then
          If VGParamSistem.stockcomp = 1 Then
             If VGParamSistem.kitvirtual = 1 Then
                FrmAyuda.BCondi = ""  ' stalma='" & Ctr_Ayuda3.xclave & "'"
              Else
                FrmAyuda.BCondi = "stalma='" & Ctr_Ayuda3.xclave & "' and stskdis-stskcom>0"
            End If
           Else
            FrmAyuda.BCondi = "stalma='" & Ctr_Ayuda3.xclave & "' and stskdis>0"
          End If
       Else
   '       FrmAyuda.BCondi = "stalma='" & Ctr_Ayuda3.xclave & "' "
          FrmAyuda.BCondi = ""
          
       End If
       FrmAyuda.BFiltro = sfiltra
   Else
       SendKeys "{tab}"
       Exit Sub
   End If
   FrmAyuda.Show 1
   If Index = 3 Then
       MBox2(1) = Escadena(nAyuda):   Label2 = Escadena(nDetalle)
       xxtipo = nDetalle
       MBox2(3).SetFocus
   ElseIf Index = 0 Then
       MBox(19) = Escadena(nDetalle)
   End If
   nAyuda = "": nDetalle = ""
End Sub

Private Sub cBoton_Click(Index As Integer)
  Dim J As Integer
  If Index = 0 Then
       Fr1(1).Visible = False
       TClie.Value = 0
       Limpiartexto MBox, 2, 9
       MBox(0).Enabled = False:  MBox(1).Enabled = False
       MBox(0).Text = g_ptoventa
       MBox(1) = g_pedserie & Right("000000000000" & TraeDataSerie("select puntovtadoccorr from vt_puntovtadocumento where documentocodigo='" & g_tipoped & "' and puntovtadocserie='" & g_pedserie & "' and puntovtacodigo='" & g_ptoventa & "' and empresacodigo='" & VGParametros.empresacodigo & "'", VGCNx), 8)
       MBox(2) = g_facserie & Right("000000000000" & TraeDataSerie("select puntovtadoccorr from vt_puntovtadocumento where documentocodigo='" & g_tipofac & "' and puntovtadocserie='" & g_facserie & "' and puntovtacodigo='" & g_ptoventa & "' and empresacodigo='" & VGParametros.empresacodigo & "'", VGCNx), 8)
       MBox(3) = g_bolserie & Right("000000000000" & TraeDataSerie("select puntovtadoccorr from vt_puntovtadocumento where documentocodigo='" & g_tipobol & "' and puntovtadocserie='" & g_bolserie & "' and puntovtacodigo='" & g_ptoventa & "' and empresacodigo='" & VGParametros.empresacodigo & "'", VGCNx), 8)
       MBox(4) = g_guiaserie & Right("000000000000" & TraeDataSerie("select puntovtadoccorr from vt_puntovtadocumento where documentocodigo='" & g_tipoguia & "' and puntovtadocserie='" & g_guiaserie & "' and puntovtacodigo='" & g_ptoventa & "' and empresacodigo='" & VGParametros.empresacodigo & "'", VGCNx), 8)
       LblTicSer.Caption = g_ticserie & Right("000000000000" & TraeDataSerie("select puntovtadoccorr from vt_puntovtadocumento where documentocodigo='" & g_tipoticket & "' and puntovtadocserie='" & g_ticserie & "' and puntovtacodigo='" & g_ptoventa & "' and empresacodigo='" & VGParametros.empresacodigo & "'", VGCNx), 8)
       
       MBox(5) = numero(0): MBox(6) = numero(0): MBox(7) = numero(0): MBox(8) = numero(TraeTipoCambio(Date, VGCNx))
       MBox(9) = Escadena(VGParamSistem.mensaje)
       MBox(19) = ""
       Ctr_Ayuda1.xnombre = Empty: Ctr_Ayuda1.xclave = Empty:
       MBox(10) = Format(VGParamSistem.FechaTrabajo, "dd/mm/yyyy")
       MBox(13) = numero(0)
       MBox(15) = numero(0)
       MBox(16) = 0: Ctr_AyuRef.xclave = Empty: Ctr_AyuRef.xnombre = Empty: MBox(18) = "0"
       'MBox (17)
       For J = 0 To 5
          MBox2(J) = ""
       Next J
       Set rsdeta = Nothing
       
       CargaGrilla

     'Se activa los parametros deventa
       Combo1.ListIndex = VerificaCombo(Combo1, VGParamSistem.moneda)     'moneda
       Combo2.ListIndex = VerificaCombo(Combo2, VGParamSistem.listapre)   'listaprecios
       'Combo2.Enabled = False
       MBox(8) = numero(VGParamSistem.tipocambio)                         'tipo de cambio
       Ctr_Ayuda3.xclave = Escadena(VGParamSistem.almacen)                'almacen
       'Ctr_Ayuda3.Filtro = "empresacodigo='" & VGParametros.empresacodigo & "' and puntovtacodigo='" & VGParametros.puntovta & "'"
       Call Ctr_Ayuda3.Ejecutar
       If Len(Trim(modoventa.almacenes)) > 0 Then
          Ctr_Ayuda3.Filtro = "taalma in (" & Trim(modoventa.almacenes) & ") "
          'Ctr_Ayuda3.Ejecutar and puntovtacodigo='" & VGParametros.puntovta & "'
       End If

       
       MBox(13).Enabled = IIf(VGParamSistem.comivende = "F", False, True)                     'comision de vendedor
       
      'Se activa los parametros de punto de venta
       MBox(2).Enabled = IIf(VGParametros.nrofactura = "1" And VGParametros.ventaauto = "0", True, False)
       MBox(3).Enabled = IIf(VGParametros.nroboleta = "1" And VGParametros.ventaauto = "0", True, False)
       MBox(4).Enabled = IIf(VGParametros.nroguia = "1" And VGParametros.ventaauto = "0", True, False)
       
     'Activamos el Tab
       Activa 1
       SSTab2.TabEnabled(2) = False
       SSTab2.Tab = 0
       'MBox(5).SetFocus
       
  ElseIf Index = 1 Then
      Fr1(1).Visible = False
  End If
 
 Text3.Enabled = IIf(VGParametros.puntovta <> "01", False, True)
 Combo3.SetFocus
End Sub

Private Sub cBoton2_Click(Index As Integer)
    If Index = 0 Then
        cOpc(0).Value = False
        cOpc(1).Value = False
    Else
        Fr5.Visible = False
        Exit Sub
    End If
    Fr5.Visible = False
    Carga_Pedido
    Activa 1
    MBox(0).Enabled = False
    MBox(1).Enabled = False
    MBox(5).SetFocus
    g_TipoMovi = 2

End Sub



Private Sub CmdAdd_Click()
Dim SQL As String
Dim nregi As Long
Dim wposi, posi As Integer

Dim wflag As Integer
Dim rssql As New ADODB.Recordset
Dim rsk As New ADODB.Recordset

'---------------------------------------------------------------------------------------------
If Len(Trim(MBox2(0))) > 0 And Not IsNumeric(MBox2(0)) Then
   MsgBox "Numero de item no valido", vbCritical, "Sistema"
   Exit Sub
End If

If dllgeneral.VerificaDatoExistente(VGCNx, "select * from stkart where stcodigo='" & MBox2(1).Text & "' and stalma='" & Ctr_Ayuda3.xclave & "'") = 0 And Len(Trim(MBox2(1))) > 0 Then
'    Call cAyuda_Click(3)
'    MBox2(1).MaxLength = 20
'   Exit Sub
Else
  wflag = verificaproducto()
   If wflag = 1 Then
      Label2 = ""
      MsgBox "Ya ingreso el producto...Verifique!!!", vbInformation, MsgTitle
      cAyuda(3).SetFocus
      Exit Sub
   End If
End If
   
If Len(Trim(MBox2(1))) = 0 Then
   MsgBox "Falta ingresar producto.", vbCritical, "Sistema"
   cAyuda(3).SetFocus
   Exit Sub
End If
      
If Not (dllgeneral.ValidaCadena(MBox2(0), "N") Or IsNumeric(MBox2(0))) Then
   MsgBox Msg29, vbInformation, "AVISO"
   Call dllgeneral.Enfoquetexto(MBox2(0))
   Exit Sub
End If

If Not IsNumeric(MBox2(3)) Then
      MsgBox "El precio unitario debe ser numerico ", vbInformation, "Sistema"
      MBox2(3).SetFocus
      Exit Sub
End If

If Not IsNumeric(MBox2(4)) Then
      MsgBox "El porcentaje de dscto debe ser numerico ", vbInformation, "Sistema"
      MBox2(4).SetFocus
      Exit Sub
End If


'---------------------------------------------------------------------------------------------


wflag = verificaproducto()
If wflag = 1 Then
    Label2 = ""
    MsgBox "Ya ingreso el producto...Verifique!!!", vbInformation, MsgTitle
    MBox2(1).SetFocus
    Exit Sub
End If
      
If Trim(MBox2(4)) = "" Then MBox2(4) = 0
If Trim(MBox2(5)) = "" Then MBox2(5) = 0

If Trim(MBox2(3)) = "" Or Trim(MBox2(4)) = "" Or Trim(MBox2(5)) = "" Then
    MsgBox Msg29, vbInformation, "AVISO"
    Call dllgeneral.Enfoquetexto(MBox2(1))
    Exit Sub
End If
      

If modoventa.nroitem < TDBGrid1.ApproxCount Then
    MsgBox "Excede el Numero de Items del Documento..!!", vbInformation, MsgTitle
    Exit Sub
End If

nregi = 0
wposi = 0

If rsdeta.RecordCount > 0 Then
    rsdeta.MoveLast
    wposi = rsdeta.Fields(0)
    posi = rsdeta.Fields(0)
    rsdeta.MoveFirst
    Do Until rsdeta.EOF
       If rsdeta.Fields(0) = MBox2(11) Then
          posi = rsdeta.Fields(0)
          Exit Do
       End If
       nregi = nregi + 1
       rsdeta.MoveNext
    Loop
End If
    
If rsdeta.RecordCount = nregi Then
  wposi = wposi + 1
  posi = wposi
  rsdeta.AddNew
End If
    
rsdeta.Fields(0) = posi
rsdeta.Fields(1) = Trim(Escadena(MBox2(1)))
rsdeta.Fields(2) = Left(Escadena(Label2) & Space(40), 40)
rsdeta.Fields(3) = Trim(MBox2(2))
rsdeta.Fields(4) = Escadena(MBox2(0))
    
If VGParamSistem.tieneigv = "1" Then
   rsdeta.Fields(5) = MBox2(3)
   'rsdeta.Fields(12) = MBox2(3).Tag
Else
   If modoventa.impuestos = "1" Then
      rsdeta.Fields(5) = (MBox2(3) / (1 + VGParamSistem.Igv))
      rsdeta.Fields(12) = MBox2(3).Tag
      SQL = " select tieneigv from  grupo,maeart where acodigo='" & rsdeta.Fields(1) & "'"
      SQL = SQL & " and afamilia=fam_codigo and alinea=lin_codigo and agrupo=gru_codigo "
      Set rssql = VGCNx.Execute(SQL)
      If rssql.RecordCount > 0 Then
         If Not rssql!tieneigv = "1" Then
            rsdeta.Fields(5) = MBox2(3)
    '        rsdeta.Fields(12) = MBox2(3).Tag
         End If
      End If
   Else
      rsdeta.Fields(5) = MBox2(3).Text
    '  rsdeta.Fields(12) = MBox2(3).Tag
   End If
End If
    
rsdeta.Fields(6) = numero(MBox2(4))
rsdeta.Fields(7) = numero(MBox2(0) * MBox2(3) - ((MBox2(0) * MBox2(3)) * (MBox2(4) / 100))) 'numero(MBox2(0) * MBox2(3))   ' IIf(VGParamSistem.tieneigv = "1", (MBox2(3) / (1 + (VGParamSistem.igv / 100))), MBox2(3)))
rsdeta.Fields(8) = numero(MBox2(5))
rsdeta.Fields(9) = IIf(Len(Trim(MBox2(12))) = 0, 0, Format(MBox2(12), "##,###,##0"))
rsdeta.Fields(10) = numero(MBox2(13))
rsdeta.Fields(11) = IIf(IsNull(MBox2(14)) Or Len(Trim(MBox2(14))) = 0, 0, MBox2(14))
'rsdeta.Fields(14) = Text3.Text
    
Set rsk = VGCNx.Execute("select * from listapre" & Trim(Combo2.Text) & " where productocodigo='" & Trim(Escadena(MBox2(1))) & "' and almacencodigo='" & Trim(Ctr_Ayuda2.xclave) & "'")
If rsk.RecordCount > 0 Then
   rsdeta.Fields(12) = rsk.Fields("productoprecvta")
Else
   rsdeta.Fields(12) = 0
End If
    
rsk.Close
Set rsk = Nothing
rsdeta!tipo = IIf(Right(RTrim(xxtipo), 1) = "*", "*", " ")
rsdeta.Update
Label2 = ""
   
ConfigGrid
Totales

MBox2(11) = wposi + 1

If MBox2(12).Enabled = True Then
  MBox2(12).SetFocus
Else
    MBox2(0).Text = Empty
  MBox2(0).SetFocus
End If

flag = 0
Exit Sub

End Sub

Private Sub cmdBotones_Click(Index As Integer)
   Dim asql As String
   Dim rswork As New ADODB.Recordset
   Dim acmd As New ADODB.Command
   Dim J, nl As Integer
   Dim nflag As Integer
   
   On Error GoTo vererror
   'On Error Resume Next
   
   nflag = 0
   Select Case Index
    Case 0
        Fr1(1).Visible = True
        Limpiartexto MBox2, 6, 10
        Fr1(0).Enabled = True
        Fr2(0).Enabled = True
        Fr3(0).Enabled = True
        TClie.Enabled = True
        Text3 = "": Text4(0) = "": Text4(1) = "": Text4(2) = ""
        Call CargarModo
        cOpc(0).Value = False: cOpc(1).Value = False: cOpc2(0).Value = False
        cOpc2(1).Value = False: cOpc2(2).Value = False
        cOpc3(0).Value = False: cOpc3(1).Value = False
        cOpc(0).SetFocus
        g_TipoMovi = 1
        masivo = 0
'        rsmasivo.Reset
        
    Case 1
       If TDBGrid2.Row >= 0 Then
          Fr1(0).Enabled = True
          Fr2(0).Enabled = True
          Fr3(0).Enabled = True
          TClie.Enabled = True
          Limpiartexto MBox2, 6, 10
          cOpc(0).Value = False: cOpc(1).Value = False: cOpc2(0).Value = False
          cOpc2(1).Value = False: cOpc2(2).Value = False
          cOpc3(0).Value = False: cOpc3(1).Value = False
          Text3 = "": Text4(0) = "": Text4(1) = "": Text4(2) = ""
          Fr5.Visible = True
          cOpc3(0).SetFocus
       End If
    Case 2
       If TDBGrid2.Row >= 0 Then
        asql = "pedidonumero='" & TDBGrid2.Columns(0).Text & "'"
        If dllgeneral.EliminaReg(VGCNx, g_DetallePuntoVta, asql) = 1 Then
           If VGParamSistem.stockcomp = 1 Then
             Set rswork = VGCNx.Execute("select productocodigo,detpedcantpedida from " & g_DetallePuntoVta & " where " & asql & "")
             If rswork.RecordCount > 0 Then
                  rswork.MoveFirst
                  Do Until rswork.EOF
                       Set acmd.ActiveConnection = VGgeneral
                       acmd.CommandType = adCmdStoredProc
                       acmd.CommandTimeout = 0
                       acmd.CommandText = "vt_actualizoalma_pro"
                       acmd.Prepared = True
                       With acmd
                             .Parameters("@basedes") = CStr(VGCNx.DefaultDatabase)
                             .Parameters("@almacen") = wCabe(19)
                             .Parameters("@tipo") = "3"
                             .Parameters("@articulo") = Trim(rswork.Fields(0))
                             .Parameters("@cantidad") = rswork.Fields(1) * -1
                      End With
                      acmd.Execute
                      Set acmd = Nothing
                      rswork.MoveNext
                Loop
             End If
           End If
            VGCNx.Execute "Delete From " & g_PedidoPuntoVta & " where " & asql
        End If
        Listado (0)
       End If
    Case 4
       Unload Me
    Case 11
        If IsNull(Ctr_Ayuda1.xclave) Or Len(Trim(Ctr_Ayuda1.xclave)) = 0 Then
           MsgBox W1TXT1, vbInformation, MsgTitle
           SSTab2.Tab = 0
           Ctr_Ayuda1.SetFocus
           Exit Sub
        End If
        If IsNull(Ctr_Ayuda2.xclave) Or Len(Trim(Ctr_Ayuda2.xclave)) = 0 Then
           MsgBox W1TXT6, vbInformation, MsgTitle
           SSTab2.Tab = 0
           Ctr_Ayuda2.SetFocus
           Exit Sub
        End If
        If IsNull(Ctr_Ayuda3.xclave) Or Len(Trim(Ctr_Ayuda3.xclave)) = 0 Then
           MsgBox W1TXT7, vbInformation, MsgTitle
           SSTab2.Tab = 1
           Ctr_Ayuda3.SetFocus
           Exit Sub
        End If
        If IsNull(MBox(8)) Or Len(Trim(MBox(8))) = 0 Or CDbl(MBox(8)) <= 0 Then
           MsgBox "Falta Tipo de Cambio", vbInformation, MsgTitle
           Call dllgeneral.Enfoquetexto(MBox(8))
           Exit Sub
        End If
        If IsNull(MBox(15)) Or Len(Trim(MBox(15))) = 0 Or CDbl(MBox(15)) < 0 Then
           MsgBox W1TXT9, vbInformation, MsgTitle
           SSTab2.Tab = 1
           Exit Sub
        End If
        If dllgeneral.VerificaDatoExistente(VGCNx, "select * from vt_cliente where clientecodigo='" & MBox3(0) & "' and clientesuspendido='1'") = 1 And MBox3(0) <> g_Eventual Then
           MsgBox W1TXT3, vbInformation, MsgTitle
           Exit Sub
        End If
'        If dllgeneral.VerificaDatoExistente(VGCNx, "select * from vt_cliente where clientecodigo='" & MBox3(0) & "' and ((clientelimitecreddolar-clientesaldodolares)*" & MBox(8) & "+ (clientelimitecredsoles-clientesaldosoles))-" & TNeto & " <=0") = 1 And MBox3(0) <> g_Eventual Then
'           MsgBox W1TXT4, vbInformation, MsgTitle
'           Exit Sub
'        End If
        If Len(Trim(Text4(0))) > 0 Then
            If dllgeneral.VerificaDatoExistente(VGCNx, "select * from vt_pedido where pedidotipofac='" & Text4(0) & "' and pedidonrofact='" & Trim(Text4(1)) & Trim(Text4(2)) & "'") = 0 Then
               MsgBox "No existe documento...Verifique!!!", vbInformation, "AVISO"
               Exit Sub
            End If
        End If
        
       If cOpc(0).Value Or cOpc3(0).Value Then
          nflag = 1
          VGCNx.BeginTrans
          If GrabarData() = 1 Then
            VGCNx.CommitTrans
            nflag = 0
            g_TipoMovi = 0
            If modoventa.emitehoja = "1" Then
               nl = IIf(modoventa.copiashoja > 0, modoventa.copiashoja, 0)
               If nl > 0 Then
                   For J = 1 To nl
                      Call DocImprimir
                   Next J
               End If
            End If
            Activa 2
            Listado (0)
            Exit Sub
          Else
             VGCNx.RollbackTrans
             nflag = 0
             g_TipoMovi = 0
             Activa 2
             Exit Sub
          End If
       Else
          If cOpc(1).Value Or cOpc3(1).Value Then
                cargar
                Fr4.Visible = True
                cOpc2(0).Value = Escadena(IIf(modoventa.documento <> g_tipofac, False, True))
                cOpc2(1).Value = Escadena(IIf(modoventa.documento <> g_tipobol, False, True))
                cOpc2(2).Value = Escadena(IIf(modoventa.documento <> g_tipoguia, False, True))
                cOpc2(3).Value = Escadena(IIf(modoventa.documento <> g_tipoticket, False, True))   'g_tipoticket
                
                TxFerimporte.Text = Trim(MBox2(10))
                TxFermoneda.Text = Left(Combo1.Text, 2)
                Ctr_Ayuoperacion.xclave = "01"
                Ctr_Ayuoperacion.xnombre = "EFECTIVO"
                Ctr_Ayutipo.Visible = False
                TxFernumero.Visible = False
                'SSTab1.Enabled = False
                Exit Sub
           Else
                g_TipoMovi = 0
                Activa 2
                Exit Sub
           End If
       End If
       g_TipoMovi = 0
       If Right(TxtHor.Text, 2) > 59 Then
        MsgBox "Hora no valida", vbInformation, "Sistema"
        TxtHor.SetFocus
End If
    Case 12
       Activa 2
       g_TipoMovi = 0
   End Select
      

   Set rsdetax = Nothing
   Exit Sub
vererror:
    If Err Then
       If nflag = 1 Then
         VGCNx.RollbackTrans
       End If
       MsgBox "Comunicarse con Sistemas ...!!" & Chr(13) & Chr(10) & Err.Number & "-" & Err.Description
       Err = 0
       Exit Sub
       Resume
    End If
End Sub

Public Function Activa(ntipo As Integer)
    If ntipo = 1 Then
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = True
        SSTab1.Tab = 1
    ElseIf ntipo = 2 Then
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = False
        SSTab1.Tab = 0
    End If
End Function

Private Sub CmdBuscar_Click()
Listado (1)
TxtNro.SetFocus
End Sub

Private Sub Cmdgrabamasivo_Click()
    Call grabamasivo
    Call dllgeneral.ActivaTab(1, 2, SSTab1)
End Sub

Private Sub Cmdsalirmasivo_Click()
   Call dllgeneral.ActivaTab(1, 2, SSTab1)
End Sub

Private Sub Combo1_Click()
   MBox(8) = TraeDataSerie("select * from ct_tipocambio where tipocambiofecha=GETDATE()", VGCNx)
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
  Seguir Combo1, KeyAscii
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
  Seguir Combo2, KeyAscii
End Sub

Private Sub Combo3_Click()
  If Combo3.ListCount >= 0 Then
     Call CargarModo
     If Len(Trim(modoventa.almacenes)) > 0 Then
          Ctr_Ayuda3.Filtro = "taalma in (" & modoventa.almacenes & ")  and empresacodigo='" & VGParametros.empresacodigo & "'"
            'Ctr_Ayuda3.Ejecutarand puntovtacodigo='" & VGParametros.puntovta & "'
      End If
  End If
End Sub

Private Sub Combo3_GotFocus()
   If Combo3.ListCount - 1 <= 0 Then
      Call dllgeneral.llenacombo(Combo3, "select modovtacodigo,modovtadescripcion from vt_modoventa", VGCNx)
      Exit Sub
   End If

End Sub


Private Sub Combo3_KeyPress(KeyAscii As Integer)
  Call Combo3_Click
  Seguir Combo3, KeyAscii
End Sub

Private Sub Combo4_GotFocus()
   If Combo4.ListCount - 1 <= 0 Then
       Call dllgeneral.llenacombo(Combo4, "select formapagocodigo,formapagodescripcion from vt_formapago", VGCNx)
      Exit Sub
   End If

End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
   Seguir Combo4, KeyAscii
End Sub

Private Sub Combo5_KeyPress(KeyAscii As Integer)
    Seguir Combo5, KeyAscii
End Sub

Private Sub Combo6_KeyPress(KeyAscii As Integer)
    Seguir Combo6, KeyAscii
End Sub

Private Sub Combo7_KeyPress(KeyAscii As Integer)
    Seguir Combo7, KeyAscii
End Sub

Private Sub Combo8_KeyPress(KeyAscii As Integer)
    Seguir Combo8, KeyAscii
End Sub

Private Sub Command1_Click()
 Call grabamasivo
 Call dllgeneral.ActivaTab(1, 1, SSTab2)
End Sub

Private Sub Command2_Click()
 Call dllgeneral.ActivaTab(1, 1, SSTab2)
   
End Sub

Private Sub Command3_Click()
Limpiartexto MBox2, 12, 12
Limpiartexto MBox2, 13, 13
Limpiartexto MBox2, 14, 14
Limpiartexto MBox2, 0, 5
Text3.Text = Empty
End Sub

Private Sub cOpc_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
    cBoton(0).SetFocus
  End If
End Sub

Private Sub cSeleccion_Click(Index As Integer)
  Dim nArchi As String
  Dim rsel As New ADODB.Recordset
  Dim nl As Integer
  Dim J As Integer
  Dim nflag As Integer
  
  'On Error GoTo nerror
  
  VGCNx.Execute "delete from gtempfile"
  VGCNx.Execute "delete from tempfile"
  
  nflag = 0
  If Index = 0 Then
  
    If TDBpagos.ApproxCount = 0 Then
        MsgBox "Primero agregue el pago que se va a realizar.", vbInformation, "ZIYAZ"
        Exit Sub
    End If
  
    If cOpc2(0).Value And dllgeneral.VerificaDatoExistente(VGCNx, "select * from vt_cliente where clientecodigo='" & Ctr_Ayuda1.xclave & "' and len(ltrim(clienteruc))=11") = 0 Then
        MsgBox "El cliente no tiene ruc valido....Verifique!!!", vbInformation, MsgTitle
        Exit Sub
    Else
        Set rsel = VGCNx.Execute("select * from vt_cliente where clientecodigo='" & Ctr_Ayuda1.xclave & "' and clientecodigo<>'" & Left(g_Eventual, Len(Trim(Ctr_Ayuda1.xclave))) & "'")
        If rsel.RecordCount > 0 Then
           MBox3(0) = Escadena(rsel!clientecodigo)
           MBox3(1) = Escadena(rsel!clienterazonsocial)
           MBox3(2) = Escadena(rsel!clienteruc)
           MBox3(3) = Escadena(rsel!clientedireccion)
           MBox3(4) = Escadena(rsel!clientedistrito)
        End If
        rsel.Close
        Set rsel = Nothing
    End If
    cSeleccion(0).Enabled = False
    nflag = 1
    'VGCNx.BeginTrans
    If (cOpc2(0).Value Or cOpc2(1).Value Or cOpc2(2).Value Or cOpc2(3).Value) And (cOpc2(0).Enabled Or cOpc2(1).Enabled Or cOpc2(2).Enabled Or cOpc2(3).Enabled) Then
      If GrabarData() = 1 Then
         'VGCNx.CommitTrans
         rsdetax.UpdateBatch adAffectAllChapters
         rsdetax.Close
         nflag = 0
         If modoventa.emitefact = "1" Or modoventa.emiteguia = "1" Then
            nl = IIf(modoventa.copiasbol > 0, modoventa.copiasbol, 0)
            If nl <= 0 Then
               If modoventa.copiasfac > 0 Then
                    nl = modoventa.copiasfac
               ElseIf modoventa.copiastic > 0 Then
                    nl = modoventa.copiastic
               End If
               'nl = IIf(modoventa.copiasfac > 0, modoventa.copiasfac, 0)
               'nl = IIf(modoventa.copiastic > 0, modoventa.copiastic, 0)
               'nl = 1
            End If
            If nl > 0 Then
                For J = 1 To nl
                   'If MsgBox("Desea imprimir por Impresora Matricial " & Chr(13) & "o Ticketera ?" & Chr(13) & Chr(13) & "Matricial [SI]  /  Ticketera [NO]", vbYesNo + vbQuestion, "Impresion de documento") = vbYes Then
                    Call DocImprimir
                   'Else
                   ' MsgBox "Impresion por tikectera en construccion", vbOKOnly + vbInformation, "Sistemas"
                   'End If
                Next J
            End If
         End If
         Listado (0)
      Else
       '  VGCNx.RollbackTrans
         nflag = 0
      End If
    Else
      VGCNx.RollbackTrans
      nflag = 0
      cSeleccion(0).Enabled = True
      MsgBox "Seleccione un tipo de Documento para Grabar...!!!", vbInformation, MsgTitle
      Exit Sub
    End If
  End If
  cSeleccion(0).Enabled = True
  Fr4.Visible = False
  Activa 2
Set rsdetax = Nothing
nerror:
   If Err Then
      If nflag = 1 Then
         VGCNx.RollbackTrans
      End If
      MsgBox "Comunicarse con Sistemas ...!!" & Chr(13) & Chr(10) & Err.Number & "-" & Err.Description
      Exit Sub
      Resume
   End If
End Sub

Private Sub Ctr_Ayuda1_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    Dim acliente As New ADODB.Recordset
    
    MBox3(0) = Trim(ColecCampos.Item(0))
    MBox3(1) = Trim(ColecCampos.Item(1))
    MBox3(2) = Trim(ColecCampos.Item(2))
    MBox(19) = Trim(ColecCampos.Item(3))
    MBox3(3) = Trim(ColecCampos.Item(3))
    MBox3(4) = ESNULO(Trim(ColecCampos.Item(4)), "")
    Text5 = Trim(ColecCampos.Item(2))
    
    If IsNull(ColecCampos.Item(10)) Or Len(Trim(ColecCampos.Item(10))) = 0 Then
       text1 = numero(0)
       Text2 = numero(0)
    Else
       text1 = numero(CDbl(Trim(ColecCampos.Item(10))))
       Text2 = numero(CDbl(Trim(ColecCampos.Item(10))) * 100)
    End If
    
    lcred(0) = numero(0)
    lcred(1) = numero(0)

    Set acliente = VGCNx.Execute("select * from vt_cliente where clientecodigo='" & Ctr_Ayuda1.xclave & "'")
    If acliente.RecordCount > 0 Then
       Combo6.ListIndex = VerificaCombo(Combo6, acliente!clientetipopersona)
       Combo7.ListIndex = VerificaCombo(Combo7, acliente!clientetipopais)
       Combo8.ListIndex = VerificaCombo(Combo8, IIf(acliente!clientemultidireccion = 1, "S", "N"))
       lcred(0) = numero(acliente!clientesaldodolares)
       lcred(1) = numero(acliente!clientelimitecreddolar)
    End If
    acliente.Close
    Set acliente = Nothing

End Sub
Private Sub cargar()
SQL = " select top 0 * from vt_pagosencaja "
rsdetax.Open SQL, VGCNx, adOpenDynamic, adLockBatchOptimistic
Set TDBpagos.DataSource = rsdetax
TDBpagos.Refresh
'rsdetax.Close
End Sub

Private Sub Ctr_Ayuda3_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
almacentipo = ESNULO(ColecCampos("tipoalmacencodigo"), 0)
End Sub

Private Sub Ctr_AyudaTc_LostFocus()
If Ctr_AyudaTc.xclave <> "" Then
    If Ctr_AyudaTc.xclave = "09" Then
        Label1(30).Visible = True
        Ctr_AyudaPro.Visible = True
        Ctr_AyudaPro.xclave = ""
        Ctr_AyudaPro.xnombre = ""
        Ctr_AyudaPro.SetFocus
        'Ctr_AyudaPro.Requerido = True
    Else
        Label1(30).Visible = False
        Ctr_AyudaPro.Visible = False
        Ctr_AyudaPro.xclave = ""
        Ctr_AyudaPro.xnombre = ""
        Ctr_AyudaPro.Requerido = False
    End If
End If

End Sub

Private Sub DtFechaDesde_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub


Private Sub DtFechaHasta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub


Private Sub Image2_Click()
cOpc(0).Value = True
End Sub

Private Sub Image3_Click()
cOpc(1).Value = True
End Sub


Private Sub Image4_Click()
cOpc3(0).Value = True
End Sub

Private Sub Image5_Click()
cOpc3(1).Value = True
End Sub


Private Sub TDBpagos_KeyDown(KeyCode As Integer, Shift As Integer)
Dim nvalor As String
Dim nvalor2 As String

If KeyCode = 46 Then
   nvalor = TDBpagos.Columns(0).Text
   nvalor2 = TDBpagos.Columns(3).Text
   
    If rsdetax.RecordCount > 0 Then
        rsdetax.MoveFirst
        Do Until rsdetax.EOF
          If rsdetax.Fields(0) = nvalor And rsdetax.Fields(6) = nvalor2 Then
            rsdetax.Delete adAffectCurrent
            rsdetax.Update
            Exit Do
          End If
          rsdetax.MoveNext
        Loop
    End If
End If

End Sub


Private Sub TxFerimporte_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then If validar() Then adicionar
 
End Sub

Function validar()
validar = False
validar = True
End Function
Private Sub adicionar()
'SQL = " select top 0 * from vt_pagosencaja "
'rsdetax.Open SQL, VGCNx, adOpenDynamic, adLockBatchOptimistic

rsdetax.AddNew

rsdetax!empresacodigo = VGParametros.empresacodigo
rsdetax!pedidonumero = MBox(1)
rsdetax!pagocodigo = Ctr_Ayuoperacion.xclave
rsdetax!pagotipocodigo = Ctr_Ayutipo.xclave
rsdetax!pagonumdoc = TxFernumero.valor
rsdetax!monedacodigo = dllgeneral.ComboDato(TxFermoneda.Text)
rsdetax!cajerocodigo = VGParametros.cajerocodigo

If rsdetax!monedacodigo = "01" And TxFermoneda.valor = "02" Then
   rsdetax!pagoimporte = Format((TxFerimporte.valor * MBox(8)), "###,###,##0.00")
Else
  rsdetax!pagoimporte = Format(Trim(TxFerimporte.valor), "###,###,##0.00")
End If
'rsdetax!pago = rsdetax.RecordCount

'rsdetax.Close
End Sub


Private Sub Ctr_Ayuoperacion_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Ctr_Ayutipo.Filtro = "pagocodigo='" & Ctr_Ayuoperacion.xclave & "'"
Ctr_Ayutipo.xclave = ""
TxFernumero.valor = ""
If ColecCampos("pagoefectivo") = True Then
   Ctr_Ayutipo.Visible = False
   TxFernumero.Visible = False
Else
   Ctr_Ayutipo.Visible = True
   TxFernumero.Visible = True
End If
End Sub
Private Sub Ctr_Ayuda3_GotFocus()
    If Len(Trim(modoventa.almacenes)) > 0 Then
       Ctr_Ayuda3.Filtro = " taalma in (" & Trim(modoventa.almacenes) & ") and empresacodigo='" & VGParametros.empresacodigo & "'"
       Ctr_Ayuda3.Ejecutar
        'and puntovtacodigo='" & VGParametros.puntovta & "'
    Else
       Ctr_Ayuda3.Filtro = " almacencodigo like '%' and puntovtacodigo='" & VGParametros.puntovta & "' and empresacodigo='" & VGParametros.empresacodigo & "'"
       Ctr_Ayuda3.Ejecutar
    End If
   
End Sub


Private Sub Form_Activate()
Listado (0)
End Sub

Private Sub Form_Load()
Call configuramasivo
MostrarForm Me, "C"
Call Ctr_Ayuoperacion.conexion(VGCNx)
Call Ctr_Ayutipo.conexion(VGCNx)
Call Ctr_AyudaTc.conexion(VGCNx)
Call Ctr_AyudaPro.conexion(VGCNx)
Call Ctr_AyuRef.conexion(VGCNx)

Ctr_AyuRef.Filtro = "entidadorden='1'"

flag = 0
'Call dllgeneral.ActivaTab(0, 1, SSTab1)
Call dllgeneral.ActivaTab(0, 1, SSTab1)

DtFechaDesde.Value = DateAdd("m", -1, Date)
DtFechaHasta.Value = Format(Date, "dd/mm/yyyy")

nLongicampo(1) = 1000:  nLongicampo(2) = 1200:   nLongicampo(3) = 6300:   nLongicampo(4) = 600:  nLongicampo(5) = 1200

MBox(1).Enabled = False: Label2 = ""
Call Cargacombo
Listado (0)
Call dllgeneral.ActivaTab(0, 2, SSTab1)

Text3.Enabled = IIf(VGParametros.puntovta = "01", True, False)
Label1(21).Visible = IIf(VGParametros.puntovta <> "01", False, True)
Ctr_AyuRef.Requerido = IIf(VGParametros.puntovta <> "01", False, True)
Ctr_AyuRef.Visible = IIf(VGParametros.puntovta <> "01", False, True)

Text3.BackColor = IIf(VGParametros.puntovta = "01", vbWhite, &H8000000F)
'&H8000000F&

cSeleccion(0).Picture = MDIPrincipal.ImageList2.ListImages("Facturado").Picture
cSeleccion(1).Picture = MDIPrincipal.ImageList2.ListImages("Retornar").Picture
cmdBotones(11).Picture = MDIPrincipal.ImageList2.ListImages("Facturar").Picture
cmdBotones(12).Picture = MDIPrincipal.ImageList2.ListImages("Retornar").Picture

cmdBotones(0).Picture = MDIPrincipal.ImageList2.ListImages("Nuevo").Picture
cmdBotones(1).Picture = MDIPrincipal.ImageList2.ListImages("Modificar").Picture
cmdBotones(2).Picture = MDIPrincipal.ImageList2.ListImages("Eliminar").Picture
cmdBotones(4).Picture = MDIPrincipal.ImageList2.ListImages("Retornar").Picture

TxtHor.Text = Format(Time, "HH:mm")

Me.Top = 0
Me.Left = 0
End Sub
Private Function configuramasivo()
   Set rsmasivo = Nothing
   Call rsmasivo.Fields.Append("item", adVarChar, 5)
   Call rsmasivo.Fields.Append("Articulo", adVarChar, 20)
   Call rsmasivo.Fields.Append("descripcion", adVarChar, 30)
   Call rsmasivo.Fields.Append("unidad", adVarChar, 10)
   Call rsmasivo.Fields.Append("saldo", adVarChar, 30)
   Call rsmasivo.Fields.Append("Tieneigv", adVarChar, 1)
   Call rsmasivo.Fields.Append("cantidad", adVarChar, 20)
   rsmasivo.Open
   masivo = 0
   TDBGrid3.Columns(0).AllowFocus = False
   TDBGrid3.Columns(1).AllowFocus = False
   TDBGrid3.Columns(2).AllowFocus = False
   TDBGrid3.Columns(3).AllowFocus = False
   TDBGrid3.Columns(4).AllowFocus = False
   TDBGrid3.Columns(5).AllowFocus = True
   'TDBGrid3.AllowUpdate = True
   TDBGrid3.Columns(5).NumberFormat = "###,##0.0000"
   Set TDBGrid3.DataSource = rsmasivo
End Function
Private Function grabamasivo()
Dim rsk As New ADODB.Recordset
Dim J As Integer
   If rsmasivo.RecordCount > 0 And masivo = 1 Then
      rsmasivo.MoveFirst
      J = 0
      Do Until rsmasivo.EOF
         If rsmasivo!cantidad > 0 Then
            J = J + 1
            rsdeta.AddNew
            rsdeta.Fields(0) = J
            rsdeta.Fields(1) = rsmasivo!articulo
            rsdeta.Fields(2) = rsmasivo!descripcion
            rsdeta.Fields(3) = rsmasivo!unidad
            If Trim(rsmasivo!cantidad) = "" Then
               rsmasivo!cantidad = 0
            End If
            rsdeta.Fields(4) = rsmasivo!cantidad
            MBox2(0) = rsmasivo!cantidad
            MBox2(2) = rsmasivo!unidad
            MBox2(1) = rsmasivo!articulo
            MBox2(4) = 0
            MBox2(5) = 0
            If Text7.Text > 0 Then
                rsdeta.Fields(12) = Text7.Text
                MBox2(3) = rsdeta.Fields(12)
              Else
                Set rsk = VGCNx.Execute("select * from listapre" & Trim(Combo2.Text) & " where productocodigo='" & rsmasivo!articulo & "' and almacencodigo='" & Trim(Ctr_Ayuda3.xclave) & "'")
                If rsk.RecordCount > 0 Then
                    rsdeta.Fields(12) = rsk.Fields("productoprecvta")
                    MBox2(3) = rsdeta.Fields(12)
                Else
                    rsdeta.Fields(12) = 0
                   MBox2(3) = 0
               End If
            End If
           If VGParamSistem.tieneigv = "1" Then
              rsdeta.Fields(5) = (MBox2(3) / (1 + VGParamSistem.Igv))
'             rsdeta.Fields(12) = MBox2(3).Tag
            Else
              If modoventa.impuestos = "1" Then
                 If rsmasivo.Fields(5) = 1 Then
                    rsdeta.Fields(5) = (MBox2(3) / (1 + VGParamSistem.Igv))
                   Else
                    rsdeta.Fields(5) = MBox2(3).Text
                  End If
          '       rsdeta.Fields(12) = MBox2(3).Tag
               Else
                 rsdeta.Fields(5) = MBox2(3).Text
           '      rsdeta.Fields(12) = MBox2(3).Tag
              End If
           End If
           rsdeta.Fields(6) = numero(MBox2(4))
           rsdeta.Fields(7) = numero(MBox2(0) * MBox2(3))   ' IIf(VGParamSistem.tieneigv = "1", (MBox2(3) / (1 + (VGParamSistem.igv / 100))), MBox2(3)))
           rsdeta.Fields(8) = rsmasivo!cantidad
           rsdeta.Fields(9) = IIf(Len(Trim(MBox2(12))) = 0, 0, Format(MBox2(12), "##,###,##0"))
           rsdeta.Fields(10) = numero(MBox2(13))
           rsdeta.Fields(11) = IIf(IsNull(MBox2(14)) Or Len(Trim(MBox2(14))) = 0, 0, MBox2(14))
           Set rsk = Nothing
           rsdeta.Update
         End If
         rsmasivo.MoveNext
      Loop
      If rsmasivo.RecordCount > 0 Then
         rsmasivo.MoveFirst
         Do Until rsmasivo.EOF
            rsmasivo.Delete
            rsmasivo.MoveNext
         Loop
      End If
      masivo = 2
   End If
   Totales
End Function
Private Function loadmasivo()
   Dim rsgrid3 As New ADODB.Recordset
   Dim SQL As String
   Dim wposi As Integer
  ' Define the Style that will be used for items that are 0
  Dim NoItem As New TrueOleDBGrid70.Style
   
   SQL = " select articulo=stcodigo,descripcion=adescri,unidad=aunidad,"
   If VGParamSistem.stockcomp = 1 Then
     SQL = SQL & "saldo=stskdis-stskcom, cantidad=0 "
   Else
      SQL = SQL & "saldo=stskdis, cantidad=0 "
   End If
   SQL = SQL & " ,tieneigv from stkart,maeart,grupo where acodigo=stcodigo and stalma='" & Trim(Ctr_Ayuda3.xclave) & "' and alinea = '001'"
   SQL = SQL & " and afamilia=fam_codigo and alinea=lin_codigo and agrupo=gru_codigo order by agrupo,acodigo"
   Set rsgrid3 = VGCNx.Execute(SQL)
    
  TDBGrid3.FetchRowStyle = True
  NoItem.BackColor = vbYellow
 
  TDBGrid3.Columns(6).AddRegexCellStyle dbgNormalCell, NoItem, "^0"
  TDBGrid3.Columns(6).AddRegexCellStyle dbgNormalCell + dbgCurrentCell, NoItem, "^0"
   Text7.Text = 0
   wposi = 0
   Text10 = 0
   If rsgrid3.RecordCount > 0 And masivo = 0 Then
      rsgrid3.MoveFirst
      Do Until rsgrid3.EOF
         rsmasivo.AddNew
         wposi = wposi + 1
         rsmasivo.Fields(0) = wposi
         rsmasivo.Fields(1) = rsgrid3!articulo
         rsmasivo.Fields(2) = Left(rsgrid3!descripcion, 30)
         rsmasivo.Fields(3) = rsgrid3!unidad
         rsmasivo.Fields(4) = rsgrid3!saldo
         rsmasivo.Fields(5) = rsgrid3!tieneigv
         rsmasivo.Fields(6) = rsgrid3!cantidad
         rsgrid3.MoveNext
     Loop
  End If
    rsmasivo.MoveFirst
  TDBGrid3.SetFocus
  TDBGrid3.Refresh
  rsgrid3.Close
  Set rsgrid3 = Nothing
  masivo = 1

End Function

Public Function Cargacombo()
   Dim J As Integer
   Dim nsql As String
   Dim rsk As New ADODB.Recordset
   
   CargaGrilla
   MBox2(11) = rsdeta.RecordCount
   If MBox2(11) > modoventa.nroitem Then
      MsgBox "No se puede Ingresar mas Items...Verifique!!!", vbInformation, MsgTitle
      Exit Function
   End If
  
  
   Call dllgeneral.llenacombo(Combo1, "select monedacodigo,monedadescripcion from gr_moneda order by monedacodigo", VGCNx)
   'If Combo1.ListCount - 1 >= 0 Then
       Combo1.ListIndex = 0
   'End If
   
   Combo2.Clear
   Set rsk = VGCNx.Execute("select right(name,1) from sysobjects where name like 'listapre%'")
   If rsk.RecordCount > 0 Then
      rsk.MoveFirst
      Do Until rsk.EOF
        Combo2.AddItem rsk.Fields(0)
        rsk.MoveNext
      Loop
   Else
     Combo2.AddItem "*ninguno"
   End If
   rsk.Close
   Set rsk = Nothing
'     Combo2.AddItem Trim(Str(J))
'   Next J
   Combo2.ListIndex = 0
   
   Call dllgeneral.llenacombo(Combo3, "select modovtacodigo,modovtadescripcion from vt_modoventa where puntovtacodigo='" & VGParametros.puntovta & "'", VGCNx)
   If Combo3.ListCount - 1 >= 0 Then
     Combo3.ListIndex = 0
   End If
   
   Call dllgeneral.llenacombo(Combo4, "select formapagocodigo,formapagodescripcion from vt_formapago", VGCNx)
   If Combo4.ListCount - 1 >= 0 Then
       Combo4.ListIndex = 0
   End If
   
   
   Call CargarTipo(Combo5, 3)
   
   Call CargarTipo(Combo6, 4)
   
   Call CargarTipo(Combo7, 5)
   
   Call CargarTipo(Combo8, 3)
   
   
   Call Ctr_Ayuda1.conexion(VGCNx)
   Call Ctr_Ayuda2.conexion(VGCNx)
   Call Ctr_Ayuda3.conexion(VGCNx)
   Call Ctr_AyuTransporte.conexion(VGCNx)
   
End Function

Public Function CargaGrilla()

   Call rsdeta.Fields.Append("Item", adInteger)
   Call rsdeta.Fields.Append("Codigo", adVarChar, 20)
   Call rsdeta.Fields.Append("Descripcion", adChar, 100)
   Call rsdeta.Fields.Append("UM", adChar, 3)
   Call rsdeta.Fields.Append("Cant", adDouble)
   Call rsdeta.Fields.Append("Precio_Vta", adDouble)
   Call rsdeta.Fields.Append("Dscto(%)", adDouble)
   Call rsdeta.Fields.Append("Total", adDouble)
   Call rsdeta.Fields.Append("%", adDouble)
   Call rsdeta.Fields.Append("CantRef", adDouble)
   Call rsdeta.Fields.Append("Factor", adDouble)
   Call rsdeta.Fields.Append("%P", adDouble)
   Call rsdeta.Fields.Append("PrecioLista", adDouble)
   Call rsdeta.Fields.Append("tipo", adChar, 1)
   Call rsdeta.Fields.Append("Glosa", adVarChar, 500)
   
   rsdeta.Open
   If rsdeta.RecordCount > 0 Then
     Totales
   End If
   ConfigGrid

End Function

Public Function ConfigGrid()
   Set TDBGrid1.DataSource = Nothing
   
   Set TDBGrid1.DataSource = rsdeta
   With TDBGrid1
      .Columns(0).Width = 600
      .Columns(0).Caption = "Item"
      .Columns(1).Width = 900
      .Columns(1).Caption = "Codigo"
      .Columns(2).Width = 4700
      .Columns(2).Caption = "Descripcion"
      .Columns(3).Width = 600
      .Columns(3).Caption = "UM"
      .Columns(4).Width = 1000
      .Columns(4).Caption = "Cant"
      .Columns(5).Width = 1300
      .Columns(5).Caption = "Precio_Vta"
      .Columns(6).Width = 1000
      .Columns(6).Caption = "Dscto(%)"
      .Columns(7).Width = 1300
      .Columns(7).Caption = "Total"
      .Columns(8).Width = 0
      .Columns(8).Caption = "%"
      .Columns(5).NumberFormat = "###,##0.0000"
      .Columns(6).NumberFormat = "###,##0.0000"
      .Columns(7).NumberFormat = "###,##0.0000"
      .Columns(8).NumberFormat = "###,##0.0000"
      .Columns(9).Width = 0
      .Columns(9).Caption = "Cant.Ref"
      .Columns(9).NumberFormat = "###,##0"
      .Columns(10).Width = 0
      .Columns(10).Caption = "Factor"
      .Columns(10).NumberFormat = "###,##0.0000"
      .Columns(11).Width = 0
      .Columns(11).NumberFormat = "###,##0.0000"
      .Columns(12).Visible = True
      .Columns(12).Width = 0
      .Columns(13).Width = 0
      .Columns(14).Width = 0
      
      .Columns(8).Visible = False
      .Columns(9).Visible = False
      .Columns(10).Visible = False
      .Columns(11).Visible = False
      .Columns(12).Visible = False
      .Columns(13).Visible = False
      .Columns(14).Visible = False
      
   End With
   TDBGrid1.Refresh
End Function
Public Function formapagos()
   Set rspagos = Nothing
   Call rspagos.Fields.Append("tipo", adVarChar, 5)
   Call rspagos.Fields.Append("Articulo", adVarChar, 20)
   Call rspagos.Fields.Append("descripcion", adVarChar, 30)
   Call rspagos.Fields.Append("unidad", adVarChar, 10)
   Call rsmasivo.Fields.Append("saldo", adVarChar, 30)
   Call rsmasivo.Fields.Append("Tieneigv", adVarChar, 1)
   Call rsmasivo.Fields.Append("cantidad", adVarChar, 20)
   rsmasivo.Open
   masivo = 0
   TDBGrid3.Columns(0).AllowFocus = False
   TDBGrid3.Columns(1).AllowFocus = False
   TDBGrid3.Columns(2).AllowFocus = False
   TDBGrid3.Columns(3).AllowFocus = False
   TDBGrid3.Columns(4).AllowFocus = False
   TDBGrid3.Columns(5).AllowFocus = True
   'TDBGrid3.AllowUpdate = True
   TDBGrid3.Columns(5).NumberFormat = "###,##0.0000"
   Set TDBGrid3.DataSource = rsmasivo

End Function
Public Function Listado(Fecha As Integer)
Dim RsListado As ADODB.Recordset
Dim SQL As String



SQL = "select pedidonumero as Pedido,pedidofecha as Fecha,pedidonotaped as Cotizacion,clienterazonsocial as [Cliente/Razon Social],pedidototneto as Total "
SQL = SQL & " from vt_tempopedido" & VGParametros.puntovta & " where empresacodigo='" & VGParametros.empresacodigo & "' and puntovtacodigo='" & VGParametros.puntovta & "' "

If Fecha = 1 Then
  If DtFechaDesde.Value > DtFechaHasta.Value Then
      MsgBox "La fecha de inicio no puede ser mayor " & Chr(13) & "a la fecha final.", vbInformation, "Sistemas"
      DtFechaDesde.SetFocus
      Exit Function
  End If
  SQL = SQL & " and pedidofecha between '" & DtFechaDesde.Value & "' and '" & DtFechaHasta.Value & "' "
End If

If Len(Trim(TxtNro.Text)) <> 0 Then SQL = " and pedidonumero='" & TxtNro.Text & "'"
If Len(Trim(TxtCliente.Text)) <> 0 Then SQL = SQL + " and clienterazonsocial like '%" & TxtCliente.Text & "%' "

SQL = SQL & " order by pedidofecha,pedidonumero"

Set RsListado = VGCNx.Execute(SQL)

LblReg.Caption = "(" & RsListado.RecordCount & ") Pedidos"
TDBGrid2.DataSource = RsListado


'Call dllgeneral.ListarEnTDBGRID(VGCNx, g_PedidoPuntoVta, TDBGrid2, "pedidonumero as Pedido,pedidofecha as Fecha,pedidonotaped as Cotizacion,clienterazonsocial as Descripcion,pedidototneto as total", "pedidofecha,pedidonumero", nLongicampo)
'TReg.Text = Format(TDBGrid2.ApproxCount, "#########0")
With TDBGrid2
  .Columns(0).Width = 1200
  .Columns(1).Width = 1200
  .Columns(2).Width = 1200
  .Columns(3).Width = 5200
  .Columns(4).Width = 2000
  .AllowUpdate = False
  .Refresh
End With

End Function

Private Sub Form_Unload(Cancel As Integer)
  Set rsdeta = Nothing
End Sub

Private Sub MBox_GotFocus(Index As Integer)
Call dllgeneral.Enfoquetexto(MBox(Index))
End Sub


Private Sub MBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 And Index >= 5 And Index < 19 Then
    If Index = 9 Then
      SSTab2.Tab = 1
      Combo3.SetFocus
    Else
      If Index Like "[567]" Then
         Totales
      End If
      SendKeys "{tab}"
    End If
  ElseIf KeyCode = 13 And (Index = 19 And Len(Trim(MBox(19))) > 0) Then
        MBox(19) = Escadena(UCase(Trim(MBox(19).ClipText)))
        If IsNull(MBox(19)) Or Len(Trim(MBox(19))) = 0 Then
           MsgBox "Falta Punto de LLegada", vbInformation, MsgTitle
           Exit Sub
        End If
        If IsNull(Ctr_Ayuda1.xclave) Or Len(Trim(Ctr_Ayuda1.xclave)) = 0 Then
           MsgBox W1TXT1, vbInformation, MsgTitle
           SSTab2.Tab = 1
           Ctr_Ayuda1.SetFocus
           Exit Sub
        End If
'        If IsNull(Ctr_Ayuda2.xclave) Or Len(Trim(Ctr_Ayuda2.xclave)) = 0 Then
'           MsgBox W1TXT6, vbInformation, MsgTitle
'           SSTab2.Tab = 0
'           Ctr_Ayuda2.SetFocus
'           Exit Sub
'        End If
        If IsNull(Ctr_Ayuda3.xclave) Or Len(Trim(Ctr_Ayuda3.xclave)) = 0 Then
           MsgBox W1TXT7, vbInformation, MsgTitle
           SSTab2.Tab = 1
           Ctr_Ayuda3.SetFocus
           Exit Sub
        End If
        If IsNull(MBox(8)) Or Len(Trim(MBox(8))) = 0 Or CDbl(MBox(8)) <= 0 Then
           MsgBox "Falta Tipo de Cambio", vbInformation, MsgTitle
           SSTab2.Tab = 0
           Call dllgeneral.Enfoquetexto(MBox(8))
           Exit Sub
        End If
        If IsNull(MBox(15)) Or Len(Trim(MBox(15))) = 0 Or CDbl(MBox(15)) < 0 Then
           MsgBox W1TXT9, vbInformation, MsgTitle
           SSTab2.Tab = 1
           Exit Sub
        End If
        If dllgeneral.VerificaDatoExistente(VGCNx, "select * from vt_cliente where clientecodigo='" & MBox3(0) & "' and clientesuspendido='1'") = 1 And MBox3(0) <> g_Eventual Then
           MsgBox W1TXT3, vbInformation, MsgTitle
           Exit Sub
        End If
'        If dllgeneral.VerificaDatoExistente(VGCNx, "select * from vt_cliente where clientecodigo='" & MBox3(0) & "' and ((clientelimitecreddolar-clientesaldodolares)*" & MBox(8) & "+ (clientelimitecredsoles-clientesaldosoles)) <=0") = 1 And MBox3(0) <> g_Eventual Then
'           MsgBox W1TXT4, vbInformation, MsgTitle
'           Exit Sub
'        End If

'        Fr1(0).Enabled = False
'        Fr2(0).Enabled = False
'        Fr3(0).Enabled = False
        TClie.Enabled = False
        Call CargarModo
        
        Ctr_Ayuda2.SetFocus
'        If Text3.Visible = True Then
'           Text3.SetFocus ' "{tab}"
'         Else
'           Text4(0).SetFocus
'        End If
  End If
End Sub


Private Sub MBox_LostFocus(Index As Integer)
'  On Error Resume Next
'  Select Case Index
'   Case 5, 6, 7, 8, 13, 15
'      If Not dllgeneral.ValidaCadena(MBox(Index), "N") Then
'         MsgBox Msg29, vbInformation, "AVISO"
'         Call dllgeneral.Enfoquetexto(MBox(Index))
'         Exit Sub
'      End If
'      MBox(Index) = Format(MBox(Index), "##,##0.0000")
'   Case 10
'      If Not dllgeneral.ValidaCadena(MBox(Index), "F") Then
'         MsgBox "Fecha No Valida", vbInformation, "AVISO"
'         Call dllgeneral.Enfoquetexto(MBox(Index))
'         Exit Sub
'      End If
'     If MBox(10).Text <> VGParamSistem.FechaTrabajo Then
'        Chkentrega.Value = 1
'      Else
'        Chkentrega.Value = 0
'     End If
'   Case 16
'      If Not dllgeneral.ValidaCadena(MBox(Index), "D") Then
'         MsgBox Msg29, vbInformation, "AVISO"
'         Call dllgeneral.Enfoquetexto(MBox(Index))
'         Exit Sub
'      End If
'      MBox(Index) = Right("000000000000" & MBox(Index), MBox(Index).MaxLength)
'   Case 19
'      MBox(19) = Escadena(UCase(Trim(MBox(19).ClipText)))
'   Case 18
'      If Not dllgeneral.ValidaCadena(MBox(Index), "D") Then
'         MsgBox Msg29, vbInformation, "AVISO"
'         Call dllgeneral.Enfoquetexto(MBox(Index))
'         Exit Sub
'      End If
'      MBox(Index) = Format(MBox(Index), "####0")
'      Exit Sub
'   Case 9
'      Call MBox_KeyDown(9, 13, 0)
'      Exit Sub
'
'   Case 2, 3, 4
'        MBox(Index) = Right("000000000000" & MBox(Index), MBox(Index).MaxLength)
'  End Select
End Sub
Private Sub MBox2_GotFocus(Index As Integer)
  On Error Resume Next
  If Ctr_Ayuda1.xclave = "" Then
     Ctr_Ayuda1.SetFocus
     Exit Sub
  End If
  If Ctr_Ayuda2.xclave = "" Then
     Ctr_Ayuda2.SetFocus
     Exit Sub
  End If
  If Ctr_Ayuda3.xclave = "" Then
     Ctr_Ayuda3.SetFocus
     Exit Sub
  End If
  If Index = 3 Then
     Call TraerProducto
  End If
   If Index Like "[234]" Then
        Fr1(0).Enabled = False
        Fr2(0).Enabled = False
        Fr3(0).Enabled = False
        TClie.Enabled = False
   End If
  Call dllgeneral.Enfoquetexto(MBox2(Index))
End Sub

Private Sub MBox2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  On Error Resume Next
  If KeyCode = 13 Then
    If Index = 12 Then
      MBox2(Index) = Format(MBox2(Index), "##,###,##0")
    ElseIf Index = 1 Then
      Call TraerProducto
   End If
    SendKeys "{tab}"
  ElseIf Index = 1 Then
      If dllgeneral.ValidaCadena(Trim(MBox2(1).ClipText), "N") = False Then
        MBox2(1).MaxLength = 64
      Else
        MBox2(1).MaxLength = 8
      End If
  End If
End Sub
Private Sub MBox2_LostFocus(Index As Integer)
Dim wflag As Integer
On Error Resume Next
  
If Ctr_Ayuda1.xclave = "" Then
   Ctr_Ayuda1.SetFocus
   Exit Sub
End If

If Ctr_Ayuda2.xclave = "" Then
   Ctr_Ayuda2.SetFocus
   Exit Sub
End If

If Ctr_Ayuda3.xclave = "" Then
   Ctr_Ayuda3.SetFocus
   Exit Sub
End If
  
  
'  Select Case Index
'   Case 0
'      If Not (dllgeneral.ValidaCadena(MBox2(Index), "N") Or IsNumeric(MBox2(Index))) Then
'         MsgBox Msg29, vbInformation, "AVISO"
'         Call dllgeneral.Enfoquetexto(MBox2(Index))
'         Exit Sub
'      End If
'   Case 1
'
'      If dllgeneral.VerificaDatoExistente(VGCNx, "select * from stkart where stcodigo='" & MBox2(Index).Text & "' and stalma='" & Ctr_Ayuda3.xclave & "'") = 0 And Len(Trim(MBox2(Index))) > 0 Then
'          Call cAyuda_Click(3)
'          MBox2(1).MaxLength = 20
'         Exit Sub
'      Else
'        wflag = verificaproducto()
'        If wflag = 1 Then
'            Label2 = ""
'            MsgBox "Ya ingreso el producto...Verifique!!!", vbInformation, MsgTitle
'            MBox2(1).SetFocus
'            Exit Sub
'         End If
'
'      End If
'   Case 3, 4, 5
'      If Index = 3 And dllgeneral.ComboDato(Combo5.Text) = "N" Then
'          Call TraerProducto
'      End If
'      If Not dllgeneral.ValidaCadena(MBox2(Index), "N") And Len(Trim(MBox2(Index))) <> 0 Then
'         MsgBox Msg29, vbInformation, "AVISO"
'         Call dllgeneral.Enfoquetexto(MBox2(Index))
'         Exit Sub
'      End If
'      If Not (dllgeneral.ValidaCadena(MBox2(0), "N") Or IsNumeric(MBox2(0))) Then
'         MsgBox Msg29, vbInformation, "AVISO"
'         Call dllgeneral.Enfoquetexto(MBox2(0))
'         Exit Sub
'      End If
'
'      If Index Like "[45]" Then
'         MBox2(Index) = Format(MBox2(Index), "######0.00000")  ' Numero(MBox2(Index))
'       Else
'         MBox2(Index) = Format(MBox2(Index), "######0.00000")
'       End If
'  End Select

End Sub

Private Sub MBox3_KeyPress(Index As Integer, KeyAscii As Integer)
   Seguir MBox3(Index), KeyAscii
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
  If SSTab1.Tab = 2 And Chkmasivo = 0 Then
     MBox2(0).SetFocus
  ElseIf SSTab1.Tab = 1 And Chkmasivo = 0 Then
     If MBox(0).Enabled = True Then
        Combo3.SetFocus
        'MBox(5).SetFocus
     Else
        Combo3.SetFocus
        'MBox(5).SetFocus
     End If
  End If
End Sub

Public Function Totales()
Dim J As Double
Dim Previo As Double
Dim rssql As New ADODB.Recordset
Dim SQL As String
Dim dct01, dct02, dct03, dct04, dct05, dct06 As Double
Dim Servicio  As Boolean
Dim RsSer As New ADODB.Recordset
  
Tbruto = 0: Tigv = 0: Tdscto = 0: TNeto = 0: TCant = 0
TImporte = 0: TSub = 0
'--Totales de Descuentos
DTGlobal = 0: DTCliente = 0: DTPPago = 0: DTOficina = 0: DTItem = 0
DTLinea = 0: DTPromo = 0: MBox2(6) = 0

   
  If rsdeta.RecordCount > 0 Then
    rsdeta.MoveFirst
    For J = 0 To rsdeta.RecordCount - 1
        If rsdeta.RecordCount > 0 Then
            Set RsSer = VGCNx.Execute("select afstock from maeart where acodigo='" & rsdeta.Fields(1) & "'")
            If RsSer.RecordCount > 0 Then
'                If RsSer.Fields("afstock") = 0 Then
'                    Servicio = True
'                Else
'                    Servicio = False
'                End If
            End If
        End If

       'IMPORTE DE MONTO BRUTO SIN IGV, ES DECIR PRECIO X CANTIDAD
    
'       If Servicio = False Then
 '           Tbruto = (Tbruto) + (rsdeta.Fields(7)) ' - MBox(7)) - ((rsdeta.Fields(7) - MBox(7)) * MBox(5) / 100)) / (1 + VGParamSistem.Igv/100)) '(rsdeta.Fields(7) / (1 + VGParamSistem.Igv/100))
'       Else
'            Tbruto = Tbruto + (rsdeta.Fields(7) / (1 + VGParamSistem.Igv / 100))
'       End If
       
       TCant = TCant + rsdeta.Fields(4)                                                      'AKI ESTABA /100
       TImporte = rsdeta.Fields(4) * IIf(VGParamSistem.tieneigv, rsdeta.Fields(5) / (1 + VGParamSistem.Igv), rsdeta.Fields(5))    '(rsdeta.Fields(7) + rsdeta.Fields(7))  'rsdeta.Fields(4) *
       '/ 100
       Tbruto = (Tbruto) + TImporte
       If IsNull(text1) Or Len(Trim(text1)) = 0 Then
           dct06 = 0
       Else
           dct06 = 0
       End If
       
       If Servicio = False Then
            dct01 = 0    ' descuento por cliente
            DTCliente = DTCliente + dct01
            
            'DESCUENTO POR ITEM
            dct02 = 0
            dct02 = (TImporte * (rsdeta.Fields(6) / 100))
            
            DTItem = DTItem + dct02
            
            'DESCUENTO ESPECIAL  :w8dct03 =(w8bruto - w8dct02-w8dct06)*w2dctpp/100
             'lo k estaba
             'dct03 = (TImporte - dct02 - dct06) * (MBox(7) / 100)
             dct03 = MBox(7)
             
            DTPPago = DTPPago + dct03
             
            'DESCUENTO POR PROMOCION  : w8dct04 =(w8bruto - w8dct02-w8dct03-w8dct06)*w2dctpr/100
            dct04 = (TImporte - dct02 - dct03 - dct06) * (MBox(6) / 100)
            DTPromo = DTPromo + dct04
             
            'DESCUENTO GENERAL : w8dct05 =(w8bruto - w8dct02-w8dct03-w8dct04-w8dct06)*w2dctgl/100
            dct05 = (TImporte - dct02 - dct03 - dct04 - dct06) * (MBox(5) / 100)
                     
            DTGlobal = DTGlobal + dct05
            
            'ACUMULADO DE TOTAL DESCUENTOS  :w8dctos = w8dct02 + w8dct03+w8dct04+w8dct05+w8dct06
             Tdscto = Tdscto + (dct01 + dct02 + dct03 + dct04 + dct05 + dct06)
        End If
    
       'ACUMULADO DE SUBTOTAL DE VENTA : w8subto = w8bruto - w8dctos
        TSub = TSub + (TImporte - (dct01 + dct02 + dct03 + dct04 + dct05 + dct06))
                
       If VGParamSistem.tieneigv = "1" Then
             Previo = (TImporte - (dct01 + dct02 + dct03 + dct04 + dct05 + dct06))
            Previo = Previo * VGParamSistem.Igv '/ 100
            Tigv = Tigv + Previo
       Else
           If modoventa.impuestos = "1" Then
                SQL = " select tieneigv from  grupo,maeart where acodigo='" & rsdeta.Fields(1) & "'"
                SQL = SQL & " and afamilia=fam_codigo and alinea=lin_codigo and agrupo=gru_codigo "
                Set rssql = VGCNx.Execute(SQL)
                If rssql.RecordCount > 0 Then
                   If rssql!tieneigv = "1" Then
                      Previo = (TImporte - (dct01 + dct02 + dct03 + dct04 + dct05 + dct06))
                      Previo = (Previo * VGParamSistem.Igv)
                      Tigv = Tigv + Previo
                      Previo = (TImporte - (dct01 + dct02 + dct03 + dct04 + dct05 + dct06))
                      Previo = (Previo * (1 + VGParamSistem.Igv / 100))
                      rsdeta.Fields(7) = Previo
                    Else
                      Previo = (TImporte - (dct01 + dct02 + dct03 + dct04 + dct05 + dct06))
                      Previo = (Previo * 0)
                      Tigv = Tigv + Previo
                      Previo = (TImporte - (dct01 + dct02 + dct03 + dct04 + dct05 + dct06))
                      Previo = (Previo * (1 + 0))
                      rsdeta.Fields(7) = Previo
                              
                   End If
                 Else
                   Previo = (TImporte - (dct01 + dct02 + dct03 + dct04 + dct05 + dct06))
                   Previo = (Previo * 0)
                   Tigv = Tigv + Previo
                
                   Previo = (TImporte - (dct01 + dct02 + dct03 + dct04 + dct05 + dct06))
                   Previo = (Previo * (1 + 0))
                   rsdeta.Fields(7) = Previo
               End If
           Else
               If rsdeta.Fields(11) > 0 Then
                    Previo = (TImporte - (dct01 + dct02 + dct03 + dct04 + dct05 + dct06))
                    rsdeta.Fields(7) = Previo * (1 + rsdeta(11))
                    Tigv = Tigv + (Previo * rsdeta(11))
               Else
                    Previo = (TImporte - (dct01 + dct02 + dct03 + dct04 + dct05 + dct06))
                    rsdeta.Fields(7) = Previo
                    'Tigv = Tigv
              End If
           End If
        End If
        rsdeta.Update
      
       rsdeta.MoveNext
    Next J
  Else
    Exit Function
  End If
  
 'IMPORTE TOTAL NETO DE FACTURA   w8tneto = w8subto + w8impto
  TNeto = Tbruto - Tdscto + Tigv
  MBox2(7) = Format(Tbruto, "#,###,##0.0000")
  MBox2(6) = numero(TCant)
  MBox2(9) = numero(Tigv)
  MBox2(8) = numero(Tdscto)
  MBox2(10) = numero(TNeto)
  
  Limpiartexto MBox2, 12, 12
  Limpiartexto MBox2, 13, 13
  Limpiartexto MBox2, 14, 14
  Limpiartexto MBox2, 0, 5
  'Text3.Text = Empty
End Function

Private Sub tclie_Click()
       
   SSTab2.TabEnabled(2) = IIf(TClie.Value = 1, 1, 0)
   If TClie.Value = 1 Then
        SSTab2.Tab = 2
        MBox3(0) = g_Eventual
        MBox3(0).Enabled = False
        MBox3(1).SetFocus
   End If
End Sub

Private Sub TDBGrid1_Click()
If rsdeta.RecordCount > 0 Then TDBGrid1.SetFocus
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim nvalor As String
On Error Resume Next
  
If KeyCode = 46 Then
   If rsdeta.RecordCount <= 0 Then
      Limpiartexto MBox2, 6, 10
      Exit Sub
   End If
   nvalor = TDBGrid1.Columns(0).Text
   If rsdeta.RecordCount > 0 Then
      rsdeta.MoveFirst
      Do Until rsdeta.EOF
        If rsdeta.Fields(0) = nvalor Then
          rsdeta.Delete adAffectCurrent
          rsdeta.Update
          Exit Do
        End If
        rsdeta.MoveNext
      Loop
   End If
   ConfigGrid
   Totales
   Exit Sub
ElseIf KeyCode = 13 Then
  Limpiartexto MBox2, 0, 5
  MBox2(11) = TDBGrid1.Columns(0).Text
  MBox2(0) = TDBGrid1.Columns(4).Text
  MBox2(1) = TDBGrid1.Columns(1).Text
  Label2 = TDBGrid1.Columns(2).Text
  MBox2(2) = Escadena(TDBGrid1.Columns(3).Text)
  MBox2(12) = Escadena(TDBGrid1.Columns(9).Text)
 Text3.Text = Escadena(TDBGrid1.Columns(14).Text)
 
  If VGParamSistem.tieneigv = "1" Then
       MBox2(3) = Format(TDBGrid1.Columns(5).Text, "######0.0000") '* (1 + (VGParamSistem.Igv)), "######0.0000")
  Else
     If modoventa.impuestos = "1" Then
         MBox2(3) = Format(IIf(IsNull(TDBGrid1.Columns(5).Text) Or Len(Trim(TDBGrid1.Columns(5).Text)) = 0, 0, TDBGrid1.Columns(5).Text) * (1 + (VGParamSistem.Igv)), "######0.0000")
     Else
         MBox2(3) = Format(TDBGrid1.Columns(5).Text, "######0.0000")
     End If
  End If
  MBox2(4) = numero(TDBGrid1.Columns(6).Text)
  MBox2(5) = numero(TDBGrid1.Columns(8).Text)
  If MBox2(12).Enabled = True Then
    MBox2(12).SetFocus
  Else
    MBox2(0).SetFocus
  End If
  flag = 1
End If
  
End Sub

Public Function Carga_Pedido()
    Dim csql As New ADODB.Recordset
    Dim acliente As New ADODB.Recordset
    Dim J As Integer
    Set csql = VGCNx.Execute("select * from " & g_PedidoPuntoVta & " where pedidonumero='" & TDBGrid2.Columns(0).Text & "'")
    If csql.RecordCount > 0 Then
       MBox(0) = Escadena(csql!puntovtacodigo)                    'Pto Venta
       MBox(1) = Escadena(csql!pedidonumero)                      'nro pedido
       If Escadena(csql!pedidotipofac) = g_tipofac Then
         MBox(2) = Escadena(csql!pedidonrofact)                     'nro factura
       Else
         MBox(2) = 0
       End If
       If Escadena(csql!pedidotipofac) = g_tipobol Then
          MBox(3) = Escadena(csql!pedidonrofact)                   'nro boleta
       Else
          MBox(3) = 0
       End If
       If Escadena(csql!pedidotipofac) = g_tipoguia Then
            MBox(4) = Escadena(csql!pedidonrofact)                   'nro guia
       Else
            MBox(4) = 0
       End If
       MBox(5) = numero(csql!pedidodsctoglobal)                   'dscto gral
       MBox(6) = numero(csql!pedidodsctoppago)                    'dscto promocional
       MBox(7) = numero(csql!pedidodsctovtaoficina)               'dscto especial
       Combo1.ListIndex = Escadena(VerificaCombo(Combo1, csql!pedidomoneda))    'moneda
       MBox(8) = numero(csql!pedidotipcambio)                             'tipo de cambio
       Combo2.ListIndex = VerificaCombo(Combo2, Trim(csql!pedidolistaprec)) 'lista precios
       MBox(9) = Escadena(csql!pedidomensaje)                            'mensajes
       If Not IsNull(csql!modovtacodigo) Then Combo3.ListIndex = VerificaCombo(Combo3, csql!modovtacodigo)       'modo de venta
       MBox(10) = Format(csql!pedidofecha, "dd/mm/yyyy")                            'fecha de atencion
       If Not IsNull(csql!formapagocodigo) Then Combo4.ListIndex = VerificaCombo(Combo4, csql!formapagocodigo) 'forma de pago
       Ctr_Ayuda1.xclave = Escadena(csql!clientecodigo)                  ' cliente MBox(11)
       
       '*****Respecto a Clientes *******
       Call Ctr_Ayuda1.Ejecutar
       
       Set acliente = VGCNx.Execute("select * from vt_cliente where clientecodigo='" & Ctr_Ayuda1.xclave & "'")
       If acliente.RecordCount > 0 Then
          Combo6.ListIndex = VerificaCombo(Combo6, acliente!clientetipopersona)
          Combo7.ListIndex = VerificaCombo(Combo7, acliente!clientetipopais)
          Combo8.ListIndex = VerificaCombo(Combo8, IIf(acliente!clientemultidireccion = 1, "S", "N"))
          lcred(0) = numero(acliente!clientesaldodolares)
          lcred(1) = numero(acliente!clientelimitecreddolar)
       End If
       acliente.Close
       Set acliente = Nothing
       
       Ctr_Ayuda2.xclave = Escadena(csql!vendedorcodigo)                    'vendedor
       Call Ctr_Ayuda2.Ejecutar
       MBox(13) = numero(csql!pedidoporccomision)                           'comision
       Ctr_Ayuda3.xclave = Escadena(csql!almacencodigo)                     'almacen
       Call Ctr_Ayuda3.Ejecutar
       'Ctr_Ayuda3.Filtro = "modovtacodigo in (" & modoventa.almacenes & ")"
       
       MBox(15) = numero(csql!pedidototalotros)                             'otros gastos
       MBox(16) = Escadena(csql!pedidonotaped)                              'nota pedido
       MBox(17) = Escadena(csql!pedidoordencompra)                          'orden de compra
       Combo5.ListIndex = VerificaCombo(Combo5, IIf(csql!pedidoautorizacion = 1, "S", "N")) 'autorizacion
       MBox(18) = Format(csql!pedidodiaspago, "##0")                        'dias pago
       MBox2(6) = numero(csql!pedidototitem)                                'Total Cantidad
       MBox2(7) = numero(csql!pedidototbruto)                               'Total Bruto
       MBox2(8) = numero(csql!pedidototalflete)                             'Total Dsctos
       MBox2(9) = numero(csql!pedidototimpuesto)                            'Total Igv
       MBox2(10) = numero(csql!pedidototneto)                               'Neto a Facturar
       MBox(19) = Escadena(csql!pedidoentrega)                             'Entrega de Pedidos
       Text3 = Escadena(csql!pedidoobserva)
       Text4(0) = Escadena(csql!pedidotiporefe)
       Text4(1) = Escadena(csql!pedidonrorefe)
       
       TClie.Value = 0
       SSTab2.Tab = 0
       SSTab2.TabEnabled(2) = True
    End If
    csql.Close
       
    Set csql = VGCNx.Execute("select detpeditem,A.productocodigo,b.adescri,a.unidadcodigo," & _
                          "detpedcantpedida,detpedmontoprecvta,detpeddsctoxitem,detpedimpbruto," & _
                          " detpedporccomis,detpedcantpedidaref,detpedfactorconv " & _
                          "from " & g_DetallePuntoVta & " A " & _
                          "inner Join " & _
                          "[" & VGCNx.DefaultDatabase & "].dbo.maeart B " & _
                          " ON A.productocodigo=b.acodigo COLLATE Modern_Spanish_CI_AI " & _
                          "where pedidonumero='" & TDBGrid2.Columns(0).Text & "' order by detpeditem ")
    
    Set rsdeta = Nothing
    Call CargaGrilla
   
    Do Until csql.EOF
       rsdeta.AddNew
       rsdeta.Fields(0) = Escadena(csql!detpeditem)
       rsdeta.Fields(1) = Escadena(csql!productocodigo)
       rsdeta.Fields(2) = Escadena(csql!adescri)
       rsdeta.Fields(3) = Escadena(csql!unidadcodigo)
       rsdeta.Fields(4) = numero(csql!detpedcantpedida)
       rsdeta.Fields(5) = numero(IIf(IsNull(csql!detpedimpbruto), 0, csql!detpedimpbruto))
       rsdeta.Fields(6) = numero(csql!detpeddsctoxitem)
       rsdeta.Fields(7) = numero(csql!detpedmontoprecvta)
       rsdeta.Fields(8) = numero(csql!detpedporccomis)
       rsdeta.Fields(9) = numero(IIf(IsNull(csql!detpedcantpedidaref), 0, csql!detpedcantpedidaref))
       rsdeta.Fields(10) = numero(IIf(IsNull(csql!detpedfactorconv), 0, csql!detpedfactorconv))
       rsdeta.Update
       csql.MoveNext
    Loop
    csql.Close
    Totales
    Call ConfigGrid
    Set csql = Nothing

End Function

Public Function GrabarData() As Integer
    Dim J As Integer
    Dim regi As Long
    Dim nsql As String
    Dim ltipo As String
    Dim lzona As String
    Dim Previo As Double
    Dim dct02, dct03, dct04, dct05, dct06 As Double
    Dim tinafecto As Double
    Dim xserie As String * 3
    Dim xfactu As Double  'String * 8
    Dim xtipofac As String * 2
    Dim rrsql As New ADODB.Recordset
    Dim CorrNs As ADODB.Recordset
    Dim acmd As New ADODB.Command
    Dim asql As New ADODB.Recordset
    Dim arbusca As New ADODB.Recordset

    'On Error GoTo vererror
   'On Error Resume Next
    
    Call CargarModo
    
    GrabarData = 0
    
    '******** CABECERA DE MOVIMIENTO *****************
    If rsdeta.RecordCount = 0 Then
      MsgBox W1TXT30, vbInformation, MsgTitle
      GrabarData = 0
      Exit Function
    End If
    'Call Totales
    For J = 1 To 29
        wCabe(J) = ""
    Next J
    wCabe(1) = MBox(0)                       'Pto Venta
    wCabe(2) = Trim(MBox(1))                       'nro pedido
    wCabe(3) = Trim(MBox(2))                        'nro factura
    wCabe(4) = Trim(MBox(3))                         'nro boleta
    wCabe(5) = Trim(MBox(4))                         'nro guia
    wCabe(6) = MBox(5)                       'dscto gral
    wCabe(7) = MBox(6)                       'dscto promocional
    wCabe(8) = MBox(7)                       'dscto especial
    wCabe(9) = dllgeneral.ComboDato(Combo1.Text)        'moneda
    wCabe(10) = MBox(8)                      'tipo de cambio
    wCabe(11) = dllgeneral.ComboDato(Combo2.Text)       'lista de precios
    wCabe(12) = Trim(Text3.Text)                  'MBox(9)  'mensajes
    wCabe(13) = dllgeneral.ComboDato(Combo3.Text)       'modo de venta
    wCabe(14) = MBox(10)                     'fecha de atencion
    wCabe(15) = dllgeneral.ComboDato(Combo4.Text)       'forma de pago
    wCabe(16) = MBox3(0)    'Ctr_Ayuda1.xclave         ' MBox(11)                     'cliente
    wCabe(17) = Ctr_Ayuda2.xclave        'MBox(12)                     'vendedor
    wCabe(18) = MBox(13)                  'comision
    wCabe(19) = Ctr_Ayuda3.xclave        'MBox(14)                     'almacen
    wCabe(20) = MBox(15)                     'otros gastos
    wCabe(21) = MBox(16)                     'nota pedido
    wCabe(22) = Ctr_AyuRef.xclave        'MBox(17)  'orden de compra
    wCabe(23) = dllgeneral.ComboDato(Combo5.Text)       'autorizacion
    wCabe(24) = numero(MBox(18))            'dias pago
    wCabe(25) = MBox2(6)                    'Total Cantidad
    wCabe(26) = Round(MBox2(7), 2)          'Total Bruto
    wCabe(27) = 0    'MBox2(8)              'total fletes --T.D.
    wCabe(28) = Round(MBox2(9), 2)          'Total Igv
    wCabe(29) = Round(MBox2(10), 2)         'Neto a Facturar
    wCabe(30) = MBox(19)                    'entrega pedido
    wCabe(31) = MBox3(1)                    'nombre cliente
    wCabe(32) = MBox3(3)                    'direccion
    wCabe(33) = MBox3(2)                    'ruc
    wCabe(34) = VGParamSistem.FechaTrabajo   'fechafactura
    wCabe(35) = DTGlobal                     'Total Descuentos Globales
    wCabe(36) = DTCliente                    'Total Descuentos Cliente
    wCabe(37) = DTOficina                    'Total Descuentos Oficina
    wCabe(38) = DTItem                       'Total Descuentos Item
    wCabe(39) = DTLinea                      'Total Descuentos Linea
    wCabe(40) = DTPromo                      'Total Descuentos x Promocion
    wCabe(41) = Trim(Text3)
    wCabe(42) = Trim(Text4(0))
    wCabe(43) = Trim(Text4(1)) & Trim(Text4(2))
    wCabe(44) = Trim(Ctr_AyudaTc.xclave)
    wCabe(45) = Trim(Ctr_AyudaPro.xclave)
    wCabe(46) = TxtHor.Text
    
    If cOpc(0).Value Or cOpc3(0).Value Then
        Set asql = VGCNx.Execute("select productocodigo,detpedcantpedida from " & g_DetallePuntoVta & " where pedidonumero='" & MBox(1) & "'")
        If asql.RecordCount > 0 Then
          If VGParamSistem.stockcomp = 1 Then
             asql.MoveFirst
             Do Until asql.EOF
                    Set acmd.ActiveConnection = VGgeneral
                    acmd.CommandType = adCmdStoredProc
                    acmd.CommandTimeout = 0
                    acmd.CommandText = "vt_actualizoalma_pro"
                    acmd.Prepared = True
                  With acmd
                        .Parameters("@basedes") = CStr(VGCNx.DefaultDatabase)
                        .Parameters("@almacen") = wCabe(19)
                        .Parameters("@tipo") = "3"
                        .Parameters("@articulo") = Trim(asql.Fields(0))
                        .Parameters("@cantidad") = asql.Fields(1) * -1
                  End With
                    acmd.Execute
                    Set acmd = Nothing
                    asql.MoveNext
              Loop
           End If
           VGCNx.Execute "Delete From " & g_DetallePuntoVta & " where pedidonumero='" & MBox(1) & "'"
           VGCNx.Execute "Delete From " & g_PedidoPuntoVta & " where pedidonumero='" & MBox(1) & "'"
        End If
        asql.Close
        nsql = "Insert Into " & g_PedidoPuntoVta & "("
    ElseIf cOpc(1).Value Or cOpc3(1).Value Then
        Set asql = VGCNx.Execute("select * from vt_detallepedido where pedidonumero='" & MBox(1) & "'")
        If asql.RecordCount > 0 Then
           VGCNx.Execute "Delete From vt_detallepedido where pedidonumero='" & MBox(1) & "'"
           VGCNx.Execute "Delete From vt_pedido where pedidonumero='" & MBox(1) & "'"
        End If
        asql.Close
        If cOpc3(1).Value Then
           VGCNx.Execute "Delete From " & g_DetallePuntoVta & " where pedidonumero='" & MBox(1) & "'"
           VGCNx.Execute "Delete From " & g_PedidoPuntoVta & " where pedidonumero='" & MBox(1) & "'"
        End If
        nsql = "Insert Into vt_Pedido ("
    End If
    Set asql = Nothing
'    VGcnx.CommitTrans
    
'    VGcnx.BeginTrans
    ' ** Verificando Numeracion de Documentos *****
    If cOpc(1).Value Or cOpc3(1).Value Then
        If cOpc2(0).Value Then
          If cOpc(1).Value Then
             MBox(1) = g_pedserie & Right("000000000000" & TraeDataSerie("select puntovtadoccorr from vt_puntovtadocumento where empresacodigo='" & VGParametros.empresacodigo & "' and documentocodigo='" & g_tipoped & "' and puntovtadocserie='" & g_pedserie & "' and puntovtacodigo='" & g_ptoventa & "'", VGCNx), 8) 'MBox(1).MaxLength)
          End If
          'wCabe(34) = Date                       'fechafactura
          MBox(2) = g_facserie & Right("000000000000" & TraeDataSerie("select puntovtadoccorr from vt_puntovtadocumento where empresacodigo='" & VGParametros.empresacodigo & "' and documentocodigo='" & g_tipofac & "' and puntovtadocserie='" & g_facserie & "' and puntovtacodigo='" & g_ptoventa & "'", VGCNx), 8) ' MBox(2).MaxLength)
          MBox(3) = "0": 'MBox(4) = "0"
          
          If dllgeneral.VerificaDatoExistente(VGCNx, "select * from vt_pedido where empresacodigo='" & VGParametros.empresacodigo & "' and   pedidonrofact='" & MBox(2) & "' and pedidotipofac='" & g_tipofac & "'") = 1 Then
            MsgBox "Ya existe Documento " & g_tipofac & "-" & MBox(2), vbInformation, MsgTitle
            GrabarData = 0
            Exit Function
          End If
        ElseIf cOpc2(1).Value Then
          If cOpc(1).Value Then
             MBox(1) = g_pedserie & Right("000000000000" & TraeDataSerie("select puntovtadoccorr from vt_puntovtadocumento where empresacodigo='" & VGParametros.empresacodigo & "' and documentocodigo='" & g_tipoped & "' and puntovtadocserie='" & g_pedserie & "' and puntovtacodigo='" & g_ptoventa & "'", VGCNx), 8) 'MBox(1).MaxLength)
          End If
          'wCabe(34) = Date                       'fechaboleta
          MBox(3) = g_bolserie & Right("00000000000" & TraeDataSerie("select puntovtadoccorr from vt_puntovtadocumento where empresacodigo='" & VGParametros.empresacodigo & "' and documentocodigo='" & g_tipobol & "' and puntovtadocserie='" & g_bolserie & "' and puntovtacodigo='" & g_ptoventa & "'", VGCNx), 8) 'MBox(3).MaxLength)
          MBox(2) = "0"
          'AKI SE CAMBIA POR NS
           Set CorrNs = VGCNx.Execute("select tanumsal from tabalm where taalma='" & Ctr_Ayuda3.xclave & "'")
           If CorrNs.RecordCount > 0 Then
               MBox(4) = Right("00000000000" & Trim(CStr(CorrNs!tanumsal)), 11)                      'nro pedido"
           End If
           CorrNs.Close
           Set CorrNs = Nothing
          'MBox(4) = ""
          
          If dllgeneral.VerificaDatoExistente(VGCNx, "select * from vt_pedido where empresacodigo='" & VGParametros.empresacodigo & "' and pedidonrofact='" & MBox(3) & "' and pedidotipofac='" & g_tipobol & "'") = 1 Then
            MsgBox "Ya existe Documento " & g_tipobol & "-" & MBox(3), vbInformation, MsgTitle
            GrabarData = 0
            Exit Function
          End If
        ElseIf cOpc2(2).Value Then
          If cOpc(1).Value Then
             MBox(1) = g_pedserie & Right("000000000000" & TraeDataSerie("select puntovtadoccorr from vt_puntovtadocumento where  empresacodigo='" & VGParametros.empresacodigo & "' and documentocodigo='" & g_tipoped & "' and puntovtadocserie='" & g_pedserie & "' and puntovtacodigo='" & g_ptoventa & "'", VGCNx), 8) ' MBox(1).MaxLength)
          End If
         ' wCabe(34) = Date                       'fechaguia
          MBox(4) = g_guiaserie & Right("000000000000" & TraeDataSerie("select puntovtadoccorr from vt_puntovtadocumento where empresacodigo='" & VGParametros.empresacodigo & "' and documentocodigo='" & g_tipoguia & "' and puntovtadocserie='" & g_guiaserie & "' and puntovtacodigo='" & g_ptoventa & "'", VGCNx), 8)  ' MBox(4).MaxLength)
          MBox(2) = "0": MBox(3) = "0"
          If dllgeneral.VerificaDatoExistente(VGCNx, "select * from vt_pedido where empresacodigo='" & VGParametros.empresacodigo & "' and pedidonrofact='" & MBox(3) & "' and pedidotipofac='" & g_tipoguia & "'") = 1 Then
            MsgBox "Ya existe Documento " & g_tipoguia & "-" & MBox(3), vbInformation, MsgTitle
            GrabarData = 0
            Exit Function
          End If
        '-------------------------------------------------------------------------------------
        ElseIf cOpc2(3).Value Then
          If cOpc(1).Value Then
             MBox(3) = g_ticserie & Right("000000000000" & TraeDataSerie("select puntovtadoccorr from vt_puntovtadocumento where  empresacodigo='" & VGParametros.empresacodigo & "' and documentocodigo='" & g_tipoticket & "' and puntovtadocserie='" & g_ticserie & "' and puntovtacodigo='" & g_ptoventa & "' ", VGCNx), 8) ' MBox(1).MaxLength)
          End If
         ' wCabe(34) = Date                       'fechaguia
          'MBox(2) = "0": MBox(3) = "0"
          If dllgeneral.VerificaDatoExistente(VGCNx, "select * from vt_pedido where empresacodigo='" & VGParametros.empresacodigo & "' and pedidonrofact='" & MBox(3) & "' and pedidotipofac='" & g_tipoticket & "'") = 1 Then
            MsgBox "Ya existe Documento " & g_tipoticket & "-" & MBox(3), vbInformation, MsgTitle
            GrabarData = 0
            Exit Function
          End If
          '-------------------------------------------------------------------------------------
        End If
    End If
    
    If cOpc(1).Value Or cOpc(0).Value Then
        '*** Verifica Serie Documentos *****
        nsql = "Update vt_puntovtadocumento " & _
                " set puntovtadoccorr='" & Right("00000000" & Trim(CStr(MBox(1) + 1)), 8) & "'" & _
                " Where documentocodigo='" & g_tipoped & "' and puntovtacodigo='" & g_ptoventa & "' and puntovtadocserie='" & g_pedserie & "'"
        nsql = nsql & " and empresacodigo='" & VGParametros.empresacodigo & "' and puntovtacodigo='" & VGParametros.puntovta & "'"
        VGCNx.Execute nsql
    End If
    
    '***** Actualizando Numeracion de Documentos*****
    If cOpc(1).Value Or cOpc3(1).Value Then
         If cOpc2(0).Value Then
             If Len(Trim(g_facserie)) = 0 Then
                MsgBox "No existe Serie de Facturas....Verifique!!", vbInformation, MsgTitle
                'VGcnx.RollbackTrans
                Exit Function
             End If
            
            nsql = "Update vt_puntovtadocumento " _
                  & " set puntovtadoccorr='" & Right("00000000" & Trim(CStr(MBox(2) + 1)), 8) & "' " _
                  & " Where documentocodigo='" & g_tipofac & "' and puntovtacodigo='" & g_ptoventa & "' " _
                  & " and puntovtadocserie='" & g_facserie & "' " _
                  & " and empresacodigo='" & VGParametros.empresacodigo & "' and puntovtacodigo='" & VGParametros.puntovta & "'"
    
         ElseIf cOpc2(1).Value Then
            If Len(Trim(g_bolserie)) = 0 Then
               MsgBox "No existe Serie de Boletas....Verifique!!", vbInformation, MsgTitle
               'VGcnx.RollbackTrans
               Exit Function
            End If
        
            nsql = "Update vt_puntovtadocumento " _
                    & " set puntovtadoccorr='" & Right("00000000" & Trim(CStr(MBox(3) + 1)), 8) & "' " _
                    & " Where documentocodigo='" & g_tipobol & "' and puntovtacodigo='" & g_ptoventa & "' " _
                    & " and puntovtadocserie='" & g_bolserie & "' and empresacodigo='" & VGParametros.empresacodigo & "' " _
                    & " and puntovtacodigo='" & VGParametros.puntovta & "'"
    
         ElseIf cOpc2(2).Value Then
             If Len(Trim(g_guiaserie)) = 0 Then
                MsgBox "No existe Serie de Guias....Verifique!!", vbInformation, MsgTitle
                'VGcnx.RollbackTrans
                Exit Function
             End If
        
        ElseIf cOpc2(3).Value Then
             If Len(Trim(g_ticserie)) = 0 Then
                MsgBox "No existe Serie de Ticket....Verifique!!", vbInformation, MsgTitle
                'VGcnx.RollbackTrans
                Exit Function
             End If
        
             nsql = "Update vt_puntovtadocumento " & _
                    "set puntovtadoccorr='" & Right("00000000" & Trim(CStr(MBox(3) + 1)), 8) & "'" & _
                    " Where documentocodigo='" & g_tipoticket & "' and puntovtacodigo='" & g_ptoventa & "' and puntovtadocserie='" & g_ticserie & "'"
             nsql = nsql & " and empresacodigo='" & VGParametros.empresacodigo & "' and puntovtacodigo='" & VGParametros.puntovta & "'"
        
        End If
        
        VGCNx.Execute nsql
        If modoventa.emiteguia = 1 Then
            nsql = "Update vt_puntovtadocumento " & _
                       "set puntovtadoccorr='" & Right("00000000" & Trim(CStr(MBox(4) + 1)), 8) & "'" & _
                       " Where documentocodigo='" & g_tipoguia & "' and puntovtacodigo='" & g_ptoventa & "' and puntovtadocserie='" & g_guiaserie & "'"
            nsql = nsql & " and empresacodigo='" & VGParametros.empresacodigo & "' and puntovtacodigo='" & VGParametros.puntovta & "'"

            VGCNx.Execute nsql
        End If
    End If
    DoEvents
    '**cambio de documentacion
    wCabe(2) = Trim(MBox(1))                         'nro pedido
    wCabe(3) = Trim(MBox(2))                         'nro factura
    wCabe(4) = Trim(MBox(3))                         'nro boleta
    'wCabe(5) = Trim(MBox(4))                         'nro guia
    
    '---------------------------------------------------------------
Dim rsql As ADODB.Recordset
Dim rsql2 As String, numsal As String
    
    Set rsql = VGCNx.Execute("select TANUMSAL from TabAlm  WHERE TAALMA='01'")
    wCabe(5) = IIf(rsql(0) = 0, 1, rsql(0))
    numsal = Format(Val(wCabe(5)) + 1, "00000000000")
    rsql2 = "Update TabAlm set TANUMsal= '" & numsal & "' where TAALMA='01' "
    VGCNx.Execute rsql2
    wCabe(5) = Format(Val(wCabe(5)), "00000000000")
    '---------------------------------------------------------------
   
    If cOpc(1).Value Or cOpc3(1).Value Then
      If cOpc2(0).Value Then
         wCabe(3) = Trim(MBox(2))
         wCabe(4) = g_tipofac
      ElseIf cOpc2(1).Value Then
         wCabe(3) = Trim(MBox(3))
         wCabe(4) = g_tipobol
      ElseIf cOpc2(2).Value Then
         wCabe(3) = Trim(MBox(4))
         wCabe(4) = g_tipoguia
      ElseIf cOpc2(3).Value Then
         wCabe(3) = Trim(MBox(3))
         wCabe(4) = g_tipoticket
      End If
    Else
        wCabe(3) = 0  ' Trim(MBox(2))                         'nro factura
        wCabe(4) = "" 'Trim(MBox(3))                         'nro boleta
        wCabe(5) = 0  'Trim(MBox(4))                         'nro guia
    End If
    
    DoEvents
    
    Set acmd.ActiveConnection = VGgeneral
    acmd.CommandType = adCmdStoredProc
    acmd.CommandText = "vt_ingresapedido_pro"
    acmd.CommandTimeout = 0
    acmd.Prepared = True
    With acmd
        .Parameters("@base") = VGCNx.DefaultDatabase
        .Parameters("@tabla") = IIf(cOpc(0).Value Or cOpc3(0).Value, g_PedidoPuntoVta, "vt_pedido")
        .Parameters("@tipo") = IIf(dllgeneral.VerificaDatoExistente(VGCNx, "select * from " & IIf(cOpc(0).Value Or cOpc3(0).Value, g_PedidoPuntoVta, "vt_pedido") & " where pedidonumero='" & wCabe(2) & "'") = 0, "1", "2") '"1"
        .Parameters("@puntovta") = wCabe(1)
        .Parameters("@numero") = wCabe(2)
        .Parameters("@factura") = wCabe(3)
        .Parameters("@boleta") = wCabe(4)
        .Parameters("@guia") = wCabe(5)
        .Parameters("@dsctoglobal") = wCabe(6)
        .Parameters("@dsctoppago") = wCabe(7)
        .Parameters("@dsctovtaofi") = wCabe(8)
        .Parameters("@moneda") = wCabe(9)
        .Parameters("@tipocambio") = wCabe(10)
        .Parameters("@listaprecio") = wCabe(11)
        .Parameters("@mensaje") = wCabe(12)
        .Parameters("@modoventa") = wCabe(13)
        .Parameters("@fecha") = wCabe(14)
        .Parameters("@formapago") = wCabe(15)
        .Parameters("@cliente") = wCabe(16)
        .Parameters("@vendedor") = wCabe(17)
        .Parameters("@porcomision") = wCabe(18)
        .Parameters("@almacen") = wCabe(19)
        .Parameters("@totalotros") = wCabe(20)
        .Parameters("@notaped") = wCabe(21)
        .Parameters("@ordencompra") = wCabe(22)
        .Parameters("@autoriza") = wCabe(23)
        .Parameters("@diaspago") = wCabe(24)
        .Parameters("@totalitem") = wCabe(25)
        .Parameters("@totalbruto") = wCabe(26)
        .Parameters("@totalflete") = wCabe(27)
        .Parameters("@totalimpuesto") = wCabe(28)
        .Parameters("@totalneto") = wCabe(29)
        .Parameters("@usuario") = g_usuario
        .Parameters("@fechaactual") = Date
        .Parameters("@totaldsctoxlinea") = wCabe(39)
        .Parameters("@montodsctoppago") = DTPPago
        .Parameters("@entregapedido") = wCabe(30)
        .Parameters("@razon") = wCabe(31)
        .Parameters("@direccion") = wCabe(32)
        .Parameters("@ruc") = wCabe(33)
        .Parameters("@fechafactura") = wCabe(34)
        .Parameters("@TDGlobal") = wCabe(35)
        .Parameters("@TDCliente") = wCabe(36)
        .Parameters("@TDOficina") = wCabe(37)
        .Parameters("@TDItem") = wCabe(38)
        .Parameters("@TDPromo") = wCabe(40)
        .Parameters("@observa") = ""   'wCabe(41)
        .Parameters("@tiporefe") = wCabe(42)
        .Parameters("@nrorefe") = wCabe(43)
        .Parameters("@nrotransporte") = Ctr_AyuTransporte.xclave
        .Parameters("@empresa") = VGParametros.empresacodigo
        '.Parameters("@TipoContacto") = Trim(wCabe(44))
'        .Parameters("@Profesional") = Trim(wCabe(45))
'        .Parameters("@hora") = Trim(wCabe(46))
        
    End With
    acmd.Execute
    Set acmd = Nothing
    DoEvents
       
  If modoventa.ctrlinventario = "1" And (cOpc3(1).Value Or cOpc(1).Value) Then
     If Chkentrega.Value = 0 Then
        If modoventa.emiteguia = 1 Then
           guias_num = g_pedserie & Right("000000000000" & TraeDataSerie("select puntovtadoccorr from vt_puntovtadocumento where empresacodigo='" & VGParametros.empresacodigo & "' and documentocodigo='GR' and puntovtadocserie='" & g_pedserie & "' and puntovtacodigo='" & g_ptoventa & "'", VGCNx), 8)
           wCabe(5) = guias_num
              
           VGCNx.Execute "Update vt_puntovtadocumento " & _
                " set puntovtadoccorr='" & Right("00000000" & Trim(CStr(Val(guias_num) + 1)), 8) & "'" & _
               " Where empresacodigo='" & VGParametros.empresacodigo & "' and documentocodigo='GR' and puntovtacodigo='" & g_ptoventa & "' and puntovtadocserie='" & g_pedserie & "'"
           
        End If
        Set acmd.ActiveConnection = VGgeneral
        acmd.CommandType = adCmdStoredProc
        acmd.CommandText = "vt_ingresoalma_pro"
        acmd.CommandTimeout = 0
        acmd.Prepared = True
        With acmd
            .Parameters("@base") = VGCNx.DefaultDatabase
            .Parameters("@tabla") = "movalmcab"
            .Parameters("@tipo") = "2"
            .Parameters("@puntovta") = wCabe(1)
            .Parameters("@numero") = wCabe(5)
            .Parameters("@factura") = wCabe(3)
            .Parameters("@boleta") = wCabe(4)
            .Parameters("@guia") = wCabe(5)
            .Parameters("@dsctoglobal") = wCabe(6)
            .Parameters("@dsctoppago") = wCabe(7)
            .Parameters("@dsctovtaofi") = wCabe(8)
            .Parameters("@moneda") = IIf(wCabe(9) = g_TipoSol, "S", "D")
            .Parameters("@tipocambio") = wCabe(10)
            .Parameters("@listaprecio") = wCabe(11)
            .Parameters("@mensaje") = wCabe(12)
            .Parameters("@modoventa") = wCabe(13)
            .Parameters("@fecha") = wCabe(14)
            .Parameters("@formapago") = wCabe(15)
            .Parameters("@cliente") = wCabe(16)
            .Parameters("@vendedor") = wCabe(17)
            .Parameters("@porcomision") = wCabe(18)
            .Parameters("@almacen") = wCabe(19)
            .Parameters("@totalotros") = wCabe(20)
            .Parameters("@notaped") = wCabe(2)
            .Parameters("@ordencompra") = wCabe(22)
            .Parameters("@autoriza") = wCabe(23)
            .Parameters("@diaspago") = wCabe(24)
            .Parameters("@totalitem") = wCabe(25)
            .Parameters("@totalbruto") = wCabe(26)
            .Parameters("@totalflete") = wCabe(27)
            .Parameters("@totalimpuesto") = wCabe(28)
            .Parameters("@totalneto") = wCabe(29)
            .Parameters("@usuario") = g_usuario
            .Parameters("@fechaactual") = Date
            .Parameters("@totaldsctoxlinea") = wCabe(39)
            .Parameters("@montodsctoppago") = DTPPago
            .Parameters("@entregapedido") = wCabe(30)
            .Parameters("@razon") = wCabe(31)
            .Parameters("@direccion") = wCabe(32)
            .Parameters("@ruc") = wCabe(33)
            .Parameters("@fechafactura") = wCabe(34)
            .Parameters("@TDGlobal") = wCabe(35)
            .Parameters("@TDCliente") = wCabe(36)
            .Parameters("@TDOficina") = wCabe(37)
            .Parameters("@TDItem") = wCabe(38)
            .Parameters("@TDPromo") = wCabe(40)
        End With
        acmd.Execute
        Set acmd = Nothing
        DoEvents
     End If
  End If
    
    If cOpc2(0).Value Or cOpc2(1).Value Or cOpc2(2).Value Then
       If wCabe(9) = g_TipoSol Then
            VGCNx.Execute "Update vt_cliente " & _
                       " Set clientesaldosoles=ISNULL(clientesaldosoles,0)+" & CDbl(wCabe(29)) & _
                       "      Where clientecodigo='" & wCabe(16) & "'"
       ElseIf wCabe(9) = g_TipoDolar Then
            VGCNx.Execute "Update vt_cliente " & _
                       " Set clientesaldodolares=ISNULL(clientesaldodolares,0)+" & CDbl(wCabe(29)) & _
                       "      Where clientecodigo='" & wCabe(16) & "'"
       End If
    End If
    DoEvents
    '********** DETALLE DE MOVIMIENTOS *****************
    rsdeta.MoveFirst
    regi = 0
    tinafecto = 0
    TImporte = 0
    Tbruto = 0
    TCant = 0
        
    Do Until rsdeta.EOF
           'IMPORTE DE MONTO BRUTO SIN IGV, ES DECIR PRECIO X CANTIDAD
           'Tbruto = Tbruto + (rsdeta.Fields(5))
           TCant = rsdeta.Fields(4)
           If VGParamSistem.tieneigv = "1" Then
             TImporte = (rsdeta.Fields(5) * rsdeta.Fields(4)) / (1 + VGParamSistem.Igv)
           End If
           'TImporte = rsdeta.Fields(5) * rsdeta.Fields(4)
           
           If IsNull(text1) Or Len(Trim(text1)) = 0 Then
                 dct06 = 0
           Else
               dct06 = TImporte * (CDbl(text1))
           End If
          
           'DESCUENTO POR ITEM
           dct02 = 0
           dct02 = (TImporte * (rsdeta.Fields(6) / 100))
           
           'DESCUENTO ESPECIAL  :w8dct03 =(w8bruto - w8dct02-w8dct06)*w2dctpp/100
            dct03 = 0
            'Lo k estaba antes
            'dct03 = (TImporte - dct02 - dct06) * (MBox(7) / 100)            '(Tbruto-dct02-dct06)
            dct03 = MBox(7)
            
           'DESCUENTO POR PROMOCION  : w8dct04 =(w8bruto - w8dct02-w8dct03-w8dct06)*w2dctpr/100
            dct04 = 0
            dct04 = (TImporte - dct02 - dct03 - dct06) * (MBox(6) / 100)
            
           'DESCUENTO GENERAL : w8dct05 =(w8bruto - w8dct02-w8dct03-w8dct04-w8dct06)*w2dctgl/100
            dct05 = 0
            dct05 = (TImporte - dct02 - dct03 - dct04 - dct06) * (MBox(5) / 100)
           
           'ACUMULADO DE TOTAL DESCUENTOS  :w8dctos = w8dct02 + w8dct03+w8dct04+w8dct05+w8dct06
            Tdscto = dct02 + dct03 + dct04 + dct05 + dct06
            
           'ACUMULADO DE SUBTOTAL DE VENTA : w8subto = w8bruto - w8dctos
           TSub = 0
           TSub = (TImporte - (dct02 + dct03 + dct04 + dct05 + dct06))
           Previo = TSub
           If VGParamSistem.tieneigv = "1" Then
              'CALCULAMOS EL IGV
              Previo = (TSub * VGParamSistem.Igv)
           Else
                If modoventa.impuestos = "1" Then
                     Previo = (TImporte - (dct02 + dct03 + dct04 + dct05 + dct06))
                     Previo = (Previo * VGParamSistem.Igv)
                Else
                    If rsdeta.Fields(11) > 0 Then
                         Previo = (TImporte - (dct02 + dct03 + dct04 + dct05 + dct06))
                         Previo = (Previo * rsdeta.Fields(11))
                    Else
                        Previo = TSub '
                        tinafecto = tinafecto + TSub
                   End If
                End If
           End If
        
        If cOpc(0).Value Or cOpc3(0).Value Then
            nsql = g_DetallePuntoVta   '"Tempodetallepedido"
        ElseIf cOpc(1).Value Or cOpc3(1).Value Then
            nsql = "vt_detallepedido"
        End If
        Set acmd.ActiveConnection = VGgeneral
        acmd.CommandType = adCmdStoredProc
        acmd.CommandTimeout = 0
        acmd.CommandText = "vt_ingresodetallepedido_pro"
        acmd.Prepared = True
        
        With acmd
            .Parameters("@base") = VGCNx.DefaultDatabase
            .Parameters("@tabla") = nsql
            .Parameters("@empresa") = VGParametros.empresacodigo
            .Parameters("@tipo") = "1"
            .Parameters("@item") = rsdeta.Fields(0)
            .Parameters("@numero") = MBox(1)
            .Parameters("@producto") = rsdeta.Fields(1)
            .Parameters("@unidad") = rsdeta.Fields(3)
            .Parameters("@cantidad") = rsdeta.Fields(4)
            .Parameters("@preciopacto") = (TImporte - Tdscto + Previo)    'rsdeta.Fields(7)
            .Parameters("@dsctoxitem") = rsdeta.Fields(6)
            .Parameters("@importebruto") = TImporte   '(rsdeta.Fields(7)) / (1 + VGParamSistem.Igv)
            .Parameters("@porcomision") = rsdeta.Fields(8)
            .Parameters("@mdsctoitem") = Tdscto
            .Parameters("@mdsctoxlinea") = 0
            .Parameters("@mdsctoxprom") = 0     '0
            .Parameters("@mimpor") = Previo   'rsdeta.Fields(7) - (rsdeta.Fields(7) / (1 + VGParamSistem.Igv)) 'Previo
            .Parameters("@unidadref") = IIf(IsNull(rsdeta.Fields(9)) Or Len(Trim(rsdeta.Fields(9))) = 0, 0, CDbl(rsdeta.Fields(9)))
            .Parameters("@preciolista") = rsdeta.Fields(5)
            .Parameters("@partida") = " "
            .Parameters("@metrica") = " "
            .Parameters("@observacion") = MBox(11)      ' rsdeta.Fields(14)
       
        End With
        acmd.Execute
        Set acmd = Nothing
            
            '******Actualizamos Saldos en Almacen *********
            If modoventa.ctrlinventario = "1" Then
            
                '--Actualizamos el archivo stkart --
               If cOpc2(0).Value Or cOpc2(1).Value Or cOpc2(2).Value Or cOpc2(3).Value Then     'PUEDE SER AKI
               
                    If cOpc2(0).Value Then
                        xserie = Left(MBox(2).Text, 3)
                        xfactu = Val(Right(MBox(2).Text, 8))
                        xtipofac = g_tipofac
                    ElseIf cOpc2(1).Value Then
                        xserie = Left(MBox(3).Text, 3)
                        xfactu = Val(Right(MBox(3).Text, 8))
                        xtipofac = g_tipobol
                    ElseIf cOpc2(2).Value Then
                        xserie = Left(MBox(4).Text, 3)
                        xfactu = Val(Right(MBox(4).Text, 8))
                        xtipofac = g_tipoguia
                    Else
                        xserie = Left(LblTicSer.Caption, 3)
                        xfactu = Val(Right(LblTicSer.Caption, 8))
                        xtipofac = g_tipoticket
                    End If
                   If Chkentrega.Value = 0 Then
                      Set acmd.ActiveConnection = VGgeneral
                      acmd.CommandType = adCmdStoredProc
                      acmd.CommandTimeout = 0
                      acmd.CommandText = "vt_ingresodetallealma_pro"
                      acmd.Prepared = True
                      With acmd
                        .Parameters("@base") = VGCNx.DefaultDatabase
                        .Parameters("@tabla") = "movalmdet" ' nsql
                        .Parameters("@tipo") = "2"
                        .Parameters("@item") = rsdeta.Fields(0)
                        .Parameters("@numero") = wCabe(5)
                        .Parameters("@producto") = Trim(rsdeta.Fields(1))
                        .Parameters("@unidad") = rsdeta.Fields(3)
                        .Parameters("@cantidad") = rsdeta.Fields(4)
                        .Parameters("@preciopacto") = rsdeta.Fields(5)
                        .Parameters("@dsctoxitem") = rsdeta.Fields(6)
                        .Parameters("@importebruto") = rsdeta.Fields(7)
                        .Parameters("@porcomision") = rsdeta.Fields(8)
                        .Parameters("@mdsctoitem") = Tdscto
                        .Parameters("@mdsctoxlinea") = 0
                        .Parameters("@mdsctoxprom") = Previo     '0
                        .Parameters("@mimpor") = rsdeta.Fields(7)       'Previo
                        .Parameters("@unidadref") = IIf(IsNull(rsdeta.Fields(9)) Or Len(Trim(rsdeta.Fields(9))) = 0, 0, CDbl(rsdeta.Fields(9)))
                        .Parameters("@almacen") = wCabe(19)
                      End With
                      acmd.Execute
                      Set acmd = Nothing
                      Set acmd.ActiveConnection = VGgeneral
                      acmd.CommandType = adCmdStoredProc
                      acmd.CommandTimeout = 0
                      acmd.CommandText = "vt_actualizoalma_pro"
                      acmd.Prepared = True
                      With acmd
                        .Parameters("@basedes") = CStr(VGCNx.DefaultDatabase)
                        .Parameters("@almacen") = wCabe(19)
                         If Chkentrega.Value = 0 Then
                             .Parameters("@tipo") = "1"
                          Else
                        
                             .Parameters("@tipo") = "3"
                        End If
                        .Parameters("@articulo") = Trim(rsdeta.Fields(1))
                        .Parameters("@cantidad") = rsdeta.Fields(4)
                      End With
                      acmd.Execute
                      Set acmd = Nothing
                    End If
              Else
                    If VGParamSistem.stockcomp = 1 Then
                       Set acmd.ActiveConnection = VGgeneral
                       acmd.CommandType = adCmdStoredProc
                       acmd.CommandTimeout = 0
                       acmd.CommandText = "vt_actualizoalma_pro"
                       acmd.Prepared = True
                       With acmd
                        .Parameters("@basedes") = CStr(VGCNx.DefaultDatabase)
                        .Parameters("@almacen") = wCabe(19)
                        .Parameters("@tipo") = "3"
                        .Parameters("@articulo") = Trim(rsdeta.Fields(1))
                        .Parameters("@cantidad") = rsdeta.Fields(4)
                      End With
                      acmd.Execute
                      Set acmd = Nothing
                   End If
                 End If
             End If
        
                
        rsdeta.MoveNext
        regi = regi + 1
    Loop
    
    '*****Actualizamos el Valor de Inafecto**********
    VGCNx.Execute "UPDATE " & g_PedidoPuntoVta & _
               " Set Pedidototinafecto=" & tinafecto & _
               " Where empresacodigo='" & VGParametros.empresacodigo & "' and pedidonumero='" & MBox(1) & "'"
    
   '*Grabar en los cargos ***ctacte ***
    
    If (cOpc2(0).Value Or cOpc2(1).Value Or cOpc2(2).Value Or cOpc2(3).Value) And modoventa.ctacte = "1" Then
        lzona = "00"
        Set asql = VGCNx.Execute("select * from vt_zonavendedor where vendedorcodigo='" & wCabe(17) & "'")
        If asql.RecordCount > 0 Then
            lzona = Escadena(asql!zonacodigo)
        End If
        asql.Close
        Set asql = Nothing
           
        ltipo = "1"
        If dllgeneral.VerificaDatoExistente(VGCNx, "select * from vt_cargo where empresacodigo='" & VGParametros.empresacodigo & "' and documentocargo='" & IIf(cOpc2(0).Value, g_tipofac, g_tipobol) & "' and cargonumdoc='" & IIf(cOpc2(0).Value, MBox(3), MBox(4)) & "'") = 0 Then
          ltipo = "1"
        Else
          ltipo = "2"
        End If
        Set acmd.ActiveConnection = VGgeneral
        acmd.CommandType = adCmdStoredProc
        acmd.CommandTimeout = 0
        acmd.CommandText = "vt_ingresacargofactura_pro"
        acmd.Prepared = True
        With acmd
            .Parameters("@base") = VGCNx.DefaultDatabase
            .Parameters("@empresa") = VGParametros.empresacodigo
            .Parameters("@tipo") = ltipo
            .Parameters("@tabla") = "vt_cargo"
            If cOpc2(0).Value = True Then
                .Parameters("@tipodocu") = g_tipofac
                .Parameters("@numero") = MBox(2)
            ElseIf cOpc2(1).Value = True Then
                .Parameters("@tipodocu") = g_tipobol
                .Parameters("@numero") = MBox(3)
            Else
                .Parameters("@tipodocu") = g_tipoguia
                .Parameters("@numero") = MBox(4)
            End If
            .Parameters("@cliente") = Escadena(wCabe(16))
            .Parameters("@vendedor") = Escadena(wCabe(17))
            .Parameters("@zona") = lzona
            .Parameters("@apefecemi") = wCabe(14)
            .Parameters("@moneda") = Escadena(wCabe(9))
            .Parameters("@apeimppag") = wCabe(29)
            .Parameters("@usuario") = g_usuario
            .Parameters("@tipocambio") = wCabe(10)
            .Parameters("@fechaact") = Date
            .Parameters("@flagcancel") = "0"
            .Parameters("@fechavenci") = CDate(wCabe(14)) + CDbl(wCabe(24))
            .Parameters("@cargoabono") = "C"
        End With
        acmd.Execute
        Set acmd = Nothing
        
    End If
    If cOpc(1).Value Or cOpc3(1).Value Then
         If cOpc2(0).Value Then
              MsgBox "Se Grabo Satisfactoriamente el Documento." & Chr(13) & Chr(10) & "FACTURA => " & MBox(2), vbInformation, MsgTitle
         ElseIf cOpc2(1).Value Then
              MsgBox "Se Grabo Satisfactoriamente el Documento." & Chr(13) & Chr(10) & "BOLETA => " & MBox(3), vbInformation, MsgTitle
         ElseIf cOpc2(2).Value Then
              MsgBox "Se Grabo Satisfactoriamente el Documento." & Chr(13) & Chr(10) & "GUIA => " & MBox(4), vbInformation, MsgTitle
         ElseIf cOpc2(3).Value Then
              MsgBox "Se Grabo Satisfactoriamente el Documento." & Chr(13) & Chr(10) & "TICKET => " & MBox(3), vbInformation, MsgTitle
         Else
              MsgBox "Se Grabo Satisfactoriamente el Documento." & Chr(13) & Chr(10) & "PEDIDO => " & MBox(1), vbInformation, MsgTitle
         End If
      Else
              MsgBox "Se Grabo Satisfactoriamente el Documento." & Chr(13) & Chr(10) & "PEDIDO => " & MBox(1), vbInformation, MsgTitle
     End If
    GrabarData = 1
    
    
'vererror:
'   If Err Then
'      MsgBox "Comunicarse con Sistemas ...!!" & Chr(13) & Chr(10) & VGcnx.Errors(0).Number & "-" & VGcnx.Errors(0).Description
'      Exit Function
'   End If
End Function

Public Function verificaproducto() As Integer
   On Error Resume Next
    verificaproducto = 0
    If rsdeta.RecordCount > 0 Then
       rsdeta.MoveFirst
       Do Until rsdeta.EOF
           If Escadena(rsdeta.Fields(1)) = MBox2(1) And flag = 0 And VGParamSistem.kitvirtual = 0 Then
              verificaproducto = 1
              Exit Do
           End If
           rsdeta.MoveNext
       Loop
    End If
End Function

Public Sub TraerProducto()
  Dim rabusca As New ADODB.Recordset
  Dim nsql As String
  Dim nvalor As Double
  Dim mone As String
  Dim nprecio As Double
  On Error Resume Next

    If Combo2.ListCount > 0 Then
       If modoventa.ctrlinventario = 0 Then
            nsql = "select *,stskdis from [" & VGCNx.DefaultDatabase & "].dbo.maeart " & _
                    "inner join [" & _
                    VGCNx.DefaultDatabase & "].dbo.stkart " & _
                    " ON acodigo=stcodigo " & _
                   " where acodigo='" & MBox2(1) & "'"
       Else
          If VGParamSistem.stockcomp = 0 Then
            nsql = "select *,stskdis from [" & VGCNx.DefaultDatabase & "].dbo.maeart " & _
                    "inner join [" & _
                    VGCNx.DefaultDatabase & "].dbo.stkart " & _
                    " ON acodigo=stcodigo " & _
                    " where acodigo='" & MBox2(1) & "' and stalma='" & Ctr_Ayuda3.xclave & "'"
           Else
             nsql = "select *,(stskdis-stskcom) as stskdis from [" & VGCNx.DefaultDatabase & "].dbo.maeart " & _
                    "inner join [" & _
                    VGCNx.DefaultDatabase & "].dbo.stkart " & _
                    " ON acodigo=stcodigo " & _
                   " where acodigo='" & MBox2(1) & "' and stalma='" & Ctr_Ayuda3.xclave & "'"
          End If
        End If
    End If
    Set rabusca = VGCNx.Execute(nsql)
    If rabusca.RecordCount > 0 Then
      If Val(rabusca.Fields("stskdis")) < Val(MBox2(0).ClipText) Then
        If modoventa.ctrlinventario = "1" Then
            nvalor = Abs(IIf(Val(rabusca.Fields("stskdis")) = 0, 1, Val(rabusca.Fields("stskdis"))))
            If Trim(MBox2(1).ClipText) = "000" Then
                MBox2(1) = Trim(MBox2(1))
            End If
            If (Abs(Val(rabusca.Fields("stskdis")) - Val(MBox2(0).ClipText)) / nvalor) > 0.00025 And Trim(MBox2(1).ClipText) <> "000" Then
              MsgBox "La cantidad disponible es ==>" & numero(rabusca.Fields("stskdis")) & "...Verifique!!!", vbInformation, "AVISO"
              rabusca.Close
              Set rabusca = Nothing
 '             Exit Sub
            End If
         End If
      End If
      Label2 = Escadena(rabusca!adescri)
      MBox2(2) = Escadena(rabusca!aunidad)
      If rabusca!acodmon = "01" Then
         mone = g_TipoSol
      ElseIf rabusca!acodmon = "92" Then
         mone = g_TipoDolar
      Else
         mone = rabusca!acodmon
      End If
      If mone <> dllgeneral.ComboDato(Combo1.Text) Then
         If dllgeneral.ComboDato(Combo1.Text) = g_TipoSol Then
            nprecio = TraePrecio(Combo2.Text, MBox2(1).Text, VGCNx, Trim(Ctr_Ayuda3.xclave))
            If nprecio > 0 Then
               MBox2(3) = numero(nprecio * CDbl(MBox(8)))
               If modoventa.unidadmedida = "R" And modoventa.usafactor = "1" Then
                  MBox2(0) = numero(0)  'rabusca!unidadfactorconv)
                  MBox2(13) = numero(0) 'rabusca!unidadfactorconv)
               ElseIf modoventa.unidadmedida = "V" And modoventa.usafactor = "1" Then
                  MBox2(13) = numero(0)  'rabusca!unidadfactorconv)
               Else
                  MBox2(13) = 1
               End If
               MBox2(14) = numero(0) 'rabusca!productoporcimpto)
            Else
               MBox2(3) = numero(TraePrecio(Combo2.Text, MBox2(1).Text, VGCNx, Trim(Ctr_Ayuda3.xclave))) 'rabusca!productoprecvta)
               If modoventa.unidadmedida = "R" And modoventa.usafactor = "1" Then
                  MBox2(0) = numero(0)  'rabusca!unidadfactorconv)
                  MBox2(13) = numero(0) 'rabusca!unidadfactorconv)
               ElseIf modoventa.unidadmedida = "V" And modoventa.usafactor = "1" Then
                  MBox2(13) = numero(0)  'rabusca!unidadfactorconv)
               Else
                  MBox2(13) = 1
               End If
               MBox2(14) = numero(0) 'rabusca!productoporcimpto)
            End If
         ElseIf dllgeneral.ComboDato(Combo1.Text) = g_TipoDolar Then
            nprecio = TraePrecio(Combo2.Text, MBox2(1).Text, VGCNx, Trim(Ctr_Ayuda3.xclave))
            If nprecio > 0 Then
               MBox2(3) = numero(nprecio / CDbl(MBox(8)))
               If modoventa.unidadmedida = "R" And modoventa.usafactor = "1" Then
                  MBox2(0) = numero(0)   'rabusca!unidadfactorconv)
                  MBox2(13) = numero(0)  'rabusca!unidadfactorconv)
               ElseIf modoventa.unidadmedida = "V" And modoventa.usafactor = "1" Then
                  MBox2(13) = numero(0)  'rabusca!unidadfactorconv)
               Else
                  MBox2(13) = 1
               End If
               MBox2(14) = numero(0)     'rabusca!productoporcimpto)
            Else
               MBox2(3) = numero(TraePrecio(Combo2.Text, MBox2(1).Text, VGCNx, Trim(Ctr_Ayuda3.xclave))) 'rabusca!productoprecvta)
               If modoventa.unidadmedida = "R" And modoventa.usafactor = "1" Then
                  MBox2(0) = numero(0)   'rabusca!unidadfactorconv)
                  MBox2(13) = numero(0)  'rabusca!unidadfactorconv)
               ElseIf modoventa.unidadmedida = "V" And modoventa.usafactor = "1" Then
                  MBox2(13) = numero(0)  'rabusca!unidadfactorconv)
               Else
                  MBox2(13) = 1
               End If
               MBox2(14) = numero(0)   'rabusca!productoporcimpto)
            End If
         End If
      Else
         MBox2(3).Text = numero(TraePrecio(Combo2.Text, MBox2(1).Text, VGCNx, Trim(Ctr_Ayuda3.xclave))) 'rabusca!productoprecvta)
         MBox2(3).Tag = numero(TraePrecio(Combo2.Text, MBox2(1).Text, VGCNx, Trim(Ctr_Ayuda3.xclave))) 'rabusca!productoprecvta)
         If modoventa.unidadmedida = "R" And modoventa.usafactor = "1" Then
            MBox2(0) = numero(0)    'rabusca!unidadfactorconv)
            MBox2(13) = numero(0)   'rabusca!unidadfactorconv)
         ElseIf modoventa.unidadmedida = "V" And modoventa.usafactor = "1" Then
            MBox2(13) = numero(0)   'rabusca!unidadfactorconv)
         Else
            MBox2(13) = 1
         End If
         MBox2(14) = numero(0)   'rabusca!productoporcimpto)
         
      End If
    End If
    MBox2(4) = numero(0)
    MBox2(5) = numero(0)
    rabusca.Close
    Set rabusca = Nothing
End Sub

Public Function DocImprimir()
Dim rf As New ADODB.Command
Dim rb As New ADODB.Recordset
Dim puntero, nguia As String
Dim cuenta As Double
Dim ntabla As String
Dim ntabla1 As String
Dim J As Integer
Dim busca As New dll_apisgen.dll_apis

SQL = "select a.*,b.* from tempfile a inner join maeart B on a.productocodigo=b.acodigo where iSNULl(b.estadodetraccion,0)=1"
Set rb = VGCNx.Execute(SQL)
Detraccion = 0
If rb.RecordCount > 0 Then Detraccion = 1
   
If (cOpc3(1).Value Or cOpc(1).Value) And cOpc2(0).Value Then
   Call imprimirfacturas
ElseIf (cOpc3(1).Value Or cOpc(1).Value) And cOpc2(1).Value Then
    Call ImprimirBoleta
ElseIf (cOpc3(1).Value Or cOpc(1).Value) And cOpc2(3).Value Then
    Call ImprimirTicket
End If

End Function

Sub ImprimirTicket()
Dim Param(5) As Variant
Dim formulas(2) As Variant

Param(0) = VGParamSistem.BDEmpresa
Param(1) = MBox(1)
Param(2) = VGParametros.empresacodigo
Param(3) = VGParametros.puntovta
Param(4) = Left(Combo1.Text, 2)

Call ImpresionRptProc("Ticket.rpt", formulas, Param, , "Impresion de Ticket")

End Sub

Private Sub imprimirfacturas()
Dim formulas(2) As Variant
Dim Param(5) As Variant
Dim reporte As String

Param(0) = VGParamSistem.BDEmpresa
Param(1) = MBox(1).Text
Param(2) = VGParametros.empresacodigo
Param(3) = VGParametros.puntovta
Param(4) = Left(Combo1.Text, 2)

formulas(0) = "@ruc='" & VGParametros.RucEmpresa & "'"
formulas(1) = "Montoletras='" & dllgeneral.NUMLET(Round(CDbl(MBox2(10)), 2)) & IIf(dllgeneral.ComboDato(Combo1.Text) = g_TipoSol, "Nuevos Soles", "Dolares Americanos") & "'"

If VGParametros.multifacturas Then
   reporte = "vt_factuimpresa_" & VGCNx.DefaultDatabase & ".rpt"
Else
   reporte = "vt_factuimpresa_" & VGCNx.DefaultDatabase & ".rpt"
End If

Call ImpresionRptProc(reporte, formulas, Param, "", "impresion de facturas")

End Sub
Private Sub impresion_pedido()
Dim contador As Double
Dim rb As New ADODB.Recordset
Dim busca As New dll_apisgen.dll_apis
Dim nguia As String
Dim SQL As String
Dim numguias As Integer
Dim k As Integer
Dim KK As Integer
nguia = "xx"
VGCNx.Execute "delete from gtempfilep2filas"
Set rb = VGCNx.Execute("select * from gtempfile inner join maeart on productocodigo=acodigo order by alinea,agrupo,productocodigo ")
If rb.RecordCount > 0 Then
   rb.MoveFirst
   If rb.RecordCount Mod 100 > 0 Then
      numguias = Int(rb.RecordCount / 100) + 1
    Else
      numguias = Int(rb.RecordCount / 100)
   End If
   contador = 0
   rb.MoveFirst
  Do While contador < numguias
       contador = contador + 1
       If contador * 100 > rb.RecordCount Then
            KK = rb.RecordCount - (contador - 1) * 100
        Else
           KK = 100
       End If
       For k = 1 To KK
           If k <= 50 Then
             TCant = (contador - 1) * 50 + k
              SQL = "INSERT INTO gtempfilep2filas(item,producto1,descripcion1,cantidad1,importe1,"
              SQL = SQL & "cantidad2,importe2)  "
              SQL = SQL & " VALUES ( '" & TCant & "','" & RTrim(rb!productocodigo) & "','" & RTrim(rb!productodescripcion) & "','" & rb!detpedcantpedida & "','" & rb!detpedimpbruto & "',"
              SQL = SQL & "0,0 )"
            Else
             TCant = (contador - 1) * 50 + k - 50
              SQL = "UPDATE gtempfilep2filas set producto2 ='" & RTrim(rb!productocodigo) & "',"
              SQL = SQL & " descripcion2='" & RTrim(rb!productodescripcion) & "',"
              SQL = SQL & "cantidad2='" & rb!detpedcantpedida & "',"
              SQL = SQL & "importe2= '" & rb!detpedimpbruto & "'"
              SQL = SQL & " where item = " & TCant & ""
           End If
           VGCNx.Execute SQL
           rb.MoveNext
        Next k
   Loop
   rb.Close
   Set rb = Nothing
End If

oCrystalReport.Reset
oCrystalReport.ReportFileName = VGParamSistem.Rutareport & "vt_pedido.rpt"
oCrystalReport.LogOnServer "pdssql.dll", _
 busca.LeerIni(App.Path & "\MARFICE.INI", "conexion", "SERVIDOR", ""), _
 busca.LeerIni(App.Path & "\MARFICE.INI", "conexion", "BDDATOS", ""), _
 busca.LeerIni(App.Path & "\MARFICE.INI", "conexion", "USUARIO", ""), _
 busca.LeerIni(App.Path & "\MARFICE.INI", "conexion", "PASSW", "")
oCrystalReport.Connect = _
 "DSN=" & busca.LeerIni(App.Path & "\MARFICE.INI", "conexion", "SERVIDOR", "") & ";" & _
 "DSQ=" & busca.LeerIni(App.Path & "\MARFICE.INI", "conexion", "BDDATOS", "") & ";" & _
 "UID=" & busca.LeerIni(App.Path & "\MARFICE.INI", "conexion", "USUARIO", "") & ";" & _
 "PWD=" & busca.LeerIni(App.Path & "\MARFICE.INI", "conexion", "PASSW", "")
                                

 oCrystalReport.Destination = crptToWindow
 oCrystalReport.WindowState = crptMaximized
 oCrystalReport.DiscardSavedData = True
 With oCrystalReport
        .WindowShowPrintSetupBtn = True
        .WindowShowExportBtn = True
        .WindowShowZoomCtl = True
        .WindowShowNavigationCtls = True
        .WindowShowPrintBtn = True
      
      .formulas(0) = "nro='" & MBox(2) & "'"
      .formulas(1) = "cliente='" & MBox3(1) & "'"
      .formulas(2) = "fecha='" & CDate(MBox(10)) & "' "
      .formulas(3) = "direccion='" & MBox3(3) & "'"
      .formulas(4) = "dni='" & MBox3(2) & "'"
      .formulas(5) = "opedido='" & MBox(1) & "'"
      .formulas(6) = "ocompra='" & Ctr_AyuRef.xclave & "'"  'MBox(17)
      .formulas(7) = "guia='" & nguia & "'"
      .formulas(8) = "distrito='" & MBox3(4).ClipText & "'"
      .formulas(9) = "destino='" & MBox(19).ClipText & "'"
      Set rb = VGCNx.Execute("select * from gr_empresa where empresacodigo='" & VGParametros.empresacodigo & "'")
      If rb.RecordCount > 0 Then
        .formulas(10) = "partida='" & Escadena(rb!empresadireccion) & "'"
       Else
         .formulas(10) = "partida=''"
      End If
      .formulas(11) = "moneda='" & Combo1.Text & "'"
      .formulas(12) = "cpago='" & Escadena(Combo4) & "'"
      .formulas(13) = "vendedor='" & Escadena(Ctr_Ayuda2.xnombre) & "'"
      rb.Close
      Set rb = Nothing
 End With
 oCrystalReport.Action = 1
 
End Sub
  

Public Sub CargarModo()
     Dim rs As New ADODB.Recordset
     Dim ncade As String
     Dim J As Integer

     On Error Resume Next
     Set rs = VGCNx.Execute("select * from vt_modoventa where modovtacodigo='" & dllgeneral.ComboDato(Combo3.Text) & "'")
     If rs.RecordCount > 0 Then
        modoventa.descuento = Escadena(rs!modovtadscto)
        modoventa.impuestos = Escadena(IIf(IsNull(rs!modovtaimpuestos) Or rs!modovtaimpuestos = 0, "0", "1"))
        modoventa.nroitem = IIf(IsNull(rs!modovtaitemxdoc), 10, rs!modovtaitemxdoc)
        modoventa.copiashoja = IIf(IsNull(rs!modovtacopiashojatrab), 1, rs!modovtacopiashojatrab)
        modoventa.copiasbol = IIf(IsNull(rs!modovtacopiasboleta), 1, rs!modovtacopiasboleta)
        modoventa.copiasfac = IIf(IsNull(rs!modovtacopiasfact), 1, rs!modovtacopiasfact)
        modoventa.copiastic = IIf(IsNull(rs!modovtacopiasticket), 1, rs!modovtacopiasticket)
        modoventa.ctacte = Escadena(IIf(IsNull(rs!modovtaactctacte) Or rs!modovtaactctacte = 0, "0", "1"))
        modoventa.ctrlinventario = Escadena(IIf(IsNull(rs!modovtactrlinventario) Or rs!modovtactrlinventario = 0, "0", "1"))
        modoventa.emitehoja = Escadena(IIf(IsNull(rs!modovtaemitehoja) Or rs!modovtaemitehoja = 0, "0", "1"))
        modoventa.emitefact = Escadena(IIf(IsNull(rs!modovtasolemitfact) Or rs!modovtasolemitfact = 0, "0", "1"))
        modoventa.emiteguia = Escadena(IIf(IsNull(rs!modovtaemiteguia) Or rs!modovtaemiteguia = 0, "0", "1"))
        modoventa.ingcliente = Escadena(IIf(IsNull(rs!modovtaingcodclie) Or rs!modovtaingcodclie = 0, "0", "1"))
        modoventa.ingforma = Escadena(IIf(IsNull(rs!modovtaingformapag) Or rs!modovtaingformapag = 0, "0", "1"))
        modoventa.ingguia = Escadena(IIf(IsNull(rs!modovtaingguiarem) Or rs!modovtaingguiarem = 0, "0", "1"))
        modoventa.inghoja = Escadena(IIf(IsNull(rs!modovtainghojatrab) Or rs!modovtainghojatrab = 0, "0", "1"))
        modoventa.ingpedido = Escadena(IIf(IsNull(rs!modovtaingpedido) Or rs!modovtaingpedido = 0, "0", "1"))
        modoventa.modificaguia = Escadena(IIf(IsNull(rs!modovtacorrguiarem) Or rs!modovtacorrguiarem = 0, "0", "1"))
        modoventa.unidadmedida = Escadena(IIf(IsNull(rs!modovtaunidadmedida) Or rs!modovtaunidadmedida = "V", "V", Escadena(rs!modovtaunidadmedida)))
        modoventa.unidadmedida = Left(modoventa.unidadmedida, 1)
        modoventa.usafactor = Escadena(IIf(IsNull(rs!modovtausafactconv) Or rs!modovtausafactconv = 0, "0", "1"))
        If Not IsNull(rs!modovtaalmacen) Then
           ncade = "'"
           For J = 1 To Len(Trim(rs!modovtaalmacen))
             If Mid(Trim(rs!modovtaalmacen), J, 1) <> "," Then
                 ncade = ncade & Mid(Trim(rs!modovtaalmacen), J, 1)
             Else
                 ncade = ncade & "','"
             End If
           Next J
           ncade = ncade & "'"
           modoventa.almacenes = ncade
        Else
           modoventa.almacenes = ""
        End If
        
        Text3.Visible = IIf(modoventa.usafactor = "1", True, False)
        MBox(1).Enabled = IIf(modoventa.documento = g_tipoped And modoventa.numeraauto <> "1" And modoventa.ingpedido = "1", True, False) 'Modo de pedido
        MBox(2).Enabled = IIf(modoventa.documento = g_tipofac And modoventa.numeraauto <> "1", True, False) 'Modo de factura
        MBox(3).Enabled = IIf(modoventa.documento = g_tipobol And modoventa.numeraauto <> "1", True, False) 'Modo de boleta
        MBox(4).Enabled = IIf(modoventa.documento = g_tipoguia And modoventa.numeraauto <> "1" And modoventa.ingguia = "1", True, False)  'Modo de Modifica
        
        modoventa.numeraauto = Escadena(IIf(IsNull(rs!modovtanumautom) Or rs!modovtanumautom = 0, "0", "1"))
        modoventa.documento = Escadena(IIf(IsNull(rs!documentocodigo), "", rs!documentocodigo))
        
        MBox2(0).Enabled = IIf(modoventa.usafactor = 0 Or (modoventa.usafactor = "1" And modoventa.unidadmedida = "V"), True, False)
     '   MBox2(12).Enabled = IIf(modoventa.usafactor = 0 Or (modoventa.usafactor = "1" And modoventa.unidadmedida = "R"), True, False)
     End If
     rs.Close
     Set rs = Nothing

End Sub

Private Sub TDBGrid2_DblClick()
cmdBotones_Click (1)
End Sub


Private Sub TDBGrid3_Click()
   If rsmasivo.RecordCount > 0 Then
      TDBGrid3.SetFocus
   End If
End Sub

Private Sub TDBGrid3_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim nvalor As Variant
  On Error Resume Next
  nvalor = KeyCode
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
   Text10 = Text10 - Val(rsmasivo!cantidad) + Val(TDBGrid3.Columns(6))
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
    rsmasivo.MoveNext
  End If
End If
End Sub


Private Sub Text3_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     SendKeys "{tab}"
  End If
End Sub

Private Sub Text4_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Index Like "[12]" Then
            Text4(Index) = Right("0000000000" & Trim(Text4(Index)), Text4(Index).MaxLength)
            If Index = 2 And Len(Trim(Text4(0))) > 0 Then
                If dllgeneral.VerificaDatoExistente(VGCNx, "select * from vt_pedido where pedidotipofac='" & Text4(0) & "' and pedidonrofact='" & Trim(Text4(1)) & Trim(Text4(2)) & "'") = 0 Then
                    MsgBox "No existe documento...Verifique!!!", vbInformation, "AVISO"
                    Exit Sub
                End If
            ElseIf Chkmasivo = 1 Then
              Call loadmasivo
              
              Call dllgeneral.ActivaTab(2, 2, SSTab1)
              TDBGrid3.SetFocus
            Else
              MBox2(0).SetFocus
            End If
        End If
        If Chkmasivo = 0 Then
            SendKeys "{tab}"
        ElseIf Index Like "[01]" Then
            SendKeys "{tab}"
        End If
    End If

End Sub

Private Sub TxtCliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub


Private Sub TxtHor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub


Private Sub TxtNro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub


