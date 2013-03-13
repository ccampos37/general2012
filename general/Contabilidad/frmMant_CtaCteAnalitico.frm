VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form frmMant_CtaCteAnalitico 
   Caption         =   "Mantenimiento de Cta Cte Analiticos - Apertura"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   11265
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4800
      Left            =   105
      TabIndex        =   0
      Top             =   420
      Width           =   11010
      _ExtentX        =   19420
      _ExtentY        =   8467
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Consulta"
      TabPicture(0)   =   "frmMant_CtaCteAnalitico.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "TDBG_CtaCte"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "framebotones"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Mantenimiento"
      TabPicture(1)   =   "frmMant_CtaCteAnalitico.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FramDetalle"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cCancela"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cAcepta"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.CommandButton cAcepta 
         Caption         =   "&Aceptar"
         Height          =   345
         Left            =   -70837
         TabIndex        =   42
         Top             =   3825
         Width           =   1290
      End
      Begin VB.CommandButton cCancela 
         Caption         =   "&Cancelar"
         Height          =   345
         Left            =   -69292
         TabIndex        =   41
         Top             =   3825
         Width           =   1290
      End
      Begin VB.Frame framebotones 
         Height          =   555
         Left            =   2640
         TabIndex        =   35
         Top             =   4125
         Width           =   5730
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Nuevo"
            Height          =   330
            Index           =   0
            Left            =   105
            TabIndex        =   40
            Top             =   165
            Width           =   1080
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "E&ditar"
            Height          =   330
            Index           =   1
            Left            =   1185
            TabIndex        =   39
            Top             =   165
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Eliminar"
            Height          =   330
            Index           =   2
            Left            =   2310
            TabIndex        =   38
            Top             =   165
            Width           =   1080
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Salir"
            Height          =   330
            Index           =   4
            Left            =   4560
            TabIndex        =   37
            Top             =   165
            Width           =   1080
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Imprimir"
            Height          =   330
            Index           =   3
            Left            =   3435
            TabIndex        =   36
            Top             =   165
            Visible         =   0   'False
            Width           =   1080
         End
      End
      Begin VB.Frame FramDetalle 
         Height          =   2460
         Left            =   -74805
         TabIndex        =   4
         Top             =   1215
         Width           =   10335
         Begin VB.ComboBox CmbID 
            Height          =   315
            ItemData        =   "frmMant_CtaCteAnalitico.frx":0038
            Left            =   8760
            List            =   "frmMant_CtaCteAnalitico.frx":0042
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   180
            Width           =   1440
         End
         Begin VB.ComboBox CmbTcambio 
            Height          =   315
            ItemData        =   "frmMant_CtaCteAnalitico.frx":005B
            Left            =   8775
            List            =   "frmMant_CtaCteAnalitico.frx":0068
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   855
            Width           =   1440
         End
         Begin TextFer.TxFer TxMonto 
            Height          =   315
            Left            =   8760
            TabIndex        =   7
            Top             =   1560
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   556
            Alignment       =   1
            BackColor       =   16384
            Object.CausesValidation=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   12648447
            MaxLength       =   15
            Text            =   "0.00"
            ColorIlumina    =   12648447
            SaltarAlEnter   =   -1  'True
            Valor           =   "0.00"
            TipoDato        =   1
            SignodeMiles    =   -1  'True
            NumeroDecimales =   3
            SignoNegativo   =   0   'False
            Formato         =   "###,###,###,###.00"
            MarcarTextoAlEnfoque=   -1  'True
            ColorTextoAlEnfocar=   16711680
         End
         Begin TextFer.TxFer TxNdoc 
            Height          =   300
            Left            =   4575
            TabIndex        =   8
            Top             =   1350
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   529
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
            MaxLength       =   10
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
         End
         Begin TextFer.TxFer TxSerie 
            Height          =   300
            Left            =   3900
            TabIndex        =   9
            Top             =   1350
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   529
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
            MaxLength       =   4
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
         End
         Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_Moneda 
            Height          =   315
            Left            =   8775
            TabIndex        =   10
            Top             =   510
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            XcodMaxLongitud =   2
            NomTabla        =   "gr_moneda"
            TituloAyuda     =   "Busqueda de Moneda"
            ListaCampos     =   "monedacodigo(1),monedadescripcion(1)"
            XcodCampo       =   "monedacodigo"
            XListCampo      =   "monedadescripcion"
            ListaCamposDescrip=   "Código,Descripción"
            ListaCamposText =   "monedacodigo,monedadescripcion"
         End
         Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_TipDoc 
            Height          =   315
            Left            =   1125
            TabIndex        =   11
            Top             =   1335
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   556
            XcodMaxLongitud =   2
            NomTabla        =   "gr_documento"
            TituloAyuda     =   "Busqueda de Tipo de  Documento"
            ListaCampos     =   "documentocodigo(1),documentodescripcion(1),documentonotacredito(2)"
            XcodCampo       =   "documentocodigo"
            XListCampo      =   "documentodescripcion"
            ListaCamposDescrip=   "Código,Descripción"
            ListaCamposText =   "documentocodigo,documentodescripcion,documentonotacredito"
         End
         Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_TipAnal 
            Height          =   420
            Left            =   5175
            TabIndex        =   12
            Top             =   945
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   741
            XcodMaxLongitud =   3
            xcodwith        =   100
            NomTabla        =   "ct_tipoanalitico"
            TituloAyuda     =   "Busqueda de Tipo de Analitico"
            ListaCampos     =   "tipoanaliticocodigo(1),tipoanaliticodescripcion(1)"
            XcodCampo       =   "tipoanaliticocodigo"
            XListCampo      =   "tipoanaliticodescripcion"
            ListaCamposDescrip=   "Código,Descripción"
            ListaCamposText =   "tipoanaliticocodigo,tipoanaliticodescripcion"
         End
         Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_Opera 
            Height          =   315
            Left            =   1125
            TabIndex        =   13
            Top             =   525
            Width           =   6465
            _ExtentX        =   11404
            _ExtentY        =   556
            XcodMaxLongitud =   2
            NomTabla        =   "ct_operacion"
            TituloAyuda     =   "Busqueda de Operacion"
            ListaCampos     =   "operacioncodigo(1),operaciondescripcion(1)"
            XcodCampo       =   "operacioncodigo"
            XListCampo      =   "operaciondescripcion"
            ListaCamposDescrip=   "Código,Descripción"
            ListaCamposText =   "operacioncodigo,operaciondescripcion"
         End
         Begin MSComCtl2.DTPicker Dtp_FechaDoc 
            Height          =   315
            Left            =   1140
            TabIndex        =   14
            Top             =   1680
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            Format          =   51838977
            CurrentDate     =   37469
         End
         Begin MSComCtl2.DTPicker DtpFech_Ven 
            Height          =   315
            Left            =   3795
            TabIndex        =   15
            Top             =   1695
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   51838977
            CurrentDate     =   37469
         End
         Begin TextFer.TxFer TxGlosa 
            Height          =   300
            Left            =   1110
            TabIndex        =   16
            Top             =   2040
            Width           =   6435
            _ExtentX        =   11351
            _ExtentY        =   529
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
            MaxLength       =   50
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
         End
         Begin ctrlayuda_f.Ctr_Ayuda CtrAy_Asiento 
            Height          =   405
            Left            =   1125
            TabIndex        =   33
            Top             =   945
            Width           =   2880
            _ExtentX        =   5080
            _ExtentY        =   714
            XcodMaxLongitud =   3
            xcodwith        =   500
            NomTabla        =   "ct_asiento"
            ListaCampos     =   "asientocodigo(1),asientodescripcion(1)"
            XcodCampo       =   "asientocodigo"
            XListCampo      =   "asientodescripcion"
            ListaCamposDescrip=   "Código,Descripción"
            ListaCamposText =   "asientocodigo,asientodescripcion"
            Requerido       =   0   'False
         End
         Begin VB.Label Label3 
            Caption         =   "Nro. Comp."
            Height          =   315
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lb_vcambio 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00F2FEFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   8775
            TabIndex        =   30
            Top             =   1230
            Width           =   1440
         End
         Begin VB.Label lbTipAnal 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Analitico :"
            Height          =   405
            Index           =   0
            Left            =   4020
            TabIndex        =   29
            Top             =   1020
            Width           =   1050
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Cod Oper. :"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   630
            Width           =   810
         End
         Begin VB.Label lbccosto 
            AutoSize        =   -1  'True
            Caption         =   "Asiento"
            Height          =   405
            Left            =   120
            TabIndex        =   27
            Top             =   1005
            Width           =   525
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H00808080&
            Height          =   2385
            Left            =   7680
            Top             =   225
            Width           =   15
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00FFFFFF&
            Height          =   2385
            Left            =   7695
            Top             =   225
            Width           =   15
         End
         Begin VB.Label lbtipdoc 
            Caption         =   "Tipo doc. :"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   1380
            Width           =   1020
         End
         Begin VB.Label lbndocum 
            AutoSize        =   -1  'True
            Caption         =   "Nº doc. :"
            Height          =   195
            Left            =   3240
            TabIndex        =   25
            Top             =   1395
            Width           =   630
         End
         Begin VB.Label lbFechaDoc 
            AutoSize        =   -1  'True
            Caption         =   "Fecha doc. :"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   1770
            Width           =   900
         End
         Begin VB.Label lbFechVen 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Venc. :"
            Height          =   195
            Left            =   2505
            TabIndex        =   23
            Top             =   1755
            Width           =   1230
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Indicador :"
            Height          =   195
            Left            =   7770
            TabIndex        =   22
            Top             =   225
            Width           =   750
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "T/Cambio :"
            Height          =   195
            Left            =   7770
            TabIndex        =   21
            Top             =   930
            Width           =   795
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "V/Cambio :"
            Height          =   195
            Left            =   7800
            TabIndex        =   20
            Top             =   1275
            Width           =   795
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Monto :"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   7785
            TabIndex        =   19
            Top             =   1605
            Width           =   540
         End
         Begin VB.Label Label23 
            Caption         =   "Glosa :"
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   2100
            Width           =   495
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Moneda :"
            Height          =   195
            Left            =   7800
            TabIndex        =   17
            Top             =   615
            Width           =   675
         End
         Begin VB.Label lblNroComprobante 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1140
            TabIndex        =   43
            Top             =   165
            Width           =   2730
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBG_CtaCte 
         Height          =   3660
         Left            =   105
         TabIndex        =   1
         Top             =   495
         Width           =   10845
         _ExtentX        =   19129
         _ExtentY        =   6456
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Asiento"
         Columns(0).DataField=   "asientocodigo"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Sub_Asiento"
         Columns(1).DataField=   "subasientocodigo"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "T.Doc"
         Columns(2).DataField=   "documentocodigo"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Num.Doc."
         Columns(3).DataField=   "ctacteanaliticonumdocumento"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "F.Doc"
         Columns(4).DataField=   "ctacteanaliticofechadoc"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Moneda"
         Columns(5).DataField=   "monedacodigo"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Debe"
         Columns(6).DataField=   "ctacteanaliticodebe"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "haber"
         Columns(7).DataField=   "ctacteanaliticohaber"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Glosa"
         Columns(8).DataField=   "ctacteanaliticoglosa"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   9
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=9"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1191"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1111"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=1773"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1693"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=1058"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=979"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=1588"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=1508"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=1535"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=1455"
         Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(21)=   "Column(5).Width=1244"
         Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=1164"
         Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(25)=   "Column(6).Width=1588"
         Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=1508"
         Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(29)=   "Column(7).Width=1931"
         Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=1852"
         Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(33)=   "Column(8).Width=5371"
         Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=5292"
         Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
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
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=66,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=54,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=58,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=62,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
         _StyleDefs(72)  =   "Named:id=33:Normal"
         _StyleDefs(73)  =   ":id=33,.parent=0"
         _StyleDefs(74)  =   "Named:id=34:Heading"
         _StyleDefs(75)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(76)  =   ":id=34,.wraptext=-1"
         _StyleDefs(77)  =   "Named:id=35:Footing"
         _StyleDefs(78)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(79)  =   "Named:id=36:Selected"
         _StyleDefs(80)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(81)  =   "Named:id=37:Caption"
         _StyleDefs(82)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(83)  =   "Named:id=38:HighlightRow"
         _StyleDefs(84)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(85)  =   "Named:id=39:EvenRow"
         _StyleDefs(86)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(87)  =   "Named:id=40:OddRow"
         _StyleDefs(88)  =   ":id=40,.parent=33"
         _StyleDefs(89)  =   "Named:id=41:RecordSelector"
         _StyleDefs(90)  =   ":id=41,.parent=34,.bgcolor=&HFFFF00&"
         _StyleDefs(91)  =   "Named:id=42:FilterBar"
         _StyleDefs(92)  =   ":id=42,.parent=33"
      End
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_CtaCtble 
      Height          =   345
      Left            =   1590
      TabIndex        =   3
      Top             =   60
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   609
      XcodMaxLongitud =   20
      xcodwith        =   800
      NomTabla        =   "ct_cuenta"
      ListaCampos     =   "cuentacodigo(1),cuentadescripcion(1)"
      XcodCampo       =   "cuentacodigo"
      XListCampo      =   "cuentadescripcion"
      ListaCamposDescrip=   "Cuenta,Descripción"
      ListaCamposText =   "cuentacodigo,cuentadescripcion"
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_Analitico 
      Height          =   300
      Left            =   7035
      TabIndex        =   31
      Top             =   60
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   529
      XcodMaxLongitud =   17
      xcodwith        =   1100
      NomTabla        =   "v_analiticoentidad"
      TituloAyuda     =   "Busqueda de Analitico"
      ListaCampos     =   "analiticocodigo(1),entidadrazonsocial(1),entidadruc(1)"
      XcodCampo       =   "analiticocodigo"
      XListCampo      =   "entidadrazonsocial"
      ListaCamposDescrip=   "Codigo,Descripcion,Ruc"
      ListaCamposText =   "analiticocodigo,entidadrazonsocial,entidadruc"
   End
   Begin VB.Label Label2 
      Caption         =   "Analitico"
      Height          =   270
      Left            =   5745
      TabIndex        =   32
      Top             =   120
      Width           =   750
   End
   Begin VB.Label Label1 
      Caption         =   "Cuenta Contable"
      Height          =   270
      Left            =   270
      TabIndex        =   2
      Top             =   135
      Width           =   1470
   End
End
Attribute VB_Name = "frmMant_CtaCteAnalitico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim modoinsert As Boolean
Dim modoedit As Boolean
Dim ClsMM1 As ClsMantMov1
Dim rs As New ADODB.Recordset

Private Sub Ctr_Analitico_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Dim dato As String
If Ctr_CtaCtble.xclave <> "" Then
   dato = "empresacodigo='" & VGParametros.empresacodigo & "' and cuentacodigo='" & Ctr_CtaCtble.xclave & "'and year(ctacteanaliticofechaconta) < " & VGParamSistem.Anoproceso & ""
   dato = dato & " and analiticocodigo='" & Ctr_Analitico.xclave & "'"
   Mostrar (dato)
End If
End Sub

Private Sub Ctr_CtaCtble_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Dim dato As String
dato = "empresacodigo='" & VGParametros.empresacodigo & "' and cuentacodigo='" & Ctr_CtaCtble.xclave & "'and year(ctacteanaliticofechaconta) < " & VGParamSistem.Anoproceso & ""

Mostrar (dato)

End Sub

Private Sub Form_Load()
  Ctr_CtaCtble.conexion VGCNx
  Ctr_CtaCtble.Filtro = " cuentanivel=" & VGnumnivelescuenta & " and cuentaestadoanalitico=1 and empresacodigo='" & VGParametros.empresacodigo & "'"
  Ctr_Analitico.conexion VGCNx
  CtrAy_Asiento.conexion VGCNx
  CtrAyu_Moneda.conexion VGCNx
  CtrAyu_Opera.conexion VGCNx
  CtrAyu_TipAnal.conexion VGCNx
  CtrAyu_TipDoc.conexion VGCNx
  SSTab1.Tab = 0
  SSTab1.TabEnabled(1) = False
  framebotones.Visible = False
End Sub

Private Sub cmdBotones_Click(Index As Integer)
  Select Case Index
    Case 0:
      Call Limpiar
      modoinsert = True
      SSTab1.TabEnabled(1) = True
      SSTab1.Tab = 1
      
    Case 1:
      If TDBG_CtaCte.Row < 0 Then
         Exit Sub
      End If
      modoedit = True
      SSTab1.TabEnabled(1) = True
      SSTab1.Tab = 1
      Call Edicion
      
    Case 2:
      Call eliminar
    
    Case 3:
      'Impresion
    
    Case 4: Unload Me
  
  End Select
  
End Sub
Sub eliminar()
Dim rssql As New ADODB.Recordset
SQL = "delete from  dbo.ct_ctacteanalitico" & VGParamSistem.Anoproceso & " where "
SQL = SQL & " empresacodigo='" & VGParametros.empresacodigo & "'"
SQL = SQL & " and cuentacodigo+asientocodigo+subasientocodigo='" & rs!cuentacodigo & rs!asientocodigo & rs!subasientocodigo & "'"
SQL = SQL & " and analiticocodigo='" & rs!analiticocodigo & "'"
SQL = SQL & " and documentocodigo+ctacteanaliticonumdocumento='" & rs!documentocodigo & rs!ctacteanaliticonumdocumento & "'"
SQL = SQL & " and cabcomprobmes=0 "
Set rssql = VGCNx.Execute(SQL)
Ctr_Analitico.Ejecutar
End Sub
Sub Limpiar()
  CtrAy_Asiento.xclave = Empty: CtrAy_Asiento.Ejecutar
  CtrAyu_Moneda.xclave = Empty: CtrAyu_Moneda.Ejecutar
  CtrAyu_Opera.xclave = Empty: CtrAyu_Opera.Ejecutar
  CtrAyu_TipAnal.xclave = Empty: CtrAyu_TipAnal.Ejecutar
  CtrAyu_TipDoc.xclave = Empty: CtrAyu_TipDoc.Ejecutar
  TxGlosa.Text = Empty
  TxMonto.Text = Empty
  TxSerie.Text = Empty: TxNdoc.Text = Empty
End Sub

Sub Edicion()
 Dim i As Integer
   'cabcomprobmes , detcomprobitem, cabcomprobnumero, subasientocodigo, asientocodigo, documentocodigo, operacioncodigo, cuentacodigo, ""
   'SQL = SQL & "ctacteanaliticofechaconta, analiticocodigo, ctacteanaliticonumdocumento, ctacteanaliticofechadoc, ctacteanaliticoglosa, ctacteanaliticodebe,"
   'SQL = SQL & "ctacteanaliticoussdebe, ctacteanaliticohaber, ctacteanaliticousshaber, ctacteanaliticocancel, ctacteanaliticofechaven,monedacodigo,ctacteanaliticosaldo "
 
 If rs.RecordCount > 0 Then
    With TDBG_CtaCte
       lblNroComprobante.Caption = .Columns(2).Value
       CtrAyu_Opera.xclave = .Columns(6).Value: CtrAyu_Opera.Ejecutar
       Ctr_CtaCtble.xclave = .Columns(7).Value: Ctr_CtaCtble.Ejecutar
       CtrAy_Asiento.xclave = .Columns(4).Value: CtrAy_Asiento.Ejecutar
       CtrAyu_TipAnal.xclave = Right(Trim$(.Columns(9).Value), 3): CtrAyu_TipAnal.Ejecutar
       Ctr_Analitico.xclave = .Columns(9).Value: Ctr_Analitico.Ejecutar
       CtrAyu_TipDoc.xclave = .Columns(5).Value: CtrAyu_TipDoc.Ejecutar
       TxSerie.Text = Left(Trim$(.Columns(10).Value), 4)
       i = InStr(1, .Columns(10).Value, "-", vbTextCompare)
       TxNdoc.Text = Trim$(Mid$(.Columns(10).Value, i + 1, Len(.Columns(10).Value) - i))
       Dtp_FechaDoc.Value = Format(.Columns(11).Value, "dd/mm/yyyy")
       DtpFech_Ven.Value = Format(.Columns(18).Value, "dd/mm/yyyy")
       TxGlosa.Text = Trim$(.Columns(12).Value)
       CtrAyu_Moneda.xclave = Trim$(.Columns(19).Value): CtrAyu_Moneda.Ejecutar
       If .Columns(13).Value > 0 Then
         CmbID.ListIndex = 0
         TxMonto.Text = .Columns(13).Value
         lb_vcambio.Caption = .Columns(13).Value / .Columns(14).Value
       Else
         CmbID.ListIndex = 1
         TxMonto.Text = .Columns(15).Value / .Columns(16).Value
       End If
    End With
 End If

End Sub

Private Sub cmdVisualizar_Click()
   Dim SQL As String
  
   SQL = "Select cabcomprobmes, detcomprobitem, cabcomprobnumero, subasientocodigo, asientocodigo, documentocodigo, operacioncodigo, cuentacodigo,"
   SQL = SQL & "ctacteanaliticofechaconta, analiticocodigo, ctacteanaliticonumdocumento, ctacteanaliticofechadoc, ctacteanaliticoglosa, ctacteanaliticodebe,"
   SQL = SQL & "ctacteanaliticoussdebe, ctacteanaliticohaber, ctacteanaliticousshaber, ctacteanaliticocancel, ctacteanaliticofechaven,monedacodigo,ctacteanaliticosaldo "
   SQL = SQL & "from  dbo.ct_ctacteanalitico" & VGParamSistem.Anoproceso & " where ctacteanaliticosaldo='" & (VGParamSistem.Anoproceso - 1) & "' AND "
   SQL = SQL & "cuentacodigo like '" & Trim$(Ctr_CtaCtble.xclave) & "%' and "
   SQL = SQL & "analiticocodigo like '" & Trim$(Ctr_Analitico.xclave) & "%'"
  
   Set rs = New ADODB.Recordset
   Set rs = VGCNx.Execute(SQL)
   
   Set TDBG_CtaCte.DataSource = rs

End Sub

Sub GrabarCtacteAnalitico()
  On Error GoTo xx
    Screen.MousePointer = 11
    VGCNx.BeginTrans
    Set VGCommandoSP = New ADODB.Command
    VGCommandoSP.ActiveConnection = VGGeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "ct_IngresaCtacteAnalitico_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@tabla") = "ct_ctacteanalitico" + VGParamSistem.Anoproceso
        .Parameters("@empresa") = VGParametros.empresacodigo
        .Parameters("@op") = IIf(modoinsert = True, "1", "2")
        .Parameters("@cabcomprobmes") = 0
        .Parameters("@cabcomprobnumero") = 0
        .Parameters("@subasientocodigo") = "0001"
        .Parameters("@asientocodigo") = Trim$(CtrAy_Asiento.xclave)
        .Parameters("@detcomprobitem") = 0
        .Parameters("@analiticocodigo") = Trim$(Ctr_Analitico.xclave)
        .Parameters("@monedacodigo") = Trim$(CtrAyu_Moneda.xclave)
        .Parameters("@documentocodigo") = Trim$(CtrAyu_TipDoc.xclave)
        .Parameters("@operacioncodigo") = Trim$(CtrAyu_Opera.xclave)
        .Parameters("@cuentacodigo") = Trim$(Ctr_CtaCtble.xclave)
        .Parameters("@detcomprobnumdocumento") = Format(Trim$(TxSerie.Text), "0000") & Format(Trim$(TxNdoc.Text), "0000000000")
        .Parameters("@detcomprobfechaemision") = Dtp_FechaDoc.Value
        .Parameters("@detcomprobfechavencimiento") = DtpFech_Ven.Value
        .Parameters("@ctacteanaliticofechacontable") = Dtp_FechaDoc.Value
        .Parameters("@detcomprobglosa") = Trim$(TxGlosa.Text)
        If CmbID.ListIndex = 0 Then
           .Parameters("@detcomprobdebe") = IIf(Trim$(CtrAyu_Moneda.xclave) = "01", TxMonto.Text, TxMonto.Text * lb_vcambio.Caption)
           .Parameters("@detcomprobussdebe") = IIf(Trim$(CtrAyu_Moneda.xclave) = "02", TxMonto.Text, TxMonto.Text / lb_vcambio.Caption)
           .Parameters("@detcomprobhaber") = 0
           .Parameters("@detcomprobusshaber") = 0
        ElseIf CmbID.ListIndex > 0 Then
           .Parameters("@detcomprobhaber") = IIf(Trim$(CtrAyu_Moneda.xclave) = "01", TxMonto.Text, TxMonto.Text * lb_vcambio.Caption)
           .Parameters("@detcomprobusshaber") = IIf(Trim$(CtrAyu_Moneda.xclave) = "02", TxMonto.Text, TxMonto.Text / lb_vcambio.Caption)
           .Parameters("@detcomprobdebe") = 0
           .Parameters("@detcomprobussdebe") = 0
        End If
        .Parameters("@detcomprobtipocambio") = lb_vcambio
        .Parameters("@ctacteanaliticofechacontable") = Dtp_FechaDoc.Value
        .Execute
    End With
    
    VGCNx.CommitTrans
    Ctr_Analitico.Ejecutar
    Screen.MousePointer = 1
    MsgBox "Se Actualizó Satisfactoriamente la Cuenta Corriente de Apertura del Año " & VGParamSistem.Anoproceso, vbInformation
    Exit Sub
xx:
    Screen.MousePointer = 1
    VGCNx.RollbackTrans
    MsgBox "No se pudo Grabar la Cuenta Corriente " & Chr(13) & err.Description, vbExclamation
End Sub

Function RecuperaTipoCambio(Fecha As String, tipo As tipocambio) As Double
  Dim Rsaux As ADODB.Recordset
  Set Rsaux = New ADODB.Recordset
  Dim Campo As String
      RecuperaTipoCambio = 0
      Select Case tipo
          Case Compra
              Campo = "tipocambiocompra"
          Case Venta
              Campo = "tipocambioventa"
          Case Promedio
              Campo = "tipocambiopromedio"
      Case Else
              Campo = "tipocambioventa"
      End Select
      Rsaux.Open "Select Valor=isnull(" & Campo & ",0)  from ct_tipocambio where tipocambiofecha ='" & Fecha & "'", VGCNx, adOpenKeyset, adLockReadOnly
      If Rsaux.RecordCount > 0 Then
          RecuperaTipoCambio = Rsaux!valor
      End If
End Function

Private Sub cAcepta_Click()

  Call GrabarCtacteAnalitico
  modoedit = False
  modoinsert = False
  SSTab1.Tab = 0
  SSTab1.TabEnabled(1) = False
End Sub

Private Sub cCancela_Click()
   SSTab1.Tab = 0
End Sub

Private Sub CmbTcambio_Click()
   Set ClsMM1 = New ClsMantMov1
   VGValorCambio = RecuperaTipoCambio(Format(Dtp_FechaDoc, "dd/mm/yyyy"), CmbTcambio.ListIndex + 1)
   lb_vcambio.Caption = Format(VGValorCambio, "#.000 ")
End Sub

Private Sub Mostrar(dato As String)
Dim rsql As New ADODB.Recordset
Dim SQL As String
SQL = "Select * from ct_ctacteanalitico" & VGParamSistem.Anoproceso & " where " & dato & ""
Set rs = VGCNx.Execute(SQL)
TDBG_CtaCte.DataSource = rs
TDBG_CtaCte.Refresh
If Ctr_CtaCtble.xclave <> "" And Ctr_Analitico.xclave <> "" Then
   framebotones.Visible = True
 Else
   framebotones.Visible = False
End If

End Sub
