VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Begin VB.Form FrmTelecreditoreportes 
   Caption         =   "Reportes"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   11550
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Opciones"
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   4560
      TabIndex        =   26
      Top             =   7200
      Visible         =   0   'False
      Width           =   3255
      Begin VB.CommandButton cmdaceptar 
         Caption         =   "Imprimir"
         Height          =   450
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   1395
      End
      Begin VB.CommandButton Cmdcancelar 
         Caption         =   "Sakir"
         Height          =   450
         Left            =   1635
         TabIndex        =   27
         Top             =   225
         Width           =   1305
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   2475
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   11175
      Begin VB.Frame Framereporte 
         Height          =   615
         Left            =   5520
         TabIndex        =   21
         Top             =   120
         Width           =   5655
         Begin VB.OptionButton OptionGenerado 
            Caption         =   "Generados"
            Height          =   255
            Left            =   360
            TabIndex        =   23
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton OptionActualizados 
            Caption         =   "Pagados"
            Height          =   255
            Left            =   1680
            TabIndex        =   22
            Top             =   240
            Width           =   1095
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuPendientes 
            Height          =   315
            Left            =   4080
            TabIndex        =   24
            Top             =   120
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   556
            XcodMaxLongitud =   0
            xcodwith        =   1000
            NomTabla        =   "te_cabecerapagosinternet"
            TituloAyuda     =   "Busqueda de pendientes"
            ListaCampos     =   "pagosnumero(1),pagosmoneda(1),pagosfecha(2)"
            XcodCampo       =   "pagosnumero"
            XListCampo      =   "pagosmoneda"
            ListaCamposDescrip=   "Nro Rendicion,Moneda, fecha rendicion"
            ListaCamposText =   "pagosnumero,pagosmoneda,pagosfecha"
            Requerido       =   0   'False
         End
         Begin VB.Label Label3 
            Caption         =   "Pendiente"
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
            Height          =   255
            Left            =   3000
            TabIndex        =   25
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Framesalida 
         BackColor       =   &H00FFC0C0&
         Height          =   1695
         Left            =   5520
         TabIndex        =   10
         Top             =   720
         Width           =   5655
         Begin MSComCtl2.DTPicker DTPfechaini 
            Height          =   375
            Left            =   3480
            TabIndex        =   11
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   " MMM - yyyy"
            DateIsNull      =   -1  'True
            Format          =   16711681
            CurrentDate     =   37513
         End
         Begin TextFer.TxFer TxFerTcambio 
            Height          =   300
            Left            =   1320
            TabIndex        =   12
            Top             =   240
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   529
            BackColor       =   16777215
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
            ForeColor       =   0
            Text            =   ""
            ColorIlumina    =   8454143
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            TipoDato        =   1
            NumeroDecimales =   3
            Formato         =   "###,###.000"
            MarcarTextoAlEnfoque=   -1  'True
         End
         Begin TextFer.TxFer TxFerreferencia 
            Height          =   300
            Left            =   1320
            TabIndex        =   13
            Top             =   600
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            BackColor       =   16777215
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
            MaxLength       =   11
            Text            =   ""
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            TipoDato        =   1
            NumeroDecimales =   2
            MarcarTextoAlEnfoque=   -1  'True
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Numero Operacion"
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
            Height          =   375
            Left            =   240
            TabIndex        =   20
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label LBLNUMERO 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label5"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   1320
            TabIndex        =   19
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label lbtot 
            Appearance      =   0  'Flat
            BackColor       =   &H00DDF7F9&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   3720
            TabIndex        =   18
            Top             =   1200
            Width           =   1605
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   0
            Left            =   4095
            TabIndex        =   17
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Referencia Banco"
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
            Height          =   495
            Left            =   240
            TabIndex        =   16
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "T.Cambio : "
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
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lbfechini 
            Caption         =   "Fecha emision"
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
            Left            =   2280
            TabIndex        =   14
            Top             =   240
            Width           =   1065
         End
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayudabanco 
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   840
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   661
         XcodMaxLongitud =   3
         xcodwith        =   400
         NomTabla        =   "gr_banco"
         TituloAyuda     =   "Ayuda de Bancos"
         ListaCampos     =   "bancocodigo(1),bancodescripcion(1)"
         XcodCampo       =   "bancocodigo"
         XListCampo      =   "bancodescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "bancocodigo,bancodescripcion"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaCuentabanco 
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   1320
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   661
         XcodMaxLongitud =   20
         xcodwith        =   2000
         NomTabla        =   "te_cuentabancos"
         TituloAyuda     =   "Busqueda de Cuenta bancaria"
         ListaCampos     =   "cbanco_numero(1),cbanco_referenciacta(1),monedacodigo(1)"
         XcodCampo       =   "cbanco_numero"
         XListCampo      =   "cbanco_referenciacta"
         ListaCamposDescrip=   "Cuenta,Descripcion,Moneda"
         ListaCamposText =   "cbanco_numero,cbanco_referenciacta,monedacodigo"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuempresa 
         Height          =   375
         Left            =   1320
         TabIndex        =   1
         Top             =   240
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   661
         XcodMaxLongitud =   3
         xcodwith        =   300
         NomTabla        =   "co_multiempresas"
         TituloAyuda     =   "Busqueda de Empresas"
         ListaCampos     =   "empresacodigo(1),empresadescripcion(1)"
         XcodCampo       =   "empresacodigo"
         XListCampo      =   "empresadescripcion"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "empresacodigo,empresadescripcion"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaMoneda 
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Top             =   1920
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   661
         Enabled         =   0   'False
         XcodMaxLongitud =   2
         xcodwith        =   300
         NomTabla        =   "gr_moneda"
         TituloAyuda     =   "Busqueda de Moneda"
         ListaCampos     =   "monedacodigo(1),monedadescripcion(1)"
         XcodCampo       =   "monedacodigo"
         XListCampo      =   "monedadescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "monedacodigo,monedadescripcion"
      End
      Begin VB.Label Label4 
         Caption         =   "Cod. Cuenta"
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
         Index           =   1
         Left            =   60
         TabIndex        =   8
         Top             =   1440
         Width           =   1125
      End
      Begin VB.Label Label4 
         Caption         =   "Banco"
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
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   945
         Width           =   885
      End
      Begin VB.Label Lblempresa1 
         AutoSize        =   -1  'True
         Caption         =   "Empresa :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   855
      End
      Begin VB.Label lbMon 
         Caption         =   "Moneda : "
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
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   2085
         Width           =   855
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBG_concil 
      Height          =   4440
      Left            =   240
      TabIndex        =   9
      Top             =   2640
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   7832
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Cod.Cliente"
      Columns(0).DataField=   "clientecodigo"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Razon Social"
      Columns(1).DataField=   "clienterazonsocial"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "T/D"
      Columns(2).DataField=   "cargodocumento"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Nº Doc"
      Columns(3).DataField=   "cargonumdoc"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Fecha"
      Columns(4).DataField=   "cargoapefecemi"
      Columns(4).NumberFormat=   "Short Date"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   4
      Columns(5)._MaxComboItems=   5
      Columns(5).ValueItems(0)._DefaultItem=   0
      Columns(5).ValueItems(0).Value=   "1"
      Columns(5).ValueItems(0).Value.vt=   8
      Columns(5).ValueItems(0).DisplayValue=   "1"
      Columns(5).ValueItems(0).DisplayValue.vt=   8
      Columns(5).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(5).ValueItems(1)._DefaultItem=   0
      Columns(5).ValueItems(1).Value=   "0"
      Columns(5).ValueItems(1).Value.vt=   8
      Columns(5).ValueItems(1).DisplayValue=   "0"
      Columns(5).ValueItems(1).DisplayValue.vt=   8
      Columns(5).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(5).ValueItems.Count=   2
      Columns(5).Caption=   "CH"
      Columns(5).DataField=   "chkconcil"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Moneda"
      Columns(6).DataField=   "monedacodigo"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Saldo"
      Columns(7).DataField=   "saldo1"
      Columns(7).EditMask=   "####,###.00"
      Columns(7).EditMaskRight=   -1  'True
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Monto Pago"
      Columns(8).DataField=   "saldo"
      Columns(8).EditMask=   "####,###.00"
      Columns(8).EditMaskRight=   -1  'True
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Deposito Pago"
      Columns(9).DataField=   "importepago"
      Columns(9).EditMask=   "#######.00"
      Columns(9).EditMaskRight=   -1  'True
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   10
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=10"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1799"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1720"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=5080"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=5001"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(2).Width=635"
      Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=556"
      Splits(0)._ColumnProps(13)=   "Column(2)._ColStyle=8196"
      Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(15)=   "Column(3).Width=2170"
      Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=2090"
      Splits(0)._ColumnProps(18)=   "Column(3)._ColStyle=8196"
      Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(20)=   "Column(4).Width=1693"
      Splits(0)._ColumnProps(21)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(4)._WidthInPix=1614"
      Splits(0)._ColumnProps(23)=   "Column(4)._ColStyle=8196"
      Splits(0)._ColumnProps(24)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(25)=   "Column(5).Width=529"
      Splits(0)._ColumnProps(26)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(5)._WidthInPix=450"
      Splits(0)._ColumnProps(28)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(29)=   "Column(6).Width=794"
      Splits(0)._ColumnProps(30)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(6)._WidthInPix=714"
      Splits(0)._ColumnProps(32)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(33)=   "Column(7).Width=1508"
      Splits(0)._ColumnProps(34)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(7)._WidthInPix=1429"
      Splits(0)._ColumnProps(36)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(37)=   "Column(8).Width=2011"
      Splits(0)._ColumnProps(38)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(8)._WidthInPix=1931"
      Splits(0)._ColumnProps(40)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(41)=   "Column(9).Width=2302"
      Splits(0)._ColumnProps(42)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(9)._WidthInPix=2223"
      Splits(0)._ColumnProps(44)=   "Column(9).Order=10"
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HE0E0E0&,.bold=0,.fontsize=825"
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
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&H344A87&"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1,.fgpicPosition=1"
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
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=54,.parent=13,.locked=-1"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=58,.parent=13,.locked=-1"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=62,.parent=13,.locked=-1"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=59,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=60,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=61,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13,.bgcolor=&HBFFFAA&"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=32,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=29,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=30,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=31,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=74,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=71,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=72,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=73,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=66,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=70,.parent=13"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=67,.parent=14"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=68,.parent=15"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=69,.parent=17"
      _StyleDefs(76)  =   "Named:id=33:Normal"
      _StyleDefs(77)  =   ":id=33,.parent=0"
      _StyleDefs(78)  =   "Named:id=34:Heading"
      _StyleDefs(79)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(80)  =   ":id=34,.wraptext=-1"
      _StyleDefs(81)  =   "Named:id=35:Footing"
      _StyleDefs(82)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(83)  =   "Named:id=36:Selected"
      _StyleDefs(84)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(85)  =   "Named:id=37:Caption"
      _StyleDefs(86)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(87)  =   "Named:id=38:HighlightRow"
      _StyleDefs(88)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(89)  =   "Named:id=39:EvenRow"
      _StyleDefs(90)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(91)  =   "Named:id=40:OddRow"
      _StyleDefs(92)  =   ":id=40,.parent=33"
      _StyleDefs(93)  =   "Named:id=41:RecordSelector"
      _StyleDefs(94)  =   ":id=41,.parent=34"
      _StyleDefs(95)  =   "Named:id=42:FilterBar"
      _StyleDefs(96)  =   ":id=42,.parent=33,.alignment=1"
   End
End
Attribute VB_Name = "FrmTelecreditoreportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents RsConcil As ADODB.Recordset
Attribute RsConcil.VB_VarHelpID = -1
Dim RsSaldoIni As ADODB.Recordset
Dim tmontosolesDebe As Double, tmontodolaresDebe As Double
Dim tmontosolesHaber As Double, tmontodolaresHaber As Double
Dim montosolesDebe As Double, montodolaresDebe As Double
Dim montosolesHaber As Double, montodolaresHaber As Double
Dim mtsoles As Double, mtdolar As Double
Public SQL As String
Public numero As String
Dim tsoles As Double, tdolar As Double
Dim montoextbanc As Double
Dim mon As String
Dim mon_descripcion As String
Dim Modificar As Integer
Dim flagcal As Boolean
Dim Fecha As Date
Dim dllgeneral As dllgeneral.dll_general

Private Sub cmdeliminar_click()
TxtNrorendicion.Enabled = True
 Modificar = 2
 Cmdcancelar.Enabled = True
 Modificar = 2
 Call Listar(Modificar)
 If MsgBox("desea Eliminar Rendicion", vbQuestion + vbYesNo) = vbYes Then
   RsConcil.MoveFirst
   If RsConcil.RecordCount() > 0 Then
      Do Until RsConcil.EOF
          RsConcil("chkconcil").Value = 0
          RsConcil.MoveNext
       Loop
    End If
    Call cmdaceptar_Click
 End If
 
End Sub

Private Sub Ctr_AyudaBanco_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Ctr_AyudaCuentabanco.Filtro = "cbanco_codigo='" & Ctr_Ayudabanco.xclave & "'"
End Sub

Private Sub Ctr_AyudaCaja_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Ctr_AyudaCuentabanco.Filtro = ""
End Sub

Private Sub Ctr_AyudaCuentabanco_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    Ctr_AyudaMoneda.xclave = ColecCampos("monedacodigo")
    Ctr_AyudaMoneda.Ejecutar

End Sub

Private Sub Ctr_AyuPendientes_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
'TxFerTcambio.valor = ESNULO(ColecCampos("pagostipodecambio"), 0)
If OptionGenerado.Value = True Then
   Call Listar(1)
  Else
   Call Listar(2)
End If
Frame2.Visible = True
LBLNUMERO.Caption = ColecCampos("pagosnumero")
cmdaceptar.Visible = True

End Sub

Private Sub Form_Load()
    Call Ctr_Ayuempresa.Conexion(VGCNx)
    Call Ctr_Ayudabanco.Conexion(VGCNx)
    Call Ctr_AyudaCuentabanco.Conexion(VGCNx)
    Call Ctr_AyudaMoneda.Conexion(VGCNx)
    Call Ctr_AyuPendientes.Conexion(VGCNx)
    DTPfechaini.Value = VGParamSistem.fechatrabajo
    TDBG_concil.FetchRowStyle = True
    Frame2.Visible = False
    TxFerTcambio.valor = 0
    OptionGenerado.Value = True
    Ctr_AyuPendientes.Filtro = " isnull(bancoestadoPendiente,'0')='1'"
    
End Sub

Private Sub cmdaceptar_Click()
Dim X As Integer
    If OptionGenerado.Value = True Then
       Call Imprimir(1)
     Else
       Call Imprimir(2)
    End If
End Sub

Private Sub CmdCancelar_Click()
    If RsConcil Is Nothing Then
        Unload Me
        Exit Sub
    End If
    
    Unload Me
End Sub
Private Sub CalcularTotales(ByVal rs As Recordset)
Dim rsaux As ADODB.Recordset
Dim vardllgen As New dllgeneral.dll_general
montosolesDebe = 0: montodolaresDebe = 0:
montosolesHaber = 0: montodolaresHaber = 0:
mtsoles = 0: mtdolar = 0
Set rsaux = rs.Clone
If rsaux.BOF = True Or rsaux.EOF = True Then Exit Sub
Dim Fecha As Double
rsaux.MoveFirst
    While Not rsaux.EOF
        If rsaux("chkconcil").Value <> 0 Then
           mtsoles = mtsoles + rsaux!importepago
        End If
        rsaux.MoveNext
    Wend
    'Soles
   lbtot(2).Caption = Format(mtsoles, "###,###,###,###.00")
End Sub

Private Sub OptionActualizados_Click()
If OptionActualizados.Value = True Then
   Ctr_AyuPendientes.Filtro = " isnull(bancoestadoPendiente,'0')='0'"
  Else
   Ctr_AyuPendientes.Filtro = " isnull(bancoestadoPendiente,'0')='1'"
End If
End Sub

Private Sub OptionGenerado_Click()
If OptionGenerado.Value = True Then
   Ctr_AyuPendientes.Filtro = " isnull(bancoestadoPendiente,'0')='1'"
  Else
   Ctr_AyuPendientes.Filtro = " isnull(bancoestadoPendiente,'0')='0'"
End If

End Sub

Private Sub Imprimir(ValorConci As String)
Dim arrform(7) As Variant, arrparm(4) As Variant
Dim NombreRep As String, CadOrden As String
Dim tipo As String
Dim fecha1 As String
If OptionGenerado.Value = True Then
   tipo = " Generados a telecredito"
    NombreRep = "te_pagosGenerados.rpt"
 Else
   tipo = " Actualizados en Tesoreria "
    NombreRep = "te_pagosActualizados.rpt"
End If
fecha1 = Format(DateSerial(DTPfechaini.Year, DTPfechaini.Month, 1), "dd/mm/yyyy")
  
    arrparm(0) = VGParamSistem.BDEmpresa
    arrparm(1) = Trim(Ctr_Ayuempresa.xclave)
    arrparm(2) = Trim(LBLNUMERO)
    arrparm(3) = ValorConci
    arrform(0) = "numero='" & LBLNUMERO & "'"
    arrform(1) = "banco='" & Ctr_Ayudabanco.xnombre & "'"
    arrform(2) = "cuenta='" & Ctr_AyudaCuentabanco.xclave & "'"
    arrform(3) = "moneda='" & Ctr_AyudaMoneda.xnombre & "'"
    arrform(4) = "fecha='" & fecha1 & "'"
    arrform(5) = "empresa='" & Ctr_Ayuempresa.xnombre & "'"
    arrform(6) = "tipo='" & tipo & "'"
        
    Call ImpresionRptProc(NombreRep, arrform, arrparm, , "Pagos en internet")
End Sub


Private Sub TDBG_concil_DblClick()
If VGmodifica = 1 Then Exit Sub
If TxFerTcambio.valor = 0 Then
   MsgBox ("Ingrese el tipo de cambio ")
   Exit Sub
End If
FramePagos.Visible = True
With RsConcil
    LblRazonsocial = !clienterazonsocial
    LblDocumento = !documentocargo + "-" + !cargonumdoc
    LblSaldo.Caption = !saldo1
    TxFerImporte.valor = !saldo
    !importepago = !saldo
    If RTrim(!tipocuenta01) <> "" Then
        Check1.Value = 1
        LblSoles = !cuenta01
    Else
        Check1.Value = 0
        LblSoles = ""
    End If
    If RTrim(!tipocuenta02) <> "" Then
        Check2.Value = 1
        LblDolar = !cuenta02
    Else
        Check2.Value = 0
        LblDolar = ""
    End If
    If !monedacodigo <> Ctr_AyudaMoneda.xclave Then
       If Ctr_AyudaMoneda.xclave = "01" Then
           !importepago = Round(!saldo * TxFerTcambio.valor, 2)
        Else
           !importepago = Round(!saldo / TxFerTcambio.valor, 2)
        End If
    End If
End With
LblDeposito.Caption = Round(Format(RsConcil!importepago, "####,###.00"), 2)
TxFerImporte.SetFocus
End Sub

Private Sub TDBG_concil_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
Dim rsclone As New ADODB.Recordset
    On Error GoTo X
     Set rsclone = RsConcil.Clone(adLockReadOnly)
    If rsclone.RecordCount = 0 Then Exit Sub
    rsclone.Bookmark = Bookmark
    If rsclone!chkconcil = 1 Then
       RowStyle.BackColor = RGB(200, 250, 100)
    End If
    flagcal = True

    Exit Sub
X:
Resume Next

End Sub

Private Sub TDBG_concil_HeadClick(ByVal ColIndex As Integer)
 TDBG_concil.Refresh
 On Error GoTo y
 With RsConcil
    If .Sort = Empty Then
        .Sort = TDBG_concil.Columns.Item(ColIndex).DataField & " asc"
    ElseIf Right(.Sort, 3) = "asc" Then
        .Sort = TDBG_concil.Columns.Item(ColIndex).DataField & " desc"
    ElseIf Right(.Sort, 4) = "desc" Then
        .Sort = TDBG_concil.Columns.Item(ColIndex).DataField & " asc"
    End If
    TDBG_concil.Refresh
 End With
y:
End Sub
Private Sub Listar(op As Integer)

Dim rs As New ADODB.Recordset
Dim vardllgen As New dllgeneral.dll_general


  If ExisteElem(0, VGCNx, "##tmp_tel") Then VGCNx.Execute ("drop table ##tmp_tel")
  SQL = " te_pagosxinternet_rpt '" & VGCNx.DefaultDatabase & "','" & Ctr_Ayuempresa.xclave & "','" & Ctr_AyuPendientes.xclave & "','" & op & "'"
   Set RsConcil = VGgeneral.Execute(SQL)
   If RsConcil.RecordCount > 0 Then
         Set TDBG_concil.DataSource = RsConcil
         TDBG_concil.Refresh
          Call CalcularTotales(RsConcil)
      End If
End Sub

