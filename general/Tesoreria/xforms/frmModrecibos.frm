VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmModrecibos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de recibos de tesoreria"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11700
   Icon            =   "frmModrecibos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   11700
   Begin TabDlg.SSTab SSTabMant 
      Height          =   8685
      Left            =   60
      TabIndex        =   0
      Top             =   -30
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   15319
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Consulta"
      TabPicture(0)   =   "frmModrecibos.frx":1272
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FrameConsulta"
      Tab(0).Control(1)=   "FrameConsul"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Mantenimiento"
      TabPicture(1)   =   "frmModrecibos.frx":128E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Shilu2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "SSTab2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "StBar"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "frameGrid"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "FrameCabecera"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin VB.Frame FrameCabecera 
         Height          =   1704
         Left            =   60
         TabIndex        =   32
         Top             =   315
         Width           =   11265
         Begin VB.CheckBox ChkActCaja 
            Alignment       =   1  'Right Justify
            Caption         =   " Caja"
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   7335
            TabIndex        =   25
            Top             =   912
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.CheckBox ChkRegComp 
            Alignment       =   1  'Right Justify
            Caption         =   "Cliente"
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   6540
            TabIndex        =   24
            Top             =   912
            Width           =   795
         End
         Begin VB.CheckBox ChkCtaCte 
            Alignment       =   1  'Right Justify
            Caption         =   "Proveedor"
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   5310
            TabIndex        =   23
            Top             =   912
            Width           =   1140
         End
         Begin VB.ComboBox CmbTcambio 
            Enabled         =   0   'False
            Height          =   288
            ItemData        =   "frmModrecibos.frx":12AA
            Left            =   8136
            List            =   "frmModrecibos.frx":12B7
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   1272
            Visible         =   0   'False
            Width           =   1755
         End
         Begin TextFer.TxFer TxNref 
            Height          =   300
            Left            =   9936
            TabIndex        =   28
            Top             =   912
            Width           =   1104
            _ExtentX        =   1958
            _ExtentY        =   529
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
            Valor           =   ""
         End
         Begin TextFer.TxFer TxNAux 
            Height          =   300
            Left            =   9912
            TabIndex        =   22
            Top             =   564
            Width           =   1092
            _ExtentX        =   1931
            _ExtentY        =   529
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
            Valor           =   ""
         End
         Begin MSComCtl2.DTPicker Dtp_FechaDoc 
            Height          =   312
            Left            =   9012
            TabIndex        =   3
            Top             =   120
            Width           =   2052
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            Format          =   85786625
            CurrentDate     =   37469
         End
         Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_ModoOperacion 
            Height          =   315
            Left            =   4725
            TabIndex        =   2
            Top             =   120
            Width           =   3090
            _ExtentX        =   5450
            _ExtentY        =   556
            XcodMaxLongitud =   2
            xcodwith        =   200
            NomTabla        =   "te_operaciongeneral"
            ListaCampos     =   "operacioncodigo(1),operaciondescripcion(1),operacionmanejactas(1),operacioncontrolaclienteprov(1),operacionvalidacajabancos(1)"
            XcodCampo       =   "operacioncodigo"
            XListCampo      =   "operaciondescripcion"
            ListaCamposDescrip=   "codigo,descripcion"
            ListaCamposText =   "operacioncodigo,operaciondescripcion,operacionmanejactas,operacioncontrolaclienteprov,operacionvalidacajabancos"
         End
         Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_Proveedor 
            Height          =   312
            Left            =   1080
            TabIndex        =   4
            Top             =   504
            Width           =   4068
            _ExtentX        =   7170
            _ExtentY        =   556
            XcodMaxLongitud =   11
            xcodwith        =   1200
            NomTabla        =   "cp_proveedor"
            TituloAyuda     =   "Busqueda de Proveedor"
            ListaCampos     =   "clientecodigo(1),clienterazonsocial(1),clienteruc(1),clientetelefono(1)"
            XcodCampo       =   "clientecodigo"
            XListCampo      =   "clienterazonsocial"
            ListaCamposDescrip=   "Codigo,Descripcion"
            ListaCamposText =   "clientecodigo,clienterazonsocial,clienteruc,clientetelefono"
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaCaja 
            Height          =   312
            Left            =   1080
            TabIndex        =   6
            Top             =   1296
            Width           =   2940
            _ExtentX        =   5186
            _ExtentY        =   556
            XcodMaxLongitud =   11
            xcodwith        =   90
            NomTabla        =   "te_codigocaja"
            TituloAyuda     =   "Busqueda de Caja"
            ListaCampos     =   "cajacodigo(1),cajadescripcion(1)"
            XcodCampo       =   "cajacodigo"
            XListCampo      =   "cajadescripcion"
            ListaCamposDescrip=   "Codigo,Descripcion"
            ListaCamposText =   "cajacodigo,cajadescripcion"
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaOficina 
            Height          =   300
            Left            =   5760
            TabIndex        =   5
            Top             =   576
            Width           =   3132
            _ExtentX        =   5530
            _ExtentY        =   529
            XcodMaxLongitud =   0
            xcodwith        =   80
            NomTabla        =   "cp_oficina"
            TituloAyuda     =   "Ayuda de Tipo Analitico"
            ListaCampos     =   "vendedorcodigo(1),vendedornombres(1)"
            XcodCampo       =   "vendedorcodigo"
            XListCampo      =   "vendedornombres"
            ListaCamposDescrip=   "Código,Descripción"
            ListaCamposText =   "vendedorcodigo,vendedornombres"
         End
         Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_Moneda 
            Height          =   312
            Left            =   5040
            TabIndex        =   7
            Top             =   1296
            Width           =   2172
            _ExtentX        =   3836
            _ExtentY        =   556
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
         Begin TextFer.TxFer TxingEgr 
            Height          =   300
            Left            =   930
            TabIndex        =   75
            Top             =   150
            Width           =   375
            _ExtentX        =   661
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
            MaxLength       =   99
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            NoCaracteres    =   "I,E"
            NoRangoCadena   =   -1  'True
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuempresa 
            Height          =   315
            Left            =   1080
            TabIndex        =   76
            Top             =   960
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   556
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
         Begin VB.Label LeModComp 
            Caption         =   "Operacion :"
            Height          =   375
            Left            =   3690
            TabIndex        =   64
            Top             =   180
            Width           =   1485
         End
         Begin VB.Label leFechaDoc 
            AutoSize        =   -1  'True
            Caption         =   "Fecha doc. :"
            Height          =   192
            Left            =   7920
            TabIndex        =   63
            Top             =   180
            Width           =   900
         End
         Begin VB.Label Le_empresa 
            AutoSize        =   -1  'True
            Caption         =   "Empresa :"
            Height          =   192
            Left            =   120
            TabIndex        =   60
            Top             =   996
            Width           =   708
         End
         Begin VB.Label Le_Caja 
            AutoSize        =   -1  'True
            Caption         =   "Caja :"
            Height          =   192
            Left            =   228
            TabIndex        =   58
            Top             =   1308
            Width           =   408
         End
         Begin VB.Label Leoficina 
            Caption         =   "Oficina :"
            Height          =   252
            Left            =   5256
            TabIndex        =   57
            Top             =   624
            Width           =   600
         End
         Begin VB.Label LeNaux 
            AutoSize        =   -1  'True
            Caption         =   "Nº Provision :"
            Height          =   192
            Left            =   8940
            TabIndex        =   56
            Top             =   624
            Width           =   960
         End
         Begin VB.Label lenref 
            AutoSize        =   -1  'True
            Caption         =   "Nº Transf. :"
            Height          =   192
            Left            =   9012
            TabIndex        =   55
            Top             =   948
            Width           =   816
         End
         Begin VB.Label Le_Proveedor 
            Caption         =   "Proveedor :"
            Height          =   252
            Left            =   156
            TabIndex        =   54
            Top             =   612
            Width           =   1020
         End
         Begin VB.Label LeTcambio 
            AutoSize        =   -1  'True
            Caption         =   "T/Cambio :"
            Height          =   192
            Left            =   7308
            TabIndex        =   52
            Top             =   1332
            Visible         =   0   'False
            Width           =   792
         End
         Begin VB.Label lb_vcambio 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00F2FEFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   9936
            TabIndex        =   27
            Top             =   1272
            Visible         =   0   'False
            Width           =   1092
         End
         Begin VB.Label Le_Mon 
            AutoSize        =   -1  'True
            Caption         =   "Moneda :"
            Height          =   192
            Left            =   4116
            TabIndex        =   51
            Top             =   1332
            Width           =   672
         End
         Begin VB.Label lendocum 
            AutoSize        =   -1  'True
            Caption         =   "TIpo ( I/E) :"
            Height          =   195
            Left            =   105
            TabIndex        =   50
            Top             =   180
            Width           =   825
         End
         Begin VB.Label leNComprob 
            AutoSize        =   -1  'True
            Caption         =   "NUMERO :"
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   1
            Left            =   1470
            TabIndex        =   33
            Top             =   180
            Width           =   810
         End
         Begin VB.Label lbNumComprobCab 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2FDFE&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0000010000"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2415
            TabIndex        =   1
            Top             =   180
            Width           =   1095
         End
      End
      Begin VB.Frame frameGrid 
         BackColor       =   &H00808080&
         Height          =   1848
         Left            =   60
         TabIndex        =   46
         Top             =   2064
         Width           =   11220
         Begin TrueOleDBGrid70.TDBGrid TDBG_Det 
            Height          =   1356
            Left            =   72
            TabIndex        =   29
            Top             =   12
            Width           =   11088
            _ExtentX        =   19553
            _ExtentY        =   2381
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Item"
            Columns(0).DataField=   "item"
            Columns(0).DataWidth=   5
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "T.Doc"
            Columns(1).DataField=   "tipodoc_concepto"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Nro.Doc"
            Columns(2).DataField=   "numdocumento"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Mon.Doc"
            Columns(3).DataField=   "monedacancela"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Doc Canc."
            Columns(4).DataField=   "tdqc"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Nro.Doc Canc."
            Columns(5).DataField=   "ndqc"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Cod CB"
            Columns(6).DataField=   "cajabanco1"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Nro. Cta.Cte"
            Columns(7).DataField=   "numctacte"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Moneda/Cuenta Bco"
            Columns(8).DataField=   "cuenta"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Imp.Soles"
            Columns(9).DataField=   "importesoles"
            Columns(9).NumberFormat=   "###,###,###,###.00"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "Imp.Dolares"
            Columns(10).DataField=   "importesoles"
            Columns(10).NumberFormat=   "###,###,###.00"
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   11
            Splits(0)._UserFlags=   0
            Splits(0).AllowRowSizing=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   13160660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=11"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=714"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=635"
            Splits(0)._ColumnProps(4)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=258"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=953"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=873"
            Splits(0)._ColumnProps(10)=   "Column(1).AllowSizing=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=260"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=1879"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1799"
            Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(17)=   "Column(3).Width=1323"
            Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=1244"
            Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(21)=   "Column(4).Width=1561"
            Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=1482"
            Splits(0)._ColumnProps(24)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(25)=   "Column(5).Width=2011"
            Splits(0)._ColumnProps(26)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(5)._WidthInPix=1931"
            Splits(0)._ColumnProps(28)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(29)=   "Column(6).Width=1191"
            Splits(0)._ColumnProps(30)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(31)=   "Column(6)._WidthInPix=1111"
            Splits(0)._ColumnProps(32)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(33)=   "Column(7).Width=2461"
            Splits(0)._ColumnProps(34)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(35)=   "Column(7)._WidthInPix=2381"
            Splits(0)._ColumnProps(36)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(37)=   "Column(8).Width=2408"
            Splits(0)._ColumnProps(38)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(39)=   "Column(8)._WidthInPix=2328"
            Splits(0)._ColumnProps(40)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(41)=   "Column(9).Width=2090"
            Splits(0)._ColumnProps(42)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(43)=   "Column(9)._WidthInPix=2011"
            Splits(0)._ColumnProps(44)=   "Column(9).AllowSizing=0"
            Splits(0)._ColumnProps(45)=   "Column(9)._ColStyle=514"
            Splits(0)._ColumnProps(46)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(47)=   "Column(10).Width=2302"
            Splits(0)._ColumnProps(48)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(49)=   "Column(10)._WidthInPix=2223"
            Splits(0)._ColumnProps(50)=   "Column(10)._ColStyle=2"
            Splits(0)._ColumnProps(51)=   "Column(10).Order=11"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            Appearance      =   2
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            RowDividerStyle =   6
            MultipleLines   =   0
            CellTips        =   2
            CellTipsWidth   =   0
            MultiSelect     =   2
            DataView        =   1
            AnimateWindow   =   2
            DeadAreaBackColor=   13160660
            RowDividerColor =   13160660
            RowSubDividerColor=   13160660
            DirectionAfterEnter=   1
            MaxRows         =   250000
            ViewColumnCaptionWidth=   1140.095
            ViewColumnWidth =   9764.788
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
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&H344A87&"
            _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(24)  =   "Splits(0).Style:id=47,.parent=1"
            _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=56,.parent=4"
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=48,.parent=2"
            _StyleDefs(27)  =   "Splits(0).FooterStyle:id=49,.parent=3"
            _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=50,.parent=5"
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=52,.parent=6"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=51,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=53,.parent=8"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=54,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=55,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=57,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=58,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=62,.parent=47,.alignment=1"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=48,.alignment=0"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=49"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=51"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=74,.parent=47"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=71,.parent=48,.alignment=0"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=72,.parent=49"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=73,.parent=51"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=47"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=48"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=49"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=51"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=70,.parent=47"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=67,.parent=48"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=68,.parent=49"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=69,.parent=51"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=16,.parent=47"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=13,.parent=48"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=14,.parent=49"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=15,.parent=51"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=28,.parent=47"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=48"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=49"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=51"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=46,.parent=47"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=43,.parent=48"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=44,.parent=49"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=45,.parent=51"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=78,.parent=47"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=75,.parent=48"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=76,.parent=49"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=77,.parent=51"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=66,.parent=47"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=48"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=49"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=51"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=158,.parent=47,.alignment=1,.bgcolor=&HF7FBA4&"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=155,.parent=48,.alignment=2"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=156,.parent=49"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=157,.parent=51,.bgcolor=&HF7FBA4&"
            _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=20,.parent=47,.alignment=1,.bgcolor=&HE1FFFF&"
            _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=17,.parent=48"
            _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=18,.parent=49"
            _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=19,.parent=51,.bgcolor=&HE1FFFF&"
            _StyleDefs(80)  =   "Named:id=33:Normal"
            _StyleDefs(81)  =   ":id=33,.parent=0"
            _StyleDefs(82)  =   "Named:id=34:Heading"
            _StyleDefs(83)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(84)  =   ":id=34,.wraptext=-1"
            _StyleDefs(85)  =   "Named:id=35:Footing"
            _StyleDefs(86)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(87)  =   "Named:id=36:Selected"
            _StyleDefs(88)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(89)  =   "Named:id=37:Caption"
            _StyleDefs(90)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(91)  =   "Named:id=38:HighlightRow"
            _StyleDefs(92)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(93)  =   "Named:id=39:EvenRow"
            _StyleDefs(94)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(95)  =   "Named:id=40:OddRow"
            _StyleDefs(96)  =   ":id=40,.parent=33"
            _StyleDefs(97)  =   "Named:id=41:RecordSelector"
            _StyleDefs(98)  =   ":id=41,.parent=34"
            _StyleDefs(99)  =   "Named:id=42:FilterBar"
            _StyleDefs(100) =   ":id=42,.parent=33"
         End
         Begin TextFer.TxFer TxTotBruto 
            Height          =   300
            Left            =   8280
            TabIndex        =   72
            Top             =   1440
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   529
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
            Valor           =   ""
         End
         Begin TextFer.TxFer TxTotIGV 
            Height          =   300
            Left            =   9684
            TabIndex        =   73
            Top             =   1440
            Width           =   1464
            _ExtentX        =   2593
            _ExtentY        =   529
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
            Valor           =   ""
         End
         Begin VB.Shape Shape10 
            BackColor       =   &H8000000B&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   90
            Left            =   0
            Top             =   -120
            Width           =   11265
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Registros :"
            ForeColor       =   &H00FFFFFF&
            Height          =   192
            Left            =   5340
            TabIndex        =   48
            Top             =   1476
            Width           =   972
         End
         Begin VB.Label lbnregdetalle 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "0 "
            Height          =   252
            Left            =   6432
            TabIndex        =   47
            Top             =   1428
            Width           =   1056
         End
      End
      Begin VB.Frame FrameConsul 
         BackColor       =   &H8000000B&
         Height          =   768
         Left            =   -74910
         TabIndex        =   35
         Top             =   375
         Width           =   11250
         Begin VB.Image Image1 
            Height          =   372
            Left            =   132
            Picture         =   "frmModrecibos.frx":12E3
            Stretch         =   -1  'True
            Top             =   216
            Width           =   456
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   " Consulta e Ingreso de Recibos de caja"
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
            Height          =   180
            Left            =   696
            TabIndex        =   45
            Top             =   300
            Width           =   5640
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            Height          =   540
            Left            =   48
            TabIndex        =   44
            Top             =   132
            Width           =   11148
         End
         Begin VB.Shape Shape6 
            BorderColor     =   &H00FFFFFF&
            Height          =   12
            Left            =   60
            Top             =   720
            Width           =   11136
         End
      End
      Begin MSComctlLib.StatusBar StBar 
         Height          =   288
         Left            =   96
         TabIndex        =   34
         Top             =   6720
         Width           =   11220
         _ExtentX        =   19791
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   3
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   6
               Object.Width           =   2547
               MinWidth        =   2547
               TextSave        =   "06/11/2012"
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   8819
               MinWidth        =   8819
               Text            =   "Comprobante Contable : "
               TextSave        =   "Comprobante Contable : "
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               AutoSize        =   1
               Object.Width           =   8334
               Picture         =   "frmModrecibos.frx":2555
               Text            =   "Estado :"
               TextSave        =   "Estado :"
            EndProperty
         EndProperty
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   2736
         Left            =   72
         TabIndex        =   30
         Top             =   3972
         Width           =   11232
         _ExtentX        =   19817
         _ExtentY        =   4815
         _Version        =   393216
         Style           =   1
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   520
         MouseIcon       =   "frmModrecibos.frx":37D7
         TabCaption(0)   =   "&Ingreso del detalle"
         TabPicture(0)   =   "frmModrecibos.frx":37F3
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "FramDetalle"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         Begin VB.Frame FramDetalle 
            Height          =   2475
            Left            =   75
            TabIndex        =   31
            Top             =   330
            Width           =   11085
            Begin TextFer.TxFer Txdolares 
               Height          =   312
               Left            =   9756
               TabIndex        =   21
               Top             =   2040
               Width           =   1164
               _ExtentX        =   2064
               _ExtentY        =   556
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
               Valor           =   ""
            End
            Begin TextFer.TxFer Txsoles 
               Height          =   312
               Left            =   7644
               TabIndex        =   20
               Top             =   2040
               Width           =   1284
               _ExtentX        =   2275
               _ExtentY        =   556
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
               Valor           =   ""
            End
            Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuFormaPago 
               Height          =   405
               Left            =   5265
               TabIndex        =   10
               Top             =   495
               Width           =   2790
               _ExtentX        =   4921
               _ExtentY        =   714
               XcodMaxLongitud =   0
               xcodwith        =   300
               NomTabla        =   "cp_tipodocumento"
               ListaCampos     =   "tdocumentocodigo(1),tdocumentodescripcion(1),tdocumentovalidabanco(1)"
               XcodCampo       =   "tdocumentocodigo"
               XListCampo      =   "tdocumentodescripcion"
               ListaCamposDescrip=   "codigo,descripcion"
               ListaCamposText =   "tdocumentocodigo,tdocumentodescripcion,tdocumentovalidabanco"
            End
            Begin TextFer.TxFer TxNroPago 
               Height          =   300
               Left            =   8160
               TabIndex        =   11
               Top             =   480
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   529
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
               Valor           =   ""
            End
            Begin TextFer.TxFer Txtglosa 
               Height          =   300
               Left            =   840
               TabIndex        =   18
               Top             =   2040
               Width           =   4212
               _ExtentX        =   7435
               _ExtentY        =   529
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
               Valor           =   ""
            End
            Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuMonedacancela 
               Height          =   315
               Left            =   840
               TabIndex        =   12
               Top             =   960
               Width           =   4215
               _ExtentX        =   7435
               _ExtentY        =   556
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
            Begin TextFer.TxFer TxNroPagado 
               Height          =   300
               Left            =   3600
               TabIndex        =   9
               Top             =   480
               Width           =   1545
               _ExtentX        =   2725
               _ExtentY        =   529
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxLength       =   14
               Text            =   ""
               Valor           =   ""
            End
            Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuBanco 
               Height          =   315
               Left            =   6195
               TabIndex        =   13
               Top             =   960
               Width           =   4785
               _ExtentX        =   8440
               _ExtentY        =   556
               XcodMaxLongitud =   2
               xcodwith        =   300
               NomTabla        =   "gr_banco"
               TituloAyuda     =   "Busqueda de Moneda"
               ListaCampos     =   "bancocodigo(1),bancodescripcion(1)"
               XcodCampo       =   "bancocodigo"
               XListCampo      =   "bancodescripcion"
               ListaCamposDescrip=   "Código,Descripción"
               ListaCamposText =   "bancocodigo,bancodescripcion"
               Requerido       =   0   'False
            End
            Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuctacte 
               Height          =   315
               Left            =   840
               TabIndex        =   14
               Top             =   1320
               Width           =   4215
               _ExtentX        =   7435
               _ExtentY        =   556
               XcodMaxLongitud =   20
               xcodwith        =   1200
               NomTabla        =   "te_cuentabancos"
               TituloAyuda     =   "Busqueda de Cuenta Bancaria"
               ListaCampos     =   "cbanco_numero(1), cbanco_referenciacta(1)   "
               XcodCampo       =   "cbanco_numero"
               XListCampo      =   "cbanco_referenciacta"
               ListaCamposDescrip=   "Cta Cte,Descripción"
               ListaCamposText =   "cbanco_numero, cbanco_referenciacta"
            End
            Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuAnalitico 
               Height          =   315
               Left            =   810
               TabIndex        =   16
               Top             =   1710
               Visible         =   0   'False
               Width           =   4260
               _ExtentX        =   7514
               _ExtentY        =   556
               XcodMaxLongitud =   11
               xcodwith        =   900
               TituloAyuda     =   "Busqueda de Centro de Costos"
               ListaCampos     =   "entidadcodigo(1),entidadrazonsocial(1)"
               XcodCampo       =   "entidadcodigo"
               XListCampo      =   "entidadrazonsocial"
               ListaCamposDescrip=   "Código,Descripción"
               ListaCamposText =   "entidadcodigo,entidadrazonsocial"
               Requerido       =   0   'False
            End
            Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayugastos 
               Height          =   285
               Left            =   6195
               TabIndex        =   15
               Top             =   1350
               Width           =   4770
               _ExtentX        =   8414
               _ExtentY        =   503
               XcodMaxLongitud =   20
               xcodwith        =   1000
               NomTabla        =   "co_gastos"
               TituloAyuda     =   "Busqueda de Cuenta de Gastos"
               ListaCampos     =   "gastoscodigo(1),gastosdescripcion(1),gastosctrlcostos(1),cuentacodigo(1),tipoanaliticocodigo(1)"
               XcodCampo       =   "gastoscodigo"
               XListCampo      =   "gastosdescripcion"
               ListaCamposDescrip=   "Código,Descripción"
               ListaCamposText =   "gastoscodigo,gastosdescripcion,gastosctrlcostos,cuentacodigo,tipoanaliticocodigo"
            End
            Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuCcosto 
               Height          =   315
               Left            =   6195
               TabIndex        =   17
               Top             =   1710
               Width           =   4770
               _ExtentX        =   8414
               _ExtentY        =   556
               XcodMaxLongitud =   10
               xcodwith        =   900
               NomTabla        =   "ct_centrocosto"
               TituloAyuda     =   "Busqueda de Centro de Costos"
               ListaCampos     =   "centrocostocodigo(1),centrocostodescripcion(1)"
               XcodCampo       =   "centrocostocodigo"
               XListCampo      =   "centrocostodescripcion"
               ListaCamposDescrip=   "Código,Descripción"
               ListaCamposText =   "centrocostocodigo,centrocostodescripcion"
               Requerido       =   0   'False
            End
            Begin MSComCtl2.DTPicker DTPFechacancelacion 
               Height          =   312
               Left            =   6324
               TabIndex        =   19
               Top             =   2064
               Width           =   1284
               _ExtentX        =   2249
               _ExtentY        =   529
               _Version        =   393216
               Format          =   85786625
               CurrentDate     =   37469
            End
            Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuconcepto 
               Height          =   315
               Left            =   120
               TabIndex        =   8
               Top             =   480
               Width           =   3375
               _ExtentX        =   5953
               _ExtentY        =   556
               XcodMaxLongitud =   2
               xcodwith        =   300
               NomTabla        =   "te_conceptocaja"
               TituloAyuda     =   "Busqueda de concepto de Caja"
               ListaCampos     =   "conceptocodigo(1),conceptodescripcion(1)"
               XcodCampo       =   "conceptocodigo"
               XListCampo      =   "conceptodescripcion"
               ListaCamposDescrip=   "Código,Descripción"
               ListaCamposText =   "conceptocodigo,conceptodescripcion"
            End
            Begin TextFer.TxFer TxtRendicion 
               Height          =   300
               Left            =   9600
               TabIndex        =   80
               Top             =   510
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   529
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
               Valor           =   ""
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "Nro. Rendicion"
               Height          =   195
               Left            =   9720
               TabIndex        =   81
               Top             =   240
               Width           =   1065
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "Concepto :"
               Height          =   195
               Left            =   600
               TabIndex        =   79
               Top             =   240
               Width           =   780
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "PAGO"
               Height          =   195
               Left            =   5400
               TabIndex        =   78
               Top             =   240
               Width           =   450
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Nro. :"
               Height          =   195
               Left            =   8520
               TabIndex        =   77
               Top             =   210
               Width           =   390
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "F. Cancelacion :"
               Height          =   192
               Left            =   5232
               TabIndex        =   74
               Top             =   2124
               Width           =   1140
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "Dolares :"
               ForeColor       =   &H00800000&
               Height          =   192
               Left            =   9048
               TabIndex        =   71
               Top             =   2112
               Width           =   660
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "Cta.Cte. :"
               Height          =   195
               Left            =   120
               TabIndex        =   70
               Top             =   1350
               Width           =   660
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Banco :"
               Height          =   195
               Left            =   5400
               TabIndex        =   69
               Top             =   990
               Width           =   555
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Doc.a  Pagar :"
               Height          =   195
               Left            =   3720
               TabIndex        =   68
               Top             =   210
               Width           =   1035
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Moneda :"
               Height          =   195
               Left            =   120
               TabIndex        =   67
               Top             =   990
               Width           =   675
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Glosa :"
               Height          =   195
               Left            =   120
               TabIndex        =   66
               Top             =   2115
               Width           =   495
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Tipo  :"
               Height          =   195
               Left            =   6720
               TabIndex        =   65
               Top             =   225
               Width           =   450
            End
            Begin VB.Label Lblanalitico 
               AutoSize        =   -1  'True
               Caption         =   "Analitico"
               Height          =   195
               Left            =   120
               TabIndex        =   62
               Top             =   1725
               Width           =   600
            End
            Begin VB.Label Lblgastos 
               AutoSize        =   -1  'True
               Caption         =   "Gastos :"
               Height          =   195
               Left            =   5400
               TabIndex        =   61
               Top             =   1440
               Width           =   585
            End
            Begin VB.Label lbccosto 
               AutoSize        =   -1  'True
               Caption         =   "C.Costo"
               Height          =   195
               Left            =   5430
               TabIndex        =   59
               Top             =   1755
               Width           =   555
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Soles :"
               ForeColor       =   &H00800000&
               Height          =   192
               Left            =   7092
               TabIndex        =   53
               Top             =   2112
               Width           =   480
            End
         End
         Begin VB.Shape Shilu1 
            BorderColor     =   &H0000FF00&
            BorderWidth     =   2
            Height          =   510
            Left            =   1605
            Top             =   -585
            Visible         =   0   'False
            Width           =   3870
         End
      End
      Begin VB.Frame FrameConsulta 
         BackColor       =   &H00808080&
         Height          =   7605
         Left            =   -74910
         TabIndex        =   36
         Top             =   1044
         Width           =   11250
         Begin TextFer.TxFer TxEjecutar 
            Height          =   300
            Left            =   120
            TabIndex        =   49
            Top             =   465
            Width           =   7485
            _ExtentX        =   13203
            _ExtentY        =   529
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
            Valor           =   ""
         End
         Begin VB.CheckBox ChkTodos 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00808080&
            Caption         =   "Todos"
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   7650
            TabIndex        =   43
            Top             =   480
            Width           =   855
         End
         Begin MSDataListLib.DataCombo Dtc_Campo 
            Height          =   315
            Left            =   9375
            TabIndex        =   42
            Top             =   435
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "nombre"
            BoundColumn     =   "codigo"
            Text            =   ""
         End
         Begin TrueOleDBGrid70.TDBGrid TDBG_Consulta 
            Height          =   5136
            Left            =   120
            TabIndex        =   37
            Top             =   816
            Width           =   11040
            _ExtentX        =   19473
            _ExtentY        =   9049
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "I/E"
            Columns(0).DataField=   "cabrec_ingsal"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Recibo"
            Columns(1).DataField=   "cabrec_numrecibo"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "F.  Doc."
            Columns(2).DataField=   "cabrec_fechadocumento"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "T,Oper"
            Columns(3).DataField=   "operacioncodigo"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Mon"
            Columns(4).DataField=   "monedacodigo"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Oficina"
            Columns(5).DataField=   "vendedorcodigo"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Client/Proveed"
            Columns(6).DataField=   "clientecodigo"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "C/B"
            Columns(7).DataField=   ""
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Cj/Bco"
            Columns(8).DataField=   "cajacodigo"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Total Soles"
            Columns(9).DataField=   "cabrec_totsoles"
            Columns(9).NumberFormat=   "###,###,###,###.00"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "Total Dolares"
            Columns(10).DataField=   "cabrec_totdolares"
            Columns(10).NumberFormat=   "###,###,###,###.00"
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(11)._VlistStyle=   0
            Columns(11)._MaxComboItems=   5
            Columns(11).Caption=   "Nro. Provision"
            Columns(11).DataField=   "cabcomprobnumero"
            Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(12)._VlistStyle=   0
            Columns(12)._MaxComboItems=   5
            Columns(12).Caption=   "Nro.Transf."
            Columns(12).DataField=   "cabrec_numreciboegreso"
            Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   13
            Splits(0)._UserFlags=   0
            Splits(0).AllowRowSizing=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   13160660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=13"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=582"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=503"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=1085"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1005"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(9)=   "Column(2).Width=1270"
            Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1191"
            Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(13)=   "Column(3).Width=1085"
            Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=1005"
            Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(17)=   "Column(4).Width=688"
            Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=609"
            Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(21)=   "Column(5).Width=1005"
            Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=926"
            Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(25)=   "Column(6).Width=2037"
            Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=1958"
            Splits(0)._ColumnProps(28)=   "Column(6)._ColStyle=2"
            Splits(0)._ColumnProps(29)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(30)=   "Column(7).Width=609"
            Splits(0)._ColumnProps(31)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(32)=   "Column(7)._WidthInPix=529"
            Splits(0)._ColumnProps(33)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(34)=   "Column(8).Width=1005"
            Splits(0)._ColumnProps(35)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(36)=   "Column(8)._WidthInPix=926"
            Splits(0)._ColumnProps(37)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(38)=   "Column(9).Width=1693"
            Splits(0)._ColumnProps(39)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(9)._WidthInPix=1614"
            Splits(0)._ColumnProps(41)=   "Column(9)._ColStyle=2"
            Splits(0)._ColumnProps(42)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(43)=   "Column(10).Width=1852"
            Splits(0)._ColumnProps(44)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(45)=   "Column(10)._WidthInPix=1773"
            Splits(0)._ColumnProps(46)=   "Column(10)._ColStyle=2"
            Splits(0)._ColumnProps(47)=   "Column(10).Order=11"
            Splits(0)._ColumnProps(48)=   "Column(11).Width=2011"
            Splits(0)._ColumnProps(49)=   "Column(11).DividerColor=0"
            Splits(0)._ColumnProps(50)=   "Column(11)._WidthInPix=1931"
            Splits(0)._ColumnProps(51)=   "Column(11).Order=12"
            Splits(0)._ColumnProps(52)=   "Column(12).Width=1614"
            Splits(0)._ColumnProps(53)=   "Column(12).DividerColor=0"
            Splits(0)._ColumnProps(54)=   "Column(12)._WidthInPix=1535"
            Splits(0)._ColumnProps(55)=   "Column(12).Order=13"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            Appearance      =   2
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            Caption         =   "Resultados de La Consulta"
            MultipleLines   =   0
            CellTips        =   2
            CellTipsWidth   =   0
            MultiSelect     =   2
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=1,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
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
            _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4,.bgcolor=&H8000000F&"
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.bgcolor=&H344A87&"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=58,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=82,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=79,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=80,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=81,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=66,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=70,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=46,.parent=13,.alignment=1,.bgcolor=&HE1FFFF&"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=43,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=44,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=45,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=86,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=83,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=84,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=85,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=78,.parent=13"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=75,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=76,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=77,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=50,.parent=13,.alignment=1,.bgcolor=&HFAF7B4&"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=47,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=48,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=49,.parent=17"
            _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=54,.parent=13,.alignment=1,.bgcolor=&HE1FFFF&"
            _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=51,.parent=14"
            _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=52,.parent=15"
            _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=53,.parent=17"
            _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=62,.parent=13"
            _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=59,.parent=14"
            _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=60,.parent=15"
            _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=61,.parent=17"
            _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=74,.parent=13"
            _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=71,.parent=14"
            _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=72,.parent=15"
            _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=73,.parent=17"
            _StyleDefs(88)  =   "Named:id=33:Normal"
            _StyleDefs(89)  =   ":id=33,.parent=0"
            _StyleDefs(90)  =   "Named:id=34:Heading"
            _StyleDefs(91)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(92)  =   ":id=34,.wraptext=-1"
            _StyleDefs(93)  =   "Named:id=35:Footing"
            _StyleDefs(94)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(95)  =   "Named:id=36:Selected"
            _StyleDefs(96)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(97)  =   "Named:id=37:Caption"
            _StyleDefs(98)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(99)  =   "Named:id=38:HighlightRow"
            _StyleDefs(100) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(101) =   "Named:id=39:EvenRow"
            _StyleDefs(102) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(103) =   "Named:id=40:OddRow"
            _StyleDefs(104) =   ":id=40,.parent=33"
            _StyleDefs(105) =   "Named:id=41:RecordSelector"
            _StyleDefs(106) =   ":id=41,.parent=34"
            _StyleDefs(107) =   "Named:id=42:FilterBar"
            _StyleDefs(108) =   ":id=42,.parent=33"
         End
         Begin VB.Shape Shape7 
            BorderColor     =   &H00404040&
            Height          =   285
            Left            =   10065
            Top             =   7110
            Width           =   1095
         End
         Begin VB.Label lbl_nregconsulta 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "0 "
            Height          =   285
            Left            =   10080
            TabIndex        =   41
            Top             =   7110
            Width           =   1080
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Registros :"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   9000
            TabIndex        =   40
            Top             =   7155
            Width           =   975
         End
         Begin VB.Label Label2 
            BackColor       =   &H00808080&
            Caption         =   "Valor :"
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   120
            TabIndex        =   39
            Top             =   210
            Width           =   2085
         End
         Begin VB.Label Label4 
            BackColor       =   &H00808080&
            Caption         =   "Criterio :"
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   8715
            TabIndex        =   38
            Top             =   510
            Width           =   570
         End
      End
      Begin VB.Shape Shilu2 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         Height          =   1908
         Left            =   12
         Top             =   2028
         Visible         =   0   'False
         Width           =   11316
      End
   End
End
Attribute VB_Name = "frmModrecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ClsMM1 As ClsMantMovimientos
Dim rscampo As ADODB.Recordset
Dim rscabecera As ADODB.Recordset
Dim WithEvents rsmantenimiento As ADODB.Recordset
Attribute rsmantenimiento.VB_VarHelpID = -1
Public IMant As Integer
Dim adReasonAux As ADODB.EventReasonEnum
Dim VlUltAccion As Integer
Dim Vlnaux As String
Public VlDocAnt As String
Public VlDocNota As String
Public VlComprob_Conta As String
Public m_cajachica As Integer

Private Sub ChkActCaja_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call SendKeys("{TAB}")
End Sub
Private Sub ChkCtaCte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call SendKeys("{TAB}")
End Sub
Private Sub ChkRegComp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call SendKeys("{TAB}")
End Sub
Private Sub ChkTodos_Click()
    If ChkTodos.Value = 1 Then
        Call EjecutarConsulta("", True)
      Else
        Call EjecutarConsulta("", False)
    End If
End Sub
Private Sub CmbTcambio_Click()
  lb_vcambio = Format(XRecuperaTipoCambio(Dtp_FechaDoc, CmbTcambio.ListIndex + 1, VGCNx), "#0.000 ")
End Sub
Private Sub CmbTcambio_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call SendKeys("{TAB}")
End Sub


Private Sub Ctr_AyuAnalitico_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
If Not VGflaglimpia Then Exit Sub
    Set ClsMM1 = New ClsMantMovimientos
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, analiticotes)
End Sub
Private Sub Ctr_AyuAnalitico_AlNoDevolverNada()
    If Not VGflaglimpia Then Exit Sub
    Set ClsMM1 = New ClsMantMovimientos
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, analiticotes)
End Sub

Private Sub CtrAyu_Ccosto_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    If Not VGflaglimpia Then Exit Sub
    Set ClsMM1 = New ClsMantMovimientos
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, ccostotes)
End Sub
Private Sub CtrAyu_ccosto_AlNoDevolverNada()
    If Not VGflaglimpia Then Exit Sub
    Set ClsMM1 = New ClsMantMovimientos
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, ccostotes)
End Sub
Private Sub CtrAyu_gastos_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    Dim SQL As String
    If Not VGflaglimpia Then Exit Sub
    Set ClsMM1 = New ClsMantMovimientos
    If ColecCampos("gastosctrlcostos") Then
        Ctr_AyuCcosto.Visible = True
        lbccosto.Visible = True
        Ctr_AyuCcosto.xclave = "": Ctr_AyuCcosto.xnombre = ""
  '      Cuentacodigo = ColecCampos("cuentacodigo")
      Else
        Ctr_AyuCcosto.Visible = False
        lbccosto.Visible = False
        Ctr_AyuCcosto.xclave = "": Ctr_AyuCcosto.xnombre = ""
    End If
    If ColecCampos("tipoanaliticocodigo") <> "00" Then
       If ExisteElem(0, VGCNx, VGComputer) Then VGCNx.Execute ("drop table " & VGComputer & "")
       Ctr_AyuAnalitico.xclave = "": Ctr_AyuAnalitico.xnombre = ""
       SQL = "select a.entidadcodigo,b.entidadrazonsocial into "
       SQL = SQL & VGComputer & " from ct_analitico a inner join "
       SQL = SQL & " ct_entidad b on a.entidadcodigo=b.entidadcodigo "
       SQL = SQL & " where a.tipoanaliticocodigo='" & ColecCampos("tipoanaliticocodigo") & "'"
       VGCNx.Execute (SQL)
       Ctr_AyuAnalitico.NomTabla = VGComputer
       Ctr_AyuAnalitico.Visible = True
       Lblanalitico.Visible = True
     Else
       Ctr_AyuAnalitico.Visible = False
       Lblanalitico.Visible = False
       Ctr_AyuCcosto.xclave = "00"
    End If
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, gastostes)
'    frameGrid.Refresh
    
End Sub
Private Sub CtrAyu_gastos_AlNoDevolverNada()
    If Not VGflaglimpia Then Exit Sub
    Set ClsMM1 = New ClsMantMovimientos
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, gastostes)
End Sub


Private Sub Ctr_AyuBanco_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
If Not VGflaglimpia Then Exit Sub
    Dim xrs As New Recordset
    Set xrs = VGCNx.Execute("select * from te_cuentabancos where cbanco_codigo='" & Ctr_AyuBanco.xclave & "' and monedacodigo='" & CtrAyu_Moneda.xclave & "'")
    If xrs.RecordCount() <= 1 Then
       MsgBox "Por lo Menos debe Existir una cuenta bancaria para este banco ", vbExclamation
        Exit Sub
    End If
    Set ClsMM1 = New ClsMantMovimientos
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, banco)
    Ctr_Ayuctacte.Filtro = "cbanco_codigo ='" & Ctr_AyuBanco.xclave & "' and monedacodigo='" & CtrAyu_Moneda.xclave & "'"

End Sub



Private Sub Ctr_AyuCcosto_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Dim x As String
    If Not VGflaglimpia Then Exit Sub
    If VGMoverRegistro Then Exit Sub
    x = "'"
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, costos)
End Sub

Private Sub Ctr_Ayuctacte_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
If Not VGflaglimpia Then Exit Sub
Set ClsMM1 = New ClsMantMovimientos
Call ClsMM1.ActualizarDetalle(rsmantenimiento, ctacte)
End Sub

Private Sub Ctr_AyuFormaPago_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
If Not VGflaglimpia Then Exit Sub
    Set ClsMM1 = New ClsMantMovimientos
     If ColecCampos("tdocumentovalidabanco") = 1 Then
       Ctr_AyuBanco.Visible = True
       Ctr_Ayuctacte.Visible = True
     Else
       Ctr_AyuBanco.Visible = False
       Ctr_Ayuctacte.Visible = False
    End If
    Ctr_AyuBanco.xclave = "": Ctr_AyuBanco.Ejecutar
    Ctr_Ayuctacte.xclave = "": Ctr_Ayuctacte.Ejecutar
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, formapago)
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, banco)
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, ctacte)
End Sub

Private Sub Ctr_Ayugastos_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    If Not VGflaglimpia Then Exit Sub
    If VGMoverRegistro Then Exit Sub
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, gastos)
End Sub

Private Sub CtrAyu_ModoOperacion_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
If ColecCampos("operacionvalidacajabancos") = "C" Then
   Ctr_AyudaCaja.Visible = True
 Else
   Ctr_AyudaCaja.Visible = False
End If
If ColecCampos("operacioncontrolaclienteprov") = "P" Then
     Ctr_Ayugastos.Visible = False
 Else
   Ctr_Ayugastos.Visible = True
End If
If ColecCampos("operacioncontrolaclienteprov") = "P" Or (ColecCampos("operacioncontrolaclienteprov") = "X" And Trim(CtrAyu_Proveedor.xclave) <> "") Then
   CtrAyu_Proveedor.Visible = True
   Ctr_Ayuconcepto.NomTabla = "cp_tipodocumento"
   Ctr_Ayuconcepto.ListaCampos = "tdocumentocodigo(1),tdocumentodescripcion(1)"
   Ctr_Ayuconcepto.ListaCamposText = "tdocumentocodigo,tdocumentodescripcion"
   Ctr_Ayuconcepto.XcodCampo = "tdocumentocodigo"
   Ctr_Ayuconcepto.XListCampo = "tdocumentodescripcion"
 ElseIf ColecCampos("operacioncontrolaclienteprov") = "X" Then
        Ctr_Ayuconcepto.NomTabla = "te_conceptocaja"
        Ctr_Ayuconcepto.ListaCampos = "conceptocodigo(1),conceptodescripcion(1)"
        Ctr_Ayuconcepto.ListaCamposText = "conceptocodigo,conceptodescripcion"
        Ctr_Ayuconcepto.XcodCampo = "conceptocodigo"
        Ctr_Ayuconcepto.XListCampo = "conceptodescripcion"
        CtrAyu_Proveedor.Visible = False
      Else
'        Ctr_Ayuconcepto.NomTabla = "cc_tipodocumento"
'        CtrAyu_Proveedor.Visible = True
     End If
 Ctr_Ayuconcepto.Ejecutar
 End Sub

Private Sub CtrAyu_Moneda_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    If CtrAyu_Moneda.xclave = "02" Then
        LeTcambio.Visible = True
        CmbTcambio.Visible = True
        lb_vcambio.Visible = True
        CmbTcambio.ListIndex = 1
      '  Call CmbTcambio_Click
      Else
        LeTcambio.Visible = False
        CmbTcambio.Visible = False
        lb_vcambio.Visible = False
    End If
    Call CmbTcambio_Click
End Sub

Private Sub Dtp_FechaDoc_Change()
    Call CmbTcambio_Click
 End Sub
Private Sub Dtp_FechaDoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call SendKeys("{TAB}")
End Sub

Private Sub Dtp_FechaDocRef_Change()
    If UCase(VlDocNota) = "A" Then Call CmbTcambio_Click
End Sub



Private Sub Form_Activate()
    MDIPrincipal.ToolComprob.Visible = True
    MDIPrincipal.mnu00.Visible = True
    Call PBoton(6)
End Sub
Private Sub Form_Load()
    Top = 0
    Left = 0
    IMant = 0
    VlUltAccion = 0
    Set VGvardllgen = New dllgeneral.dll_general
    Set rscabecera = New ADODB.Recordset
    Set ClsMM1 = New ClsMantMovimientos
    ClsMM1.CargarAyudas (1)
    Set TDBG_Consulta.DataSource = Nothing
    TDBG_Det.FetchRowStyle = True
    Call PrepararTemporalDetalle
 '   If rsmantenimiento.RecordCount = 0 Then
 '       Call HabilitarDetalle(False, FramDetalle, Me)
 '    Else
 '       Call HabilitarDetalle(True, FramDetalle, Me)
 '   End If
    Call VGvardllgen.ActivaTab(0, 1, SSTabMant)
    Call GetCamposdeConsulta
    
End Sub
Private Sub GetCamposdeConsulta()
    Set rscampo = New ADODB.Recordset
    Call rscampo.Fields.Append("codigo", adVarChar, 60)
    Call rscampo.Fields.Append("Nombre", adVarChar, 50)
    rscampo.Open
    rscampo.AddNew
    rscampo!codigo = "cabrec_numrecibo"
    rscampo!nombre = "Nro. de Recibo"
    rscampo.Update
    rscampo.AddNew
    rscampo!codigo = "convert(varchar(10),cabrec_fechadocumento,103)"
    rscampo!nombre = "Fecha Recibo"
    rscampo.Update
    rscampo.AddNew
    rscampo!codigo = "clientecodigo"
    rscampo!nombre = "Ruc Proveedor"
    rscampo.Update
    rscampo.AddNew
    rscampo!codigo = "vendedorcodigo"
    rscampo!nombre = "Codigo Oficina"
    rscampo.Update
    rscampo.AddNew
    rscampo!codigo = "cajacodigo"
    rscampo!nombre = "Codigo caja"
    rscampo.Update
    rscampo.AddNew
    rscampo!codigo = "operacioncodigo"
    rscampo!nombre = "Codigo Operacion"
    rscampo.Update
    rscampo.AddNew
    rscampo!codigo = "cabrec_ingsal"
    rscampo!nombre = "(I)ngresos /(E)gresos"
    rscampo.Update
    Set Dtc_Campo.RowSource = rscampo
    Dtc_Campo.BoundText = "cabrec_numrecibo"
End Sub
Public Sub AlMoverRegistro()
Dim vardllgen As New dllgeneral.dll_general
Dim pos As Integer
 '   If VGactulizodoc Then Exit Sub 'Estoy Actualizando documentos
    VGMoverRegistro = True
    On Error Resume Next
    With rsmantenimiento
        Ctr_AyuAnalitico.xclave = !entidad: Ctr_AyuAnalitico.Ejecutar
        TxNroPagado.Text = !numdocumento
        Ctr_Ayuctacte.xclave = !numctacte
        Ctr_AyuFormaPago.xclave = !tdqc: Ctr_AyuFormaPago.Ejecutar
        TxNroPago.Text = !ndqc
        
        Ctr_AyuMonedacancela.xclave = !monedacancela: Ctr_AyuMonedacancela.Ejecutar
        Ctr_Ayuconcepto.xclave = !tipodoc_concepto: Ctr_Ayuconcepto.Ejecutar
        Txsoles.Text = Format(!importesoles, "###,###,###.00"): Txsoles.valor = Format(!inafecto, "#0.00")
        Txdolares.Text = Format(!importedolares, "###,###,###.00"): Txdolares.valor = Format(!Impcompra, "#0.00")
        Txtglosa.Text = !observacion
        Ctr_AyuCcosto.xclave = Trim(!costos): Ctr_AyuCcosto.Ejecutar
        Ctr_Ayugastos.xclave = Trim(!gastos): Ctr_Ayugastos.Ejecutar
        Ctr_AyuBanco.xclave = !cajabanco1: Ctr_AyuBanco.Ejecutar
        Ctr_Ayuctacte.xclave = Trim(!numctacte)
        DTPFechacancelacion.Value = !fechacancela
        TxtRendicion.Text = !rendicionnumero
        
     '   : Ctr_Ayuctacte.Ejecutar
       
    End With
    VGMoverRegistro = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    MDIPrincipal.ToolComprob.Visible = False
    MDIPrincipal.mnu00.Visible = False
End Sub

Private Sub rsmantenimiento_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If (adReason = adRsnMove Or adReason = adRsnMoveNext) And pRecordset.RecordCount > 0 And adReasonAux <> adRsnAddNew Then
        Call AlMoverRegistro
    End If
    If adReasonAux = adRsnAddNew Then adReasonAux = adRsnMove
End Sub
Private Sub rsmantenimiento_RecordChangeComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, ByVal pError As ADODB.error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    adReasonAux = adReason
End Sub
Private Sub SSTabMant_Click(PreviousTab As Integer)
    If PreviousTab = 0 Then
        Ctr_AyuCcosto.Requerido = True
        CtrAyu_Moneda.Requerido = True
        CtrAyu_Proveedor.Requerido = True
        CtrAyu_ModoOperacion.Requerido = True
        Ctr_AyudaCaja.Requerido = True
        Ctr_AyudaOficina.Requerido = True
        Ctr_AyuMonedacancela.Requerido = True
        Ctr_AyuFormaPago.Requerido = True
        Ctr_AyuMonedacancela.Requerido = True
        Ctr_Ayugastos.Requerido = True
        Ctr_Ayuconcepto.Requerido = True
    
'        MDIPrincipal.mnu00_01(9).Visible = True
       ' le_Mes.Caption = Format(VGParamSistem.mesproceso, "00")
      Else
        Ctr_AyuCcosto.Requerido = False
        If VGParametros.sistemamultiempresas = True Then
           Ctr_Ayuempresa.Visible = True
         Else
           Ctr_Ayuempresa.xclave = "00"
           Ctr_Ayuempresa.Visible = False
        End If
        CtrAyu_Moneda.Requerido = False
        CtrAyu_Proveedor.Requerido = False
        CtrAyu_ModoOperacion.Requerido = False
        Ctr_AyudaCaja.Requerido = False
        Ctr_AyudaOficina.Requerido = False
        Ctr_Ayugastos.Requerido = True
        Ctr_Ayuconcepto.Requerido = True

        If TxEjecutar.Enabled And Me.Visible Then TxEjecutar.SetFocus
'        MDIPrincipal.mnu00_01(9).Visible = False
    End If
End Sub
Private Sub TDBG_Consulta_DblClick()
    Call Modificar
End Sub
Private Sub TDBG_Consulta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call Modificar
    End If
End Sub
Private Sub TDBG_Det_GotFocus()
    'frameGrid.BackColor = &H628837
    Shilu1.Visible = True: Shilu2.Visible = True
End Sub
Private Sub TDBG_Det_LostFocus()
    Shilu1.Visible = False: Shilu2.Visible = False
    'frameGrid.BackColor = &H808080
End Sub
Private Sub TxEjecutar_KeyDown(KeyCode As Integer, Shift As Integer)
Dim cad As String
    If KeyCode = 13 Then
        cad = Dtc_Campo.BoundText & " like '" & Trim(TxEjecutar.Text) & "%'"
        Call EjecutarConsulta(cad, False)
    End If
End Sub
Private Sub EjecutarConsulta(ByVal criterio As String, Optional ByVal todos As Boolean)
Dim cad As String
Dim order As String
Dim sqlcad As String, xasiento As String, xsubasiento As String
    Set rscabecera = New ADODB.Recordset
    Set VGvardllgen = New dllgeneral.dll_general
    If criterio = "" Then
        criterio = " where 1=1 "
        order = " order by cabrec_fechadocumento,cabrec_numreciboegreso "
      Else
         criterio = " where " + criterio
        order = " "
    End If
    If todos Then cad = " "
    sqlcad = "select * from te_cabecerarecibos " & " " & cad & " " & criterio & order
    rscabecera.Open sqlcad, VGCNx, adOpenKeyset, adLockReadOnly
    
    If rscabecera.RecordCount > 0 Then
        lbl_nregconsulta.Caption = Format(rscabecera.RecordCount, "0 ")
        TDBG_Consulta.SetFocus
      Else
        lbl_nregconsulta.Caption = Format(0, "0 ")
        TxEjecutar.SetFocus
    End If
    Set TDBG_Consulta.DataSource = rscabecera
End Sub
Public Sub CalcularTotales(ByVal rs As Recordset)
Dim rsaux As ADODB.Recordset
Dim vardllgen As New dllgeneral.dll_general
Set rsaux = rs.Clone(adLockReadOnly)
If rsaux Is Nothing And IMant <> 1 Then Exit Sub
Dim ximpbruto As Double, xigv As Double
Dim xinafecto As Double, ximpcompra As Double

ximpbruto = 0: xigv = 0:
xinafecto = 0: ximpcompra = 0:
rsaux.MoveFirst
    While Not rsaux.EOF
        ximpbruto = ximpbruto + vardllgen.ESNULO(rsaux!impbruto, 0)
        xigv = xigv + vardllgen.ESNULO(rsaux!igv, 0)
        xinafecto = xinafecto + vardllgen.ESNULO(rsaux!inafecto, 0)
        ximpcompra = ximpcompra + vardllgen.ESNULO(rsaux!Impcompra, 0)
        rsaux.MoveNext
    Wend
    TxTotBruto.Text = Format(ximpbruto, "###,###,###,###.00 ") ' Debe
    TxTotBruto.valor = Format(ximpbruto, "#0.00 ") ' Debe
    
    TxTotIGV.Text = Format(xigv, "###,###,###,###.00 ") ' Debe
    TxTotIGV.valor = Format(xigv, "#0.00 ") ' Debe
    
End Sub
Private Sub Mostrar()
    If rscabecera.State = 0 Then Exit Sub
    If rscabecera.RecordCount = 0 Then Exit Sub
    Set ClsMM1 = New ClsMantMovimientos
    Set VGvardllgen = New dllgeneral.dll_general
    Call ClsMM1.MostrarCabecera(rscabecera.Fields, 1)
    Call ClsMM1.Limpia(1)
    Call PrepararTemporalDetalle
    Call ClsMM1.MostrarDetalle(rsmantenimiento, 1)
    Call VGvardllgen.ActivaTab(1, 1, SSTabMant)
    VlUltAccion = 4
    Call PBoton(VlUltAccion)
    'Comprobante Contable :
    StBar.Panels(2).Text = " Comprobante Contable : " & VlComprob_Conta
    Vlnaux = Trim(TxNAux.Text)
'    VlDocAnt = Trim(CtrAyu_Proveedor.xclave) & "-" & Trim(CtrAyu_TipDoc.xclave) & "-" & Trim(TxSerie.Text) & IIf(Trim(TxSerie.Text) = "", "", "-") & Trim(TxNdoc.Text)
End Sub
Private Sub PrepararTemporalDetalle()
    Set rsmantenimiento = New ADODB.Recordset
    Call ClsMM1.CreaRsTempDetalle(rsmantenimiento)
    rsmantenimiento.Open
    Set TDBG_Det.DataSource = rsmantenimiento
End Sub
Public Sub Botones(ByRef tool As Toolbar, Nuevo As Boolean, Grabar As Boolean, Eliminar As Boolean, _
                   Modificar As Boolean, Cancelar As Boolean, Anadet As Boolean, EliDet As Boolean)
    With tool.Buttons
        .Item(1).Enabled = Nuevo
        .Item(2).Enabled = Grabar
        .Item(3).Enabled = Eliminar
        .Item(4).Enabled = Modificar
        .Item(5).Enabled = Cancelar
        .Item(6).Visible = True
        .Item(7).Visible = True
        .Item(8).Visible = True
        .Item(7).Enabled = Anadet
        .Item(8).Enabled = EliDet
    End With
    With MDIPrincipal
        .mnu00_01(1).Enabled = Nuevo
        .mnu00_01(2).Enabled = Grabar
        .mnu00_01(3).Enabled = Eliminar
        .mnu00_01(4).Enabled = Modificar
        .mnu00_01(5).Enabled = Cancelar
        .mnu00_01(6).Visible = True
        .mnu00_01(7).Visible = True
        .mnu00_01(6).Enabled = Anadet
        .mnu00_01(7).Enabled = EliDet
    End With
End Sub
Public Sub Xnuevo()
    'Validacion
    Call PrepararTemporalDetalle
    Set VGvardllgen = New dllgeneral.dll_general
    Call ClearControlsInframe(FrameCabecera, Me)
    lbnregdetalle.Caption = "0 "
    Set ClsMM1 = New ClsMantMovimientos
    Call ClsMM1.LimpiarCab
    Call ClsMM1.Limpia
    Call VGvardllgen.ActivaTab(1, 1, SSTabMant)
    Call HabilitarDetalle(True, FramDetalle, Me)
    IMant = 1
    VlUltAccion = 1
    Ctr_AyuCcosto.Visible = False
    lbccosto.Visible = False
    TxingEgr.SetFocus
End Sub
Public Sub Grabar()
Dim xnumerocompro As String, nnumerocorrcomprob As Double
Dim xnumerocomprolibro As String, nnumerocorrcomproblibro As Double
Dim Existelibro As Boolean

Dim varnerror As Integer
Set VGvardllgen = New dllgeneral.dll_general
On Error GoTo ErrorGrabar
Dim xcon As Long
VGvarVerifica = True
VGErrorString = ""
varnerror = 0
    Set ClsMM1 = New ClsMantMovimientos
    If Not ClsMM1.ValidarGrabarCabecera(rsmantenimiento.RecordCount) Then Exit Sub
    If Not ClsMM1.ValidarRsDetalle(rsmantenimiento) Then Exit Sub
    xcon = rsmantenimiento.RecordCount
    If rsmantenimiento.RecordCount < 1 Then
        MsgBox "Por lo Menos debe Existir un registro con valores ", vbExclamation
        Exit Sub
    End If
    If rsmantenimiento.RecordCount <> xcon Then
        If MsgBox("Esta Seguro de Grabar ? " & Chr(13) & _
                  "Al momento de grabar se eliminaran lo registro ceros ", vbQuestion + vbOKCancel) = vbCancel _
                  Then
            rsmantenimiento.Filter = 0
            Exit Sub
        End If
    End If
    VGCNx.BeginTrans 'Inicio la transaccion
    Screen.MousePointer = vbHourglass
    Dim xnumero As Long
          
    ' Actualizando numero de comprobante
    
    If frmModrecibos.IMant = 1 Then
       xnumero = ActualizaNumeroAuto("te_parametroempresa", 1, VGCNx, 1)
     Else
       xnumero = CDbl(frmModrecibos.lbNumComprobCab)
    End If
    
    
    '2=>Paso Grabo la Cabecera del Comprobante
    Dim Xnumtesor As String
    Call ClsMM1.Grabaren_Tesoreria(4, xnumero, rsmantenimiento, Xnumtesor)
    If Not VGvarVerifica Then varnerror = 2: GoTo ErrorGrabar
        
    
    '3=>Generar Asiento en Linea segun parametro
    
'    Call ClsMM1.GeneraAsientoenLine(IMant, xnumero, VlComprob_Conta)
    If Not VGvarVerifica Then varnerror = 4: GoTo ErrorGrabar
                               
    VGCNx.CommitTrans 'Acepto toda la transaccion porque es correcta
    If IMant = 1 Then
        MsgBox "Se grabo Satisfactoriamente  El numero de Comprobante Generado Es :" & Chr(13) & _
           "Nro: " & xnumero
      Else
        MsgBox "Se Actualizo Satisfactoriamente  ", vbInformation
    End If
    
    Screen.MousePointer = vbDefault
    IMant = 0
    If rscabecera.State = 1 Then
        rscabecera.Requery
    End If
    Call Cancelar(1)
    Exit Sub
    
ErrorGrabar:
    Select Case varnerror
        Case 1
            MsgBox "No se Genero Correctamente el numero del Comprobante" & Chr(13) & VGErrorString, vbExclamation
        Case 2, 3, 4, 5, 6
            VGgeneral.RollbackTrans
            MsgBox "Hubo Errores al Grabar" & Chr(13) & VGErrorString, vbExclamation
            Call Cancelar(1)
            
        Case Else
            MsgBox "Errores Desconocidos " & Chr(13) & Err.Description
    End Select
    Screen.MousePointer = vbDefault
    Exit Sub
    Resume
End Sub
Public Sub Modificar()
    IMant = 2
    Call Mostrar
End Sub
Public Sub Eliminar()
    If MsgBox("Esta Seguro que desea Eliminar este Comprobante", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Set ClsMM1 = New ClsMantMovimientos
    'Verificar si es que tiene abonos
    VGgeneral.BeginTrans
    Screen.MousePointer = vbHourglass
    
 '   If ChkActCaja.Value = 1 Then 'Este en el caso que actualice caja
        Call ClsMM1.Grabaren_Tesoreria(3, Int(Trim(lbNumComprobCab.Caption)))
 '   End If
 '   If ChkCtaCte.Value = 1 Then 'Y esto para actualizar cuenta corriente
 '       Call ClsMM1.GrabarCP_Cargo(3, Int(Trim(lbNumComprobCab.Caption)))
 '   End If
    Screen.MousePointer = vbHourglass
    
    Dim sqlcad As String
    sqlcad = "" & _
    " Update dbo.ct_cabcomprob" & Year(VGParamSistem.fechatrabajo) & _
    " Set cabcomprobtotdebe=0, " & _
    "     cabcomprobtothaber=0," & _
    "     cabcomprobtotussdebe=0, " & _
    "     cabcomprobtotusshaber = 0 " & _
    " Where cabcomprobnumero='" & VlComprob_Conta & "' " & Chr(13) & _
    " Update dbo.ct_detcomprob" & Year(VGParamSistem.fechatrabajo) & _
    "   Set detcomprobdebe=0, " & _
    "   detcomprobhaber=0, " & _
    "   detcomprobussdebe=0, " & _
    "   detcomprobusshaber = 0 " & _
    " Where cabcomprobnumero='" & VlComprob_Conta & "' "


' OJO
    
'    VGCnxCT.Execute sqlcad
    
    
    VGgeneral.CommitTrans
    If rscabecera.State = 1 Then
        rscabecera.Requery
    End If
    
    'Anulo el comprobante Generado en contabilidad
    'blanqueando el asiento
    
          
    MsgBox "El Registro se Elimino Correctamente"
    Call Cancelar(1)
    Screen.MousePointer = vbDefault
    VlUltAccion = 3
End Sub
Public Sub Cancelar(Optional op As Integer)
Set VGvardllgen = New dllgeneral.dll_general

    If op <> 1 Then
        If MsgBox("Esta Seguro que Desea Cancelar la Operación ", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
            'Resolver el problema que el cursor debe parpadear donde se ha quedado
            Exit Sub
        End If
    End If
        
    If SSTabMant.Tab = 1 Then
        Call VGvardllgen.ActivaTab(0, 1, SSTabMant)
        VlUltAccion = 6
        Set rsmantenimiento = Nothing
    End If
    
End Sub
Public Sub AñadirDetalle()
    Set ClsMM1 = New ClsMantMovimientos
    If rsmantenimiento.RecordCount > 0 Then
        If Not ClsMM1.ValidarGrabarDetalle Then Exit Sub
    End If
    Call HabilitarDetalle(True, FramDetalle, Me)
    Call ClsMM1.AñadiralDetalle(rsmantenimiento)
    lbnregdetalle.Caption = Format(rsmantenimiento.RecordCount, "0 ")
End Sub
Public Sub EliminarDetalle()
Dim reg As Long
Dim num As Integer
    Screen.MousePointer = 11
    On Error Resume Next
    If rsmantenimiento.State = 0 Then Exit Sub
    If rsmantenimiento.RecordCount = 0 Then Exit Sub
    Set ClsMM1 = New ClsMantMovimientos
    If rsmantenimiento.RecordCount = 1 Then
        ClsMM1.Limpia
    End If
    num = CInt(rsmantenimiento!Item)
    reg = rsmantenimiento.RecordCount
    rsmantenimiento.Delete
    If num = reg Then rsmantenimiento.MoveNext
    
    Call CalcularTotales(rsmantenimiento)
    Screen.MousePointer = 1
    
End Sub
Public Sub Imprimir()
    If rscabecera Is Nothing Then Exit Sub
    If rscabecera.State = 0 Then Exit Sub
    If rscabecera.RecordCount = 0 Then Exit Sub
    Call ImprimirComprob(rscabecera(0), rscabecera(1))
End Sub
Private Sub ImprimirComprob(Ncomprob As String, mes As String)
Dim arrform(0) As Variant, arrparm(6) As Variant
Screen.MousePointer = 11
    arrparm(0) = Trim(VGParamSistem.Servidor)
    arrparm(1) = Trim(VGParamSistem.BDEmpresa)
    arrparm(2) = Trim(VGParamSistem.AnoProceso)
    arrparm(3) = CInt(Trim(VGParamSistem.MesProceso))
    arrparm(4) = Trim(Ncomprob)
    Call ImpresionRptProc("rptVoucherComprobCompra.rpt", arrform, arrparm)
Screen.MousePointer = 1
End Sub

Public Sub PMant(Index As Integer)
    Select Case Index
        Case 1
            Call Xnuevo
        Case 2
            Call Grabar
        Case 3 'Eliminar
            Call Eliminar
        Case 4 'Modificar
            Call Modificar
        Case 5
            Call Cancelar
        Case 6
            Call AñadirDetalle
        Case 7
            Call EliminarDetalle
        Case 8
            Call Imprimir
    End Select
    Call PBoton(VlUltAccion)
End Sub
Private Sub PBoton(Index As Integer)
    Select Case Index
        Case 0, 5
            Call Botones(MDIPrincipal.ToolComprob, True, False, False, True, False, False, False)
        Case 1 'nuevo
            Call Botones(MDIPrincipal.ToolComprob, False, True, False, False, True, True, True)
        Case 3 'Eliminar
            Call Botones(MDIPrincipal.ToolComprob, True, False, False, True, False, False, False)
        Case 4 'Modificar
            Call Botones(MDIPrincipal.ToolComprob, False, True, True, False, True, True, True)
        Case 6 'Modificar
            Call Botones(MDIPrincipal.ToolComprob, False, False, False, True, False, False, False)
        Case 7 'salir
            Call Botones(MDIPrincipal.ToolComprob, False, False, False, False, True, False, False)
    End Select
End Sub

Private Sub TxIngEgr_LostFocus()
  lbNumComprobCab.Caption = ActualizaNumeroAuto("te_parametrosempresa", 1, VGCNx)
End Sub


Private Sub TxNroPagado_Change()
    If Not VGflaglimpia Then Exit Sub
    If VGMoverRegistro Then Exit Sub
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, docpago)
End Sub

Private Sub TxNroPago_Change()
    If Not VGflaglimpia Then Exit Sub
    If VGMoverRegistro Then Exit Sub
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, cheque)
End Sub

Private Sub Txtglosa_Change()
    If Not VGflaglimpia Then Exit Sub
    If VGMoverRegistro Then Exit Sub
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, glosates)
End Sub
Private Sub Txglosa_KeyDown(KeyCode As Integer, Shift As Integer)
Dim cad As String
If KeyCode = 13 Then
    Call EjecutarConsulta(cad, False)
End If
End Sub

Function consultadoctesor(Proveedor As String, tD As String, Ndocumento As String) As Boolean
Dim sqlcad As String
Dim rsaux As ADODB.Recordset
    Set rsaux = New ADODB.Recordset
    sqlcad = " Select * From " & _
             VGParamSistem.BDEmpresa & ".dbo.te_cabecerarecibos A, " & _
             VGParamSistem.BDEmpresa & ".dbo.te_detallerecibos B " & _
             " Where A.cabrec_numrecibo=b.cabrec_numrecibo and " & _
             " A.clientecodigo='" & Trim(Proveedor) & "' and " & _
             " B.detrec_tipodoc_concepto='" & Trim(tD) & "'  and " & _
             " B.detrec_numdocumento=dbo.fn_coviertenumdoc('" & Ndocumento & "') and isnull(cabrec_estadoreg,0)<> 1 "
    rsaux.Open sqlcad, VGgeneral, adOpenKeyset, adLockReadOnly
    If rsaux.RecordCount >= 1 Then
        consultadoctesor = True
        Else
        consultadoctesor = False
    End If
End Function


