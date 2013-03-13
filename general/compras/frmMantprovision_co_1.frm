VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "textfer.ocx"
Begin VB.Form frmMantprovision 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Provisiones de Compra"
   ClientHeight    =   8580
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   11520
   Icon            =   "frmMantprovision_co.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   11448
   ScaleWidth      =   12216
   Begin TabDlg.SSTab SSTabMant 
      Height          =   7224
      Left            =   60
      TabIndex        =   0
      Top             =   12
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   12742
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Consulta"
      TabPicture(0)   =   "frmMantprovision_co.frx":1272
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameConsulta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrameConsul"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Mantenimiento"
      TabPicture(1)   =   "frmMantprovision_co.frx":128E
      Tab(1).ControlEnabled=   0   'False
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
      Tab(1).Control(5)=   "framTotales"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin VB.Frame framTotales 
         BackColor       =   &H00C0C0C0&
         Height          =   336
         Left            =   -74940
         TabIndex        =   80
         Top             =   3276
         Width           =   11235
         Begin TextFer.TxFer TxTotBruto 
            Height          =   300
            Left            =   5088
            TabIndex        =   33
            Top             =   0
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   529
            Alignment       =   1
            Appearance      =   0
            BackColor       =   14679546
            Object.CausesValidation=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   8388608
            MaxLength       =   15
            Text            =   "0.00"
            ColorIlumina    =   14679546
            SaltarAlEnter   =   -1  'True
            Valor           =   "0.00"
            TipoDato        =   1
            SignodeMiles    =   -1  'True
            NumeroDecimales =   3
            SignoNegativo   =   0   'False
            Formato         =   "###,###,###,###.00"
            MarcarTextoAlEnfoque=   -1  'True
         End
         Begin TextFer.TxFer TxTotIGV 
            Height          =   300
            Left            =   6612
            TabIndex        =   34
            Top             =   0
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
            Alignment       =   1
            Appearance      =   0
            BackColor       =   16777152
            Object.CausesValidation=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   8388608
            MaxLength       =   15
            Text            =   "0.00"
            ColorIlumina    =   16777152
            SaltarAlEnter   =   -1  'True
            Valor           =   "0.00"
            TipoDato        =   1
            SignodeMiles    =   -1  'True
            NumeroDecimales =   3
            SignoNegativo   =   0   'False
            Formato         =   "###,###,###,###.00"
            MarcarTextoAlEnfoque=   -1  'True
         End
         Begin TextFer.TxFer TxTotInafecto 
            Height          =   300
            Left            =   7788
            TabIndex        =   35
            Top             =   0
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   529
            Alignment       =   1
            Appearance      =   0
            BackColor       =   14679546
            Object.CausesValidation=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   8388608
            MaxLength       =   15
            Text            =   "0.00"
            ColorIlumina    =   14679546
            SaltarAlEnter   =   -1  'True
            Valor           =   "0.00"
            TipoDato        =   1
            SignodeMiles    =   -1  'True
            NumeroDecimales =   3
            SignoNegativo   =   0   'False
            Formato         =   "###,###,###,###.00"
            MarcarTextoAlEnfoque=   -1  'True
         End
         Begin TextFer.TxFer TxTotImpCompra 
            Height          =   300
            Left            =   9312
            TabIndex        =   36
            Top             =   0
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   529
            Alignment       =   1
            Appearance      =   0
            BackColor       =   14679546
            Object.CausesValidation=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   8388608
            MaxLength       =   15
            Text            =   "0.00"
            ColorIlumina    =   14679546
            SaltarAlEnter   =   -1  'True
            Valor           =   "0.00"
            TipoDato        =   1
            SignodeMiles    =   -1  'True
            NumeroDecimales =   3
            SignoNegativo   =   0   'False
            Formato         =   "###,###,###,###.00"
            MarcarTextoAlEnfoque=   -1  'True
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Totales Generales ........."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   216
            Left            =   468
            TabIndex        =   82
            Top             =   120
            Width           =   2808
         End
         Begin VB.Label Label17 
            BackColor       =   &H00344A87&
            Height          =   384
            Left            =   72
            TabIndex        =   81
            Top             =   -72
            Width           =   11076
         End
      End
      Begin VB.Frame FrameCabecera 
         Height          =   2904
         Left            =   -74940
         TabIndex        =   40
         Top             =   315
         Width           =   11265
         Begin VB.TextBox TxTesor 
            Height          =   285
            Left            =   5760
            TabIndex        =   90
            Top             =   165
            Visible         =   0   'False
            Width           =   1470
         End
         Begin MSComCtl2.DTPicker DTPFechaCaja 
            Height          =   252
            Left            =   1308
            TabIndex        =   19
            Top             =   1968
            Visible         =   0   'False
            Width           =   2172
            _ExtentX        =   3831
            _ExtentY        =   445
            _Version        =   393216
            Format          =   60424193
            CurrentDate     =   37617
         End
         Begin VB.CheckBox ChkOperGrab 
            Caption         =   "Operación Grabada"
            ForeColor       =   &H00000080&
            Height          =   270
            Left            =   3750
            TabIndex        =   86
            Top             =   195
            Width           =   1830
         End
         Begin VB.Frame FrmChk 
            Enabled         =   0   'False
            Height          =   516
            Left            =   180
            TabIndex        =   78
            Top             =   1368
            Width           =   4260
            Begin VB.CheckBox ChkCtaCte 
               Alignment       =   1  'Right Justify
               Caption         =   "Cuenta Cte."
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   90
               TabIndex        =   15
               Top             =   195
               Width           =   1140
            End
            Begin VB.CheckBox ChkRegComp 
               Alignment       =   1  'Right Justify
               Caption         =   "Regist. Compra"
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   1320
               TabIndex        =   16
               Top             =   195
               Width           =   1395
            End
            Begin VB.CheckBox ChkActCaja 
               Alignment       =   1  'Right Justify
               Caption         =   "Actualiza Caja"
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   2835
               TabIndex        =   17
               Top             =   195
               Visible         =   0   'False
               Width           =   1305
            End
         End
         Begin VB.ComboBox CmbTcambio 
            Enabled         =   0   'False
            Height          =   288
            ItemData        =   "frmMantprovision_co.frx":12AA
            Left            =   4410
            List            =   "frmMantprovision_co.frx":12B7
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   2244
            Visible         =   0   'False
            Width           =   1755
         End
         Begin TextFer.TxFer TxNdoc 
            Height          =   300
            Left            =   9168
            TabIndex        =   9
            Top             =   768
            Width           =   1992
            _ExtentX        =   3514
            _ExtentY        =   529
            Object.CausesValidation=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   8
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
         End
         Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_TipDoc 
            Height          =   312
            Left            =   8760
            TabIndex        =   7
            Top             =   468
            Width           =   2412
            _ExtentX        =   4255
            _ExtentY        =   550
            XcodMaxLongitud =   2
            NomTabla        =   "cp_tipodocumento"
            TituloAyuda     =   "Busqueda de Tipo de  Documento"
            ListaCampos     =   "tdocumentocodigo(1),tdocumentodescripcion(1),tdocumentotipo(1)"
            XcodCampo       =   "tdocumentocodigo"
            XListCampo      =   "tdocumentodescripcion"
            ListaCamposDescrip=   "Código,Descripción,CargoAbono"
            ListaCamposText =   "tdocumentocodigo,tdocumentodescripcion,tdocumentotipo"
         End
         Begin MSComCtl2.DTPicker Dtp_FechaDoc 
            Height          =   312
            Left            =   8760
            TabIndex        =   10
            Top             =   1092
            Width           =   2412
            _ExtentX        =   4255
            _ExtentY        =   550
            _Version        =   393216
            Format          =   60424193
            CurrentDate     =   37469
         End
         Begin MSComCtl2.DTPicker DtpFech_Ven 
            Height          =   312
            Left            =   8760
            TabIndex        =   11
            Top             =   1416
            Width           =   2412
            _ExtentX        =   4255
            _ExtentY        =   550
            _Version        =   393216
            Format          =   60424193
            CurrentDate     =   37469
         End
         Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_Moneda 
            Height          =   312
            Left            =   1308
            TabIndex        =   21
            Top             =   2244
            Width           =   2172
            _ExtentX        =   3831
            _ExtentY        =   550
            XcodMaxLongitud =   2
            NomTabla        =   "gr_moneda"
            TituloAyuda     =   "Busqueda de Moneda"
            ListaCampos     =   "monedacodigo(1),monedadescripcion(1)"
            XcodCampo       =   "monedacodigo"
            XListCampo      =   "monedadescripcion"
            ListaCamposDescrip=   "Código,Descripción"
            ListaCamposText =   "monedacodigo,monedadescripcion"
         End
         Begin TextFer.TxFer txRuc 
            Height          =   300
            Left            =   1296
            TabIndex        =   4
            Top             =   996
            Width           =   2172
            _ExtentX        =   3831
            _ExtentY        =   529
            Object.CausesValidation=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   11
            Locked          =   -1  'True
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            NoCaracteres    =   "0123456789"
            MarcarTextoAlEnfoque=   -1  'True
            NoRangoCadena   =   -1  'True
         End
         Begin TextFer.TxFer TxSerie 
            Height          =   300
            Left            =   8748
            TabIndex        =   8
            Top             =   756
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   529
            Object.CausesValidation=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   3
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
         End
         Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_Proveedor 
            Height          =   264
            Left            =   1308
            TabIndex        =   3
            Top             =   732
            Width           =   5988
            _ExtentX        =   10562
            _ExtentY        =   466
            XcodMaxLongitud =   11
            xcodwith        =   1400
            NomTabla        =   "cp_proveedor"
            TituloAyuda     =   "Busqueda de Proveedor"
            ListaCampos     =   "clientecodigo(1),clienterazonsocial(1),clienteruc(1),clientetelefono(1)"
            XcodCampo       =   "clientecodigo"
            XListCampo      =   "clienterazonsocial"
            ListaCamposDescrip=   "Codigo,Descripcion"
            ListaCamposText =   "clientecodigo,clienterazonsocial,clienteruc,clientetelefono"
         End
         Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_Modoprovi 
            Height          =   312
            Left            =   1308
            TabIndex        =   1
            Top             =   432
            Width           =   2376
            _ExtentX        =   4191
            _ExtentY        =   550
            XcodMaxLongitud =   2
            NomTabla        =   "co_modoprovi"
            TituloAyuda     =   "Busqueda de Modo de Compra"
            ListaCampos     =   "modoprovicod(1), modoprovidesc(1),modoprovictacte(3), modoproviregcom(3), modoprovitesor(3)"
            XcodCampo       =   "modoprovicod"
            XListCampo      =   "modoprovidesc"
            ListaCamposDescrip=   "Código,Descripción"
            ListaCamposText =   "modoprovicod, modoprovidesc,modoprovictacte, modoproviregcom, modoprovitesor"
         End
         Begin MSComCtl2.DTPicker DTPFechaContab 
            Height          =   300
            Left            =   5085
            TabIndex        =   2
            Top             =   465
            Width           =   2205
            _ExtentX        =   3895
            _ExtentY        =   529
            _Version        =   393216
            Format          =   60424193
            CurrentDate     =   37489
         End
         Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_TipRef 
            Height          =   312
            Left            =   8748
            TabIndex        =   12
            Top             =   1728
            Width           =   2448
            _ExtentX        =   4318
            _ExtentY        =   550
            XcodMaxLongitud =   2
            NomTabla        =   "cp_tipodocumento"
            TituloAyuda     =   "Busqueda de Tipo de  Documento"
            ListaCampos     =   "tdocumentocodigo(1),tdocumentodescripcion(1)"
            XcodCampo       =   "tdocumentocodigo"
            XListCampo      =   "tdocumentodescripcion"
            ListaCamposDescrip=   "Código,Descripción"
            ListaCamposText =   "tdocumentocodigo,tdocumentodescripcion"
            Requerido       =   0   'False
         End
         Begin TextFer.TxFer TxNref 
            Height          =   300
            Left            =   8736
            TabIndex        =   13
            Top             =   2040
            Width           =   2448
            _ExtentX        =   4318
            _ExtentY        =   529
            Object.CausesValidation=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   20
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
         End
         Begin MSComCtl2.DTPicker Dtp_FechaDocRef 
            Height          =   288
            Left            =   8748
            TabIndex        =   14
            Top             =   2376
            Width           =   2448
            _ExtentX        =   4318
            _ExtentY        =   508
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   60424193
            CurrentDate     =   37601
         End
         Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_TipCompra 
            Height          =   312
            Left            =   1308
            TabIndex        =   23
            Top             =   2532
            Width           =   2556
            _ExtentX        =   4509
            _ExtentY        =   550
            XcodMaxLongitud =   2
            NomTabla        =   "co_tipocompra"
            TituloAyuda     =   "Busqueda de Tipo de Compra"
            ListaCampos     =   "tipocompracodigo(1), tipocompradesc(1)"
            XcodCampo       =   "tipocompracodigo"
            XListCampo      =   "tipocompradesc"
            ListaCamposDescrip=   "Código,Descripción"
            ListaCamposText =   "tipocompracodigo, tipocompradesc"
         End
         Begin TextFer.TxFer TxNAux 
            Height          =   300
            Left            =   5712
            TabIndex        =   18
            Top             =   1452
            Width           =   1572
            _ExtentX        =   2773
            _ExtentY        =   529
            Object.CausesValidation=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   5
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
         End
         Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_TipSubAsi 
            Height          =   312
            Left            =   4680
            TabIndex        =   24
            Top             =   2544
            Visible         =   0   'False
            Width           =   2640
            _ExtentX        =   4657
            _ExtentY        =   550
            XcodMaxLongitud =   3
            xcodwith        =   100
            NomTabla        =   "co_tiposubasi"
            TituloAyuda     =   "Busqueda de Tipo de Sub Asiento"
            ListaCampos     =   "tiposubasicodigo(1), tiposubasidesc(1)"
            XcodCampo       =   "tiposubasicodigo"
            XListCampo      =   "tiposubasidesc"
            ListaCamposDescrip=   "Código,Descripción"
            ListaCamposText =   "tiposubasicodigo, tiposubasidesc"
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaOficina 
            Height          =   300
            Left            =   8748
            TabIndex        =   6
            Top             =   156
            Width           =   2412
            _ExtentX        =   4255
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
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaCaja 
            Height          =   312
            Left            =   4392
            TabIndex        =   20
            Top             =   1932
            Visible         =   0   'False
            Width           =   2940
            _ExtentX        =   5186
            _ExtentY        =   550
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
         Begin VB.Label LeCaja 
            AutoSize        =   -1  'True
            Caption         =   "Caja :"
            Height          =   192
            Left            =   3588
            TabIndex        =   89
            Top             =   1980
            Visible         =   0   'False
            Width           =   408
         End
         Begin VB.Label LeFechCaja 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Caja :"
            Height          =   192
            Left            =   156
            TabIndex        =   88
            Top             =   2016
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Label Leoficina 
            Caption         =   "Oficina :"
            Height          =   252
            Left            =   7536
            TabIndex        =   87
            Top             =   168
            Width           =   840
         End
         Begin VB.Label le_Mes 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   288
            Left            =   5352
            TabIndex        =   85
            Top             =   1464
            Width           =   360
         End
         Begin VB.Label leSubAsi 
            AutoSize        =   -1  'True
            Caption         =   "Tip. Sub :"
            Height          =   192
            Left            =   3888
            TabIndex        =   84
            Top             =   2592
            Visible         =   0   'False
            Width           =   696
         End
         Begin VB.Label LeNaux 
            AutoSize        =   -1  'True
            Caption         =   "Nº Aux :"
            Height          =   240
            Left            =   4740
            TabIndex        =   83
            Top             =   1512
            Width           =   540
         End
         Begin VB.Label Lebel16 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Ref. :"
            Height          =   192
            Left            =   7596
            TabIndex        =   79
            Top             =   2460
            Width           =   888
         End
         Begin VB.Label letipref 
            Caption         =   "T.D. Ref. :"
            Height          =   252
            Left            =   7560
            TabIndex        =   77
            Top             =   1740
            Width           =   1020
         End
         Begin VB.Label lenref 
            AutoSize        =   -1  'True
            Caption         =   "Nº Ref. :"
            Height          =   192
            Left            =   7572
            TabIndex        =   76
            Top             =   2148
            Width           =   612
         End
         Begin VB.Label LeTipComp 
            AutoSize        =   -1  'True
            Caption         =   "Tip. Compra :"
            Height          =   240
            Left            =   132
            TabIndex        =   75
            Top             =   2568
            Width           =   948
         End
         Begin VB.Shape Shape11 
            BorderColor     =   &H00FFFFFF&
            Height          =   2568
            Left            =   7416
            Top             =   288
            Width           =   12
         End
         Begin VB.Shape Shape5 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   2
            Height          =   12
            Left            =   60
            Top             =   1356
            Width           =   7260
         End
         Begin VB.Label leFecha 
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Contable :"
            Height          =   225
            Left            =   3750
            TabIndex        =   74
            Top             =   540
            Width           =   1320
         End
         Begin VB.Label lbTelef 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   252
            Left            =   5088
            TabIndex        =   5
            Top             =   1008
            Width           =   2196
         End
         Begin VB.Label LeTelf 
            AutoSize        =   -1  'True
            Caption         =   "Telefono :"
            Height          =   192
            Left            =   4188
            TabIndex        =   73
            Top             =   1068
            Width           =   720
         End
         Begin VB.Label LeModComp 
            Caption         =   "Modo Compra :"
            Height          =   252
            Left            =   156
            TabIndex        =   72
            Top             =   480
            Width           =   1128
         End
         Begin VB.Label Le_Proveedor 
            Caption         =   "Proveedor :"
            Height          =   204
            Left            =   156
            TabIndex        =   71
            Top             =   780
            Width           =   1020
         End
         Begin VB.Label Leruc 
            AutoSize        =   -1  'True
            Caption         =   "Nº. de R.U.C. :"
            Height          =   192
            Left            =   156
            TabIndex        =   66
            Top             =   1068
            Width           =   1056
         End
         Begin VB.Label LeTcambio 
            AutoSize        =   -1  'True
            Caption         =   "T/Cambio :"
            Height          =   192
            Left            =   3588
            TabIndex        =   65
            Top             =   2256
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
            Left            =   6216
            TabIndex        =   64
            Top             =   2256
            Visible         =   0   'False
            Width           =   1092
         End
         Begin VB.Label LeMon 
            AutoSize        =   -1  'True
            Caption         =   "Moneda :"
            Height          =   240
            Left            =   156
            TabIndex        =   63
            Top             =   2304
            Width           =   672
         End
         Begin VB.Label leFechVen 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Venc. :"
            Height          =   192
            Left            =   7548
            TabIndex        =   62
            Top             =   1488
            Width           =   1008
         End
         Begin VB.Label leFechaDoc 
            AutoSize        =   -1  'True
            Caption         =   "Fecha doc. :"
            Height          =   192
            Left            =   7548
            TabIndex        =   61
            Top             =   1164
            Width           =   900
         End
         Begin VB.Label letipdoc 
            Caption         =   "Tipo Doc. :"
            Height          =   252
            Left            =   7512
            TabIndex        =   60
            Top             =   528
            Width           =   840
         End
         Begin VB.Label lendocum 
            AutoSize        =   -1  'True
            Caption         =   "Nº doc. :"
            Height          =   192
            Left            =   7548
            TabIndex        =   59
            Top             =   864
            Width           =   636
         End
         Begin VB.Label leNComprob 
            AutoSize        =   -1  'True
            Caption         =   "NUMERO :"
            ForeColor       =   &H00000080&
            Height          =   144
            Index           =   1
            Left            =   156
            TabIndex        =   42
            Top             =   192
            Width           =   816
         End
         Begin VB.Label lbNumComprobCab 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2FDFE&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0000010000"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   1332
            TabIndex        =   41
            Top             =   180
            Width           =   2292
         End
      End
      Begin VB.Frame frameGrid 
         BackColor       =   &H00808080&
         Height          =   1848
         Left            =   -74940
         TabIndex        =   55
         Top             =   3456
         Width           =   11220
         Begin TrueOleDBGrid70.TDBGrid TDBG_Det 
            Height          =   1632
            Left            =   72
            TabIndex        =   32
            Top             =   180
            Width           =   11040
            _ExtentX        =   19473
            _ExtentY        =   2879
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
            Columns(1).Caption=   "Cuenta"
            Columns(1).DataField=   "cuentacodigo"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Descripcion"
            Columns(2).DataField=   "CuentaDes"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "CC"
            Columns(3).DataField=   "ccosto"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Imp. Bruto"
            Columns(4).DataField=   "impbruto"
            Columns(4).NumberFormat=   "###,###,###,###.00"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "I.G.V."
            Columns(5).DataField=   "igv"
            Columns(5).NumberFormat=   "###,###,###,###.00"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Inafecto"
            Columns(6).DataField=   "Inafecto"
            Columns(6).NumberFormat=   "###,###,###.00"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Imp. Compra"
            Columns(7).DataField=   "impcompra"
            Columns(7).NumberFormat=   "###,###,###.00"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   8
            Splits(0)._UserFlags=   0
            Splits(0).AllowRowSizing=   0   'False
            Splits(0).RecordSelectorWidth=   508
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   13160660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=8"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1037"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=953"
            Splits(0)._ColumnProps(4)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=258"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=1545"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1461"
            Splits(0)._ColumnProps(10)=   "Column(1).AllowSizing=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=260"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=4339"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=4255"
            Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(17)=   "Column(3).Width=1439"
            Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=1355"
            Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(21)=   "Column(4).Width=2731"
            Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=514"
            Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(26)=   "Column(5).Width=2074"
            Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=1990"
            Splits(0)._ColumnProps(29)=   "Column(5).AllowSizing=0"
            Splits(0)._ColumnProps(30)=   "Column(5)._ColStyle=514"
            Splits(0)._ColumnProps(31)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(32)=   "Column(6).Width=2731"
            Splits(0)._ColumnProps(33)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(6)._WidthInPix=2646"
            Splits(0)._ColumnProps(35)=   "Column(6)._ColStyle=2"
            Splits(0)._ColumnProps(36)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(37)=   "Column(7).Width=2731"
            Splits(0)._ColumnProps(38)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(39)=   "Column(7)._WidthInPix=2646"
            Splits(0)._ColumnProps(40)=   "Column(7)._ColStyle=2"
            Splits(0)._ColumnProps(41)=   "Column(7).Order=8"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=7.8,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=7.8,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
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
            ViewColumnWidth =   9768.189
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
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=16,.parent=47"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=13,.parent=48"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=14,.parent=49"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=15,.parent=51"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=47"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=48"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=49"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=51"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=154,.parent=47,.alignment=1,.bgcolor=&H80000018&"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=151,.parent=48,.alignment=2"
            _StyleDefs(54)  =   ":id=151,.bgcolor=&H8000000F&"
            _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=152,.parent=49"
            _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=153,.parent=51,.bgcolor=&HE1FFFF&"
            _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=158,.parent=47,.alignment=1,.bgcolor=&HF7FBA4&"
            _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=155,.parent=48,.alignment=2"
            _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=156,.parent=49"
            _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=157,.parent=51,.bgcolor=&HF7FBA4&"
            _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=20,.parent=47,.alignment=1,.bgcolor=&HE1FFFF&"
            _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=17,.parent=48"
            _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=18,.parent=49"
            _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=19,.parent=51,.bgcolor=&HE1FFFF&"
            _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=24,.parent=47,.alignment=1,.bgcolor=&HE1FFFF&"
            _StyleDefs(66)  =   "Splits(0).Columns(7).HeadingStyle:id=21,.parent=48"
            _StyleDefs(67)  =   "Splits(0).Columns(7).FooterStyle:id=22,.parent=49"
            _StyleDefs(68)  =   "Splits(0).Columns(7).EditorStyle:id=23,.parent=51,.bgcolor=&HE1FFFF&"
            _StyleDefs(69)  =   "Named:id=33:Normal"
            _StyleDefs(70)  =   ":id=33,.parent=0"
            _StyleDefs(71)  =   "Named:id=34:Heading"
            _StyleDefs(72)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(73)  =   ":id=34,.wraptext=-1"
            _StyleDefs(74)  =   "Named:id=35:Footing"
            _StyleDefs(75)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(76)  =   "Named:id=36:Selected"
            _StyleDefs(77)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(78)  =   "Named:id=37:Caption"
            _StyleDefs(79)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(80)  =   "Named:id=38:HighlightRow"
            _StyleDefs(81)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(82)  =   "Named:id=39:EvenRow"
            _StyleDefs(83)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(84)  =   "Named:id=40:OddRow"
            _StyleDefs(85)  =   ":id=40,.parent=33"
            _StyleDefs(86)  =   "Named:id=41:RecordSelector"
            _StyleDefs(87)  =   ":id=41,.parent=34"
            _StyleDefs(88)  =   "Named:id=42:FilterBar"
            _StyleDefs(89)  =   ":id=42,.parent=33"
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
            Left            =   8940
            TabIndex        =   57
            Top             =   2016
            Width           =   972
         End
         Begin VB.Label lbnregdetalle 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "0 "
            Height          =   204
            Left            =   10032
            TabIndex        =   56
            Top             =   1980
            Width           =   1056
         End
         Begin VB.Shape Shape9 
            BorderColor     =   &H00404040&
            Height          =   285
            Left            =   10005
            Top             =   2205
            Width           =   1095
         End
      End
      Begin VB.Frame FrameConsul 
         BackColor       =   &H8000000B&
         Height          =   1005
         Left            =   90
         TabIndex        =   44
         Top             =   375
         Width           =   11250
         Begin VB.Image Image1 
            Height          =   465
            Left            =   135
            Picture         =   "frmMantprovision_co.frx":12E3
            Stretch         =   -1  'True
            Top             =   210
            Width           =   450
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   " Consulta e Ingreso de Provision de Compras"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   690
            TabIndex        =   54
            Top             =   495
            Width           =   5640
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            Height          =   735
            Left            =   45
            TabIndex        =   53
            Top             =   135
            Width           =   11145
         End
         Begin VB.Shape Shape6 
            BorderColor     =   &H00FFFFFF&
            Height          =   15
            Left            =   60
            Top             =   915
            Width           =   11130
         End
      End
      Begin MSComctlLib.StatusBar StBar 
         Height          =   288
         Left            =   -74904
         TabIndex        =   43
         Top             =   6816
         Width           =   11220
         _ExtentX        =   19791
         _ExtentY        =   508
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   3
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   6
               Object.Width           =   2547
               MinWidth        =   2547
               TextSave        =   "13/08/2003"
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
               Object.Width           =   8340
               Picture         =   "frmMantprovision_co.frx":2555
               Text            =   "Estado :"
               TextSave        =   "Estado :"
            EndProperty
         EndProperty
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   1596
         Left            =   -74928
         TabIndex        =   37
         Top             =   5268
         Width           =   11232
         _ExtentX        =   19812
         _ExtentY        =   2815
         _Version        =   393216
         Style           =   1
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   520
         MouseIcon       =   "frmMantprovision_co.frx":37D7
         TabCaption(0)   =   "&Ingreso del detalle"
         TabPicture(0)   =   "frmMantprovision_co.frx":37F3
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Shilu1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "FramDetalle"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         Begin VB.Frame FramDetalle 
            Height          =   1224
            Left            =   75
            TabIndex        =   38
            Top             =   330
            Width           =   11085
            Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_Cuenta 
               Height          =   312
               Left            =   948
               TabIndex        =   25
               Top             =   120
               Width           =   10008
               _ExtentX        =   17653
               _ExtentY        =   550
               XcodMaxLongitud =   20
               xcodwith        =   1000
               NomTabla        =   "ct_cuenta"
               TituloAyuda     =   "Busqueda de Cuenta"
               ListaCampos     =   $"frmMantprovision_co.frx":380F
               XcodCampo       =   "cuentacodigo"
               XListCampo      =   "cuentadescripcion"
               ListaCamposDescrip=   "Código,Descripción"
               ListaCamposText =   "cuentacodigo,cuentadescripcion,cuentaestadoccostos,cuentaestadoanalitico,cuentadocumento,tipoanaliticocodigo,tipoajuste"
            End
            Begin TextFer.TxFer TxImpCompra 
               Height          =   312
               Left            =   9012
               TabIndex        =   29
               Top             =   456
               Width           =   1908
               _ExtentX        =   3366
               _ExtentY        =   550
               Alignment       =   1
               BackColor       =   16777215
               Object.CausesValidation=   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   128
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
            Begin TextFer.TxFer TxImpBruto 
               Height          =   300
               Left            =   936
               TabIndex        =   26
               Top             =   468
               Width           =   1908
               _ExtentX        =   3366
               _ExtentY        =   529
               Alignment       =   1
               BackColor       =   16777215
               Object.CausesValidation=   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   128
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
            Begin TextFer.TxFer TxIGV 
               Height          =   312
               Left            =   3396
               TabIndex        =   27
               Top             =   456
               Width           =   1908
               _ExtentX        =   3366
               _ExtentY        =   550
               Alignment       =   1
               BackColor       =   16777215
               Object.CausesValidation=   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   128
               MaxLength       =   15
               Text            =   "0.00"
               ColorIlumina    =   12648447
               SaltarAlEnter   =   -1  'True
               Valor           =   "0.00"
               TipoDato        =   1
               SignodeMiles    =   -1  'True
               NumeroDecimales =   3
               Formato         =   "###,###,###,###.00"
               MarcarTextoAlEnfoque=   -1  'True
               ColorTextoAlEnfocar=   16711680
            End
            Begin TextFer.TxFer TxInafecto 
               Height          =   312
               Left            =   6036
               TabIndex        =   28
               Top             =   456
               Width           =   1908
               _ExtentX        =   3366
               _ExtentY        =   550
               Alignment       =   1
               BackColor       =   16777215
               Object.CausesValidation=   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   128
               MaxLength       =   15
               Text            =   "0.00"
               ColorIlumina    =   12648447
               SaltarAlEnter   =   -1  'True
               Valor           =   "0.00"
               TipoDato        =   1
               SignodeMiles    =   -1  'True
               NumeroDecimales =   3
               Formato         =   "###,###,###,###.00"
               MarcarTextoAlEnfoque=   -1  'True
               ColorTextoAlEnfocar=   16711680
            End
            Begin TextFer.TxFer Txtglosa 
               Height          =   300
               Left            =   912
               TabIndex        =   30
               Top             =   852
               Width           =   5532
               _ExtentX        =   9758
               _ExtentY        =   529
               Object.CausesValidation=   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
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
            End
            Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_Ccosto 
               Height          =   312
               Left            =   7332
               TabIndex        =   31
               Top             =   876
               Width           =   3612
               _ExtentX        =   6371
               _ExtentY        =   550
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
            Begin VB.Label lbccosto 
               AutoSize        =   -1  'True
               Caption         =   "C.Costo"
               Height          =   192
               Left            =   6636
               TabIndex        =   92
               Top             =   912
               Width           =   552
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Glosa :"
               Height          =   192
               Left            =   108
               TabIndex        =   91
               Top             =   936
               Width           =   492
            End
            Begin VB.Shape Shape14 
               BorderColor     =   &H00FFFFFF&
               Height          =   12
               Left            =   72
               Top             =   816
               Width           =   10920
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Imp. Compra :"
               ForeColor       =   &H00800000&
               Height          =   192
               Left            =   7992
               TabIndex        =   70
               Top             =   528
               Width           =   972
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Inafecto :"
               ForeColor       =   &H00800000&
               Height          =   192
               Left            =   5340
               TabIndex        =   69
               Top             =   528
               Width           =   672
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "I.G.V. :"
               ForeColor       =   &H00800000&
               Height          =   192
               Left            =   2880
               TabIndex        =   68
               Top             =   528
               Width           =   492
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               Caption         =   "Imp. Bruto :"
               ForeColor       =   &H00800000&
               Height          =   192
               Left            =   108
               TabIndex        =   67
               Top             =   528
               Width           =   816
            End
            Begin VB.Shape Shape1 
               BorderColor     =   &H00FFFFFF&
               Height          =   12
               Left            =   96
               Top             =   444
               Width           =   10920
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Cuenta :"
               Height          =   192
               Left            =   120
               TabIndex        =   39
               Top             =   180
               Width           =   600
            End
         End
         Begin VB.Shape Shilu1 
            BorderColor     =   &H0000FF00&
            BorderWidth     =   2
            Height          =   36
            Left            =   1608
            Top             =   12
            Visible         =   0   'False
            Width           =   9636
         End
      End
      Begin VB.Frame FrameConsulta 
         BackColor       =   &H00808080&
         Height          =   5664
         Left            =   48
         TabIndex        =   45
         Top             =   1380
         Width           =   11250
         Begin TextFer.TxFer TxEjecutar 
            Height          =   300
            Left            =   120
            TabIndex        =   58
            Top             =   465
            Width           =   7485
            _ExtentX        =   13208
            _ExtentY        =   529
            Object.CausesValidation=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            ColorIlumina    =   -2147483624
            Valor           =   ""
         End
         Begin VB.CheckBox ChkTodos 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00808080&
            Caption         =   "Todos"
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   7650
            TabIndex        =   52
            Top             =   480
            Width           =   855
         End
         Begin MSDataListLib.DataCombo Dtc_Campo 
            Height          =   315
            Left            =   9375
            TabIndex        =   51
            Top             =   435
            Width           =   1785
            _ExtentX        =   3154
            _ExtentY        =   508
            _Version        =   393216
            Style           =   2
            ListField       =   "nombre"
            BoundColumn     =   "codigo"
            Text            =   ""
         End
         Begin TrueOleDBGrid70.TDBGrid TDBG_Consulta 
            Height          =   4368
            Left            =   120
            TabIndex        =   46
            Top             =   816
            Width           =   11040
            _ExtentX        =   19473
            _ExtentY        =   7705
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Nº Provisión"
            Columns(0).DataField=   "cabprovinumero"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Fecha Cont."
            Columns(1).DataField=   "cabprovifchconta"
            Columns(1).NumberFormat=   "Short Date"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Total Bruto"
            Columns(2).DataField=   "cabprovitotbru"
            Columns(2).NumberFormat=   "###,###,###,###.00"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Total IGV"
            Columns(3).DataField=   "cabprovitotigv"
            Columns(3).NumberFormat=   "###,###,###,###.00"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Total Infecto"
            Columns(4).DataField=   "cabprovitotinaf"
            Columns(4).NumberFormat=   "###,###,###,###.00"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Total Compra"
            Columns(5).DataField=   "cabprovitotal"
            Columns(5).NumberFormat=   "###,###,###,###.00"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Numero Auxiliar"
            Columns(6).DataField=   "cabprovinumaux"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   7
            Splits(0)._UserFlags=   0
            Splits(0).AllowRowSizing=   0   'False
            Splits(0).RecordSelectorWidth=   508
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   13160660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=7"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2011"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1926"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2180"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2096"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(9)=   "Column(2).Width=2942"
            Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2858"
            Splits(0)._ColumnProps(12)=   "Column(2)._ColStyle=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(14)=   "Column(3).Width=3027"
            Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=2942"
            Splits(0)._ColumnProps(17)=   "Column(3)._ColStyle=2"
            Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(19)=   "Column(4).Width=2794"
            Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=2709"
            Splits(0)._ColumnProps(22)=   "Column(4)._ColStyle=2"
            Splits(0)._ColumnProps(23)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(24)=   "Column(5).Width=2836"
            Splits(0)._ColumnProps(25)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(26)=   "Column(5)._WidthInPix=2752"
            Splits(0)._ColumnProps(27)=   "Column(5)._ColStyle=2"
            Splits(0)._ColumnProps(28)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(29)=   "Column(6).Width=2731"
            Splits(0)._ColumnProps(30)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(31)=   "Column(6)._WidthInPix=2646"
            Splits(0)._ColumnProps(32)=   "Column(6).Order=7"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=7.8,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=7.8,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=1,.bgcolor=&HE1FFFF&"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=1,.bgcolor=&HFAF7B4&"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=1,.bgcolor=&HE1FFFF&"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=1,.bgcolor=&HE1FFFF&"
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
         Begin VB.Shape Shape8 
            BackColor       =   &H8000000B&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   96
            Left            =   0
            Top             =   -36
            Width           =   11268
         End
         Begin VB.Shape Shape7 
            BorderColor     =   &H00404040&
            Height          =   288
            Left            =   10068
            Top             =   5244
            Width           =   1092
         End
         Begin VB.Label lbl_nregconsulta 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "0 "
            Height          =   288
            Left            =   10080
            TabIndex        =   50
            Top             =   5244
            Width           =   1080
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Registros :"
            ForeColor       =   &H00FFFFFF&
            Height          =   192
            Left            =   9000
            TabIndex        =   49
            Top             =   5280
            Width           =   972
         End
         Begin VB.Label Label2 
            BackColor       =   &H00808080&
            Caption         =   "Valor :"
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   120
            TabIndex        =   48
            Top             =   210
            Width           =   2085
         End
         Begin VB.Label Label4 
            BackColor       =   &H00808080&
            Caption         =   "Criterio :"
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   8715
            TabIndex        =   47
            Top             =   510
            Width           =   570
         End
      End
      Begin VB.Shape Shilu2 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         Height          =   2505
         Left            =   -63705
         Top             =   4350
         Visible         =   0   'False
         Width           =   30
      End
   End
End
Attribute VB_Name = "frmMantprovision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ClsMM1 As ClsMantMov1
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
    If UCase(VlDocNota) <> "A" Then
        lb_vcambio = Format(XRecuperaTipoCambio(Dtp_FechaDoc, CmbTcambio.ListIndex + 1, VGcnxCT), "#0.000 ")
      Else
        If IsNull(Dtp_FechaDocRef) Then
            MsgBox "La Fecha del Documento de Referencia esta en nulo", vbInformation
            Dtp_FechaDocRef.SetFocus
            Exit Sub
        End If
        lb_vcambio = Format(XRecuperaTipoCambio(Dtp_FechaDocRef, CmbTcambio.ListIndex + 1, VGcnxCT), "#0.000 ")
    End If
End Sub
Private Sub CmbTcambio_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call SendKeys("{TAB}")
End Sub
Private Sub CtrAyu_Ccosto_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    If Not VGflaglimpia Then Exit Sub
    Set ClsMM1 = New ClsMantMov1
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, Ccosto)
End Sub
Private Sub CtrAyu_Ccosto_AlNoDevolverNada()
    If Not VGflaglimpia Then Exit Sub
    Set ClsMM1 = New ClsMantMov1
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, cuenta)
End Sub
Private Sub CtrAyu_Cuenta_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    If Not VGflaglimpia Then Exit Sub
    Set ClsMM1 = New ClsMantMov1
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, cuenta)
    If ColecCampos("cuentaestadoccostos") Then
        CtrAyu_Ccosto.Visible = True
        lbccosto.Visible = True
        CtrAyu_Ccosto.xclave = "": CtrAyu_Ccosto.xnombre = ""
      Else
        CtrAyu_Ccosto.Visible = False
        lbccosto.Visible = False
        CtrAyu_Ccosto.xclave = "": CtrAyu_Ccosto.xnombre = ""
    End If
End Sub
Private Sub CtrAyu_Cuenta_AlNoDevolverNada()
    If Not VGflaglimpia Then Exit Sub
    Set ClsMM1 = New ClsMantMov1
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, cuenta)
End Sub
Private Sub CtrAyu_Modoprovi_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Set VGvardllgen = New dllgeneral.dll_general
    ChkCtaCte.Value = IIf(VGvardllgen.ESNULO(ColecCampos("modoprovictacte").Value, 0) = 0, 0, 1)
    ChkRegComp.Value = IIf(VGvardllgen.ESNULO(ColecCampos("modoproviregcom").Value, 0) = 0, 0, 1)
    ChkActCaja.Value = IIf(VGvardllgen.ESNULO(ColecCampos("modoprovitesor").Value, 0) = 0, 0, 1)
    If ChkActCaja.Value = 1 Then
        DTPFechaCaja.Visible = True
        Ctr_AyudaCaja.Visible = True
        LeFechCaja.Visible = True
        LeCaja.Visible = True
      Else
        DTPFechaCaja.Visible = False
        Ctr_AyudaCaja.Visible = False
        LeFechCaja.Visible = False
        LeCaja.Visible = False
    End If
End Sub
Private Sub CtrAyu_Modoprovi_AlNoDevolverNada()
Set VGvardllgen = New dllgeneral.dll_general
    ChkCtaCte.Value = 0
    ChkRegComp.Value = 0
    ChkActCaja.Value = 0
End Sub
Private Sub CtrAyu_Moneda_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    'If CtrAyu_Moneda.xclave = "02" Then
        LeTcambio.Visible = True
        CmbTcambio.Visible = True
        lb_vcambio.Visible = True
        CmbTcambio.ListIndex = 1
        'Call CmbTcambio_Click
'      Else
'        LeTcambio.Visible = False
'        CmbTcambio.Visible = False
'        lb_vcambio.Visible = False
'    End If
End Sub
Private Sub CtrAyu_Proveedor_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    Set VGvardllgen = New dllgeneral.dll_general
    txRuc.Text = VGvardllgen.ESNULO(ColecCampos("clienteruc").Value, "")
    lbTelef.Caption = VGvardllgen.ESNULO(ColecCampos("clientetelefono").Value, "")
End Sub
Private Sub CtrAyu_Proveedor_AlNoDevolverNada()
    txRuc.Text = ""
    lbTelef.Caption = ""
End Sub
Private Sub CtrAyu_TipCompra_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    CtrAyu_TipSubAsi.xclave = "": CtrAyu_TipSubAsi.xnombre = ""
    CtrAyu_TipSubAsi.Filtro = "tipocompracodigo='" & CtrAyu_TipCompra.xclave & "'"
End Sub

Private Sub CtrAyu_TipDoc_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    Set VGvardllgen = New dllgeneral.dll_general
    VlDocNota = VGvardllgen.ESNULO(ColecCampos("tdocumentotipo").Value, "")
End Sub

Private Sub CtrAyu_TipDoc_AlNoDevolverNada()
    Set VGvardllgen = New dllgeneral.dll_general
    VlDocNota = ""
End Sub

Private Sub Dtp_FechaDoc_Change()
    Call CmbTcambio_Click
    DtpFech_Ven.Value = Dtp_FechaDoc
End Sub
Private Sub Dtp_FechaDoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call SendKeys("{TAB}")
End Sub

Private Sub Dtp_FechaDocRef_Change()
    If UCase(VlDocNota) = "A" Then Call CmbTcambio_Click
End Sub

Private Sub Dtp_FechaDocRef_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call SendKeys("{TAB}")
End Sub
Private Sub DtpFech_Ven_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call SendKeys("{TAB}")
End Sub
Private Sub DTPFechaContab_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call SendKeys("{TAB}")
End Sub
Private Sub Form_Activate()
    MDIPrincipal.ToolComprob.Visible = True
    MDIPrincipal.mnu00.Visible = True
    Call PBoton(VlUltAccion)
End Sub
Private Sub Form_Load()
    Top = 0
    Left = 0
    'Inicializo la fechas
    DTPFechaContab.Value = VGParamSistem.FechaTrabajo
    Dtp_FechaDoc.Value = VGParamSistem.FechaTrabajo
    DtpFech_Ven.Value = VGParamSistem.FechaTrabajo
    IMant = 0
    VlUltAccion = 0
    Set VGvardllgen = New dllgeneral.dll_general
    Set rscabecera = New ADODB.Recordset
    Set ClsMM1 = New ClsMantMov1
    ClsMM1.CargarAyudas
    Set TDBG_Consulta.DataSource = Nothing
    TDBG_Det.FetchRowStyle = True
    Call PrepararTemporalDetalle
    If rsmantenimiento.RecordCount = 0 Then
        Call HabilitarDetalle(False, FramDetalle, Me)
     Else
        Call HabilitarDetalle(True, FramDetalle, Me)
    End If
    Call VGvardllgen.ActivaTab(0, 1, SSTabMant)
    Call GetCamposdeConsulta
End Sub
Private Sub GetCamposdeConsulta()
    Set rscampo = New ADODB.Recordset
    Call rscampo.Fields.Append("codigo", adVarChar, 60)
    Call rscampo.Fields.Append("Nombre", adVarChar, 50)
    rscampo.Open
    rscampo.AddNew
    rscampo!codigo = "cabprovinumero"
    rscampo!nombre = "Nro. de Provision"
    rscampo.Update
    rscampo.AddNew
    rscampo!codigo = "convert(varchar(10),cabprovifchconta,103)"
    rscampo!nombre = "Fecha Contable"
    rscampo!codigo = "cabprovinumaux"
    rscampo!nombre = "Nro. Auxiliar"
    rscampo.Update
    Set Dtc_Campo.RowSource = rscampo
    Dtc_Campo.BoundText = "cabprovinumero"
End Sub
Public Sub AlMoverRegistro()
Dim vardllgen As New dllgeneral.dll_general
Dim pos As Integer
    If VGactulizodoc Then Exit Sub 'Estoy Actualizando documentos
    VGMoverRegistro = True
    On Error Resume Next
    With rsmantenimiento
        CtrAyu_Cuenta.xclave = Trim(!cuentacodigo): CtrAyu_Cuenta.Ejecutar
        TxImpBruto.Text = Format(!impbruto, "###,###,###.00"): TxImpBruto.valor = Format(!impbruto, "#0.00")
        TxIGV.Text = Format(!IGV, "###,###,###.00"): TxIGV.valor = Format(!IGV, "#0.00")
        TxInafecto.Text = Format(!Inafecto, "###,###,###.00"): TxInafecto.valor = Format(!Inafecto, "#0.00")
        TxImpCompra.Text = Format(!ImpCompra, "###,###,###.00"): TxImpCompra.valor = Format(!ImpCompra, "#0.00")
        Txtglosa.Text = !glosa
        CtrAyu_Ccosto.xclave = Trim(vardllgen.ESNULO(!Ccosto, "00")): CtrAyu_Ccosto.Ejecutar
    End With
    VGMoverRegistro = False
End Sub
Private Sub form_Unload(Cancel As Integer)
    MDIPrincipal.ToolComprob.Visible = False
    MDIPrincipal.mnu00.Visible = False
End Sub
Private Sub rsmantenimiento_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If (adReason = adRsnMove Or adReason = adRsnMoveNext) And pRecordset.RecordCount > 0 And adReasonAux <> adRsnAddNew Then
        Call AlMoverRegistro
    End If
    If adReasonAux = adRsnAddNew Then adReasonAux = adRsnMove
End Sub
Private Sub rsmantenimiento_RecordChangeComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    adReasonAux = adReason
End Sub
Private Sub SSTabMant_Click(PreviousTab As Integer)
    If PreviousTab = 0 Then
        CtrAyu_Ccosto.Requerido = False
        CtrAyu_Cuenta.Requerido = True
        CtrAyu_TipDoc.Requerido = True
        CtrAyu_TipRef.Requerido = False
        CtrAyu_Moneda.Requerido = True
        CtrAyu_Proveedor.Requerido = True
        CtrAyu_TipCompra.Requerido = True
        CtrAyu_Modoprovi.Requerido = True
        Ctr_AyudaCaja.Requerido = True
        Ctr_AyudaOficina.Requerido = True
        MDIPrincipal.mnu00_01(9).Visible = True
        le_Mes.Caption = Format(VGParamSistem.Mesproceso, "00")
      Else
        CtrAyu_Ccosto.Requerido = False
        CtrAyu_Cuenta.Requerido = False
        CtrAyu_TipDoc.Requerido = False
        CtrAyu_TipRef.Requerido = False
        CtrAyu_Moneda.Requerido = False
        CtrAyu_Proveedor.Requerido = False
        CtrAyu_TipCompra.Requerido = False
        CtrAyu_Modoprovi.Requerido = False
        Ctr_AyudaCaja.Requerido = False
        Ctr_AyudaOficina.Requerido = False
        If TxEjecutar.Enabled And Me.Visible Then TxEjecutar.SetFocus
        MDIPrincipal.mnu00_01(9).Visible = False
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
Dim sqlcad As String, xasiento As String, xsubasiento As String
    Set rscabecera = New ADODB.Recordset
    Set VGvardllgen = New dllgeneral.dll_general
    If criterio = "" Then
        cad = " where 1=0 "
      Else
        cad = " where  cabprovimes=" & CInt(VGParamSistem.Mesproceso) & " and "
    End If
    If todos Then cad = " where  cabprovimes=" & CInt(VGParamSistem.Mesproceso) & "  "
    sqlcad = "select * from " & VGParamSistem.TablaCabcomprob & " " & cad & " " & criterio
    rscabecera.Open sqlcad, VGcnx, adOpenKeyset, adLockReadOnly
    
    If rscabecera.RecordCount > 0 Then
        lbl_nregconsulta.Caption = Format(rscabecera.RecordCount, "0 ")
        TDBG_Consulta.SetFocus
      Else
        lbl_nregconsulta.Caption = Format(0, "0 ")
        TxEjecutar.SetFocus
    End If
    Set TDBG_Consulta.DataSource = rscabecera
End Sub
Public Sub CalcularTotales(ByVal Rs As Recordset)
Dim rsaux As ADODB.Recordset
Dim vardllgen As New dllgeneral.dll_general
Set rsaux = Rs.Clone(adLockReadOnly)
If rsaux Is Nothing And IMant <> 1 Then Exit Sub
Dim ximpbruto As Double, xigv As Double
Dim xinafecto As Double, ximpcompra As Double

ximpbruto = 0: xigv = 0:
xinafecto = 0: ximpcompra = 0:
rsaux.MoveFirst
    While Not rsaux.EOF
        ximpbruto = ximpbruto + vardllgen.ESNULO(rsaux!impbruto, 0)
        xigv = xigv + vardllgen.ESNULO(rsaux!IGV, 0)
        xinafecto = xinafecto + vardllgen.ESNULO(rsaux!Inafecto, 0)
        ximpcompra = ximpcompra + vardllgen.ESNULO(rsaux!ImpCompra, 0)
        rsaux.MoveNext
    Wend
    TxTotBruto.Text = Format(ximpbruto, "###,###,###,###.00 ") ' Debe
    TxTotBruto.valor = Format(ximpbruto, "#0.00 ") ' Debe
    
    TxTotIGV.Text = Format(xigv, "###,###,###,###.00 ") ' Debe
    TxTotIGV.valor = Format(xigv, "#0.00 ") ' Debe
    
    TxTotInafecto.Text = Format(xinafecto, "###,###,###,###.00 ") ' Debe
    TxTotInafecto.valor = Format(xinafecto, "#0.00 ") ' Debe
    
    TxTotImpCompra.Text = Format(ximpcompra, "###,###,###,###.00 ") ' Debe
    TxTotImpCompra.valor = Format(ximpcompra, "#0.00 ") ' Debe
    
End Sub
Private Sub Mostrar()
    If rscabecera.State = 0 Then Exit Sub
    If rscabecera.RecordCount = 0 Then Exit Sub
    Set ClsMM1 = New ClsMantMov1
    Set VGvardllgen = New dllgeneral.dll_general
    Call ClearControlsInframe(FrameCabecera, Me)
    Call ClsMM1.MostrarCabecera(rscabecera.Fields)
    Call ClsMM1.Limpia
    Call PrepararTemporalDetalle
    Call ClsMM1.MostrarDetalle(rsmantenimiento)
    Call HabilitarDetalle(True, FramDetalle, Me)
    Call VGvardllgen.ActivaTab(1, 1, SSTabMant)
    VlUltAccion = 4
    Call PBoton(VlUltAccion)
    If VGParamCompra.Auxaut Then
        TxNAux.Locked = True
     Else
        TxNAux.Locked = False
    End If
    'Comprobante Contable :
    StBar.Panels(2).Text = " Comprobante Contable : " & VlComprob_Conta
    Vlnaux = Trim(TxNAux.Text)
    VlDocAnt = Trim(CtrAyu_Proveedor.xclave) & "-" & Trim(CtrAyu_TipDoc.xclave) & "-" & Trim(TxSerie.Text) & IIf(Trim(TxSerie.Text) = "", "", "-") & Trim(TxNdoc.Text)
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
    Set ClsMM1 = New ClsMantMov1
    Call ClsMM1.LimpiarCab
    Call ClsMM1.Limpia
    Call VGvardllgen.ActivaTab(1, 1, SSTabMant)
    Call HabilitarDetalle(False, FramDetalle, Me)
    lbNumComprobCab.Caption = UltNumeroAuto(VGParamSistem.TablaCabcomprob, 1, VGcnx) + 1
    IMant = 1
    VlUltAccion = 1
    
    If VGParamCompra.Auxaut Then
        TxNAux.Locked = True
     Else
        TxNAux.Locked = False
    End If
    TxNAux.Text = ClsMM1.NumeroAuxiliar(CInt(VGParamSistem.Mesproceso))
    CtrAyu_Ccosto.Visible = False
    lbccosto.Visible = False
    CtrAyu_Modoprovi.SetFocus
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
    Set ClsMM1 = New ClsMantMov1
    If Not ClsMM1.ValidarGrabarCabecera(rsmantenimiento.RecordCount) Then Exit Sub
    If Not ClsMM1.ValidarRsDetalle(rsmantenimiento) Then Exit Sub
    xcon = rsmantenimiento.RecordCount
    rsmantenimiento.Filter = "(Impcompra<>0)"
    If rsmantenimiento.RecordCount < 1 Then
        MsgBox "Por lo Menos debe Existir dos registro con valores ", vbExclamation
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
    VGcnxMarfice.BeginTrans 'Inicio la transaccion
    Screen.MousePointer = vbHourglass
    '1=>Paso Genera el Correlativo del Comprobante
    Dim xnumero As Long
    If IMant = 1 Then
        If VGParamCompra.Auxaut Then
            xnumerocompro = ClsMM1.NumeroAuxiliar(CInt(VGParamSistem.Mesproceso), xnumero)
          Else
            xnumerocompro = Trim(TxNAux.Text)
            'Validar si el Numero ya ha sido ingresado
            If ExisteSQL(VGcnx, "Select * From co_cabprovi" & VGParamSistem.Anoproceso & _
                               " Where cabprovinumaux='" & xnumerocompro & "'") Then
                MsgBox "El Numero de Comprobante Auxiliar ya ha sido ingresado", vbExclamation
                TxNAux.SetFocus
                Exit Sub
            End If
            
        End If
        '2=>Paso Actualizo el Correlativo en la Tabla SubAsiento si es que ingrese un nuevo
        'Comprobante
        Call ClsMM1.ActualizaCorrelAuxiliar(xnumero)
        If Not VGvarVerifica Then varnerror = 6: GoTo ErrorGrabar
      Else
        'Validar si el Numero ya ha sido ingresado cuando esta siendo modificado
        If Vlnaux <> Trim(TxNAux.Text) Then
            If ExisteSQL(VGcnx, "Select * From co_cabprovi" & VGParamSistem.Anoproceso & _
                               " Where cabprovinumaux='" & Trim(TxNAux.Text) & "'") Then
                MsgBox "El Numero de Comprobante Auxiliar ya ha sido ingresado", vbExclamation
                TxNAux.SetFocus
                Exit Sub
            End If
        End If
        xnumerocompro = Trim(TxNAux.Text)
    End If
    If Not VGvarVerifica Then varnerror = 1: GoTo ErrorGrabar
    
    '2=>Paso Grabo la Cabecera del Comprobante
    Dim Xnumtesor As String
    If ChkActCaja.Value = 1 Then
        Call ClsMM1.Grabaren_Tesoreria(IMant, UltNumeroAuto(VGParamSistem.TablaCabcomprob, 1, VGcnx), Xnumtesor)
    End If
        
    Call ClsMM1.GrabarCabecera(IMant, 0, Format(CInt(VGParamSistem.Mesproceso), "00") & xnumerocompro, Xnumtesor)
    If Not VGvarVerifica Then varnerror = 2: GoTo ErrorGrabar
    
    If ChkCtaCte.Value = 1 Then
        Call ClsMM1.GrabarCP_Cargo(IMant, UltNumeroAuto(VGParamSistem.TablaCabcomprob, 1, VGcnx))
    End If
    
    '3=>Paso Grabo los Detalle del Comprobante
    Call ClsMM1.GrabarDetalle(rsmantenimiento)
    If Not VGvarVerifica Then varnerror = 3: GoTo ErrorGrabar
    
    '4=>Generar Asiento en Linea segun parametro
    Call ClsMM1.GeneraAsientoenLine(IMant, UltNumeroAuto(VGParamSistem.TablaCabcomprob, 1, VGcnx), VlComprob_Conta)
    If Not VGvarVerifica Then varnerror = 4: GoTo ErrorGrabar
                
    VGcnxMarfice.CommitTrans 'Acepto toda la transaccion porque es correcta
    If IMant = 1 Then
        MsgBox "Se grabo Satisfactoriamente  El numero de Comprobante Generado Es :" & Chr(13) & _
           "Nro: " & UltNumeroAuto(VGParamSistem.TablaCabcomprob, 1, VGcnx) & Chr(13) & _
           "El Numero Auxiliar Generado es : " & Format(CInt(VGParamSistem.Mesproceso), "00") & xnumerocompro
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
            VGcnxMarfice.RollbackTrans
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
    Set ClsMM1 = New ClsMantMov1
    'Verificar si es que tiene abonos
    If consultadoctesor(Trim(CtrAyu_Proveedor.xclave), CtrAyu_TipDoc.xclave, Trim(TxSerie.Text) & "-" & Trim(TxNdoc.Text)) Then
        MsgBox "No se puede eliminar este documento porque tiene abonos " & Chr(13) & _
               "Anula primero desde el sistema de tesoreria  ", vbExclamation
        Exit Sub
    End If
    VGcnxMarfice.BeginTrans
    Screen.MousePointer = vbHourglass
    If ChkActCaja.Value = 1 Then 'Este en el caso que actualice caja
        Call ClsMM1.Grabaren_Tesoreria(3, CInt(Trim(lbNumComprobCab.Caption)))
    End If
    If ChkCtaCte.Value = 1 Then 'Y esto para actualizar cuenta corriente
        Call ClsMM1.GrabarCP_Cargo(3, CInt(Trim(lbNumComprobCab.Caption)))
    End If
    Call ClsMM1.GrabarCabecera(3, Trim(lbNumComprobCab.Caption))
    Screen.MousePointer = vbHourglass
    
    Dim sqlcad As String
    sqlcad = "" & _
    " Update dbo.ct_cabcomprob" & VGParamSistem.Anoproceso & _
    " Set cabcomprobtotdebe=0, " & _
    "     cabcomprobtothaber=0," & _
    "     cabcomprobtotussdebe=0, " & _
    "     cabcomprobtotusshaber = 0 " & _
    " Where cabcomprobnumero='" & VlComprob_Conta & "' " & Chr(13) & _
    " Update dbo.ct_detcomprob" & VGParamSistem.Anoproceso & _
    "   Set detcomprobdebe=0, " & _
    "   detcomprobhaber=0, " & _
    "   detcomprobussdebe=0, " & _
    "   detcomprobusshaber = 0 " & _
    " Where cabcomprobnumero='" & VlComprob_Conta & "' "
    VGcnxCT.Execute sqlcad
    
    
    VGcnxMarfice.CommitTrans
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
Public Sub Cancelar(Optional OP As Integer)
Set VGvardllgen = New dllgeneral.dll_general

    If OP <> 1 Then
        If MsgBox("Esta Seguro que Desea Cancelar la Operación ", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
            'Resolver el problema que el cursor debe parpadear donde se ha quedado
            Exit Sub
        End If
    End If
        
    If SSTabMant.Tab = 1 Then
        Call VGvardllgen.ActivaTab(0, 1, SSTabMant)
        VlUltAccion = 5
        Set rsmantenimiento = Nothing
    End If
    
End Sub
Public Sub AñadirDetalle()
    Set ClsMM1 = New ClsMantMov1
    If rsmantenimiento.RecordCount > 0 Then
        If Not ClsMM1.ValidarGrabarDetalle Then Exit Sub
    End If
    Call HabilitarDetalle(True, FramDetalle, Me)
    Call ClsMM1.AñadiralDetalle(rsmantenimiento)
    lbnregdetalle.Caption = Format(rsmantenimiento.RecordCount, "0 ")
    CtrAyu_Cuenta.SetFocus
End Sub
Public Sub EliminarDetalle()
Dim reg As Long
Dim num As Integer
    Screen.MousePointer = 11
    On Error Resume Next
    If rsmantenimiento.State = 0 Then Exit Sub
    If rsmantenimiento.RecordCount = 0 Then Exit Sub
    Set ClsMM1 = New ClsMantMov1
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
Public Sub imprimir()
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
    arrparm(2) = Trim(VGParamSistem.Anoproceso)
    arrparm(3) = CInt(Trim(VGParamSistem.Mesproceso))
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
            Call imprimir
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
    End Select
End Sub

Private Sub TxIGV_Change()
    If Not VGflaglimpia Then Exit Sub
    If VGMoverRegistro Then Exit Sub
    Set ClsMM1 = New ClsMantMov1
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, IGV)
    Call CalculoCompra
    Call CalcularTotales(rsmantenimiento)
End Sub

Private Sub TxImpBruto_Change()
    If Not VGflaglimpia Then Exit Sub
    If VGMoverRegistro Then Exit Sub
    Set VGvardllgen = New dllgeneral.dll_general
    Set ClsMM1 = New ClsMantMov1
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, impbruto)
    TxIGV.valor = CDbl(VGvardllgen.ESNULO(TxImpBruto.valor, 0)) * VGParamCompra.IGV
    TxIGV.Text = Format(TxIGV.valor, "###,###,###.00")
    Call CalculoCompra
    Call CalcularTotales(rsmantenimiento)
End Sub
Private Sub CalculoCompra()
    TxImpCompra.valor = CDbl(VGvardllgen.ESNULO(Espunto(TxImpBruto.valor), 0)) + CDbl(VGvardllgen.ESNULO(Espunto(TxIGV.valor), 0)) + CDbl(VGvardllgen.ESNULO(Espunto(TxInafecto.valor), 0))
    TxImpCompra.Text = Format(Espunto(TxImpCompra.valor), "###,###,###.00")
End Sub

Private Sub TxImpCompra_Change()
    If Not VGflaglimpia Then Exit Sub
    If VGMoverRegistro Then Exit Sub
    Set ClsMM1 = New ClsMantMov1
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, ImpCompra)
    Call CalcularTotales(rsmantenimiento)
End Sub

Private Sub TxInafecto_Change()
    If Not VGflaglimpia Then Exit Sub
    If VGMoverRegistro Then Exit Sub
    Set ClsMM1 = New ClsMantMov1
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, Inafecto)
    Call CalculoCompra
    Call CalcularTotales(rsmantenimiento)
End Sub

Private Sub TxNref_LostFocus()
    Dim DocAct As String
    If VlDocNota <> "A" Then Exit Sub
    DocAct = Trim(CtrAyu_Proveedor.xclave) & "-" & Trim(CtrAyu_TipRef.xclave) & "-" & Trim(TxNref.Text)
    If Not ExisteSQL(VGcnx, " Select * From dbo.co_cabprovi" & VGParamSistem.Anoproceso & _
                     " Where " & _
                     " isnull(proveedorcodigo,'')+'-'+isnull(documetocodigo,'')+'-'+cabprovinumdoc='" & DocAct & "'") Then
           MsgBox "El Documento de Referencia No Existe", vbExclamation
           'TxNref.Text = ""
           Exit Sub
    End If
End Sub
Private Sub Txtglosa_Change()
    If Not VGflaglimpia Then Exit Sub
    If VGMoverRegistro Then Exit Sub
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, glosa)
End Sub
Function consultadoctesor(Proveedor As String, Td As String, Ndocumento As String) As Boolean
Dim sqlcad As String
Dim rsaux As ADODB.Recordset
    Set rsaux = New ADODB.Recordset
    sqlcad = " Select * From " & _
             VGParamSistem.BDEmpresa & ".dbo.te_cabecerarecibos A, " & _
             VGParamSistem.BDEmpresa & ".dbo.te_detallerecibos B " & _
             " Where A.cabrec_numrecibo=b.cabrec_numrecibo and " & _
             " A.clientecodigo='" & Trim(Proveedor) & "' and " & _
             " B.detrec_tipodoc_concepto='" & Trim(Td) & "'  and " & _
             " B.detrec_numdocumento=dbo.fn_coviertenumdoc('" & Ndocumento & "') and isnull(cabrec_estadoreg,0)<> 1 "
    rsaux.Open sqlcad, VGcnxMarfice, adOpenKeyset, adLockReadOnly
    If rsaux.RecordCount >= 1 Then
        consultadoctesor = True
        Else
        consultadoctesor = False
    End If
End Function
