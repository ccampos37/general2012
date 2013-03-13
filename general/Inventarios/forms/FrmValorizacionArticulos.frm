VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Begin VB.Form FrmValorizacionArticulos 
   Caption         =   "Form1"
   ClientHeight    =   7200
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   6975
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   11445
      _ExtentX        =   20188
      _ExtentY        =   12303
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Documentos a Valorizar"
      TabPicture(0)   =   "FrmValorizacionArticulos.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Datos por Documento"
      TabPicture(1)   =   "FrmValorizacionArticulos.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "TDBG_concil"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "FrameCabecera"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "framTotales"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame5"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Datos Complementarios"
      TabPicture(2)   =   "FrmValorizacionArticulos.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.Frame Frame5 
         Height          =   975
         Left            =   7920
         TabIndex        =   79
         Top             =   5880
         Width           =   2055
         Begin VB.CommandButton Cmdgraba 
            Caption         =   "&Grabar"
            Height          =   675
            Left            =   240
            Picture         =   "FrmValorizacionArticulos.frx":0054
            Style           =   1  'Graphical
            TabIndex        =   80
            Top             =   240
            Width           =   775
         End
         Begin VB.CommandButton CmdSalir2 
            Caption         =   "&Salir"
            Height          =   675
            Left            =   1080
            Picture         =   "FrmValorizacionArticulos.frx":0496
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   240
            Width           =   775
         End
      End
      Begin VB.Frame framTotales 
         BackColor       =   &H00C0C0C0&
         Height          =   384
         Left            =   0
         TabIndex        =   71
         Top             =   3480
         Width           =   11235
         Begin TextFer.TxFer TxTotBruto 
            Height          =   300
            Left            =   6645
            TabIndex        =   72
            Top             =   45
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
            Alignment       =   1
            Appearance      =   0
            BackColor       =   14679546
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
            Left            =   7815
            TabIndex        =   73
            Top             =   45
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   529
            Alignment       =   1
            Appearance      =   0
            BackColor       =   16777152
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
            Left            =   8655
            TabIndex        =   74
            Top             =   45
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   529
            Alignment       =   1
            Appearance      =   0
            BackColor       =   14679546
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
         Begin TextFer.TxFer TxTotTotal 
            Height          =   300
            Left            =   9930
            TabIndex        =   75
            Top             =   45
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
            Alignment       =   1
            Appearance      =   0
            BackColor       =   14679546
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
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "Totales Generales ........."
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
            Height          =   210
            Left            =   465
            TabIndex        =   77
            Top             =   360
            Width           =   2805
         End
         Begin VB.Label Label24 
            BackColor       =   &H00344A87&
            Height          =   2760
            Left            =   75
            TabIndex        =   76
            Top             =   90
            Width           =   11430
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Totales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   975
         Left            =   120
         TabIndex        =   66
         Top             =   5880
         Width           =   6615
         Begin TextFer.TxFer TxtImpbruto 
            Height          =   300
            Left            =   240
            TabIndex        =   18
            Top             =   480
            Width           =   1140
            _ExtentX        =   2011
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
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            TipoDato        =   1
            NumeroDecimales =   4
            SignoNegativo   =   0   'False
            MarcarTextoAlEnfoque=   -1  'True
         End
         Begin TextFer.TxFer Txtimpigv 
            Height          =   300
            Left            =   1800
            TabIndex        =   19
            Top             =   480
            Width           =   1140
            _ExtentX        =   2011
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
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            TipoDato        =   1
            NumeroDecimales =   4
            SignoNegativo   =   0   'False
         End
         Begin TextFer.TxFer Txtimpinafecto 
            Height          =   300
            Left            =   3240
            TabIndex        =   20
            Top             =   480
            Width           =   1500
            _ExtentX        =   2646
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
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            TipoDato        =   1
            NumeroDecimales =   4
         End
         Begin TextFer.TxFer Txtimptotal 
            Height          =   300
            Left            =   5280
            TabIndex        =   21
            Top             =   480
            Width           =   1260
            _ExtentX        =   2223
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
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            TipoDato        =   1
            NumeroDecimales =   4
            SignoNegativo   =   0   'False
         End
         Begin VB.Label Label23 
            Caption         =   "Valor Neto"
            Height          =   255
            Index           =   3
            Left            =   5400
            TabIndex        =   70
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label23 
            Caption         =   "Valor Inafecto"
            Height          =   255
            Index           =   2
            Left            =   3360
            TabIndex        =   69
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label23 
            Caption         =   "Valor IGV"
            Height          =   255
            Index           =   1
            Left            =   1800
            TabIndex        =   68
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label23 
            Caption         =   "Valor Imponible"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   67
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame FrameCabecera 
         Height          =   2955
         Left            =   0
         TabIndex        =   30
         Top             =   480
         Width           =   11265
         Begin VB.CheckBox Check1 
            Caption         =   "Principal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   360
            TabIndex        =   78
            Top             =   0
            Width           =   1575
         End
         Begin VB.ComboBox CmbTcambio 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "FrmValorizacionArticulos.frx":08D8
            Left            =   4530
            List            =   "FrmValorizacionArticulos.frx":08E5
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   2490
            Visible         =   0   'False
            Width           =   1755
         End
         Begin VB.CheckBox ChkOperGrab 
            Caption         =   "Operación Grabada"
            ForeColor       =   &H00000080&
            Height          =   270
            Left            =   3750
            TabIndex        =   35
            Top             =   435
            Value           =   1  'Checked
            Width           =   1830
         End
         Begin VB.TextBox TxTesor 
            Height          =   285
            Left            =   5760
            TabIndex        =   34
            Top             =   405
            Visible         =   0   'False
            Width           =   1470
         End
         Begin VB.CheckBox ChkCtaCte 
            Alignment       =   1  'Right Justify
            Caption         =   "Cuenta Cte."
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   144
            TabIndex        =   33
            Top             =   1770
            Width           =   1140
         End
         Begin VB.CheckBox ChkRegComp 
            Alignment       =   1  'Right Justify
            Caption         =   "Regist. Compra"
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   1380
            TabIndex        =   32
            Top             =   1770
            Width           =   1395
         End
         Begin VB.CheckBox ChkActCaja 
            Alignment       =   1  'Right Justify
            Caption         =   "Actualiza Caja"
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   2892
            TabIndex        =   31
            Top             =   1770
            Visible         =   0   'False
            Width           =   1305
         End
         Begin MSComCtl2.DTPicker DTPFechaCaja 
            Height          =   300
            Left            =   1305
            TabIndex        =   6
            Top             =   2115
            Visible         =   0   'False
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   529
            _Version        =   393216
            Format          =   17563649
            CurrentDate     =   37617
         End
         Begin TextFer.TxFer TxNdoc 
            Height          =   300
            Left            =   9120
            TabIndex        =   15
            Top             =   1950
            Width           =   1995
            _ExtentX        =   3519
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
            MaxLength       =   8
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
         End
         Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_TipDoc 
            Height          =   315
            Left            =   8640
            TabIndex        =   13
            Top             =   1590
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            XcodMaxLongitud =   2
            NomTabla        =   "cp_tipodocumento"
            TituloAyuda     =   "Busqueda de Tipo de  Documento"
            ListaCampos     =   "tdocumentocodigo(1),tdocumentodescripcion(1),tdocumentotipo(1),documentoretencion(1)"
            XcodCampo       =   "tdocumentocodigo"
            XListCampo      =   "tdocumentodescripcion"
            ListaCamposDescrip=   "Código,Descripción,CargoAbono,Retencion"
            ListaCamposText =   "tdocumentocodigo,tdocumentodescripcion,tdocumentotipo,documentoretencion"
         End
         Begin MSComCtl2.DTPicker Dtp_FechaDoc 
            Height          =   315
            Left            =   8640
            TabIndex        =   16
            Top             =   2265
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            Format          =   17563649
            CurrentDate     =   37469
         End
         Begin MSComCtl2.DTPicker DtpFech_Ven 
            Height          =   315
            Left            =   8640
            TabIndex        =   17
            Top             =   2595
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            Format          =   17563649
            CurrentDate     =   37469
         End
         Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_Moneda 
            Height          =   315
            Left            =   1065
            TabIndex        =   8
            Top             =   2520
            Width           =   2415
            _ExtentX        =   4260
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
         Begin TextFer.TxFer txRuc 
            Height          =   300
            Left            =   1290
            TabIndex        =   4
            Top             =   1425
            Width           =   1455
            _ExtentX        =   2566
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
            Left            =   8625
            TabIndex        =   14
            Top             =   1935
            Width           =   420
            _ExtentX        =   741
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
            MaxLength       =   3
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
         End
         Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_Proveedor 
            Height          =   315
            Left            =   1305
            TabIndex        =   3
            Top             =   1065
            Width           =   4065
            _ExtentX        =   7170
            _ExtentY        =   556
            XcodMaxLongitud =   11
            xcodwith        =   1200
            NomTabla        =   "cp_proveedor"
            TituloAyuda     =   "Busqueda de Proveedor"
            ListaCampos     =   "clientecodigo(1),clienterazonsocial(1),clienteruc(1),clientetelefono(1),proveedorcontribuyente(2)"
            XcodCampo       =   "clientecodigo"
            XListCampo      =   "clienterazonsocial"
            ListaCamposDescrip=   "Codigo,Descripcion"
            ListaCamposText =   "clientecodigo,clienterazonsocial,clienteruc,clientetelefono,proveedorcontribuyente"
         End
         Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_Modoprovi 
            Height          =   315
            Left            =   1305
            TabIndex        =   1
            Top             =   720
            Width           =   2370
            _ExtentX        =   4180
            _ExtentY        =   556
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
            Top             =   705
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   529
            _Version        =   393216
            Format          =   17563649
            CurrentDate     =   37489
         End
         Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_TipCompra 
            Height          =   315
            Left            =   8640
            TabIndex        =   10
            Top             =   420
            Width           =   2550
            _ExtentX        =   4498
            _ExtentY        =   556
            XcodMaxLongitud =   2
            NomTabla        =   "co_tipocompra"
            TituloAyuda     =   "Busqueda de Tipo de Compra"
            ListaCampos     =   "tipocompracodigo(1), tipocompradesc(1),tipocomprainafecta(1)"
            XcodCampo       =   "tipocompracodigo"
            XListCampo      =   "tipocompradesc"
            ListaCamposDescrip=   "Código,Descripción"
            ListaCamposText =   "tipocompracodigo, tipocompradesc,tipocomprainafecta"
         End
         Begin TextFer.TxFer TxNAux 
            Height          =   300
            Left            =   5712
            TabIndex        =   36
            Top             =   1548
            Width           =   1572
            _ExtentX        =   2778
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
            MaxLength       =   5
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
         End
         Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_TipSubAsi 
            Height          =   315
            Left            =   8640
            TabIndex        =   11
            Top             =   795
            Visible         =   0   'False
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   556
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
            Left            =   8640
            TabIndex        =   12
            Top             =   1200
            Width           =   2415
            _ExtentX        =   4260
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
            Height          =   315
            Left            =   4395
            TabIndex        =   7
            Top             =   2130
            Visible         =   0   'False
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
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuempresa 
            Height          =   315
            Left            =   3720
            TabIndex        =   5
            Top             =   1440
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   556
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
         Begin VB.Shape Shape3 
            BorderColor     =   &H00808080&
            Height          =   2628
            Left            =   7392
            Top             =   276
            Width           =   12
         End
         Begin VB.Label lbNumComprobCab 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2FDFE&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0000010000"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1335
            TabIndex        =   59
            Top             =   420
            Width           =   2295
         End
         Begin VB.Label leNComprob 
            AutoSize        =   -1  'True
            Caption         =   "NUMERO :"
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   58
            Top             =   435
            Width           =   810
         End
         Begin VB.Label lendocum 
            AutoSize        =   -1  'True
            Caption         =   "Nº doc. :"
            Height          =   195
            Left            =   7560
            TabIndex        =   57
            Top             =   2040
            Width           =   630
         End
         Begin VB.Label letipdoc 
            Caption         =   "Tipo Doc. :"
            Height          =   255
            Left            =   7560
            TabIndex        =   56
            Top             =   1650
            Width           =   840
         End
         Begin VB.Label leFechaDoc 
            AutoSize        =   -1  'True
            Caption         =   "Fecha doc. :"
            Height          =   195
            Left            =   7545
            TabIndex        =   55
            Top             =   2340
            Width           =   900
         End
         Begin VB.Label leFechVen 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Venc. :"
            Height          =   195
            Left            =   7545
            TabIndex        =   54
            Top             =   2670
            Width           =   1005
         End
         Begin VB.Label LeMon 
            AutoSize        =   -1  'True
            Caption         =   "Moneda :"
            Height          =   195
            Left            =   150
            TabIndex        =   53
            Top             =   2550
            Width           =   675
         End
         Begin VB.Label lb_vcambio 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00F2FEFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   6360
            TabIndex        =   52
            Top             =   2490
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label LeTcambio 
            AutoSize        =   -1  'True
            Caption         =   "T/Cambio :"
            Height          =   195
            Left            =   3585
            TabIndex        =   51
            Top             =   2550
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.Label Leruc 
            AutoSize        =   -1  'True
            Caption         =   "Nº. de R.U.C. :"
            Height          =   195
            Left            =   150
            TabIndex        =   50
            Top             =   1500
            Width           =   1050
         End
         Begin VB.Label Le_Proveedor 
            Caption         =   "Proveedor :"
            Height          =   255
            Left            =   150
            TabIndex        =   49
            Top             =   1110
            Width           =   1020
         End
         Begin VB.Label LeModComp 
            Caption         =   "Modo Compra :"
            Height          =   255
            Left            =   150
            TabIndex        =   48
            Top             =   765
            Width           =   1125
         End
         Begin VB.Label LeTelf 
            AutoSize        =   -1  'True
            Caption         =   "Telef:"
            Height          =   195
            Left            =   5385
            TabIndex        =   47
            Top             =   1140
            Width           =   405
         End
         Begin VB.Label lbTelef 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   5805
            TabIndex        =   46
            Top             =   1080
            Width           =   1470
         End
         Begin VB.Label leFecha 
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Contable :"
            Height          =   225
            Left            =   3750
            TabIndex        =   45
            Top             =   780
            Width           =   1320
         End
         Begin VB.Shape Shape11 
            BorderColor     =   &H00FFFFFF&
            Height          =   4350
            Left            =   7410
            Top             =   -2835
            Width           =   15
         End
         Begin VB.Label LeTipComp 
            AutoSize        =   -1  'True
            Caption         =   "Tip. Compra :"
            Height          =   195
            Left            =   7455
            TabIndex        =   44
            Top             =   450
            Width           =   945
         End
         Begin VB.Label LeNaux 
            AutoSize        =   -1  'True
            Caption         =   "Nº Aux :"
            Height          =   195
            Left            =   4740
            TabIndex        =   43
            Top             =   1845
            Width           =   585
         End
         Begin VB.Label leSubAsi 
            AutoSize        =   -1  'True
            Caption         =   "Tip. Sub :"
            Height          =   195
            Left            =   7485
            TabIndex        =   42
            Top             =   840
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.Label le_Mes 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5355
            TabIndex        =   41
            Top             =   1800
            Width           =   360
         End
         Begin VB.Label Leoficina 
            Caption         =   "Oficina :"
            Height          =   255
            Left            =   7530
            TabIndex        =   40
            Top             =   1245
            Width           =   840
         End
         Begin VB.Label LeFechCaja 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Caja :"
            Height          =   195
            Left            =   150
            TabIndex        =   39
            Top             =   2160
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Label LeCaja 
            AutoSize        =   -1  'True
            Caption         =   "Caja :"
            Height          =   195
            Left            =   3585
            TabIndex        =   38
            Top             =   2175
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.Label Leempresa 
            AutoSize        =   -1  'True
            Caption         =   "Empresa :"
            Height          =   195
            Left            =   2880
            TabIndex        =   37
            Top             =   1500
            Width           =   705
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   5760
         Left            =   -74880
         TabIndex        =   23
         Top             =   360
         Width           =   11100
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   4680
            Style           =   2  'Dropdown List
            TabIndex        =   64
            Top             =   240
            Width           =   2415
         End
         Begin VB.TextBox TxtBuscar 
            Height          =   315
            Left            =   600
            TabIndex        =   63
            Top             =   240
            Width           =   1455
         End
         Begin VB.Frame Frame1 
            Height          =   2535
            Left            =   9600
            TabIndex        =   60
            Top             =   2640
            Width           =   1215
            Begin VB.CommandButton cmdNuevo 
               Caption         =   "&Nuevo"
               Height          =   675
               Left            =   240
               Picture         =   "FrmValorizacionArticulos.frx":0911
               Style           =   1  'Graphical
               TabIndex        =   62
               Top             =   480
               Width           =   775
            End
            Begin VB.CommandButton Command7 
               Caption         =   "&Salir"
               Height          =   675
               Left            =   240
               Picture         =   "FrmValorizacionArticulos.frx":0D53
               Style           =   1  'Graphical
               TabIndex        =   61
               Top             =   1560
               Width           =   775
            End
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "FrmValorizacionArticulos.frx":1195
            Left            =   8880
            List            =   "FrmValorizacionArticulos.frx":119F
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   210
            Width           =   1455
         End
         Begin TrueOleDBGrid70.TDBGrid TDBNota 
            Height          =   1680
            Left            =   240
            TabIndex        =   24
            Top             =   645
            Width           =   10650
            _ExtentX        =   18785
            _ExtentY        =   2963
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "ALM"
            Columns(0).DataField=   "dealma"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "T.Doc"
            Columns(1).DataField=   "td"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Nro.Doc"
            Columns(2).DataField=   "denumdoc"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "F.Doc"
            Columns(3).DataField=   "cafecdoc"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Cod.Prov"
            Columns(4).DataField=   "cacodpro"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Razon Social"
            Columns(5).DataField=   "proveedor"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Doc Ref"
            Columns(6).DataField=   "doc_refe"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Nro Refe"
            Columns(7).DataField=   "nro_refe"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   8
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=8"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=794"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=714"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=953"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=873"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(9)=   "Column(2).Width=2275"
            Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2196"
            Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(13)=   "Column(3).Width=2090"
            Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2011"
            Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(17)=   "Column(4).Width=2408"
            Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2328"
            Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(21)=   "Column(5).Width=5450"
            Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=5371"
            Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(25)=   "Column(6).Width=1429"
            Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=1349"
            Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(29)=   "Column(7).Width=2725"
            Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=2646"
            Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
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
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=66,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=63,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=64,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=65,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=62,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=58,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=46,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=43,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=44,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=45,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=32,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=29,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=30,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=31,.parent=17"
            _StyleDefs(68)  =   "Named:id=33:Normal"
            _StyleDefs(69)  =   ":id=33,.parent=0"
            _StyleDefs(70)  =   "Named:id=34:Heading"
            _StyleDefs(71)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(72)  =   ":id=34,.wraptext=-1"
            _StyleDefs(73)  =   "Named:id=35:Footing"
            _StyleDefs(74)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(75)  =   "Named:id=36:Selected"
            _StyleDefs(76)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(77)  =   "Named:id=37:Caption"
            _StyleDefs(78)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(79)  =   "Named:id=38:HighlightRow"
            _StyleDefs(80)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(81)  =   "Named:id=39:EvenRow"
            _StyleDefs(82)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(83)  =   "Named:id=40:OddRow"
            _StyleDefs(84)  =   ":id=40,.parent=33"
            _StyleDefs(85)  =   "Named:id=41:RecordSelector"
            _StyleDefs(86)  =   ":id=41,.parent=34"
            _StyleDefs(87)  =   "Named:id=42:FilterBar"
            _StyleDefs(88)  =   ":id=42,.parent=33"
         End
         Begin MSFlexGridLib.MSFlexGrid FG 
            Height          =   3210
            Left            =   240
            TabIndex        =   26
            Top             =   2400
            Width           =   9300
            _ExtentX        =   16404
            _ExtentY        =   5662
            _Version        =   393216
            AllowUserResizing=   1
         End
         Begin VB.Label Label21 
            Caption         =   "Indice"
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
            Left            =   7920
            TabIndex        =   29
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label22 
            Caption         =   "Filtro"
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
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label11 
            Caption         =   "Almacen"
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
            Left            =   3840
            TabIndex        =   27
            Top             =   240
            Width           =   735
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBG_concil 
         Height          =   1935
         Left            =   120
         TabIndex        =   65
         Top             =   3840
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   3413
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
         Columns(1).Caption=   "codigo"
         Columns(1).DataField=   "vodigo"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Descripcion"
         Columns(2).DataField=   "descripcion"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "familia"
         Columns(3).DataField=   "familia"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Cod.Gastos"
         Columns(4).DataField=   "Gastos"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Cantidad"
         Columns(5).DataField=   "cantidad"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
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
         Columns(6).Caption=   "valor imponible"
         Columns(6).DataField=   "IMPBRUTO"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Valor Igv"
         Columns(7).DataField=   "impigv"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Valor Inafecto"
         Columns(8).DataField=   "IMPINAFECTO"
         Columns(8).NumberFormat=   "###,###,###.00"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Valor Total"
         Columns(9).DataField=   "imptotal"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   10
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=10"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=847"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=767"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=1482"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1402"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8196"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=4657"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=4577"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=8196"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=1191"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1111"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=8196"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=1376"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=1296"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=8196"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=1376"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=1296"
         Splits(0)._ColumnProps(29)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(30)=   "Column(6).Width=2143"
         Splits(0)._ColumnProps(31)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(32)=   "Column(6)._WidthInPix=2064"
         Splits(0)._ColumnProps(33)=   "Column(6)._ColStyle=1"
         Splits(0)._ColumnProps(34)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(35)=   "Column(6)._FootDivider=0"
         Splits(0)._ColumnProps(36)=   "Column(7).Width=1429"
         Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=1349"
         Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=1"
         Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(41)=   "Column(8).Width=2249"
         Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=2170"
         Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=8193"
         Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(46)=   "Column(9).Width=2143"
         Splits(0)._ColumnProps(47)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(48)=   "Column(9)._WidthInPix=2064"
         Splits(0)._ColumnProps(49)=   "Column(9)._ColStyle=1"
         Splits(0)._ColumnProps(50)=   "Column(9).Order=10"
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
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=46,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=50,.parent=13,.alignment=2,.bgcolor=&HFFFFFF&"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=32,.parent=13,.alignment=2"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=29,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=30,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=31,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=82,.parent=13,.alignment=2,.locked=-1"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=79,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=80,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=81,.parent=17"
         _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=86,.parent=13,.alignment=2"
         _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=83,.parent=14"
         _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=84,.parent=15"
         _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=85,.parent=17"
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
         _StyleDefs(96)  =   ":id=42,.parent=33"
      End
   End
End
Attribute VB_Name = "FrmValorizacionArticulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSQL As String
Dim precio As Double
Dim flaggraba As Integer
Dim VGvarVerifica As Boolean
Dim CANTIDAD As Double
Dim tipcam As Double
Dim rs As Recordset
Public Sqltabla As String
Public sqltabla1 As String
Dim WithEvents rsmantenimiento As ADODB.Recordset
Attribute rsmantenimiento.VB_VarHelpID = -1
Public rsdetalle As New ADODB.Recordset
Public Rs2 As New ADODB.Recordset
Public rscabecera As New ADODB.Recordset
Public rstotal As New ADODB.Recordset
Dim rsNota As ADODB.Recordset

Dim mRsql As String
Dim mRsql1 As String
Dim totdoc As Double
Dim sCodMon As String
Public IMant As Integer
Public Emiteretencion As String
Public tipoinafecto As Integer
Public comprainafecta As String
Public documentoinafecto As String
Public tipodetraccion As Integer
Public emitedetraccion As String
Public buencontribuyente As String

Public VlDocAnt As String
Public estadorendicion As Integer
Public fecharendicion As String
Public numerorendicion As String
Public numerorecibo As String
Public modoproviold As Integer

Public VlDocNota As String
Dim Fecha As Date   'Fecha del documento
Public VGDllGeneral As dllgeneral.dll_general
Dim i0 As Integer
Dim xAlma As String
Dim xDescri_alma As String
Dim rsSTKART As New ADODB.Recordset


Private Sub Cmdgrabaanulacion_Click()

End Sub

Private Sub Cmdgraba_Click()
Call grabar
End Sub

Private Sub cmdNuevo_Click()
Dim SQL As String
Dim Sql1 As String


Set rscabecera = Nothing
Sqltabla = "##valoriza" & ComputerName & ""
sqltabla1 = "##valoriza1" & ComputerName & ""

If ExisteElem(0, VGCNx, Sqltabla) Then
   Set rsmantenimiento = VGCNx.Execute(" drop table " & Sqltabla & "")
End If

Set rsmantenimiento = New ADODB.Recordset
SQL = " select deitem as item,decodigo as codigo,left(adescri,25) as descripcion,afamilia as familia"
SQL = SQL & ",gastoscodigo as Gastos,decantid as cantidad,isnull(deprecio,0)  as Impbruto,"
SQL = SQL & " isnull(deigv,0) as ImpIgv ,isnull(DEPRECI1,0) as ImpInafecto,isnull(deprevta,0) as ImpTotal "
SQL = SQL & " into " & Sqltabla & " from movalmdet a inner join maeart b on a.decodigo=b.acodigo"
SQL = SQL & " left join familia c on b.afamilia=c.fam_codigo"
SQL = SQL & " where dealma='" & rsNota!dealma & "' and detd='" & rsNota!tD & "' and "
SQL = SQL & " denumdoc='" & rsNota!denumdoc & "'"
IMant = 1
Set rsmantenimiento = VGCNx.Execute(SQL)
Set rsmantenimiento = New ADODB.Recordset
SQL = " select * from " & Sqltabla & ""

flaggraba = 0
rsmantenimiento.Open (SQL), VGCNx, adOpenDynamic, adLockBatchOptimistic

SQL = "##provision" & ComputerName
If ExisteElem(0, VGCNx, SQL) Then
   Set rscabecera = VGCNx.Execute(" drop table " & SQL & "")
End If

Sql1 = "Select top 0 * into ##provision" & ComputerName & " from co_cabeceraprovisiones"
SQL = "Select top 0 * into ##provision1" & ComputerName & " from co_cabeceraprovisiones"

Set rscabecera = New ADODB.Recordset

rscabecera.Open (Sql1), VGCNx, adOpenDynamic, adLockBatchOptimistic

TDBG_concil.DataSource = rsmantenimiento
DTPFechaContab.Value = Date
Dtp_FechaDoc.Value = Date
DtpFech_Ven.Value = Date
DTPFechaCaja.Value = Date
Ctr_AyudaOficina.xclave = VGparametros.CpOficina: Ctr_AyudaOficina.Ejecutar
TxSerie.text = Left(ESNULO(rsNota!nro_refe, ""), 3)
TxNdoc.text = Right(ESNULO(rsNota!nro_refe, ""), 8)
CtrAyu_TipDoc.xclave = ESNULO(rsNota!doc_refe, ""): CtrAyu_TipDoc.Ejecutar
TxNAux.text = NumeroAuxiliar(Month(DTPFechaContab))
CtrAyu_Proveedor.xclave = rsNota!CACODPRO
Call CtrAyu_Proveedor.Ejecutar
lbNumComprobCab.Caption = UltNumeroAuto(VGParamSistem.TablaCabcomprob, 1, VGCNx)
Call VGDllGeneral.ActivaTab(1, 1, SSTab1)
End Sub
Public Sub CargarAyudas()
With FrmValorizacionArticulos
 '   Call .CtrAyu_Cuenta.Conexion(VGcnxCT): .CtrAyu_Cuenta.filtro = "(cuentanivel=" & VGnumniveles & " and cuentacodigo <>'00') and (" & VGparametros.ctascompra & ")"
 '   Call .CtrAyu_gastos.Conexion(VGCNx): .CtrAyu_gastos.filtro = "(gastosnivel=" & VGnumnivgas & " and gastoscodigo <>'00') "
    Call .CtrAyu_TipDoc.Conexion(VGCNx): .CtrAyu_TipDoc.filtro = "tdocumentocodigo<>'00'"
    Call .CtrAyu_moneda.Conexion(VGcnxCT): .CtrAyu_moneda.filtro = "monedacodigo<>'00'"
    Call .CtrAyu_Modoprovi.Conexion(VGCNx): .CtrAyu_Modoprovi.filtro = "modoprovicod <>'00'"
    Call .CtrAyu_TipCompra.Conexion(VGCNx)
    Call .CtrAyu_TipSubAsi.Conexion(VGCNx)
 '   Call .Ctr_AyuAnalitico.Conexion(VGCNx)
    Call .CtrAyu_Proveedor.Conexion(VGCNx): .CtrAyu_Proveedor.filtro = "clientecodigo <>'00'"
    Call .Ctr_AyudaCaja.Conexion(VGCNx)
    Call .Ctr_AyudaOficina.Conexion(VGCNx)
 '   Call .CtrAyu_Ccosto.Conexion(VGcnxCT)
'    .CtrAyu_Ccosto.filtro = "centrocostonivel=" & VGnumnivcos & " and centrocostocodigo<>'00' "
    Call .Ctr_Ayuempresa.Conexion(VGCNx): .Ctr_Ayuempresa.filtro = "empresacodigo<>'00'"
End With
End Sub

Private Sub CmdSalir2_Click()
Call VGDllGeneral.ActivaTab(0, 1, SSTab1)
End Sub

Private Sub Combo2_Click()
    Call cargar_grid
End Sub

Private Sub Command7_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Set VGDllGeneral = New dllgeneral.dll_general
    Dim rsc As New ADODB.Recordset
    IMant = 1
    Call CargarAyudas
    Set rsc = VGCNx.Execute("Select  TAALMA,TADESCRI  from  tabalm")
        If rsc.RecordCount > 0 Then
        Combo2.Clear
        rsc.MoveFirst
        Do Until rsc.EOF
            Combo2.AddItem rsc.Fields(0) & "-" & IIf(IsNull(rsc.Fields(1)), "****", rsc.Fields(1))
            rsc.MoveNext
        Loop
    End If
    rsc.Close
   TDBG_concil.FetchRowStyle = True
   Set rsc = Nothing
  central FrmValArtPed
  
  Combo1.ListIndex = 0
  Combo2.ListIndex = 0
  TDBNota.FetchRowStyle = True
  i0 = InStr(Combo2.text, "-")
  xDescri_alma = Left(Combo2.text, i0 - 1)
  Call cargar_grid
  Call VGDllGeneral.ActivaTab(0, 1, SSTab1)
End Sub

Private Sub Combo1_Click()
     FG.Col = Combo1.ListIndex
     FG.Sort = 5
End Sub

Private Sub Form_Unload(Cancel As Integer)
     Unload Me
End Sub


Private Sub TDBNota_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  totdoc = 1
   Call cargar_grilla2
End Sub
Private Sub TDBNota_HeadClick(ByVal ColIndex As Integer)
 With rsNota
    If .Sort = Empty Then
        .Sort = TDBNota.Columns.item(ColIndex).DataField & " asc"
    ElseIf Right(.Sort, 3) = "asc" Then
        .Sort = TDBNota.Columns.item(ColIndex).DataField & " desc"
    ElseIf Right(.Sort, 4) = "desc" Then
        .Sort = TDBNota.Columns.item(ColIndex).DataField & " asc"
    End If
    TDBNota.Refresh
 End With
End Sub

Public Sub grabaalmacen(ByVal numero As Integer)
   Dim criterio As String
   Dim cadena As String
   Dim auxdisp As Double
   Dim AUXPRECIO As Double
   Dim RSAUX As New ADODB.Recordset
   Set RSAUX = rsmantenimiento.Clone
   Dim rxaux As New ADODB.Recordset
   Do Until RSAUX.EOF()
      SQL = " update movalmdet set deprecio=" & RSAUX!impbruto & "/" & RSAUX!CANTIDAD & " where dealma='" & rsNota!dealma & "' and detd='" & rsNota!tD & "' and "
      SQL = SQL & " denumdoc='" & rsNota!denumdoc & "' and deitem='" & RSAUX!item & "'"
      Set rxaux = VGCNx.Execute(SQL)
      SQL = " update stkart set stkpreult=" & RSAUX!impbruto & "/" & RSAUX!CANTIDAD & " where stalma='" & rsNota!dealma & "' and stcodigo='" & RSAUX!codigo & "'"
      Set rxaux = VGCNx.Execute(SQL)
      RSAUX.MoveNext
   Loop
   SQL = " update movalmcab set cabprovinumero=" & numero & " , estadoprovision=1 where caalma='" & rsNota!dealma & "' and catd='" & rsNota!tD & "' and "
   SQL = SQL & " canumdoc='" & rsNota!denumdoc & "'"
   Set rxaux = VGCNx.Execute(SQL)
      
  
End Sub

Private Sub limpiaGrid()
Dim I As Integer
 If FG.Rows = 1 Then Exit Sub
 I = FG.RowSel
 If FG.Rows > 2 Then
        FG.RemoveItem I
 Else
        FG.Clear
        FG.Rows = 1
        FG.FormatString = "Cod. Articulo.|Descripcion| Tr| Num.Doc."
        FG.Row = 0
        FG.ColWidth(0) = 950
        FG.ColWidth(1) = 3700
        FG.ColWidth(2) = 450
        FG.ColWidth(3) = 1300
        FG.ColWidth(4) = 2
        FG.ColWidth(5) = 2
  End If
End Sub

Public Sub cargar_grid()

   i0 = InStr(Combo2.text, "-")
   xDescri_alma = Left(Combo2.text, i0 - 1)
       '****************************************************RMM 07/07/2001
  Set rsSTKART = New ADODB.Recordset

  rsSTKART.Open "Select * from STKART WHERE STALMA='" & xDescri_alma & "'", VGCNx, adOpenDynamic, adLockOptimistic

  Dim SQL As String
  
  SQL = "select  n.dealma,N.DETD as TD,n.DENUMDOC , m.cafecdoc,M.CACODPRO,"
  SQL = SQL & "p.clienterazonsocial as Proveedor ,m.CARFTDOC as 'Doc_Refe', "
  SQL = SQL & "m.CARFNDOC as 'Nro_Refe' from MovAlmCab m, MovAlmDet n ,MaeArt , cp_proveedor p"
  SQL = SQL & " Where  m.CAALMA ='" & xDescri_alma & "' AND n.DEALMA = m.CAALMA and "
  SQL = SQL & " CATD='NI' and ACODIGO  = n.DECODIGO And n.DENUMDOC = m.CANUMDOC"
  SQL = SQL & " and n.DETD= m.CATD and M.CACODPRO=p.clientecodigo and  isnull(m.casitgui,'')<>'A' "
  SQL = SQL & " and isnull(estadoprovision,0)<>1 group by n.dealma,N.DETD,n.DENUMDOC,"
  SQL = SQL & " M.CACODPRO, m.CANOMPRO ,m.CARFTDOC, m.CARFNDOC,m.cafecdoc,p.clienterazonsocial order by 1,2"
  
  Set rsNota = New ADODB.Recordset
  Set rsNota = VGCNx.Execute(SQL)
  
    If rsNota.RecordCount = 0 Then
        MsgBox "No hay Notas de Ingreso/Salida", vbInformation, Caption
        rsNota.Close
        Set TDBNota.DataSource = Nothing
        FG.Clear
        'Unload Me
        Exit Sub
    End If

  
  Set TDBNota.DataSource = rsNota
  
 
End Sub

Public Sub cargar_grilla2()

          mRsql = "select  n.DECODIGO, ADESCRI, N.DETD,n.DENUMDOC, m.CANOMPRO, m.CARFTDOC, m.CARFNDOC from MovAlmCab m, MovAlmDet n ,MaeArt  Where  m.CAALMA ='" & xDescri_alma & _
                             "' AND n.DEALMA = m.CAALMA and (CATD='NI' OR CATD='NC' )  and ACODIGO  = n.DECODIGO      And   n.DENUMDOC = m.CANUMDOC  and n.DETD= m.CATD and isnull(m.CASITGUI,'')<>'A'  AND "
          mRsql = mRsql & "n.dealma='" & TDBNota.Columns(0) & "' and n.DETD='" & TDBNota.Columns(1).Value & "' and n.DENUMDOC='" & TDBNota.Columns(2).Value & "' ORDER BY m.CANUMDOC"
        
          Set rs = VGCNx.Execute(mRsql)
          If rs.RecordCount = 0 Then
                     MsgBox "No hay Artículos por Valorizar que esten Pendientes", vbExclamation, mensaje1
                     FG.Clear
          Else
                    
                    Call limpiar_grilla2
                    FG.Rows = 1
                    rs.MoveFirst
                    FG.Visible = False
                    While Not rs.EOF
                            FG.AddItem (rs(0) & vbTab & Trim(rs(1)) & vbTab & rs(2) & vbTab & rs(3) & vbTab & rs(4) & vbTab & rs(5) & vbTab & rs(6))
                            rs.MoveNext
                    Wend
                    rs.Close
                    FG.Visible = True
          End If
End Sub

Public Sub limpiar_grilla2()

    FG.Clear
    FG.Cols = 7
    'FG.FormatString = "Codigo Art.|Descripcion| TD |Num.Doc| |"
    FG.Row = 0
    FG.ColWidth(0) = 1400
    FG.ColWidth(1) = 5100
    FG.ColWidth(2) = 500
    FG.ColWidth(3) = 1000
    FG.ColWidth(4) = 1000
    FG.ColWidth(5) = 800
    FG.ColWidth(6) = 1000
    
    'FG.FormatString = "Codigo Art.|Descripcion| TD |Num.Doc| |"
    FG.ColAlignment(0) = 1
    FG.ColAlignment(1) = 1
    
    FG.Row = 0
    FG.Col = 0
    Dim cabecera(1, 6)
    Dim I As Integer
    I = 0
    cabecera(1, 0) = "Codigo"
    cabecera(1, 1) = "Descripcion"
    cabecera(1, 2) = "TD"
    cabecera(1, 3) = "Num. Doc" '--"Nro. Documento"
    cabecera(1, 4) = "Proveedor" '---"Proveedor"
    cabecera(1, 5) = "Doc. Ref" '--"Doc.]REf"
    cabecera(1, 6) = "Num. Ref." '- -"Num ref"
    
    
    For I = 0 To FG.Cols - 1
        FG.Col = I
        FG.text = cabecera(1, I)
    Next I

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
    If CtrAyu_moneda.xclave = "02" Then
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
Private Sub CtrAyu_Proveedor_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    Set VGvardllgen = New dllgeneral.dll_general
    txRuc.text = VGvardllgen.ESNULO(ColecCampos("clienteruc").Value, "")
    buencontribuyente = VGvardllgen.ESNULO(ColecCampos("proveedorcontribuyente").Value, 0)
    lbTelef.Caption = VGvardllgen.ESNULO(ColecCampos("clientetelefono").Value, "")
End Sub
Private Sub CtrAyu_Proveedor_AlNoDevolverNada()
    txRuc.text = ""
    lbTelef.Caption = ""
End Sub
Private Sub CtrAyu_TipCompra_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    Set VGvardllgen = New dllgeneral.dll_general
    CtrAyu_TipSubAsi.xclave = "": CtrAyu_TipSubAsi.xnombre = ""
    CtrAyu_TipSubAsi.filtro = "tipocompracodigo='" & CtrAyu_TipCompra.xclave & "'"
    comprainafecta = VGvardllgen.ESNULO(ColecCampos("tipocomprainafecta").Value, 0)
'    If comprainafecta = 1 Then
'       TxImpBruto.Visible = False
'       TxIGV.Visible = False
'     Else
'       TxImpBruto.Visible = True
'       TxIGV.Visible = True
'     End If
End Sub
Private Sub Ctr_Ayuempresa_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    Set VGvardllgen = New dllgeneral.dll_general
    tipoinafecto = VGvardllgen.ESNULO(ColecCampos("agentederetencion").Value, 0)
    If tipoinafecto = 0 Then
       tipoinafecto = 1
     Else
       tipoinafecto = 0
    End If
End Sub

Private Sub CtrAyu_TipDoc_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    Set VGvardllgen = New dllgeneral.dll_general
    VlDocNota = VGvardllgen.ESNULO(ColecCampos("tdocumentotipo").Value, "")
    documentoinafecto = VGvardllgen.ESNULO(ColecCampos("documentoretencion").Value, 0)
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
Private Sub DtpFech_Ven_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call SendKeys("{TAB}")
End Sub
Private Sub DTPFechaContab_KeyDown(KeyCode As Integer, Shift As Integer)
   TxNAux.text = NumeroAuxiliar(Month(DTPFechaContab))
   If KeyCode = 13 Then Call SendKeys("{TAB}")
End Sub
Public Function NumeroAuxiliar(mes As Integer, Optional ByRef numero As Long) As String
On Error GoTo Errnum
Dim RSAUX As ADODB.Recordset
Dim cad As String
    Set RSAUX = New ADODB.Recordset
    cad = "Select isnull(mes" & Trim(Format(mes, "00")) & ",0)+1 as numcorrelativo   From co_correlames " & _
          "Where Ano='" & Right(Str(Year(DTPFechaContab)), 4) & "'"
          
    RSAUX.Open cad, VGCNx, adOpenKeyset, adLockReadOnly
               
    If RSAUX.RecordCount > 0 Then
       NumeroAuxiliar = Trim(Format(RSAUX!numcorrelativo, "00000"))
       numero = RSAUX!numcorrelativo
       Else
        NumeroAuxiliar = "00"
        numero = 0
    End If
    Exit Function
Errnum:
    MsgBox ("Error en Numero de Comprobante " & Chr(13) & Err.Description)
End Function
Private Sub CmbTcambio_Click()
lb_vcambio = Format(XRecuperaTipoCambio(Dtp_FechaDoc, CmbTcambio.ListIndex + 1, VGcnxCT), "#0.000 ")
End Sub
Private Sub CmbTcambio_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call SendKeys("{TAB}")
End Sub

Private Sub TDBG_concil_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
Dim rsclone As New ADODB.Recordset
    On Error Resume Next
    Set rsclone = rsmantenimiento.Clone(adLockBatchOptimistic)
    If rsclone.RecordCount = 0 Then Exit Sub
    rsclone.Bookmark = Bookmark
    If numero(rsclone!impbruto) = 0 And numero(rsclone!impinafecto) = 0 Then
       RowStyle.BackColor = RGB(254, 251, 218)
       '185,251,210
     Else
       RowStyle.BackColor = RGB(200, 250, 100)
    End If
    If flaggraba = 0 Then
       TxtImpbruto.text = rsmantenimiento!impbruto
       Txtimpigv.text = rsmantenimiento!impIgv
       Txtimpinafecto.text = rsmantenimiento!impinafecto
       Txtimptotal.text = rsmantenimiento!imptotal
   End If
   TxtImpbruto.SetFocus
End Sub
Private Sub rsmantenimiento_FieldChangeComplete(ByVal cFields As Long, ByVal Fields As Variant, ByVal pError As ADODB.error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Static Cont As Integer
On Error GoTo X
    If Cont = 1 Then
        Cont = 0
        Exit Sub
    End If
        Call CalcularTotales(rsmantenimiento, rstotal)
    Cont = 1
    TDBG_concil.Refresh
    Exit Sub
X:
End Sub
Private Sub CalcularTotales(ByVal rs As Recordset, rstotal As Recordset)
Dim RSAUX As ADODB.Recordset
Dim rsauxtot As ADODB.Recordset
Dim vardllgen As New dllgeneral.dll_general
Dim txtbruto As Double, TxtIGV As Double, Txtinafecto As Double, TxtTotal As Double
Set RSAUX = rs.Clone(adLockReadOnly)
If RSAUX.BOF = True Or RSAUX.EOF = True Then Exit Sub

If ExisteElem(0, VGCNx, sqltabla1) Then
   Set rstotal = VGCNx.Execute(" drop table " & sqltabla1 & "")
End If

SQL = " select * into " & sqltabla1 & " from " & Sqltabla & ""
Set rstotal = VGCNx.Execute(SQL)
SQL = " select * from " & sqltabla1 & ""

rstotal.Open (SQL), VGCNx, adOpenDynamic, adLockBatchOptimistic
RSAUX.MoveFirst
rstotal.MoveFirst
    While Not RSAUX.EOF
        rstotal!impbruto = vardllgen.ESNULO(RSAUX!impbruto, 0) * RSAUX!CANTIDAD
        rstotal!impIgv = vardllgen.ESNULO(RSAUX!impIgv, 0) * RSAUX!CANTIDAD
        rstotal!impinafecto = vardllgen.ESNULO(RSAUX!impinafecto, 0) * RSAUX!CANTIDAD
        rstotal!imptotal = vardllgen.ESNULO(RSAUX!imptotal, 0) * RSAUX!CANTIDAD
        
        txtbruto = txtbruto + rstotal!impbruto
        TxtIGV = TxtIGV + rstotal!impIgv
        Txtinafecto = Txtinafecto + rstotal!impinafecto
        TxtTotal = TxtTotal + rstotal!imptotal
        RSAUX.MoveNext
        rstotal.MoveNext
    Wend
    TxTotBruto.text = Round(txtbruto, 2)
    TxTotIGV.text = Round(TxtIGV, 2)
    TxTotInafecto.text = Round(Txtinafecto, 2)
        TxTotTotal.text = Round(Round(txtbruto, 2) + Round(TxtIGV, 2) + Round(Txtinafecto, 2), 2)
 End Sub

Private Sub TxtImpbruto_KeyPress(KeyAscii As Integer)
If KeyAscii = 0 Then
   Txtimpigv.text = numero((TxtImpbruto.text) * 0.19)
   Txtimptotal.text = numero(TxtImpbruto.text) + numero(Txtimpigv.text) + numero(Txtimpinafecto.text)
   '+ numero(Txtimpinafecto.text)
End If
End Sub
Private Sub TxtImpinafecto_KeyPress(KeyAscii As Integer)
If KeyAscii = 0 Then
   Txtimptotal.text = numero(TxtImpbruto.text) + numero(Txtimpigv.text) + numero(Txtimpinafecto.text)
End If
End Sub
Private Sub grabagrilla()
Dim posi As Integer
flaggraba = 1
posi = rsmantenimiento.Bookmark

rsmantenimiento!impbruto = numero(TxtImpbruto.text)
rsmantenimiento!impIgv = numero(Txtimpigv.text)
rsmantenimiento!impinafecto = numero(Txtimpinafecto.text)
rsmantenimiento!imptotal = numero(Txtimptotal.text)
Call CalcularTotales(rsmantenimiento, rstotal)
flaggraba = 0
    rsmantenimiento.MoveNext
If rsmantenimiento.EOF() Then rsmantenimiento.MoveFirst
If posi = rsmantenimiento.RecordCount() Then
   posi = 1
 Else
   posi = posi + 1
End If
        TxtImpbruto.SetFocus
End Sub

Private Sub Txtimptotal_KeyPress(KeyAscii As Integer)
If KeyAscii = 0 Then
   Call grabagrilla
End If
End Sub

Public Sub grabar()
Dim xnumerocompro As String, nnumerocorrcomprob As Double
Dim xnumerocomprolibro As String, nnumerocorrcomproblibro As Double
Dim Existelibro As Boolean
Dim SQL As String
Dim xsql As New ADODB.Recordset
Dim op2 As Integer
Dim varnerror As Integer
Dim sqltes As String
Dim Vlnaux As String
Dim estadorendicion As Integer
Dim VlComprob_Conta As String
Dim sqlaux As ADODB.Recordset
Set VGvardllgen = New dllgeneral.dll_general
On Error GoTo ErrorGrabar
Dim xcon As Long
Dim datoold As String
Dim datonuevo As String
Dim modoproviold As Integer
Dim vgerrorstring As String
VGvarVerifica = True
vgerrorstring = ""
varnerror = 0
emitedetraccion = 0
    If Not ValidarGrabarCabecera(rsmantenimiento.RecordCount) Then Exit Sub
    Set rsdetalle = New ADODB.Recordset
    If Not ValidarRsDetalle(rstotal, rsdetalle) Then Exit Sub
    xcon = rsmantenimiento.RecordCount
    rsmantenimiento.Filter = "(ImpTotal =0)"
    If rsmantenimiento.RecordCount > 0 Then
        MsgBox "Por lo Menos un registro sin valores ", vbExclamation
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
    VGgeneral.BeginTrans 'Inicio la transaccion
    Screen.MousePointer = vbHourglass
    '1=>Paso Genera el Correlativo del Comprobante
    Dim xnumero As Long
    If IMant = 1 Then
        If VGparametros.Auxaut Then
            xnumerocompro = NumeroAuxiliar(CInt(VGParamSistem.Mesproceso), xnumero)
          Else
            xnumerocompro = Trim(TxNAux.text)
            'Validar si el Numero ya ha sido ingresado
            If ExisteSQL(VGCNx, "Select * From co_cabprovi" & VGParamSistem.Anoproceso & _
                               " Where cabprovinumaux='" & xnumerocompro & "'") Then
                MsgBox "El Numero de Comprobante Auxiliar ya ha sido ingresado", vbExclamation
                TxNAux.SetFocus
                Exit Sub
            End If
            
        End If
        '2=>Paso Actualizo el Correlativo en la Tabla SubAsiento si es que ingrese un nuevo
        'Comprobante
        Call ActualizaCorrelAuxiliar(xnumero, VGParamSistem.FechaTrabajo)
        If Not VGvarVerifica Then varnerror = 6: GoTo ErrorGrabar
      Else
        If Month(DTPFechaContab) < Val(VGParamSistem.Mesproceso) Then
           If VGparametros.Auxaut Then
              xnumerocompro = NumeroAuxiliar(Month(DTPFechaContab), xnumero)
              Call ActualizaCorrelAuxiliar(xnumero, VGParamSistem.FechaTrabajo)
              If Not VGvarVerifica Then varnerror = 6: GoTo ErrorGrabar
            End If
        Else
           'Validar si el Numero ya ha sido ingresado cuando esta siendo modificado
           If Vlnaux <> Trim(TxNAux.text) Then
              If ExisteSQL(VGCNx, "Select * From co_cabprovi" & VGParamSistem.Anoproceso & _
                                " Where cabprovinumaux='" & Trim(TxNAux.text) & "'") Then
                 MsgBox "El Numero de Comprobante Auxiliar ya ha sido ingresado", vbExclamation
                 TxNAux.SetFocus
                 Exit Sub
              End If
          End If
        End If
        xnumerocompro = Trim(TxNAux.text)
    End If
    If Not VGvarVerifica Then varnerror = 1: GoTo ErrorGrabar
    
        
    ' Actualizando numero de comprobante
    
    If IMant = 1 Then
       xnumero = UltNumeroAuto(VGParamSistem.TablaCabcomprob, 1, VGCNx)
       VGCNx.Execute ("Update co_sistema SET cabprovinumero=" & xnumero + 1)
     Else
       xnumero = CDbl(lbNumComprobCab)
      
    End If
    
    
    '2=>Paso Grabo la Cabecera del Comprobante
    Dim Xnumtesor As String
    If ChkActCaja.Value = 1 Or modoproviold = 1 Then
        Call Grabaren_Tesoreria(IMant, xnumero, rsdetalle, Xnumtesor)
    End If
    
    Call GrabarCabecera(IMant, xnumero, Format(CInt(VGParamSistem.Mesproceso), "00") & xnumerocompro, Xnumtesor)
    If Not VGvarVerifica Then varnerror = 2: GoTo ErrorGrabar
    
    If ChkCtaCte.Value = 1 Then
        Set xsql = New ADODB.Recordset
        SQL = "select * from cp_cargo where clientecodigo='" & Trim(FrmValorizacionArticulos.CtrAyu_Proveedor.xclave) & "'"
        SQL = SQL & " and documentocargo='" & Trim(FrmValorizacionArticulos.CtrAyu_TipDoc.xclave) & "'"
        SQL = SQL & " and cargonumdoc='" & Format(Trim(FrmValorizacionArticulos.TxSerie.text), "000") & Format(Left(Trim(FrmValorizacionArticulos.TxNdoc.text), 8), "00000000") & "'"
        Set xsql = VGCNx.Execute(SQL)
        op2 = IMant
        If xsql.RecordCount() = 0 Then
           IMant = 1
        End If
        datoold = FrmValorizacionArticulos.CtrAyu_Proveedor.Tag & FrmValorizacionArticulos.CtrAyu_TipDoc.Tag
        datoold = datoold & Format(Trim(FrmValorizacionArticulos.TxSerie.Tag), "000") & Format(Left(Trim(FrmValorizacionArticulos.TxNdoc.Tag), 8), "00000000")
        
        datonuevo = FrmValorizacionArticulos.CtrAyu_Proveedor.xclave & FrmValorizacionArticulos.CtrAyu_TipDoc.xclave
        datonuevo = datonuevo & Format(Trim(FrmValorizacionArticulos.TxSerie.text), "000") & Format(Left(Trim(FrmValorizacionArticulos.TxNdoc.text), 8), "00000000")
        
        If op2 = 2 And datoold <> datonuevo Then
           IMant = 2
        End If
        
        Call GrabarCP_Cargo(IMant, xnumero)
        IMant = op2
    End If
    
    '3=>Paso Grabo los Detalle del Comprobante
    
    Call GrabarDetalle(rsdetalle, xnumero)
    If Not VGvarVerifica Then varnerror = 3: GoTo ErrorGrabar
    
    '4=>Generar Asiento en Linea segun parametro
    If VGparametros.sistemaasientoenlinea Then
       Call GeneraAsientoenLine(IMant, xnumero, VlComprob_Conta)
       If Not VGvarVerifica Then varnerror = 4: GoTo ErrorGrabar
    End If
                  
    Call grabaalmacen(xnumero)
    Call cargar_grid
                     
    VGgeneral.CommitTrans 'Acepto toda la transaccion porque es correcta
    If FrmValorizacionArticulos.estadorendicion = 1 Then
       sqltes = " update te_detallerecibos set chkconcil=" & FrmValorizacionArticulos.estadorendicion & ",fechconcil='" & FrmValorizacionArticulos.fecharendicion & "'"
       sqltes = sqltes & ",rendicionnumero='" & FrmValorizacionArticulos.numerorendicion & "' where cabrec_numrecibo='" & FrmValorizacionArticulos.numerorecibo & "'"
       Set sqlaux = VGCNx.Execute(sqltes)
    End If
    If IMant = 1 Then
        MsgBox "Se grabo Satisfactoriamente  El numero de Comprobante Generado Es :" & Chr(13) & _
           "Nro: " & xnumero & Chr(13) & _
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
            MsgBox "No se Genero Correctamente el numero del Comprobante" & Chr(13) & vgerrorstring, vbExclamation
        Case 2, 3, 4, 5, 6
            VGgeneral.RollbackTrans
            MsgBox "Hubo Errores al Grabar" & Chr(13) & vgerrorstring, vbExclamation
            Call Cancelar(1)
            
        Case Else
            MsgBox "Errores Desconocidos " & Chr(13) & Err.Description
    End Select
    Screen.MousePointer = vbDefault
    Exit Sub
    Resume
End Sub

Public Sub Grabaren_Tesoreria(ByVal op As Integer, Optional ByVal Numeroprovi As Long = 0, Optional ByVal rs As Recordset, Optional ByRef XNum As String)
'On Error GoTo ErrorGrabaTesore
Dim numero As String
Set VGvardllgen = New dllgeneral.dll_general
Dim rb As ADODB.Recordset
Dim item As Integer
   'Obtener el Ultimo Numero Correlativo de las cajas
    Dim opaux As Integer
    opaux = op
    Set VGCommandoSP = New ADODB.Command
    Set VGvardllgen = New dllgeneral.dll_general
    VGCommandoSP.ActiveConnection = VGgeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "te_abonadocumento_pro"
    VGCommandoSP.Parameters.Refresh
    If op = 2 Then
        If Trim(FrmValorizacionArticulos.TxTesor.text) = "" Then
            op = 1
        End If
    End If
    With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@tipo") = IIf(op = 3, 2, op)
        If op = 2 Or op = 3 Then
            'Set rb = VGcnx.Execute("Select cabprovinumtesor  From " & VGParamSistem.TablaCabcomprob & " Where cabprovinumero=" & FrmValorizacionArticulos.lbNumComprobCab.Caption)
            numero = FrmValorizacionArticulos.TxTesor.text
            
          Else
            Set rb = VGCNx.Execute("select * from te_parametroempresa where empresacodigo='01'")
            If rb.RecordCount > 0 Then
                numero = Format(CDbl(VGvardllgen.ESNULO(rb!empresanumegreso, "0")) + 1, "000000")
                VGCNx.Execute "Update te_parametroempresa Set empresanumegreso='" & numero & "' where empresacodigo='01'"
                'VGcnx.Execute "Update " & VGParamSistem.TablaCabcomprob & " Set cabprovinumtesor='" & Numero & _
                '              "' Where cabprovinumero=" & Numeroprovi
            End If
        End If
        XNum = numero
        .Parameters("@estadoreg") = ""
        .Parameters("@numrecibo") = numero
        If op = 3 Then
            .Execute
        End If
        If op = 2 Or FrmValorizacionArticulos.modoproviold = 1 Then
            .Execute
           If FrmValorizacionArticulos.ChkActCaja = 1 Then
              .Parameters("@tipo") = 1
              op = 1
            Else
              XNum = ""
           End If
        End If
        
        'Este para que al eliminar no utilizar estos parametros
         If op = 1 Then
            .Parameters("@controlctacte") = "N"
            .Parameters("@vendedorcodigo") = FrmValorizacionArticulos.Ctr_AyudaOficina.xclave
            .Parameters("@cajacodigo") = FrmValorizacionArticulos.Ctr_AyudaCaja.xclave
            .Parameters("@clientecodigo") = FrmValorizacionArticulos.CtrAyu_Proveedor.xclave
            .Parameters("@descripcion") = ""
            .Parameters("@operacion") = FrmValorizacionArticulos.CtrAyu_Modoprovi.xclave
            .Parameters("@monedacodigo") = FrmValorizacionArticulos.CtrAyu_moneda.xclave
            .Parameters("@ingsal") = "E"
            .Parameters("@tipocambio") = CDbl(FrmValorizacionArticulos.lb_vcambio.Caption)
            .Parameters("@totsoles") = IIf(FrmValorizacionArticulos.CtrAyu_moneda.xclave = "01", CDbl(FrmValorizacionArticulos.TxTotTotal.valor), Round(CDbl(FrmValorizacionArticulos.TxTotTotal.valor) * CDbl(FrmValorizacionArticulos.lb_vcambio.Caption), 2))
            .Parameters("@totdolares") = IIf(FrmValorizacionArticulos.CtrAyu_moneda.xclave <> "01", CDbl(FrmValorizacionArticulos.TxTotTotal.valor), Round(CDbl(FrmValorizacionArticulos.TxTotTotal.valor) / CDbl(FrmValorizacionArticulos.lb_vcambio.Caption), 2))
            .Parameters("@fechadocumento") = FrmValorizacionArticulos.DTPFechaCaja.Value
            .Parameters("@observa") = ""
            .Parameters("@transferauto") = ""
            .Parameters("@numreciboegreso") = ""
            .Parameters("@usuario") = VGUsuario
            .Parameters("@fechaact") = Now
            If VGparametros.sistemamultiempresas Then
               .Parameters("@empresa") = FrmValorizacionArticulos.Ctr_Ayuempresa.xclave
             Else
                .Parameters("@empresa") = "01"
             End If
            .Parameters("@cabprovinumero") = Numeroprovi
            .Parameters("@saldodocxrendir") = IIf(FrmValorizacionArticulos.CtrAyu_moneda.xclave = "01", CDbl(FrmValorizacionArticulos.TxTotTotal.valor), Round(CDbl(FrmValorizacionArticulos.TxTotTotal.valor) * CDbl(FrmValorizacionArticulos.lb_vcambio.Caption), 2))
            .Execute
        End If
    End With
op = opaux
Set VGCommandoSP = New ADODB.Command
VGCommandoSP.ActiveConnection = VGgeneral
VGCommandoSP.CommandType = adCmdStoredProc
VGCommandoSP.CommandText = "te_abonadetalledocumento_pro"
VGCommandoSP.Parameters.Refresh
If op = 3 Then
   With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@tipo") = IIf(op = 3, 2, op)
        .Parameters("@numrecibo") = numero
        If op = 3 Then
            .Execute
        End If
   End With
Else
  rs.MoveFirst
  item = 1
  While Not rs.EOF
      With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@tipo") = IIf(op = 3, 2, op)
        .Parameters("@numrecibo") = numero
        If op = 3 Then
            .Execute
        End If
        If op = 2 Or FrmValorizacionArticulos.modoproviold = 1 Then
            If item = 1 Then .Execute
            If FrmValorizacionArticulos.ChkActCaja = 1 Then
               .Parameters("@tipo") = 1
               op = 1
            End If
        End If
        If op = 1 Then
            .Parameters("@estadoreg") = ""
            .Parameters("@item") = item
            .Parameters("@emisioncheque") = "C"
            .Parameters("@tipodocconcepto") = FrmValorizacionArticulos.CtrAyu_TipDoc.xclave
            .Parameters("@numdocumento") = Format(Trim(FrmValorizacionArticulos.TxSerie.text), "000") & Format(Left(Trim(FrmValorizacionArticulos.TxNdoc.text), 8), "00000000")
            .Parameters("@carabo") = FrmValorizacionArticulos.VlDocNota
            .Parameters("@formacan") = ""
            .Parameters("@tdqc") = ""
            .Parameters("@ndqc") = ""
            .Parameters("@tipocajabanco") = "C"
            .Parameters("@cajabanco") = FrmValorizacionArticulos.Ctr_AyudaCaja.xclave
            .Parameters("@numctacte") = ""
            .Parameters("@adicionactacte") = "P"
            .Parameters("@monedadocumento") = FrmValorizacionArticulos.CtrAyu_moneda.xclave
            .Parameters("@monedacancela") = FrmValorizacionArticulos.CtrAyu_moneda.xclave
            .Parameters("@importesoles") = IIf(FrmValorizacionArticulos.CtrAyu_moneda.xclave = "01", CDbl(rs.Fields("total")), Round(CDbl(rs.Fields("total")) * CDbl(FrmValorizacionArticulos.lb_vcambio.Caption), 2))
            .Parameters("@importedolares") = IIf(FrmValorizacionArticulos.CtrAyu_moneda.xclave <> "01", CDbl(rs.Fields("total")), Round(CDbl(rs.Fields("total") / CDbl(FrmValorizacionArticulos.lb_vcambio.Caption)), 2))
            .Parameters("@contabledisponi") = "S"
            .Parameters("@fechacancela") = FrmValorizacionArticulos.DTPFechaCaja.Value
            .Parameters("@observacion") = ""
            .Parameters("@gastos") = ESNULO(rs.Fields("gastos"), "0213")
            .Parameters("@usuario") = VGUsuario
            .Parameters("@fechaact") = Now
            .Parameters("@entidad") = "" 'rs.Fields("analitico")
            .Parameters("@centrocosto") = "" 'rs.Fields("ccosto")
            .Execute
         End If
    End With
    item = item + 1
    rs.MoveNext
  Wend
End If

Exit Sub

ErrorGrabaTesore:
    VGvarVerifica = False
    MsgBox ("Error en Grabar en Cuentas por Pagar " & Chr(13) & Err.Description)
End Sub
Public Sub Cancelar(Optional op As Integer)
Set VGvardllgen = New dllgeneral.dll_general

    If op <> 1 Then
        If MsgBox("Esta Seguro que Desea Cancelar la Operación ", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
            'Resolver el problema que el cursor debe parpadear donde se ha quedado
            Exit Sub
        End If
    End If
        
    If SSTab1.Tab = 1 Then
        Call VGvardllgen.ActivaTab(0, 1, SSTab1)
        Set rsmantenimiento = Nothing
    End If
    
End Sub
