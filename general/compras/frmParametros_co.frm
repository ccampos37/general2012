VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Begin VB.Form frmParametros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parámetros Generales"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   10410
   Begin VB.Frame Frame6 
      Height          =   1455
      Left            =   8730
      TabIndex        =   48
      Top             =   6945
      Width           =   1575
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Aceptar"
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   51
         Top             =   240
         Width           =   1185
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Cancelar"
         Height          =   360
         Index           =   2
         Left            =   120
         TabIndex        =   50
         Top             =   960
         Width           =   1185
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Imprimir"
         Height          =   360
         Index           =   1
         Left            =   120
         TabIndex        =   49
         Top             =   600
         Width           =   1185
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1305
      Left            =   3450
      TabIndex        =   41
      Top             =   5625
      Width           =   6855
      Begin VB.CheckBox Chkempresas 
         Caption         =   "Multiermpresas"
         Height          =   495
         Left            =   90
         TabIndex        =   60
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtcomprobante 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Height          =   375
         Left            =   2970
         TabIndex        =   59
         Text            =   "Text1"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Txtminimoretencion 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Height          =   375
         Left            =   5730
         TabIndex        =   58
         Text            =   "Text1"
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox Chkbancarizacion 
         Caption         =   "Bancarizacion"
         Height          =   375
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TxtMinimobancarizacion01 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   43
         Text            =   "TxtMinimoBancarizacion01"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox TxtMinimobancarizacion02 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Height          =   375
         Left            =   5640
         TabIndex        =   42
         Text            =   "TxtMinimoBancarizacion02"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lbl 
         Caption         =   "Ultimo comprobante"
         Height          =   570
         Index           =   3
         Left            =   1770
         TabIndex        =   62
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label lbl 
         Caption         =   "Monto Minimo Retencion"
         Height          =   570
         Index           =   15
         Left            =   4530
         TabIndex        =   61
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label lbl 
         Caption         =   "Monto Minimo Soles"
         Height          =   570
         Index           =   17
         Left            =   1800
         TabIndex        =   45
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label lbl 
         Caption         =   "Monto Minimo Dolares"
         Height          =   570
         Index           =   16
         Left            =   4320
         TabIndex        =   44
         Top             =   240
         Width           =   1125
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1164
      Left            =   96
      TabIndex        =   28
      Top             =   1200
      Width           =   10230
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         Caption         =   "Controla Saldos en rendiciones"
         Height          =   195
         Index           =   4
         Left            =   7560
         TabIndex        =   46
         Top             =   744
         Width           =   2595
      End
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo cambio para entrar al sistema "
         Height          =   324
         Index           =   1
         Left            =   48
         TabIndex        =   31
         Top             =   744
         Width           =   2820
      End
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         Caption         =   "Asientos Contables en linea"
         Height          =   276
         Index           =   3
         Left            =   5070
         TabIndex        =   30
         Top             =   744
         Width           =   2355
      End
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         Caption         =   "Activa Centros de Gastos"
         Height          =   330
         Index           =   2
         Left            =   2970
         TabIndex        =   29
         Top             =   744
         Width           =   2115
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaMon 
         Height          =   312
         Left            =   1488
         TabIndex        =   32
         Top             =   144
         Width           =   3996
         _ExtentX        =   7038
         _ExtentY        =   556
         XcodMaxLongitud =   0
         xcodwith        =   500
         NomTabla        =   "gr_moneda"
         TituloAyuda     =   "Ayuda de Moneda"
         ListaCampos     =   "monedacodigo(1),monedadescripcion(1)"
         XcodCampo       =   "monedacodigo"
         XListCampo      =   "monedadescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "monedacodigo,monedadescripcion"
         Requerido       =   0   'False
      End
      Begin TextFer.TxFer txt 
         Height          =   300
         Index           =   4
         Left            =   8910
         TabIndex        =   33
         Top             =   165
         Width           =   930
         _ExtentX        =   1640
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
         Text            =   ""
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         TipoDato        =   1
         NumeroDecimales =   2
         SignoNegativo   =   0   'False
         MarcarTextoAlEnfoque=   -1  'True
      End
      Begin TextFer.TxFer txt 
         Height          =   300
         Index           =   5
         Left            =   1464
         TabIndex        =   34
         Top             =   444
         Width           =   5556
         _ExtentX        =   9790
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
         MaxLength       =   250
         Text            =   ""
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         NoCaracteres    =   "0123456789%"
         MarcarTextoAlEnfoque=   -1  'True
         NoRangoCadena   =   -1  'True
      End
      Begin VB.Label lbl 
         Caption         =   "Moneda Base"
         Height          =   216
         Index           =   7
         Left            =   108
         TabIndex        =   37
         Top             =   180
         Width           =   2208
      End
      Begin VB.Label lbl 
         Caption         =   "Valor IGV (%)"
         Height          =   210
         Index           =   9
         Left            =   7545
         TabIndex        =   36
         Top             =   165
         Width           =   2205
      End
      Begin VB.Label lbl 
         Caption         =   "Ctas de Compras"
         Height          =   216
         Index           =   5
         Left            =   108
         TabIndex        =   35
         Top             =   468
         Width           =   2208
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1164
      Left            =   144
      TabIndex        =   18
      Top             =   48
      Width           =   10188
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         Caption         =   "Descr.larga para imprimir"
         Height          =   324
         Index           =   0
         Left            =   5940
         TabIndex        =   19
         Top             =   444
         Width           =   2124
      End
      Begin TextFer.TxFer txt 
         Height          =   300
         Index           =   0
         Left            =   2448
         TabIndex        =   20
         Top             =   144
         Width           =   5652
         _ExtentX        =   9975
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
         MaxLength       =   40
         Text            =   ""
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         MarcarTextoAlEnfoque=   -1  'True
      End
      Begin TextFer.TxFer txt 
         Height          =   300
         Index           =   1
         Left            =   2448
         TabIndex        =   21
         Top             =   444
         Width           =   3408
         _ExtentX        =   6006
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
         MaxLength       =   15
         Text            =   ""
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         MarcarTextoAlEnfoque=   -1  'True
      End
      Begin TextFer.TxFer txt 
         Height          =   300
         Index           =   3
         Left            =   8196
         TabIndex        =   22
         Top             =   432
         Visible         =   0   'False
         Width           =   2112
         _ExtentX        =   3731
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
         Text            =   " "
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   " "
         TipoDato        =   1
         SignoNegativo   =   0   'False
         MarcarTextoAlEnfoque=   -1  'True
      End
      Begin TextFer.TxFer txt 
         Height          =   300
         Index           =   2
         Left            =   2436
         TabIndex        =   23
         Top             =   756
         Width           =   5664
         _ExtentX        =   10001
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
         MarcarTextoAlEnfoque=   -1  'True
      End
      Begin VB.Label lbl 
         Caption         =   "Descripción Empresa (Larga)"
         Height          =   216
         Index           =   0
         Left            =   96
         TabIndex        =   27
         Top             =   192
         Width           =   2208
      End
      Begin VB.Label lbl 
         Caption         =   "Descripción Empresa (Abrev.)"
         Height          =   216
         Index           =   1
         Left            =   96
         TabIndex        =   26
         Top             =   492
         Width           =   2208
      End
      Begin VB.Label lbl 
         Caption         =   "Dirección Empresa"
         Height          =   216
         Index           =   2
         Left            =   96
         TabIndex        =   25
         Top             =   852
         Width           =   2208
      End
      Begin VB.Label lbl 
         Caption         =   "RUC Empresa"
         Height          =   216
         Index           =   4
         Left            =   8256
         TabIndex        =   24
         Top             =   180
         Visible         =   0   'False
         Width           =   2208
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Generasión de Asientos a Contabilidad"
      Height          =   3195
      Left            =   108
      TabIndex        =   3
      Top             =   2364
      Width           =   10152
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaCtaIGV 
         Height          =   300
         Left            =   165
         TabIndex        =   4
         Top             =   930
         Width           =   5580
         _ExtentX        =   9843
         _ExtentY        =   529
         XcodMaxLongitud =   0
         xcodwith        =   1000
         NomTabla        =   "ct_cuenta"
         TituloAyuda     =   "Ayuda de Cuenta IGV"
         ListaCampos     =   "cuentacodigo(1),cuentadescripcion(1)"
         XcodCampo       =   "cuentacodigo"
         XListCampo      =   "cuentadescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "cuentacodigo,cuentadescripcion"
      End
      Begin TextFer.TxFer txt 
         Height          =   300
         Index           =   6
         Left            =   144
         TabIndex        =   5
         Top             =   384
         Width           =   756
         _ExtentX        =   1323
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
         MaxLength       =   40
         Text            =   ""
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         NoCaracteres    =   "0123456789%"
         MarcarTextoAlEnfoque=   -1  'True
         NoRangoCadena   =   -1  'True
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaLibro 
         Height          =   300
         Left            =   1104
         TabIndex        =   6
         Top             =   384
         Width           =   4716
         _ExtentX        =   8308
         _ExtentY        =   529
         XcodMaxLongitud =   0
         xcodwith        =   1000
         NomTabla        =   "ct_libro"
         TituloAyuda     =   "Ayuda de Libro"
         ListaCampos     =   "librocodigo(1),librodescripcion(1)"
         XcodCampo       =   "librocodigo"
         XListCampo      =   "librodescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "librocodigo,librodescripcion"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaTipAnal 
         Height          =   300
         Left            =   6150
         TabIndex        =   7
         Top             =   390
         Width           =   3810
         _ExtentX        =   6720
         _ExtentY        =   529
         XcodMaxLongitud =   0
         xcodwith        =   1000
         NomTabla        =   "ct_tipoanalitico"
         TituloAyuda     =   "Ayuda de Tipo Analitico"
         ListaCampos     =   "tipoanaliticocodigo(1),tipoanaliticodescripcion(1)"
         XcodCampo       =   "tipoanaliticocodigo"
         XListCampo      =   "tipoanaliticodescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "tipoanaliticocodigo,tipoanaliticodescripcion"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaCtaTotal 
         Height          =   300
         Left            =   6120
         TabIndex        =   8
         Top             =   1485
         Width           =   3900
         _ExtentX        =   6879
         _ExtentY        =   529
         XcodMaxLongitud =   0
         xcodwith        =   1000
         NomTabla        =   "ct_cuenta"
         TituloAyuda     =   "Ayuda de Cuenta Total"
         ListaCampos     =   "cuentacodigo(1),cuentadescripcion(1)"
         XcodCampo       =   "cuentacodigo"
         XListCampo      =   "cuentadescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "cuentacodigo,cuentadescripcion"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaCtaIES 
         Height          =   405
         Left            =   6150
         TabIndex        =   9
         Top             =   900
         Width           =   3810
         _ExtentX        =   6720
         _ExtentY        =   714
         XcodMaxLongitud =   0
         xcodwith        =   1000
         NomTabla        =   "ct_cuenta"
         TituloAyuda     =   "Ayuda de Cuenta IGV"
         ListaCampos     =   "cuentacodigo(1),cuentadescripcion(1)"
         XcodCampo       =   "cuentacodigo"
         XListCampo      =   "cuentadescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "cuentacodigo,cuentadescripcion"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaCtaRTA 
         Height          =   315
         Left            =   165
         TabIndex        =   10
         Top             =   1485
         Width           =   5580
         _ExtentX        =   9843
         _ExtentY        =   556
         XcodMaxLongitud =   0
         xcodwith        =   1000
         NomTabla        =   "ct_cuenta"
         TituloAyuda     =   "Ayuda de Cuenta IGV"
         ListaCampos     =   "cuentacodigo(1),cuentadescripcion(1)"
         XcodCampo       =   "cuentacodigo"
         XListCampo      =   "cuentadescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "cuentacodigo,cuentadescripcion"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaPercepcion 
         Height          =   300
         Left            =   165
         TabIndex        =   52
         Top             =   2025
         Width           =   5520
         _ExtentX        =   9737
         _ExtentY        =   529
         XcodMaxLongitud =   0
         xcodwith        =   1000
         NomTabla        =   "co_gastos"
         TituloAyuda     =   "Ayuda de Cuenta Percepción"
         ListaCampos     =   "gastoscodigo(1),gastosdescripcion(1)"
         XcodCampo       =   "gastoscodigo"
         XListCampo      =   "gastosdescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "gastoscodigo,gastosdescripcion"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_TipoDoc 
         Height          =   300
         Left            =   6120
         TabIndex        =   54
         Top             =   2025
         Width           =   3900
         _ExtentX        =   6879
         _ExtentY        =   529
         XcodMaxLongitud =   0
         xcodwith        =   1000
         NomTabla        =   "cp_tipodocumento"
         TituloAyuda     =   "Ayuda de Documentos"
         ListaCampos     =   "tdocumentocodigo(1),tdocumentodescripcion(1)"
         XcodCampo       =   "tdocumentocodigo"
         XListCampo      =   "tdocumentodescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "tdocumentocodigo,tdocumentodescripcion"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuDocRet 
         Height          =   300
         Left            =   150
         TabIndex        =   56
         Top             =   2625
         Width           =   5520
         _ExtentX        =   9737
         _ExtentY        =   529
         XcodMaxLongitud =   0
         xcodwith        =   1000
         NomTabla        =   "cp_tipodocumento"
         TituloAyuda     =   "Ayuda de Documentos"
         ListaCampos     =   "tdocumentocodigo(1),tdocumentodescripcion(1)"
         XcodCampo       =   "tdocumentocodigo"
         XListCampo      =   "tdocumentodescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "tdocumentocodigo,tdocumentodescripcion"
         Requerido       =   0   'False
      End
      Begin VB.Label lbl 
         Caption         =   "Doc de retencion Honorarios"
         Height          =   210
         Index           =   19
         Left            =   195
         TabIndex        =   57
         Top             =   2430
         Width           =   2550
      End
      Begin VB.Label lbl 
         Caption         =   "Doc. de A cuenta"
         Height          =   210
         Index           =   18
         Left            =   6435
         TabIndex        =   55
         Top             =   1830
         Width           =   1320
      End
      Begin VB.Label Label3 
         Caption         =   "Codigo Percepción"
         Height          =   240
         Left            =   195
         TabIndex        =   53
         Top             =   1845
         Width           =   1455
      End
      Begin VB.Label lbl 
         Caption         =   "Cuenta Total"
         Height          =   210
         Index           =   11
         Left            =   6405
         TabIndex        =   17
         Top             =   1290
         Width           =   1200
      End
      Begin VB.Label lbl 
         Caption         =   "Cuenta IGV"
         Height          =   216
         Index           =   10
         Left            =   192
         TabIndex        =   16
         Top             =   672
         Width           =   1200
      End
      Begin VB.Label lbl 
         Caption         =   "Sub Asiento "
         Height          =   216
         Index           =   6
         Left            =   108
         TabIndex        =   15
         Top             =   192
         Width           =   1200
      End
      Begin VB.Label lbl 
         Caption         =   "Libro "
         Height          =   216
         Index           =   8
         Left            =   1344
         TabIndex        =   14
         Top             =   192
         Width           =   2112
      End
      Begin VB.Label lbl 
         Caption         =   "Tipo Analitico "
         Height          =   210
         Index           =   12
         Left            =   6390
         TabIndex        =   13
         Top             =   195
         Width           =   1200
      End
      Begin VB.Label lbl 
         Caption         =   "Cuenta IES"
         Height          =   210
         Index           =   13
         Left            =   6495
         TabIndex        =   12
         Top             =   690
         Width           =   1200
      End
      Begin VB.Label lbl 
         Caption         =   "Cuenta Renta"
         Height          =   216
         Index           =   14
         Left            =   192
         TabIndex        =   11
         Top             =   1248
         Width           =   1200
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Parametros para generar registros en cuentas por pagar"
      Height          =   924
      Left            =   4080
      TabIndex        =   0
      Top             =   7155
      Width           =   4620
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaTipoPlan 
         Height          =   300
         Left            =   1116
         TabIndex        =   39
         Top             =   216
         Width           =   3372
         _ExtentX        =   5953
         _ExtentY        =   529
         XcodMaxLongitud =   0
         xcodwith        =   500
         NomTabla        =   "cp_tipoplanilla"
         TituloAyuda     =   "Ayuda de Tipo de Planilla"
         ListaCampos     =   "tplanillacodigo(1),tplanilladescripcion(1)"
         XcodCampo       =   "tplanillacodigo"
         XListCampo      =   "tplanilladescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "tplanillacodigo,tplanilladescripcion"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaOficina 
         Height          =   300
         Left            =   1116
         TabIndex        =   40
         Top             =   528
         Width           =   3372
         _ExtentX        =   5953
         _ExtentY        =   529
         XcodMaxLongitud =   0
         xcodwith        =   500
         NomTabla        =   "cp_oficina"
         TituloAyuda     =   "Ayuda de Tipo Analitico"
         ListaCampos     =   "vendedorcodigo(1),vendedornombres(1)"
         XcodCampo       =   "vendedorcodigo"
         XListCampo      =   "vendedornombres"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "vendedorcodigo,vendedornombres"
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Planilla"
         Height          =   288
         Left            =   36
         TabIndex        =   2
         Top             =   276
         Width           =   2292
      End
      Begin VB.Label Label2 
         Caption         =   "Oficina"
         Height          =   288
         Left            =   396
         TabIndex        =   1
         Top             =   564
         Width           =   2148
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2640
      Left            =   90
      TabIndex        =   38
      Top             =   5595
      Width           =   3330
      _ExtentX        =   5874
      _ExtentY        =   4657
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As New ADODB.Recordset
Dim FlagNUEVO As Boolean
Dim ValorMoneda As String * 2

Private Sub Ctr_AyudaCuentaAjuste_AlDevolverDato(Index As Integer, ByVal ColecCampos As ADODB.Fields)
  cmdBotones(0).Enabled = ValidaBoton()
End Sub

Private Sub chk_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call SendKeys("{TAB}")
End Sub

Private Sub Chkempresas_Click()
cmdBotones(0).Enabled = ValidaBoton()
End Sub
Private Sub Chkbancarizacion_click()
 cmdBotones(0).Enabled = ValidaBoton()
 If Chkbancarizacion.Value = 1 Then
    Frame5.Visible = True
  Else
    Frame5.Visible = False
 End If
End Sub

Private Sub Ctr_AyudaCtaIES_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    cmdBotones(0).Enabled = ValidaBoton()
End Sub

Private Sub Ctr_AyudaCtaIGV_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    cmdBotones(0).Enabled = ValidaBoton()
End Sub

Private Sub Ctr_AyudaCtaRTA_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    cmdBotones(0).Enabled = ValidaBoton()
End Sub

Private Sub Ctr_AyudaCtaTotal_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    cmdBotones(0).Enabled = ValidaBoton()
End Sub

Private Sub Ctr_AyudaLibro_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    cmdBotones(0).Enabled = ValidaBoton()
End Sub

Private Sub Ctr_AyudaOficina_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    cmdBotones(0).Enabled = ValidaBoton()
End Sub

Private Sub Ctr_AyudaPercepcion_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    cmdBotones(0).Enabled = ValidaBoton()
End Sub

Private Sub Ctr_AyudaTipAnal_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    cmdBotones(0).Enabled = ValidaBoton()
End Sub

Private Sub Ctr_AyudaTipoPlan_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
        cmdBotones(0).Enabled = ValidaBoton()
End Sub

Private Sub Form_Load()
Dim rsaux As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Call Ctr_AyudaMon.conexion(VGcnxCT)
    Call Ctr_AyudaLibro.conexion(VGcnxCT)
    Call Ctr_AyudaTipAnal.conexion(VGcnxCT)
    Call Ctr_AyudaCtaIGV.conexion(VGcnxCT)
    Call Ctr_AyudaCtaIES.conexion(VGcnxCT)
    Call Ctr_AyudaCtaRTA.conexion(VGcnxCT)
    Call Ctr_AyudaCtaTotal.conexion(VGcnxCT)
    Call Ctr_AyudaTipoPlan.conexion(VGCNx)
    Call Ctr_AyudaOficina.conexion(VGCNx)
    Call Ctr_AyudaPercepcion.conexion(VGCNx)
    Call Ctr_TipoDoc.conexion(VGCNx)
    Call Ctr_AyuDocRet.conexion(VGCNx)
    Set rsaux = New ADODB.Recordset
    rsaux.Open "Select UltNivel=isnull(sistemaultimonivel,0) from co_sistema ", VGCNx, adOpenKeyset, adLockReadOnly
    If rsaux.RecordCount > 0 Then
        VGnumnivgas = rsaux!UltNivel
      Else
        VGnumnivgas = 1
    End If
    Ctr_AyudaCtaIGV.Filtro = "cuentanivel=" & VGnumniveles & " and cuentacodigo like '401%'"
    Ctr_AyudaCtaIES.Filtro = "cuentanivel=" & VGnumniveles & " and cuentacodigo like '4%'"
    Ctr_AyudaCtaRTA.Filtro = "cuentanivel=" & VGnumniveles & " and cuentacodigo like '4%'"
    
    Ctr_AyudaCtaTotal.Filtro = "cuentanivel=" & VGnumniveles & " and cuentacodigo like '42%' "
    
    Call CargarData
    CentrarForm MDIPrincipal, Me
    
 
    cmdBotones(0).Enabled = False
End Sub
Sub CargarData()
  Dim SQL As String
  Set rs = New ADODB.Recordset
  rs.Open "co_sistema", VGCNx
  If rs.RecordCount = 0 Then
    FlagNUEVO = True
    Call LlenarLista
    Exit Sub
  End If
  Call MuestraData
  Call LlenarLista
  Call MarcarLista
End Sub
Sub MuestraData()
Dim I As Integer
Set VGvardllgen = New dllgeneral.dll_general
   txt(0).Text = VGvardllgen.ESNULO(rs!sistemadescripcionempresa, "")
   txt(1).Text = VGvardllgen.ESNULO(rs!sistemadescrcortaempresa, "")
   chk(0).Value = IIf(rs!sistemaesttipodescrempresa = 0, 0, 1)
   chk(1).Value = IIf(VGvardllgen.ESNULO(rs!permite_tc, 0) = 0, 0, 1)
   chk(2).Value = IIf(VGvardllgen.ESNULO(rs!sistemactrlgastos, 0) = 0, 0, 1)
   chk(3).Value = IIf(VGvardllgen.ESNULO(rs!sistemaasientoenlinea, 0) = 0, 0, 1)
   txt(2).Text = VGvardllgen.ESNULO(rs!sistemadireccionempresa, "")
   Ctr_AyudaMon.xclave = rs!monedacodigo: Ctr_AyudaMon.Ejecutar
   txt(5).Text = VGvardllgen.ESNULO(rs!sistemactacomp, "")
   txt(4).Text = VGvardllgen.ESNULO(rs!sistemaigv, 0)
   txt(6).Text = VGvardllgen.ESNULO(rs!sistemasubasiento, "")
   txtcomprobante = ESNULO(rs!cabprovinumero, 1)
   Txtminimoretencion = ESNULO(rs!sistemaminimoretencion, 999999)
   
   Ctr_AyudaLibro.xclave = VGvardllgen.ESNULO(rs!sistemalibro, "00"): Ctr_AyudaLibro.Ejecutar
   Ctr_AyudaTipAnal.xclave = VGvardllgen.ESNULO(rs!sistematipanal, "00"): Ctr_AyudaTipAnal.Ejecutar
   Ctr_AyudaCtaTotal.xclave = VGvardllgen.ESNULO(rs!sistemactatotal, "00"): Ctr_AyudaCtaTotal.Ejecutar
   Ctr_AyudaCtaIGV.xclave = VGvardllgen.ESNULO(rs!sistemactaIGV, "00"): Ctr_AyudaCtaIGV.Ejecutar
   Ctr_AyudaCtaIES.xclave = VGvardllgen.ESNULO(rs!sistemactaIES, "00"): Ctr_AyudaCtaIES.Ejecutar
   Ctr_AyudaCtaRTA.xclave = VGvardllgen.ESNULO(rs!sistemactaRTA, "00"): Ctr_AyudaCtaRTA.Ejecutar
   Ctr_AyudaTipoPlan.xclave = VGvardllgen.ESNULO(rs!sistematipoplan, "00"): Ctr_AyudaTipoPlan.Ejecutar
   Ctr_AyudaOficina.xclave = VGvardllgen.ESNULO(rs!sistemaoficina, "00"): Ctr_AyudaOficina.Ejecutar
   Ctr_AyudaPercepcion.xclave = VGvardllgen.ESNULO(rs!codigopercepcion, "00"): Ctr_AyudaPercepcion.Ejecutar
   Ctr_TipoDoc.xclave = ESNULO(rs!tipodocacuenta, "")
   
   If ESNULO(rs!tipodocacuenta, "") <> "" Then Ctr_TipoDoc.Ejecutar
   
   Chkempresas.Value = IIf(VGvardllgen.ESNULO(rs!sistemamultiempresas, 0) = 0, 0, 1)
   Chkbancarizacion.Value = IIf(VGvardllgen.ESNULO(rs!bancarizacion, 0) = 0, 0, 1)
   TxtMinimobancarizacion01.Text = IIf(VGvardllgen.ESNULO(rs!minimobancarizacion01, 0) = 0, 999999, rs!minimobancarizacion01)
   TxtMinimobancarizacion02.Text = IIf(VGvardllgen.ESNULO(rs!minimobancarizacion02, 0) = 0, 999999, rs!minimobancarizacion02)
   
   
End Sub
Private Sub cmdBotones_Click(Index As Integer)
  Select Case Index
    Case 0:
     If ValidarData() = True Then
       Call GrabarData
       Call CargarParametrosCompras
     End If
    
    Case 2: Unload Me
  End Select
  
End Sub

Function ValidarData() As Boolean
 Dim I As Integer
 Dim flagList As Boolean
 Dim nC As Integer
 Dim SQL As String
 Dim rsX As ADODB.Recordset
  ValidarData = False
  ValidarData = True
End Function

Sub GrabarData()
On Error GoTo X
Dim SQL  As String
Dim strvalor As String
    ValorMoneda = Ctr_AyudaMon.xclave
    Set VGvardllgen = New dllgeneral.dll_general
    If FlagNUEVO = True Then
        Call Grabaparam(1)
        
    Else
        Call Grabaparam(2)
    End If
    
    cmdBotones(0).Enabled = False
    Exit Sub
X:
  MsgBox "Error inesperado: " & Err.Description & "  " & Err.Number
  Exit Sub
  Resume
End Sub
Private Sub Grabaparam(OP As Integer)
Dim strvalor As String
Dim RSQL As New ADODB.Recordset
strvalor = NivelCuenta()
    Set VGCommandoSP = New ADODB.Command
    Set VGvardllgen = New dllgeneral.dll_general
    VGCommandoSP.ActiveConnection = VGGeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "co_grabaparam_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@Op") = OP '1- Inserta ; 2 Actualiza
        .Parameters("@sistemadescripcionempresa") = Trim(txt(0).Text)
        .Parameters("@sistemadescrcortaempresa") = Trim(txt(1).Text)
        .Parameters("@sistemaesttipodescrempresa") = chk(0).Value
        .Parameters("@sistemadireccionempresa") = Trim(txt(2).Text)
        .Parameters("@sistemaempresaruc") = Trim(txt(3).Text)
        .Parameters("@monedacodigo") = ValorMoneda
        .Parameters("@sistemactacomp") = Trim(txt(5).Text)
        .Parameters("@sistemaigv") = txt(4).Text
        .Parameters("@usuariocodigo") = VGUsuario
        .Parameters("@fechaact") = Now
        .Parameters("@sistemasubasiento") = Trim(txt(6).Text)
        .Parameters("@sistemalibro") = Ctr_AyudaLibro.xclave
        .Parameters("@sistematipanal") = Ctr_AyudaTipAnal.xclave
        .Parameters("@sistemactatotal") = Ctr_AyudaCtaTotal.xclave
        .Parameters("@sistemactaIGV") = Ctr_AyudaCtaIGV.xclave
        .Parameters("@sistemactaIES") = Ctr_AyudaCtaIES.xclave
        .Parameters("@sistemactaRTA") = Ctr_AyudaCtaRTA.xclave
        .Parameters("@sistematipoplan") = Ctr_AyudaTipoPlan.xclave
        .Parameters("@sistemaoficina") = Ctr_AyudaOficina.xclave
        .Parameters("@permite_tc") = chk(1).Value
        .Parameters("@sistemaactivaccostos") = chk(2).Value
        .Parameters("@sistemaasientoenlinea") = chk(3).Value
        .Parameters("@sistemamultiempresas") = Chkempresas.Value
        .Parameters("@sistemaactivagastos") = chk(2).Value
        .Parameters("@sistemaconfiguragastos") = strvalor
        .Parameters("@cabprovinumero") = txtcomprobante
        .Parameters("@minimoretencion") = Txtminimoretencion
        .Parameters("@bancarizacion") = Chkbancarizacion
        .Parameters("@minimosoles") = TxtMinimobancarizacion01.Text
        .Parameters("@minimodolares") = TxtMinimobancarizacion02.Text
        .Parameters("@codigopercepcion") = Ctr_AyudaPercepcion.xclave
'        .Parameters("@tipodocacuenta") = Ctr_TipoDoc.xclave
'        .Parameters("@tipodocRetencion") = Ctr_AyuDocRet.xclave
       .Execute
    End With
    Call Parametrogastos
    Set RSQL = VGCNx.Execute("update co_sistema set sistemaultimonivel=" & VGnumnivgas)
End Sub

Function NivelCuenta() As String
 Dim I As Integer
 Dim Valor As String
    Valor = Empty
    For I = 1 To 9
      If ListView1.ListItems.Item(I).Checked = True Then
          Valor = Valor & ListView1.ListItems.Item(I).Text & "*"
      End If
    Next
    NivelCuenta = Left(Valor, Len(Valor) - 1)
End Function

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
  cmdBotones(0).Enabled = ValidaBoton()
End Sub

Private Sub txt_Change(Index As Integer)
  cmdBotones(0).Enabled = ValidaBoton()
End Sub

Private Sub chk_Click(Index As Integer)
  cmdBotones(0).Enabled = ValidaBoton()
End Sub

Function ValidaBoton() As Boolean
 Dim I As Integer
   For I = 0 To 6
     If txt(I).Text = Empty Then
       ValidaBoton = False
       Exit Function
     End If
   Next
   ValidaBoton = True
End Function

Private Sub txt_LostFocus(Index As Integer)
  txt(Index).Text = UCase(txt(Index).Text)
End Sub

Sub LlenarLista()
 Dim I As Integer
 Dim itmX As ListItem
   ListView1.ColumnHeaders.Clear
   ListView1.ListItems.Clear
   ListView1.ColumnHeaders.Add , , "Número Dígitos", ListView1.Width / 1
   ListView1.View = lvwReport
   For I = 1 To 9
     Set itmX = ListView1.ListItems.Add(, , I)
   Next

End Sub

Sub MarcarLista()
 Dim I As Integer
 Dim j As Integer
      
   Call ParametroCuentagastos
   For I = 1 To VGnumnivgas
     For j = 1 To 9
       If ListView1.ListItems.Item(j).Text = VG_gNIVELES(I - 1) Then
          ListView1.ListItems.Item(j).Checked = True
       End If
     Next
   Next

End Sub

Private Sub txtcomprobante_Change()
cmdBotones(0).Enabled = ValidaBoton()
End Sub

Private Sub TxtMinimobancarizacion01_Change()
cmdBotones(0).Enabled = ValidaBoton()
End Sub

Private Sub TxtMinimobancarizacion02_Change()
cmdBotones(0).Enabled = ValidaBoton()
End Sub

Private Sub Txtminimoretencion_Change()
cmdBotones(0).Enabled = ValidaBoton()
End Sub
