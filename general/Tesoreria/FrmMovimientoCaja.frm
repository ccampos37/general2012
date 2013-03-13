VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmMovimientoCaja 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Otros Movimientos"
   ClientHeight    =   9150
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   11655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9150
   ScaleWidth      =   11655
   ShowInTaskbar   =   0   'False
   Begin ctrlayuda_f.Ctr_Ayuda CtrAy_Concepto 
      Height          =   390
      Left            =   330
      TabIndex        =   63
      Top             =   7665
      Visible         =   0   'False
      Width           =   3030
      _ExtentX        =   5345
      _ExtentY        =   688
      Enabled         =   0   'False
      XcodMaxLongitud =   0
      xcodwith        =   300
      NomTabla        =   "te_conceptocaja"
      ListaCampos     =   "conceptocodigo(1),conceptodescripcion(1)"
      XcodCampo       =   "conceptocodigo"
      XListCampo      =   "conceptodescripcion"
      ListaCamposDescrip=   "Código,Descripción"
      ListaCamposText =   "conceptocodigo,conceptodescripcion"
      Requerido       =   0   'False
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C0C0&
      Height          =   990
      Left            =   3480
      TabIndex        =   56
      Top             =   7785
      Width           =   4230
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Salir"
         Height          =   690
         Index           =   7
         Left            =   3255
         Picture         =   "FrmMovimientoCaja.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   210
         Width           =   870
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Eliminar"
         Height          =   690
         Index           =   6
         Left            =   2250
         Picture         =   "FrmMovimientoCaja.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   210
         Width           =   825
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Grabar"
         Height          =   690
         Index           =   5
         Left            =   1230
         Picture         =   "FrmMovimientoCaja.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   210
         Width           =   870
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Nuevo"
         Height          =   690
         Index           =   4
         Left            =   180
         Picture         =   "FrmMovimientoCaja.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   210
         Width           =   870
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8850
      Left            =   210
      TabIndex        =   11
      Top             =   0
      Width           =   11160
      Begin VB.Frame Frame2 
         Height          =   2085
         Left            =   180
         TabIndex        =   42
         Top             =   120
         Width           =   10875
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   990
            MaxLength       =   50
            TabIndex        =   6
            Top             =   1320
            Width           =   4740
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00400000&
            BorderStyle     =   0  'None
            Height          =   525
            Left            =   5910
            TabIndex        =   43
            Top             =   1320
            Width           =   4755
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               Enabled         =   0   'False
               Height          =   285
               Index           =   1
               Left            =   3390
               TabIndex        =   48
               Top             =   180
               Width           =   1365
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               Enabled         =   0   'False
               Height          =   285
               Index           =   0
               Left            =   1440
               TabIndex        =   47
               Top             =   180
               Width           =   1365
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "US$"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Index           =   2
               Left            =   2910
               TabIndex        =   46
               Top             =   240
               Width           =   435
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "S/."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Index           =   1
               Left            =   1080
               TabIndex        =   45
               Top             =   240
               Width           =   345
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackColor       =   &H00004000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "TOTAL"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0080FF80&
               Height          =   390
               Index           =   0
               Left            =   120
               TabIndex        =   44
               Top             =   120
               Width           =   915
            End
         End
         Begin VB.CommandButton cayuda 
            Caption         =   "..."
            Height          =   315
            Index           =   0
            Left            =   7860
            TabIndex        =   9
            Top             =   240
            Width           =   225
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   7368
            MaxLength       =   2
            TabIndex        =   3
            Top             =   240
            Width           =   465
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   3
            Left            =   4920
            MaxLength       =   10
            TabIndex        =   8
            Top             =   1680
            Width           =   855
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1680
            Width           =   1965
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   3228
            MaxLength       =   6
            TabIndex        =   1
            Top             =   216
            Width           =   1008
         End
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Left            =   768
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   180
            Width           =   1512
         End
         Begin MSMask.MaskEdBox MBox1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   3
            EndProperty
            Height          =   312
            Left            =   5268
            TabIndex        =   2
            Top             =   216
            Width           =   1212
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuempresa 
            Height          =   315
            Left            =   1005
            TabIndex        =   5
            Top             =   960
            Width           =   4815
            _ExtentX        =   8493
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
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaCaja 
            Height          =   315
            Left            =   1005
            TabIndex        =   4
            Top             =   600
            Width           =   3900
            _ExtentX        =   6879
            _ExtentY        =   556
            XcodMaxLongitud =   11
            xcodwith        =   400
            NomTabla        =   "te_codigocaja"
            TituloAyuda     =   "Busqueda de Caja"
            ListaCampos     =   "cajacodigo(1),cajadescripcion(1),cajarendiciones(2)"
            XcodCampo       =   "cajacodigo"
            XListCampo      =   "cajadescripcion"
            ListaCamposDescrip=   "Codigo,Descripcion,controla Rendicion"
            ListaCamposText =   "cajacodigo,cajadescripcion,cajarendiciones"
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayutransf 
            Height          =   315
            Left            =   6720
            TabIndex        =   71
            Top             =   600
            Visible         =   0   'False
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   556
            XcodMaxLongitud =   7
            xcodwith        =   800
            NomTabla        =   "te_cabecerarecibos"
            TituloAyuda     =   "Busqueda de Documentos x rendir"
            ListaCampos     =   "cabrec_numreciboegreso(1),cabrec_descripcion(1),SaldoDocxRendir(1),cabrec_fechadocumento(2),clientecodigo(1)"
            XcodCampo       =   "cabrec_numreciboegreso"
            XListCampo      =   "cabrec_descripcion"
            ListaCamposDescrip=   "Nro.transferencia,descripcion,Saldo,Fecha docuemnto, clientecodigo"
            ListaCamposText =   "cabrec_numreciboegreso,cabrec_descripcion,SaldoDocxRendir,cabrec_fechadocumento, clientecodigo"
         End
         Begin VB.Label LeReferencia 
            AutoSize        =   -1  'True
            Caption         =   "Nro.Transf."
            Height          =   195
            Left            =   5880
            TabIndex        =   70
            Top             =   660
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.Label Lblempresa 
            AutoSize        =   -1  'True
            Caption         =   "Empresa :"
            Height          =   195
            Left            =   180
            TabIndex        =   68
            Top             =   990
            Width           =   705
         End
         Begin VB.Label Label1 
            Caption         =   "Glosa"
            Height          =   255
            Index           =   4
            Left            =   180
            TabIndex        =   62
            Top             =   1335
            Width           =   1110
         End
         Begin VB.Label Label2 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   0
            Left            =   8100
            TabIndex        =   10
            Top             =   240
            Width           =   2565
         End
         Begin VB.Label Label1 
            Caption         =   "Moneda"
            Height          =   255
            Index           =   7
            Left            =   180
            TabIndex        =   55
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Cambio"
            Height          =   255
            Index           =   6
            Left            =   3360
            TabIndex        =   54
            Top             =   1680
            Width           =   1005
         End
         Begin VB.Label Label1 
            Caption         =   "Cod. Caja"
            Height          =   375
            Index           =   5
            Left            =   180
            TabIndex        =   53
            Top             =   660
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Operacion"
            Height          =   252
            Index           =   3
            Left            =   6540
            TabIndex        =   52
            Top             =   240
            Width           =   792
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Doc."
            Height          =   252
            Index           =   2
            Left            =   4392
            TabIndex        =   51
            Top             =   240
            Width           =   852
         End
         Begin VB.Label Label1 
            Caption         =   "No.Recibo"
            Height          =   252
            Index           =   1
            Left            =   2328
            TabIndex        =   50
            Top             =   216
            Width           =   828
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo"
            Height          =   252
            Index           =   0
            Left            =   180
            TabIndex        =   49
            Top             =   216
            Width           =   732
         End
      End
      Begin VB.Frame Frame3 
         Height          =   5625
         Left            =   180
         TabIndex        =   12
         Top             =   2070
         Width           =   10890
         Begin VB.Frame Frame4 
            Height          =   2025
            Left            =   90
            TabIndex        =   14
            Top             =   3450
            Width           =   10575
            Begin VB.TextBox Text2 
               Height          =   300
               Index           =   10
               Left            =   3210
               MaxLength       =   50
               TabIndex        =   31
               Top             =   1560
               Width           =   7110
            End
            Begin VB.CommandButton cayuda 
               Caption         =   "..."
               Height          =   285
               Index           =   7
               Left            =   2850
               TabIndex        =   30
               Top             =   1560
               Width           =   195
            End
            Begin VB.TextBox Text2 
               Height          =   300
               Index           =   9
               Left            =   120
               MaxLength       =   30
               TabIndex        =   29
               Top             =   1560
               Width           =   2715
            End
            Begin VB.CommandButton cayuda 
               Caption         =   "..."
               Height          =   285
               Index           =   6
               Left            =   7530
               TabIndex        =   24
               Top             =   405
               Width           =   195
            End
            Begin VB.CommandButton cayuda 
               Caption         =   "..."
               Height          =   315
               Index           =   5
               Left            =   5400
               TabIndex        =   21
               Top             =   405
               Width           =   255
            End
            Begin VB.CommandButton cayuda 
               Caption         =   "..."
               Height          =   315
               Index           =   4
               Left            =   4575
               TabIndex        =   19
               Top             =   405
               Width           =   255
            End
            Begin VB.CommandButton cayuda 
               Caption         =   "..."
               Height          =   315
               Index           =   2
               Left            =   3705
               TabIndex        =   17
               Top             =   390
               Width           =   270
            End
            Begin VB.TextBox Text2 
               Height          =   285
               Index           =   8
               Left            =   7770
               MaxLength       =   10
               TabIndex        =   25
               Top             =   420
               Width           =   1140
            End
            Begin VB.TextBox Text2 
               Height          =   285
               Index           =   7
               Left            =   7050
               MaxLength       =   2
               TabIndex        =   23
               Top             =   405
               Width           =   465
            End
            Begin VB.TextBox Text2 
               Height          =   285
               Index           =   6
               Left            =   5655
               MaxLength       =   10
               TabIndex        =   22
               Top             =   405
               Width           =   1395
            End
            Begin VB.TextBox Text2 
               Height          =   285
               Index           =   5
               Left            =   4845
               MaxLength       =   2
               TabIndex        =   20
               Top             =   405
               Width           =   510
            End
            Begin VB.TextBox Text2 
               Height          =   285
               Index           =   4
               Left            =   3990
               MaxLength       =   2
               TabIndex        =   18
               Top             =   405
               Width           =   555
            End
            Begin VB.TextBox Text2 
               Height          =   285
               Index           =   1
               Left            =   615
               MaxLength       =   2
               TabIndex        =   16
               Top             =   390
               Width           =   390
            End
            Begin VB.TextBox Text2 
               BackColor       =   &H80000000&
               Enabled         =   0   'False
               Height          =   285
               Index           =   0
               Left            =   120
               MaxLength       =   2
               TabIndex        =   15
               Top             =   390
               Width           =   465
            End
            Begin MSMask.MaskEdBox MBox2 
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
               Left            =   9000
               TabIndex        =   26
               Top             =   480
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_gastos 
               Height          =   330
               Left            =   120
               TabIndex        =   27
               Top             =   915
               Visible         =   0   'False
               Width           =   3075
               _ExtentX        =   5424
               _ExtentY        =   582
               XcodMaxLongitud =   10
               xcodwith        =   700
               NomTabla        =   "co_gastos"
               ListaCampos     =   "gastoscodigo(1),gastosdescripcion(1),tipoanaliticocodigo(1),gastosctrlcostos(1)"
               XcodCampo       =   "gastoscodigo"
               XListCampo      =   "gastosdescripcion"
               ListaCamposDescrip=   "Codigo,Descripcion"
               ListaCamposText =   "gastoscodigo,gastosdescripcion,tipoanaliticocodigo,gastosctrlcostos"
            End
            Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuAnalitico 
               Height          =   315
               Left            =   3465
               TabIndex        =   28
               Top             =   915
               Visible         =   0   'False
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   556
               XcodMaxLongitud =   10
               xcodwith        =   900
               NomTabla        =   "v_analiticoentidad"
               TituloAyuda     =   "Busqueda de Centro de Costos"
               ListaCampos     =   "entidadcodigo(1),entidadrazonsocial(1),entidaddireccion(1)"
               XcodCampo       =   "entidadcodigo"
               XListCampo      =   "entidadrazonsocial"
               ListaCamposDescrip=   "Código,Descripción, cliente"
               ListaCamposText =   "entidadcodigo,entidadrazonsocial,entidaddireccion"
               Requerido       =   0   'False
            End
            Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_Ccosto 
               Height          =   315
               Left            =   7200
               TabIndex        =   69
               Top             =   915
               Visible         =   0   'False
               Width           =   3135
               _ExtentX        =   5530
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
            Begin VB.Label Lblanalitico 
               AutoSize        =   -1  'True
               Caption         =   "Analitico"
               Height          =   195
               Left            =   4560
               TabIndex        =   67
               Top             =   680
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.Label Lbccosto 
               AutoSize        =   -1  'True
               Caption         =   "C.Costos :"
               Height          =   195
               Left            =   8040
               TabIndex        =   66
               Top             =   680
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.Label Lblgastos 
               AutoSize        =   -1  'True
               Caption         =   "Gastos :"
               Height          =   195
               Left            =   1080
               TabIndex        =   65
               Top             =   680
               Visible         =   0   'False
               Width           =   585
            End
            Begin VB.Label Label2 
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   2
               Left            =   1020
               TabIndex        =   64
               Top             =   390
               Width           =   2625
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "Observaciones"
               Height          =   180
               Index           =   11
               Left            =   3240
               TabIndex        =   41
               Top             =   1305
               Width           =   6165
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "Nro. Cuenta Corriente"
               Height          =   180
               Index           =   10
               Left            =   180
               TabIndex        =   40
               Top             =   1290
               Width           =   2625
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "Fec. Cancela"
               Height          =   210
               Index           =   9
               Left            =   9090
               TabIndex        =   39
               Top             =   150
               Width           =   975
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "Importe"
               Height          =   210
               Index           =   8
               Left            =   7860
               TabIndex        =   38
               Top             =   165
               Width           =   765
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "Mon."
               Height          =   180
               Index           =   7
               Left            =   7110
               TabIndex        =   37
               Top             =   165
               Width           =   465
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "Numero"
               Height          =   180
               Index           =   6
               Left            =   5820
               TabIndex        =   36
               Top             =   165
               Width           =   915
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "Banco"
               Height          =   180
               Index           =   5
               Left            =   4800
               TabIndex        =   35
               Top             =   165
               Width           =   615
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "TD.Canc."
               Height          =   180
               Index           =   4
               Left            =   3990
               TabIndex        =   34
               Top             =   165
               Width           =   645
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "Tipo Concepto"
               Height          =   180
               Index           =   1
               Left            =   945
               TabIndex        =   33
               Top             =   150
               Width           =   1140
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "Item"
               Height          =   180
               Index           =   0
               Left            =   15
               TabIndex        =   32
               Top             =   150
               Width           =   645
            End
         End
         Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
            Height          =   3165
            Left            =   120
            TabIndex        =   13
            Top             =   0
            Width           =   10590
            _ExtentX        =   18680
            _ExtentY        =   5583
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
            AllowUpdate     =   0   'False
            Appearance      =   2
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=236,.bold=0,.fontsize=825,.italic=0"
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
      End
   End
   Begin MSComctlLib.StatusBar Panel 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   61
      Top             =   8805
      Width           =   11655
      _ExtentX        =   20558
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
Attribute VB_Name = "FrmMovimientoCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adll As New dllgeneral.dll_general
Dim rsdetat As New ADODB.Recordset
Dim fecharendicion As Date
Dim controlarendicion As Boolean
Dim m_fondofijo As Integer
Dim m_docxrendir As Integer
Dim saldodocxrendir As Double
Dim clientecodigo As String

Property Let docxrendir(valor As String)
   m_docxrendir = valor
End Property
Property Let fondofijo(valor As String)
   m_fondofijo = valor
End Property


Public Sub ConfigGrid()
   With TDBGrid1
       .Columns(0).Width = 600
       .Columns(1).Width = 600
       .Columns(2).Width = 1500
       .Columns(3).Width = 1000
       .Columns(4).Width = 600
       .Columns(5).Width = 700
       .Columns(6).Width = 700
       .Columns(7).Width = 1500
       .Columns(8).Width = 600
       .Columns(9).HeadAlignment = dbgCenter
       .Columns(10).Width = 1300
       .Columns(11).NumberFormat = "##,###,##0.00"
       .Columns(12).Width = 1000
       .Columns(13).Width = 1000
       '.Columns(14).Width = 2000
       .Refresh
   End With
End Sub
Public Sub cargar_grilla()
   Set rsdetat = Nothing
   Call rsdetat.Fields.Append("Item", adChar, 3)
   Call rsdetat.Fields.Append("Tipo", adChar, 2)
   Call rsdetat.Fields.Append("Numero", adChar, 14)
   Call rsdetat.Fields.Append("T/P", adChar, 1)
   Call rsdetat.Fields.Append("T.Canc", adChar, 2)
   Call rsdetat.Fields.Append("Banco", adChar, 2)
   Call rsdetat.Fields.Append("Numero Doc", adChar, 20)
   Call rsdetat.Fields.Append("Mnda", adChar, 2)
   Call rsdetat.Fields.Append("Importe", adDouble)
   Call rsdetat.Fields.Append("Fecha Canc", adDate)
   Call rsdetat.Fields.Append("Cta Cte", adChar, 30)
   Call rsdetat.Fields.Append("Observaciones", adChar, 50)
   Call rsdetat.Fields.Append("analitico", adVarChar, 11)
   Call rsdetat.Fields.Append("entidad", adVarChar, 10)
   Call rsdetat.Fields.Append("ccosto", adVarChar, 10)
   
   
   rsdetat.Open
   Set TDBGrid1.DataSource = rsdetat
   TDBGrid1.Refresh
   Call ConfigGrid
   
End Sub

Private Sub cAyuda_Click(Index As Integer)
 Dim rb As New ADODB.Recordset
 nAyuda = "": nDetalle = ""
  If Index = 0 Then
         If Len(Trim(Text1(1))) > 0 Then
           SendKeys "{tab}"
           Exit Sub
         End If
         Dim dfiltra(1, 2) As String
         dfiltra(1, 1) = "Codigo": dfiltra(1, 2) = "operacioncodigo"
         FrmAyudaTes.TipoForma = 1
         FrmAyudaTes.BConexion = VGCNx
         FrmAyudaTes.Bdata = "0"
         FrmAyudaTes.BTabla = "te_operaciongeneral"
         FrmAyudaTes.BCampos = "operacioncodigo as Codigo,operaciondescripcion as Descripcion"
         FrmAyudaTes.BOrden = "operacioncodigo"
         FrmAyudaTes.BCondi = "operacioncontrolaclienteprov<>'P' and  operacioncontrolaclienteprov<>'C' and operacionvalidacajabancos<>'X'"
         FrmAyudaTes.BFiltro = dfiltra
         FrmAyudaTes.Show 1
         Text1(1).Text = nAyuda
         Label2(0) = nDetalle
         Call Text1_KeyPress(1, 13)
         
    ElseIf Index = 1 Then
         Set rb = VGCNx.Execute("select * from te_operaciongeneral where operacioncodigo='" & Escadena(Text1(1)) & "' and operacioncontrolaclienteprov='" & IIf(adll.ComboDato(Combo1) = "I", "C", "C") & "'")
         If rb.RecordCount > 0 Then
            If Escadena(rb!operacionvalidacajabancos) = "B" Then
                Ctr_AyudaCaja.Enabled = False
                cayuda(1).Enabled = False
                Combo2.SetFocus
                rb.Close
                Set rb = Nothing
                Exit Sub
            Else
                Ctr_AyudaCaja.Enabled = True
                cayuda(1).Enabled = True
                Ctr_AyudaCaja.SetFocus
            End If
        End If
        rb.Close
        Set rb = Nothing
         
        If Len(Trim(Ctr_AyudaCaja.xclave)) > 0 Then
           SendKeys "{tab}"
           Exit Sub
        End If
        
        Dim gfiltra(1, 2) As String
        gfiltra(1, 1) = "Codigo": gfiltra(1, 2) = "cajacodigo"
        FrmAyudaTes.TipoForma = 1
        FrmAyudaTes.BConexion = VGCNx
        FrmAyudaTes.Bdata = "0"
        FrmAyudaTes.BTabla = "te_codigocaja"
        FrmAyudaTes.BCampos = "cajacodigo as Codigo,cajadescripcion as Descripcion"
        FrmAyudaTes.BOrden = "cajacodigo"
        FrmAyudaTes.BCondi = ""
        FrmAyudaTes.BFiltro = gfiltra
        FrmAyudaTes.Show 1
        Ctr_AyudaCaja.xclave = nAyuda
        Label2(1) = nDetalle
        SendKeys "{tab}"
     ElseIf Index = 2 Then
       If Len(Trim(Text2(1))) > 0 Then
          SendKeys "{tab}"
          Exit Sub
        End If

        If adll.VerificaDatoExistente(VGCNx, "select * from te_conceptocaja") = 1 Then
            Dim zfiltra(1, 2) As String
            zfiltra(1, 1) = "Código": zfiltra(1, 2) = "conceptocodigo"
            FrmAyudaTes.TipoForma = 1
            FrmAyudaTes.BConexion = VGCNx
            FrmAyudaTes.Bdata = "0"
            FrmAyudaTes.BTabla = "te_conceptocaja"
            FrmAyudaTes.BCampos = "conceptocodigo as Codigo,conceptodescripcion as Descripcion"
            FrmAyudaTes.BOrden = "conceptocodigo"
            FrmAyudaTes.BCondi = Empty
            FrmAyudaTes.BFiltro = zfiltra
            FrmAyudaTes.Show 1
            Text2(1).Text = nAyuda
            Label2(2).Caption = nDetalle
          '  Ctrayu_ccostos.Filtro = FiltroCcosto(nAyuda, flag)
          '  Ctrayu_ccostos.Visible = flag
          '  lbccosto.Visible = flag

         Else
             nAyuda = "": nDetalle = ""
             MsgBox "No existen Conceptos en Tesorería...", vbInformation, MsgTitle
             Exit Sub
         End If

    ElseIf Index = 3 Then
        If Len(Trim(Text2(2))) > 0 Then
           SendKeys "{tab}"
           Exit Sub
         End If
  ElseIf Index = 4 Then   'Tipo de cancelacion
    If Len(Trim(Text2(4))) > 0 Then
      SendKeys "{tab}"
      Exit Sub
    End If
    If adll.VerificaDatoExistente(VGCNx, "select * from cp_tipodocumento where tdocumentotipo='A' and tdocumentoingcobra='1'") = 1 Then
        Dim ffiltra(1, 2) As String
        ffiltra(1, 1) = "Documento": ffiltra(1, 2) = "tdocumentocodigo"
        FrmAyudaTes.TipoForma = 1
        FrmAyudaTes.BConexion = VGCNx
        FrmAyudaTes.Bdata = "0"
        FrmAyudaTes.BTabla = "cp_tipodocumento"
        FrmAyudaTes.BCampos = "tdocumentocodigo as Codigo,tdocumentodescripcion as Descripcion"
        FrmAyudaTes.BOrden = "tdocumentocodigo"
        FrmAyudaTes.BCondi = "tdocumentotipo='A' and tdocumentocancela='1'"
        FrmAyudaTes.BFiltro = ffiltra
        FrmAyudaTes.Show 1
        Text2(4) = nAyuda
     Else
         nAyuda = "": nDetalle = ""
         MsgBox "No existen Documentos...", vbInformation, MsgTitle
         Exit Sub
     End If
     Exit Sub
   ElseIf Index = 5 Then    'Tipo de Banco
        If Len(Trim(Text2(5))) > 0 Then
           SendKeys "{tab}"
           Exit Sub
        End If
        Dim tfiltra(1, 2) As String
        tfiltra(1, 1) = "Banco": tfiltra(1, 2) = "bancodescripcion"
        FrmAyudaTes.TipoForma = 1
        FrmAyudaTes.BConexion = VGCNx
        FrmAyudaTes.Bdata = "0"
        FrmAyudaTes.BTabla = "gr_banco a INNER JOIN te_cuentabancos b ON a.bancocodigo=b.cbanco_codigo"
        FrmAyudaTes.BCampos = "DISTINCT a.bancocodigo as Codigo,a.bancodescripcion as Descripcion"
        FrmAyudaTes.BOrden = "a.bancocodigo"
        FrmAyudaTes.BCondi = "b.empresacodigo='" & Ctr_Ayuempresa.xclave & "'"
        FrmAyudaTes.BFiltro = tfiltra
        FrmAyudaTes.Show 1
        Text2(5) = nAyuda
   ElseIf Index = 6 Then    'Tipo de Moneda
        If Len(Trim(Text2(7))) > 0 Then
           SendKeys "{tab}"
           Exit Sub
        End If
        Dim pfiltra(1, 2) As String
        pfiltra(1, 1) = "Codigo": pfiltra(1, 2) = "monedacodigo"
        FrmAyudaTes.TipoForma = 1
        FrmAyudaTes.BConexion = VGCNx
        FrmAyudaTes.Bdata = "0"
        FrmAyudaTes.BTabla = "gr_moneda"
        FrmAyudaTes.BCampos = "monedacodigo as Codigo,monedadescripcion as Descripcion"
        FrmAyudaTes.BOrden = "monedacodigo"
        FrmAyudaTes.BCondi = ""
        FrmAyudaTes.BFiltro = pfiltra
        FrmAyudaTes.Show 1
        Text2(7) = nAyuda
   ElseIf Index = 7 Then    'Nro Cuenta Corriente
        If Len(Trim(Text2(9))) > 0 Then
           SendKeys "{tab}"
           Exit Sub
        End If
        Dim qfiltra(1, 2) As String
        qfiltra(1, 1) = "Banco": qfiltra(1, 2) = "bancocodigo"
        FrmAyudaTes.TipoForma = 1
        FrmAyudaTes.BConexion = VGCNx
        FrmAyudaTes.Bdata = "0"
        FrmAyudaTes.BTabla = "te_cuentabancos inner join gr_banco on te_cuentabancos.cbanco_codigo=gr_banco.bancocodigo"
        FrmAyudaTes.BCampos = "cbanco_numero as NoCtaCte,monedacodigo as Moneda,bancocodigo as CodBan,bancodescripcion as Banco"
        FrmAyudaTes.BOrden = "gr_banco.bancocodigo"
        FrmAyudaTes.BCondi = "gr_banco.bancocodigo='" & Text2(5).Text & "' and te_cuentabancos.monedacodigo='" & Text2(7).Text & "' and te_cuentabancos.empresacodigo='" & Ctr_Ayuempresa.xclave & "'"
        FrmAyudaTes.BFiltro = qfiltra
        FrmAyudaTes.Show 1
        Text2(9).Text = nAyuda
   End If
   nAyuda = "": nDetalle = ""
End Sub

Public Function GrabarData() As Integer
On Error GoTo X
  Dim acmd As New ADODB.Command
  Dim rb As New ADODB.Recordset
  Dim xabono, xzona, xmone, xcuenta, xtipo As String
  Dim xnumplan, ximpsol, xtcam, xnumpag As Double
  VGCNx.BeginTrans
    GrabarData = 0
    'Actualizamos el numerador de tipo de ingreso
    Set rb = VGCNx.Execute("select * from te_parametroempresa where empresacodigo='" & VGCodEmpresa & "'")
    If rb.RecordCount > 0 Then
     If adll.ComboDato(Combo1.Text) = "I" Then
        Text1(0) = Right("0000000000" & Trim(CStr(CDbl(IIf(IsNull(rb!empresanumeingreso) Or Len(Trim(rb!empresanumeingreso)) = 0, 1, rb!empresanumeingreso + 1)))), 6)
        VGCNx.Execute "Update te_parametroempresa Set empresanumeingreso='" & Right("0000000000" & Trim(CStr(Val(Text1(0)))), 6) & "' where empresacodigo='" & VGCodEmpresa & "'"
         
     ElseIf adll.ComboDato(Combo1.Text) = "E" Then
        Text1(0) = Right("0000000000" & Trim(CStr(CDbl(IIf(IsNull(rb!empresanumegreso) Or Len(Trim(rb!empresanumegreso)) = 0, 1, rb!empresanumegreso + 1)))), 6)
        VGCNx.Execute "Update te_parametroempresa Set empresanumegreso='" & Right("0000000000" & Trim(CStr(Val(Text1(0)))), 6) & "' where empresacodigo='" & VGCodEmpresa & "'"
     End If
    End If
    rb.Close
    Set rb = Nothing
    
VGCNx.CommitTrans
VGCNx.BeginTrans
    Set acmd.ActiveConnection = VGGeneral
    acmd.CommandType = adCmdStoredProc
    acmd.CommandText = "te_abonadocumento_pro"
    acmd.CommandTimeout = 0
    acmd.Prepared = True
    With acmd
        .Parameters("@base") = VGCNx.DefaultDatabase
        .Parameters("@tipo") = "1"
        .Parameters("@numrecibo") = Escadena(Text1(0))
        .Parameters("@estadoreg") = ""
        .Parameters("@controlctacte") = "0"
        .Parameters("@vendedorcodigo") = VGoficina
        .Parameters("@cajacodigo") = Escadena(Ctr_AyudaCaja.xclave)
        .Parameters("@clientecodigo") = ""
        .Parameters("@descripcion") = Escadena(Text1(4).Text)
        .Parameters("@operacion") = Escadena(Text1(1))
        .Parameters("@monedacodigo") = adll.ComboDato(Combo2)
        .Parameters("@ingsal") = adll.ComboDato(Combo1)
        .Parameters("@tipocambio") = CDbl(Text1(3))
        .Parameters("@totsoles") = CDbl(Label5(0))
        .Parameters("@totdolares") = CDbl(Label5(1))
        .Parameters("@fechadocumento") = MBox1.Text
        .Parameters("@observa") = ""
        .Parameters("@transferauto") = ""
        .Parameters("@numreciboegreso") = Ctr_Ayutransf.xclave
        .Parameters("@usuario") = VGUsuario
        .Parameters("@fechaact") = Now
        .Parameters("@empresa") = Escadena(Ctr_Ayuempresa.xclave)
        If Ctr_Ayutransf.Visible = True Then
           .Parameters("@NumeroDocXRendir") = Ctr_Ayutransf.xclave
           .Parameters("@responsablectasxrendir") = clientecodigo
        End If
     End With
     acmd.Execute
     Set acmd = Nothing
       If rsdetat.RecordCount > 0 Then
          rsdetat.MoveLast
          rsdetat.MoveFirst
          Do Until rsdetat.EOF
             xabono = "": xcuenta = "": xtipo = ""
             
             ' Registramos datos en Tesoreria
             Dim VlDllgeneral As New dll_general
             
             xtcam = CDbl(Text1(3))
             Set acmd.ActiveConnection = VGGeneral
             acmd.CommandType = adCmdStoredProc
             acmd.CommandText = "te_abonadetalledocumento_pro"
             acmd.CommandTimeout = 0
             acmd.Prepared = True
             With acmd
                 .Parameters("@base") = VGCNx.DefaultDatabase
                 .Parameters("@tipo") = "1"
                 .Parameters("@numrecibo") = Text1(0).Text
                 .Parameters("@estadoreg") = ""
                 .Parameters("@item") = rsdetat.Fields(0)
                 .Parameters("@emisioncheque") = IIf(Len(Trim(Ctr_AyudaCaja.xclave)) = 0, "B", "C") ' ver si es cheque
                 .Parameters("@tipodocconcepto") = rsdetat.Fields(1)
                 .Parameters("@numdocumento") = ""
                 .Parameters("@carabo") = xabono
                 .Parameters("@formacan") = ""
                 .Parameters("@tdqc") = rsdetat.Fields(4)
                 .Parameters("@ndqc") = Trim(rsdetat.Fields(6))
                 If Ctr_Ayutransf.Visible = True Then
                    .Parameters("@tdqc") = "20"
                    .Parameters("@ndqc") = Ctr_Ayutransf.xclave
                 End If
                 .Parameters("@tipocajabanco") = IIf(Len(Trim(Ctr_AyudaCaja.xclave)) = 0, "B", "C")
                 .Parameters("@cajabanco") = IIf(Len(Trim(Ctr_AyudaCaja.xclave)) = 0, Trim(rsdetat.Fields(5)), Trim(Ctr_AyudaCaja.xclave))
                 .Parameters("@numctacte") = Escadena(rsdetat.Fields(10))    'numero de cuenta corriente
                 .Parameters("@adicionactacte") = ""
                 .Parameters("@monedadocumento") = ""
                 .Parameters("@monedacancela") = Escadena(rsdetat.Fields(7))
                 .Parameters("@importesoles") = CDbl(IIf(rsdetat.Fields(7) = g_tiposol, rsdetat.Fields(8), (rsdetat.Fields(8) * xtcam)))
                 .Parameters("@importedolares") = CDbl(IIf(rsdetat.Fields(7) = g_tiposol, (rsdetat.Fields(8) / xtcam), rsdetat.Fields(8)))
       '          .Parameters("@contabledisponi") = Escadena(VGParametros.saldocontadispo)      'sale de empresas
                 .Parameters("@fechacancela") = Format(rsdetat.Fields(9), "dd/mm/yyyy")
                 .Parameters("@observacion") = Escadena(rsdetat.Fields(11))
                 .Parameters("@gastos") = Escadena(rsdetat.Fields(12))
                 .Parameters("@usuario") = VGUsuario
                 .Parameters("@fechaact") = Now
                 .Parameters("@entidad") = Escadena(rsdetat.Fields(13))
                 .Parameters("@Centrocosto") = VlDllgeneral.ESNULO(rsdetat!Ccosto, "00")
             End With
             acmd.Execute
             Set acmd = Nothing
             DoEvents
             rsdetat.MoveNext
         Loop
    End If
    Set rsdetat = Nothing
'    If VGParametros.controlaestadosrendicion Then
'       If controlarendicion Then Call Actualizasaldorendicion
'    End If
VGCNx.CommitTrans
    GrabarData = 1
    MsgBox "Los datos han sido grabados satisfactoriamente...!!!", vbInformation, MsgTitle
    Exit Function
X:
  GrabarData = 0
 VGCNx.RollbackTrans
  MsgBox "No se pudo Grabar " & Err.Description & " - " & Err.Number, vbInformation, Caption
  Exit Function
End Function

Private Sub cmdBotones_Click(Index As Integer)
  Dim nvalor As String
  
  On Error Resume Next
  
  Select Case Index
    Case 4
       Frame4.Enabled = True
       Call Limpiartexto(Text2, 0, 8)
       Frame4.Enabled = False
       Frame2.Enabled = True
       Call Limpiartexto(Text1, 0, 3)
       Set rsdetat = Nothing
       Call cargar_grilla
       Call ConfigGrid
       Combo1.SetFocus
       Label2(0).Caption = Empty
       Label2(1).Caption = Empty
       cmdBotones(5).Enabled = True
       cmdBotones(4).Enabled = False
       Call Combo1_Click
    
    Case 5
       If ValidarGrabacion() = 1 Then
          'Grabamos Cabecera de Tesoreria
          Screen.MousePointer = 11
          If GrabarData() = 1 Then
             'Generando el Asiento Contable en Linea
             If VGParametros.sistemaasientoenlinea Then
                Call GeneraAsientoEnlineaTesor(CDate(MBox1.Text), Ctr_Ayuempresa.xclave, "X", Escadena(Text1(0)), 1, "''''", adll.ComboDato(Combo2), IIf(Len(Trim(Ctr_AyudaCaja.xclave)) = 0, "B", "C"), adll.ComboDato(Combo1))
             End If
             If MsgBox("Desea Imprimir el Recibo ", vbQuestion + vbOKCancel) = vbOK Then
               Screen.MousePointer = 1
               Call ImprimirRecibo(Escadena(Text1(0).Text))
             End If
          Else
             Screen.MousePointer = 1
             MsgBox "Error al Grabar: " & Err.Description & " - " & Err.Number, vbInformation, MsgTitle
          End If
          Call Combo1_Click
          Screen.MousePointer = 1
          cmdBotones(5).Enabled = False
          Frame2.Enabled = True
          Call Limpiartexto(Text1, 0, 3)
          Combo1.SetFocus
          cmdBotones(4).Enabled = True
       End If
         
    Case 6
      If rsdetat.RecordCount > 0 Then
       nvalor = TDBGrid1.Columns(0).Text
       If rsdetat.RecordCount > 0 Then
          rsdetat.MoveFirst
          Do Until rsdetat.EOF
            If rsdetat.Fields(0) = nvalor Then
              rsdetat.Delete adAffectCurrent
              rsdetat.Update
              Exit Do
            End If
            rsdetat.MoveNext
          Loop
       End If
      End If
      TDBGrid1.Refresh
      Call Totales
      'Call ConfigGrid
      
    Case 7
      Unload Me
  End Select
  
End Sub

Function ValidarGrabacion() As Integer
Dim rsaux As New ADODB.Recordset
Dim xrendicion As String
ValidarGrabacion = 0
If rsdetat.RecordCount <= 0 Then
     MsgBox "Falta añadir el Detalle a la Ventana del Browse", vbInformation, Caption
     Exit Function
End If
If VGParametros.sistemamultiempresas Then
   If Ctr_Ayuempresa.xclave = "" Then
      MsgBox "Debe ingresar codigo de empresa ", vbInformation
      Exit Function
   End If
End If
Set VGvardllgen = New dllgeneral.dll_general
If VGParametros.sistemactrlgastos Then
   If (CtrAyu_gastos.Enabled And CtrAyu_gastos.Visible) And (Trim(CtrAyu_gastos.xclave) = "" Or Trim(CtrAyu_gastos.xclave) = "00") Then
      MsgBox "Debe ingresar la Cuenta de gastos ", vbInformation
      CtrAyu_gastos.SetFocus
      Exit Function
    End If
End If
If (CtrAyu_Ccosto.Enabled And CtrAyu_Ccosto.Visible) And (Trim(CtrAyu_Ccosto.xclave) = "" Or Trim(CtrAyu_Ccosto.xclave) = "00") Then
    MsgBox "Debe ingresar El cento de costo", vbInformation
    CtrAyu_Ccosto.SetFocus
    Exit Function
End If
  If Ctr_Ayutransf.Visible = True And Label5(0) > saldodocxrendir Then
       MsgBox "Monto del documento ingresado excede al saldo de doc. a rendir, que es  -- > " & saldodocxrendir & " , verifique ", vbInformation
      Ctr_Ayutransf.SetFocus
      Exit Function
    End If

' If VGParametros.controlaestadosrendicion Then
'      If controlarendicion Then
'          SQL = "select numero=max(rendicionnumero) from te_rendiciones where oficinacodigo='" & VGoficina & "'"
'          SQL = SQL & " and codigocaja='" & Ctr_AyudaCaja.xclave & "'"
'          Set rsaux = VGCNx.Execute(SQL)
'          xrendicion = rsaux!numero
'          SQL = " select rendicionfecha,rendicionsaldofinal=rendicionsaldoinicial+rendicioningresos-rendicionegresos + isnull(saldoacumuladoxrendir,0)"
'          SQL = SQL & " from te_rendiciones where oficinacodigo='" & VGoficina & "'"
'          SQL = SQL & " and codigocaja='" & Ctr_AyudaCaja.xclave & "' and rendicionnumero='" & xrendicion & "'"
'          Set rsaux = VGCNx.Execute(SQL)
'          If numero(rsaux!rendicionsaldofinal) - Label5(0) < 0 Then
'             MsgBox " Monto de cancelacion Origina Saldos Negativos, Monto permitido es de  -- > " & rsaux!saldofinal, vbInformation
'             Ctr_AyudaCaja.SetFocus
'            Exit Function
'          End If
'       End If
'   End If
If (Ctr_AyuAnalitico.Enabled And Ctr_AyuAnalitico.Visible) And (Trim(Ctr_AyuAnalitico.xclave) = "" Or Trim(Ctr_AyuAnalitico.xclave) = "00") Then
      MsgBox "Debe ingresar Codigo de Analitico ", vbInformation
      Ctr_AyuAnalitico.SetFocus
      Exit Function
End If

ValidarGrabacion = 1
End Function

Private Sub Combo1_Click()
  Dim rs As New ADODB.Recordset
  
  Set rs = VGCNx.Execute("select * from te_parametroempresa where empresacodigo='" & VGCodEmpresa & "'")
  If rs.RecordCount > 0 Then
    If adll.ComboDato(Combo1.Text) = "I" Then
        Text1(0) = Right("0000000000" & Trim(CStr(CDbl(IIf(IsNull(rs!empresanumeingreso) Or Len(Trim(rs!empresanumeingreso)) = 0, 1, rs!empresanumeingreso + 1)))), 6)
    ElseIf adll.ComboDato(Combo1.Text) = "E" Then
        Text1(0) = Right("0000000000" & Trim(CStr(CDbl(IIf(IsNull(rs!empresanumegreso) Or Len(Trim(rs!empresanumegreso)) = 0, 1, rs!empresanumegreso + 1)))), 6)
    End If
  End If
  rs.Close
  Set rs = Nothing
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    Call Seguir(Combo1, KeyAscii)
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
    Call Seguir(Combo2, KeyAscii)
End Sub

Private Sub Ctr_AyudaCaja_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Dim rsaux As New ADODB.Recordset
Dim xrendicion As String
If ColecCampos!cajarendiciones Then
   SQL = "select numero=max(rendicionnumero) from te_rendiciones where oficinacodigo='" & VGoficina & "'"
   SQL = SQL & " and codigocaja='" & Ctr_AyudaCaja.xclave & "'"
   Set rsaux = VGCNx.Execute(SQL)
   xrendicion = ESNULO(rsaux!numero, "")
   SQL = " select rendicionfecha from te_rendiciones where oficinacodigo='" & VGoficina & "'"
   SQL = SQL & " and codigocaja='" & Ctr_AyudaCaja.xclave & "' and rendicionnumero='" & xrendicion & "'"
   Set rsaux = VGCNx.Execute(SQL)
   If xrendicion > 0 Then
      fecharendicion = rsaux!rendicionfecha - VGParametros.diasatrazorendicion
      controlarendicion = ColecCampos!cajarendiciones
    Else
     controlarendicion = False
    End If
End If
If m_docxrendir = 1 Or m_fondofijo = 1 Then
   Ctr_Ayutransf.Visible = True
   LeReferencia.Visible = True
   SQL = " isnull(estadodocxrendir,0)<2 and cajacodigo='" & Ctr_AyudaCaja.xclave & "' "
   SQL = SQL & " and cabrec_transferenciaautomatico=1 "
   Ctr_Ayutransf.filtro = SQL

 Else
   Ctr_Ayutransf.Visible = False
   LeReferencia.Visible = False
End If
End Sub

Private Sub Ctr_Ayuempresa_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
If VGParametros.sistemamultiempresas Then
  CtrAyu_Ccosto.filtro = "empresacodigo='" & Ctr_Ayuempresa.xclave & "' and centrocostotipo=" & VGnumnivcos & " and centrocostocodigo<>'00' "
End If
End Sub

Private Sub Ctr_Ayutransf_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
saldodocxrendir = ColecCampos("saldodocxrendir")
clientecodigo = ColecCampos("clientecodigo")
End Sub

Private Sub CtrAyu_gastos_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    Dim SQL As String
    If ColecCampos("gastosctrlcostos") Then
        CtrAyu_Ccosto.Visible = True
        lbccosto.Visible = True
        CtrAyu_Ccosto.xclave = "": CtrAyu_Ccosto.xnombre = ""
    '    Cuentacodigo = ColecCampos("cuentacodigo")
      Else
        CtrAyu_Ccosto.Visible = False
        lbccosto.Visible = False
        CtrAyu_Ccosto.xclave = "00"
    End If
    If ColecCampos("tipoanaliticocodigo") <> "00" Then
       Ctr_AyuAnalitico.xclave = "": Ctr_AyuAnalitico.xnombre = ""
       Ctr_AyuAnalitico.filtro = " tipoanaliticocodigo='" & VGParamSistem.tipoanaliticocodigo & "' and isnull(proyectocierre,0)=0"
       Ctr_AyuAnalitico.Visible = True
       Lblanalitico.Visible = True
     Else
       Ctr_AyuAnalitico.Visible = False
       Lblanalitico.Visible = False
    End If

End Sub

Private Sub Form_Load()
   MostrarForm Me, "C"
   Combo1.Clear
   Combo1.AddItem "I- INGRESOS"
   Combo1.AddItem "E- EGRESOS"
   Combo1.ListIndex = 0
   
   'Call Ctr_Ayuda2.Conexion(VGcnx)
    
   Text1(0).Enabled = False
   Call adll.llenacombo(Combo2, "select monedacodigo,monedadescripcion from gr_moneda", VGCNx)
   Combo2.ListIndex = 0
   
   Frame4.Enabled = False
   
   MBox1 = Format(VGParamSistem.fechatrabajo, "dd/mm/yyyy")
'   MBox1.Text = VGParamSistem.fechatrabajo
   Text1(3) = DatoTipoCambio(VGcnxCT, MBox1.Text)
   Call cargar_grilla
   Call ConfigGrid
   Call CtrAyu_Ccosto.conexion(VGcnxCT)
   CtrAyu_Ccosto.filtro = "centrocostotipo=" & VGnumnivcos & " and centrocostocodigo<>'00' "
   Call CtrAyu_gastos.conexion(VGCNx): CtrAyu_gastos.filtro = "(gastosnivel=" & VGnumnivgas & " and gastoscodigo <>'00') "
   Call Ctr_AyuAnalitico.conexion(VGCNx)

   SQL = " isnull(CajaCuentaxRendir,0)=" & m_docxrendir & " and isnull(Cajafondofijo,0)=" & m_fondofijo
   If VGParametros.listacajas <> "" Then SQL = SQL & " and cajacodigo in (" & VGParametros.listacajas & ")"
   Call Ctr_AyudaCaja.conexion(VGCNx)
   Ctr_AyudaCaja.filtro = SQL
   Call Ctr_Ayutransf.conexion(VGCNx)
   
   Call Ctr_Ayuempresa.conexion(VGCNx)
   If VGParametros.sistemamultiempresas Then
      Ctr_Ayuempresa.Visible = True
    Else
      Ctr_Ayuempresa.xclave = "01"
      Ctr_Ayuempresa.Visible = False
      Lblempresa.Visible = False
   End If
   If VGParametros.sistemactrlgastos Then
      CtrAyu_gastos.Visible = True
    Else
      CtrAyu_gastos.xclave = "00"
   End If
   cmdBotones(4).Enabled = False
   
End Sub

Private Sub MBox1_KeyDown(KeyCode As Integer, Shift As Integer)
  Call Seguir(MBox1, KeyCode)
End Sub

Private Sub MBox1_LostFocus()
 If IsDate(MBox1.Text) Then Text1(3).Text = DatoTipoCambio(VGcnxCT, MBox1.Text)
End Sub

Private Sub MBox2_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      SendKeys "{tab}"
   End If
End Sub

Public Sub grabacion()
   Dim rb As New ADODB.Recordset
    If Len(Trim(Ctr_AyudaCaja.xclave)) = 0 Then
       Set rb = VGCNx.Execute("select * from gr_banco where bancocodigo='" & Text2(5) & "'")
       If rb.RecordCount = 0 Then
         MsgBox "No existe el banco indicado .... Verifique!!", vbInformation, MsgTitle
         rb.Close
         Set rb = Nothing
         Text2(5).SetFocus
         Exit Sub
       End If
       rb.Close
       Set rb = Nothing
    End If

    Set rb = VGCNx.Execute("select * from gr_moneda where monedacodigo='" & Text2(7) & "'")
    If rb.RecordCount = 0 Then
      MsgBox "No existe moneda .... Verifique!!", vbInformation, MsgTitle
      rb.Close
      Set rb = Nothing
      Text2(7).SetFocus
      Exit Sub
    End If
    rb.Close
    Set rb = Nothing
    Set VGvardllgen = New dllgeneral.dll_general
   If VGParametros.sistemactrlgastos Then
      If (CtrAyu_gastos.Enabled And CtrAyu_gastos.Visible) And (Trim(CtrAyu_gastos.xclave) = "" Or Trim(CtrAyu_gastos.xclave) = "00") Then
         MsgBox "Debe ingresar la Cuenta de gastos ", vbInformation
         CtrAyu_gastos.SetFocus
         Exit Sub
      End If
  End If
  If (CtrAyu_Ccosto.Enabled And CtrAyu_Ccosto.Visible) And (Trim(CtrAyu_Ccosto.xclave) = "" Or Trim(CtrAyu_Ccosto.xclave) = "00") Then
      MsgBox "Debe ingresar El cento de costo", vbInformation
      CtrAyu_Ccosto.SetFocus
      Exit Sub
  End If


    Text2(8) = numero(Text2(8))
    
    rsdetat.AddNew
    rsdetat.Fields(0) = Escadena(Text2(0))
    rsdetat.Fields(1) = Escadena(Text2(1))
    'rsdetat.Fields(2) = ""
    rsdetat.Fields(4) = Escadena(Text2(4))
    rsdetat.Fields(5) = Escadena(Text2(5))
    rsdetat.Fields(6) = Escadena(Text2(6))
    rsdetat.Fields(7) = Escadena(Text2(7))
    rsdetat.Fields(8) = numero(Text2(8))
    rsdetat.Fields(9) = Format(MBox2, "dd/mm/yyyy")
    rsdetat.Fields(10) = Escadena(Text2(9).Text)
    rsdetat.Fields(11) = Escadena(Text2(10).Text)
    rsdetat.Fields(12) = IIf(CtrAyu_gastos.Visible, CtrAyu_gastos.xclave, "00")
    rsdetat.Fields(13) = IIf(Ctr_AyuAnalitico.Visible, Ctr_AyuAnalitico.xclave, "00")
    rsdetat.Fields(14) = IIf(CtrAyu_Ccosto.Visible, CtrAyu_Ccosto.xclave, "00")
    rsdetat.Update
    TDBGrid1.Refresh
    Call ConfigGrid
    
    Call Limpiartexto(Text2, 0, 1)
    Call Limpiartexto(Text2, 4, 8)
    MBox2 = Format(VGParamSistem.fechatrabajo, "dd/mm/yyyy")
    Text2(0) = CStr(CDbl(rsdetat.Fields(0)) + 1)
    Call Totales
    Text2(1).SetFocus
End Sub

Private Sub MBox2_LostFocus()
If MBox2.Text <> MBox1.Text Then
   MsgBox (" Fecha diferente al recibo, verifique  ")
End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   Call adll.Enfoquetexto(Text1(Index))
   If Index = 1 Then
      Ctr_Ayutransf.Visible = False
      LeReferencia.Visible = False
   End If
MBox2.Text = MBox1.Text
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
  Dim rb As New ADODB.Recordset
  On Error Resume Next
  
  If KeyAscii = 13 Then
     If Index = 1 Then
         Set rb = VGCNx.Execute("select * from te_operaciongeneral where operacioncodigo='" & Escadena(Text1(1)) & "' and operacioncontrolaclienteprov='" & IIf(adll.ComboDato(Combo1) = "I", "X", "X") & "'")
         If rb.RecordCount > 0 Then
            Text1(1).Text = Escadena(rb!operacioncodigo)
            Label2(0).Caption = Escadena(rb!operaciondescripcion)
            If Escadena(rb!operacionvalidacajabancos) = "B" Then
                Ctr_AyudaCaja.Enabled = True
                cayuda(1).Enabled = True
                Text2(9).Enabled = True
                Ctr_AyudaCaja.xclave = Empty: Label2(1).Caption = ""
                Ctr_AyudaCaja.Enabled = False
                cayuda(1).Enabled = False
                rb.Close
                Set rb = Nothing
                Combo2.SetFocus
                Ctr_AyudaCaja.Visible = False
                Label1(5).Visible = False
                Exit Sub
            Else
                Ctr_AyudaCaja.Visible = True
                Label1(5).Visible = True
                Ctr_AyudaCaja.Enabled = True
                cayuda(1).Enabled = True
                Text2(6).Enabled = False
                Text2(9).Text = Empty
                Text2(9).Enabled = False
                cayuda(5).Enabled = False
                Text2(5).Enabled = False
                
                Ctr_AyudaCaja.SetFocus
                Set rb = Nothing
                Exit Sub
            End If
         Else
            Ctr_AyudaCaja.Enabled = True
            cayuda(1).Enabled = True
            Text1(1).Text = Empty: Label2(0).Caption = Empty: Ctr_AyudaCaja.xclave = Empty: Label2(1).Caption = Empty
         End If
         rb.Close
         Set rb = Nothing
     ElseIf Index = 2 Then
        Set rb = VGCNx.Execute("select * from te_codigocaja where cajacodigo='" & Ctr_AyudaCaja.xclave & "'")
        If rb.RecordCount > 0 Then
            Ctr_AyudaCaja.xclave = Escadena(rb!cajacodigo)
            Label2(1) = Escadena(rb!cajadescripcion)
        Else
            Ctr_AyudaCaja.xclave = ""
            Label2(1).Caption = ""
        End If
        rb.Close
        Set rb = Nothing
     ElseIf Index = 3 Then
        Call Totales
        
        If Not IsDate(MBox1) Then
            MsgBox "Fecha no valida...Verifique!!", vbInformation, MsgTitle
            MBox1.SetFocus
            Exit Sub
        End If
        If Len(Trim(Text1(1))) = 0 Then
            MsgBox "Falta Ingresar Tipo de Operacion...Verifique!!", vbInformation, MsgTitle
            Text1(1).SetFocus
            Exit Sub
        End If
        If Len(Trim(Text1(3))) = 0 Then
            MsgBox "Falta Ingresar Tipo de Cambio..Verifique!!", vbInformation, MsgTitle
            Text1(3).SetFocus
            Exit Sub
        End If
        
        Frame4.Enabled = True
        Call Limpiartexto(Text2, 0, 8)
        MBox2 = Format(MBox1.Text, "dd/mm/yyyy")
        If rsdetat.RecordCount = 0 Then
          Text2(0) = 1
        Else
          rsdetat.MoveLast
          Text2(0) = CStr(CDbl(rsdetat.Fields(0)) + 1)
        End If
        Frame2.Enabled = False
        If Len(Ctr_AyudaCaja.xclave) > 0 Then
           Text2(5).Enabled = False
           Text2(9).Enabled = False
           cayuda(5).Enabled = False
           cayuda(7).Enabled = False
        End If
        Text2(1).SetFocus
        Exit Sub
     End If
     Call Seguir(Text1(Index), 13)
  End If
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
 Dim rb As New ADODB.Recordset
 
  If KeyAscii = 13 Then
    Text2(Index) = UCase(Text2(Index))
    If Index = 1 Then
       Set rb = VGCNx.Execute("select * from te_conceptocaja")
       If rb.RecordCount = 0 Then
         MsgBox "No existe Conceptos de Tesorería...Verifique!!", vbInformation, MsgTitle
         rb.Close
         Set rb = Nothing
         Exit Sub
       End If
       rb.Close
       Set rb = Nothing
    ElseIf Index = 4 Then   'Tipo de cancelacion
       Set rb = VGCNx.Execute("select * from cp_tipodocumento where tdocumentotipo='A' and tdocumentoingcobra='1' and tdocumentocodigo='" & Text2(4).Text & "'")
       If rb.RecordCount = 0 Then
         MsgBox "No existe tipo de documento...Verifique!!", vbInformation, MsgTitle
         rb.Close
         Set rb = Nothing
         Exit Sub
       End If
       rb.Close
       Set rb = Nothing
    ElseIf Index = 5 Then
       Set rb = VGCNx.Execute("select * from gr_banco where bancocodigo='" & Text2(5) & "'")
       If rb.RecordCount = 0 Then
         MsgBox "No existe el banco indicado .... Verifique!!", vbInformation, MsgTitle
         rb.Close
         Set rb = Nothing
         Exit Sub
       End If
       rb.Close
       Set rb = Nothing
    ElseIf Index = 7 Then
       Set rb = VGCNx.Execute("select * from gr_moneda where monedacodigo='" & Text2(7) & "'")
       If rb.RecordCount = 0 Then
         MsgBox "No existe moneda .... Verifique!!", vbInformation, MsgTitle
         rb.Close
         Set rb = Nothing
         Exit Sub
       End If
       rb.Close
       Set rb = Nothing
    ElseIf Index = 8 Then
       Text2(8) = numero(Text2(8))
       If Text2(8) < 0 Then
        MsgBox "El importe debe ser mayor que cero. Se corregirá el importe", vbInformation, "Aviso"
        Text2(8) = numero(Text2(8) * (-1))
       End If
    ElseIf Index = 9 Then
       Set rb = VGCNx.Execute("select * from te_cuentabancos inner join gr_banco on te_cuentabancos.cbanco_codigo=gr_banco.bancocodigo where gr_banco.bancocodigo='" & Escadena(Text2(5)) & "' and te_cuentabancos.monedacodigo='" & Text2(7) & "' and te_cuentabancos.cbanco_numero='" & Trim(Text2(9)) & "'")
       If rb.RecordCount = 0 Then
         MsgBox "No existe la cuenta corriente del banco indicado .... Verifique!!", vbInformation, MsgTitle
         rb.Close
         Set rb = Nothing
         Exit Sub
       End If
       rb.Close
       Set rb = Nothing
    ElseIf Index = 10 Then
       Call grabacion
       Exit Sub
    End If
    Call Seguir(Text2(Index), KeyAscii)
  End If
End Sub

Public Function Totales()
Dim sumas, sumad As Double
    Dim Tsumas, Tsumad As Double
    
    sumas = 0: sumad = 0: Tsumas = 0: Tsumad = 0
    If rsdetat.RecordCount > 0 Then
        rsdetat.MoveFirst
        Do Until rsdetat.EOF
           If rsdetat.Fields(7) = g_tipodolar Then
               sumad = sumad + CDbl(rsdetat.Fields(8))
           ElseIf rsdetat.Fields(7) = g_tiposol Then
               sumas = sumas + CDbl(rsdetat.Fields(8))
           End If
           rsdetat.MoveNext
        Loop
    End If
    If Text1(3) = 0 Or Len(Trim(Text1(3))) = 0 Then Text1(3) = numero(1)
    Tsumad = sumad + (sumas / CDbl(Text1(3)))
    Tsumas = sumad * CDbl(Text1(3)) + sumas

    Label5(0) = numero(Tsumas): Label5(1) = numero(Tsumad)
        
End Function
Private Sub Actualizasaldorendicion()
Dim rsaux As New ADODB.Recordset
Dim xrendicion As String
Dim xsaldo As Double
SQL = "select numero=max(rendicionnumero) from te_rendiciones where oficinacodigo='" & VGoficina & "'"
SQL = SQL & " and codigocaja='" & Ctr_AyudaCaja.xclave & "'"
Set rsaux = VGCNx.Execute(SQL)
xrendicion = rsaux!numero
If Ctr_AyudaCaja.xclave = "02" Then
   xsaldo = Label5(0) * Text1(3).Text
 Else
   xsaldo = Label5(0)
End If
SQL = " update te_rendiciones set saldoacumuladoxrendir=isnull(saldoacumuladoxrendir,0)-" & xsaldo
SQL = SQL & " where oficinacodigo='" & VGoficina & "'"
SQL = SQL & " and codigocaja='" & Ctr_AyudaCaja.xclave & "' and rendicionnumero='" & xrendicion & "'"
Set rsaux = VGCNx.Execute(SQL)
End Sub


