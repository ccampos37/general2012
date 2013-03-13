VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Begin VB.Form frmRequerimientosPedidos 
   Caption         =   "Generacion de requerimientos"
   ClientHeight    =   6840
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   11505
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Fradatos 
      Height          =   2145
      Left            =   45
      TabIndex        =   41
      Top             =   1005
      Width           =   11100
      Begin VB.TextBox txtEntE 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   10365
         MaxLength       =   50
         TabIndex        =   44
         Top             =   975
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtObs 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1185
         MaxLength       =   80
         TabIndex        =   43
         Top             =   1725
         Width           =   7800
      End
      Begin VB.CheckBox ChkTraslado 
         Caption         =   "Traslado Fisico"
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
         Left            =   3120
         TabIndex        =   42
         Top             =   1350
         Width           =   1635
      End
      Begin MSComCtl2.DTPicker txtEmi 
         Height          =   315
         Left            =   1185
         TabIndex        =   45
         Top             =   570
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Format          =   21692417
         CurrentDate     =   37015
      End
      Begin MSComCtl2.DTPicker txtEnt 
         Height          =   315
         Left            =   4020
         TabIndex        =   46
         Top             =   570
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         _Version        =   393216
         Format          =   21692417
         CurrentDate     =   37015
      End
      Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_Proveedor 
         Height          =   315
         Left            =   1185
         TabIndex        =   47
         Top             =   195
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   556
         XcodMaxLongitud =   11
         xcodwith        =   1100
         NomTabla        =   "cp_proveedor"
         TituloAyuda     =   "Busqueda de Proveedor"
         ListaCampos     =   "clientecodigo(1),clienterazonsocial(1),clienteruc(1)"
         XcodCampo       =   "clientecodigo"
         XListCampo      =   "clienterazonsocial"
         ListaCamposDescrip=   "Codigo,Descripcion,Ruc"
         ListaCamposText =   "clientecodigo,clienterazonsocial,clienteruc"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_almacenO 
         Height          =   315
         Left            =   1185
         TabIndex        =   48
         Top             =   960
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   556
         XcodMaxLongitud =   0
         xcodwith        =   250
         NomTabla        =   "tabalm"
         ListaCampos     =   "taalma(1),tadescri(1)"
         XcodCampo       =   "taalma"
         XListCampo      =   "tadescri"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "taalma,tadescri"
         Requerido       =   0   'False
      End
      Begin TextFer.TxFer lblRuc 
         Height          =   315
         Left            =   6690
         TabIndex        =   49
         Top             =   195
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
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
      Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_solicitante 
         Height          =   315
         Left            =   6690
         TabIndex        =   50
         Top             =   585
         Width           =   3750
         _ExtentX        =   6615
         _ExtentY        =   556
         XcodMaxLongitud =   3
         xcodwith        =   300
         NomTabla        =   "co_solicitantes"
         TituloAyuda     =   "Busqueda de Solicitante"
         ListaCampos     =   "solicitantecodigo(1),solicitantenombre(1)"
         XcodCampo       =   "solicitantecodigo"
         XListCampo      =   "solicitantenombre"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "solicitantecodigo,solicitantenombre"
      End
      Begin TextFer.TxFer TxFcot 
         Height          =   315
         Left            =   1170
         TabIndex        =   51
         Top             =   1320
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
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
         ColorIlumina    =   8454143
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         MarcarTextoAlEnfoque=   -1  'True
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_almacenD 
         Height          =   315
         Left            =   5865
         TabIndex        =   52
         Top             =   960
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   556
         XcodMaxLongitud =   0
         xcodwith        =   250
         NomTabla        =   "tabalm"
         ListaCampos     =   "taalma(1),tadescri(1)"
         XcodCampo       =   "taalma"
         XListCampo      =   "tadescri"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "taalma,tadescri"
         Requerido       =   0   'False
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
         Left            =   8775
         TabIndex        =   53
         Top             =   240
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   5
         Format          =   "HH:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Solicitante :"
         Height          =   195
         Left            =   5730
         TabIndex        =   64
         Top             =   630
         Width           =   825
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Entregar en   :"
         Height          =   195
         Left            =   9300
         TabIndex        =   63
         Top             =   990
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Emisión :"
         Height          =   195
         Left            =   90
         TabIndex        =   62
         Top             =   600
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor :"
         Height          =   195
         Left            =   90
         TabIndex        =   61
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "R.U.C.  :"
         Height          =   195
         Left            =   6000
         TabIndex        =   60
         Top             =   210
         Width           =   555
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "F. Entrega :"
         Height          =   195
         Left            =   2970
         TabIndex        =   59
         Top             =   630
         Width           =   825
      End
      Begin VB.Label Label12 
         Caption         =   "Observación :"
         Height          =   255
         Left            =   90
         TabIndex        =   58
         Top             =   1740
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cotizacion :"
         Height          =   195
         Left            =   90
         TabIndex        =   57
         Top             =   1365
         Width           =   825
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Almac Salida :"
         Height          =   195
         Left            =   90
         TabIndex        =   56
         Top             =   1020
         Width           =   1005
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Almac. Ingreso:"
         Height          =   195
         Left            =   4710
         TabIndex        =   55
         Top             =   1020
         Width           =   1095
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Hora :"
         Height          =   195
         Left            =   8280
         TabIndex        =   54
         Top             =   225
         Width           =   435
      End
   End
   Begin VB.Frame fraTotales 
      Height          =   975
      Left            =   135
      TabIndex        =   30
      Top             =   4095
      Visible         =   0   'False
      Width           =   9708
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Importe      :"
         Height          =   195
         Left            =   720
         TabIndex        =   40
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Descuento :"
         Height          =   195
         Left            =   720
         TabIndex        =   39
         Top             =   600
         Width           =   870
      End
      Begin VB.Label lblImp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "892,760.00"
         Height          =   285
         Left            =   1680
         TabIndex        =   38
         Top             =   240
         Width           =   1110
      End
      Begin VB.Label lblDes 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "892,760.00"
         Height          =   285
         Left            =   1680
         TabIndex        =   37
         Top             =   600
         Width           =   1110
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Total  :"
         Height          =   195
         Left            =   3600
         TabIndex        =   36
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblTot 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "892,760.00"
         Height          =   285
         Left            =   4200
         TabIndex        =   35
         Top             =   240
         Width           =   1230
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "I.G.V.   :"
         Height          =   195
         Left            =   6360
         TabIndex        =   34
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Compra :"
         Height          =   195
         Left            =   6360
         TabIndex        =   33
         Top             =   600
         Width           =   630
      End
      Begin VB.Label lblIgv 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "892,760.00"
         Height          =   285
         Left            =   7080
         TabIndex        =   32
         Top             =   240
         Width           =   1110
      End
      Begin VB.Label lblCom 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "892,760.00"
         Height          =   285
         Left            =   7080
         TabIndex        =   31
         Top             =   600
         Width           =   1110
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1020
      Left            =   45
      TabIndex        =   23
      Top             =   75
      Width           =   11100
      Begin ctrlayuda_f.Ctr_Ayuda Ctrayu_tipoorden 
         Height          =   330
         Left            =   5670
         TabIndex        =   24
         Top             =   180
         Width           =   3270
         _ExtentX        =   5768
         _ExtentY        =   582
         XcodMaxLongitud =   11
         xcodwith        =   200
         NomTabla        =   "co_tipodeorden"
         TituloAyuda     =   "Busqueda de Tipo de Orden"
         ListaCampos     =   "tipoordencodigo(1),tipoordendescripcion(1),tipoordennumeracion(2)"
         XcodCampo       =   "tipoordencodigo"
         XListCampo      =   "tipoordendescripcion"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "tipoordencodigo,tipoordendescripcion,tipoordennumeracion"
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Estado  :"
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
         Left            =   3060
         TabIndex        =   29
         Top             =   600
         Width           =   780
      End
      Begin VB.Label lblEst 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3900
         TabIndex        =   28
         Top             =   540
         Width           =   1650
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Número  :"
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
         Left            =   180
         TabIndex        =   27
         Top             =   600
         Width           =   840
      End
      Begin VB.Label lblNum 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1110
         TabIndex        =   26
         Top             =   540
         Width           =   1560
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Orden :"
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
         Left            =   4500
         TabIndex        =   25
         Top             =   210
         Width           =   1080
      End
   End
   Begin VB.CommandButton CmdSalir2 
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
      Height          =   1065
      Left            =   8025
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5385
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton cmdGra 
      Caption         =   "&Grabar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5385
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton cmdEdi2 
      Caption         =   "&Editar"
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
      Height          =   1065
      Left            =   3735
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5385
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton cmdEli2 
      Caption         =   "&Quitar"
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
      Height          =   1065
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5385
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton cmdNue2 
      Caption         =   "&Agregar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   2310
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5385
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   585
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4485
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame FrameInicio 
      Height          =   6645
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   11355
      Begin VB.CommandButton cmdImp 
         Caption         =   "&Imprimir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   5400
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.CommandButton cmdEdi 
         Caption         =   "&Editar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   2655
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   5400
         Width           =   1170
      End
      Begin VB.CommandButton CmdEli 
         Caption         =   "&Anular"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   5400
         Width           =   1170
      End
      Begin VB.CommandButton cmdNue 
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
         Height          =   1065
         Left            =   1305
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   5400
         Width           =   1170
      End
      Begin VB.CommandButton CmdSalir 
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
         Height          =   1065
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   5400
         Width           =   1170
      End
      Begin VB.CommandButton Cmdatendido 
         Cancel          =   -1  'True
         Caption         =   "&Atendido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   5400
         Width           =   1170
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   5055
         Left            =   120
         TabIndex        =   16
         Top             =   150
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   8916
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "T.de Orden"
         Columns(0).DataField=   "tipoordencodigo"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Numero"
         Columns(1).DataField=   "oc_cnumord"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Razon Social"
         Columns(2).DataField=   "oc_crazsoc"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "F.Emision"
         Columns(3).DataField=   "oc_dfecdoc"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Estado"
         Columns(4).DataField=   "estado"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "estadodescripcion"
         Columns(5).DataField=   "estadoocdescripcion"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   16
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Solicitante"
         Columns(6).DataField=   "solicitantenombre"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Almacen Origen"
         Columns(7).DataField=   "almacenorigen"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Almacen Destino"
         Columns(8).DataField=   "almacendestino"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   9
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   873
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   15790320
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=9"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1667"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1588"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2011"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1931"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=3916"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=3836"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=1667"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=1588"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=1349"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=1270"
         Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(21)=   "Column(5).Width=3545"
         Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=3466"
         Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(25)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(29)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(33)=   "Column(8).Width=4657"
         Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=4577"
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
         Caption         =   "Relacion de requerimientos en estado de atencion"
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   15790320
         RowDividerColor =   15790320
         RowSubDividerColor=   15790320
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
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=70,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=67,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=68,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=69,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=62,.parent=13"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=59,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=60,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=61,.parent=17"
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
         _StyleDefs(90)  =   ":id=41,.parent=34"
         _StyleDefs(91)  =   "Named:id=42:FilterBar"
         _StyleDefs(92)  =   ":id=42,.parent=33"
      End
   End
   Begin VB.Frame FrameAtendido 
      Height          =   2055
      Left            =   3075
      TabIndex        =   0
      Top             =   2355
      Visible         =   0   'False
      Width           =   5295
      Begin VB.TextBox Textorden 
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Textnumorden 
         Height          =   375
         Left            =   1440
         TabIndex        =   4
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Textrazonsocial 
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   1440
         Width           =   3375
      End
      Begin VB.CommandButton CmdAceptaAtendido 
         Caption         =   "Aceptar"
         Height          =   495
         Left            =   3600
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CmdSaliirAtendido 
         Caption         =   "Salir"
         Height          =   495
         Left            =   3600
         TabIndex        =   1
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "Tipo de Orden"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label19 
         Caption         =   "Numero"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label22 
         Caption         =   "Razon Social"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   1095
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex1 
      Height          =   2145
      Left            =   45
      TabIndex        =   17
      Top             =   3135
      Visible         =   0   'False
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   3784
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      RowHeightMin    =   240
      BackColorSel    =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   "^Código|Fab|Descripción|Undadi|Cantidad|C.Costo|Analitico|Comentario 1|Comentario 2|Familia"
      BandDisplay     =   1
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Bindings        =   "frmRequerimientosPedidos.frx":0000
      Left            =   585
      Top             =   3975
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmRequerimientosPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Colex As New Collection
Dim adodc1 As ADODB.Recordset
Public VGvardllgen As dllgeneral.dll_general
Dim cSql1 As String
Dim nT As Integer       'Ingreso,Modificación,Ficha Tecnica
Dim cCod As String
Dim nTra As Integer
Dim Mensaje As String
Dim unum As String
Sub OculObj02(ntipo As Boolean)
    cmdGra.Visible = ntipo
    CmdSalir2.Visible = ntipo
End Sub
Sub OculObj03(ntipo As Boolean)
    Fradatos.Visible = ntipo
    fraTotales.Visible = ntipo
End Sub
Sub OculObj04(ntipo As Boolean)
    cmdNue.Visible = ntipo
    cmdEdi.Visible = ntipo
    CmdEli.Visible = ntipo
    cmdImp.Visible = ntipo
    CmdSalir.Visible = ntipo
End Sub
Sub OculObj06(ntipo As Boolean)
    TDBGrid1.Visible = ntipo
End Sub
Sub Abre_Tabla_OCs()
Dim SQL As String
Set VGvardllgen = New dllgeneral.dll_general
Set adodc1 = New ADODB.Recordset

SQL = "SELECT estado=case when a.oc_estadoorden='1' then 'Anulado' else '' end,* FROM co_cabordcompra a "
SQL = SQL & " inner join co_estadorequerimiento b on a.estadooccodigo= b.estadooccodigo"
SQL = SQL & " inner join co_tipodeorden c on a.tipoordencodigo=c.tipoordencodigo "
SQL = SQL & " inner join co_solicitantes d on oc_csolict=solicitantecodigo where estadoocatendido<>1  and empresacodigo='" & VGparametros.empresacodigo & "' and puntovtacodigo='" & VGparametros.puntovta & "' "
SQL = SQL & " and b.nivelrequerimiento<= 4 and flagrequerimientosPedidos=1 and estadoordencodigo<>2 and a.oc_estadoorden<>'1' ORDER BY oc_dfecdoc "
Set adodc1 = VGCNx.Execute(SQL)
' antes iba flagrequerimientosCODIGO
Set TDBGrid1.DataSource = adodc1
    
End Sub

Private Sub CmdAceptaAtendido_Click()
Dim rrsql As New ADODB.Recordset
SQL = " update co_cabordcompra  set estadooccodigo='" & VGparametros.Valorestadooccodigo & "'"
SQL = SQL & " where tipoordencodigo='" & adodc1!tipoordencodigo & "' and OC_CNUMORD='" & adodc1!oc_cnumord & "'"
Set rrsql = VGCNx.Execute(SQL)
TDBGrid1.Refresh
FrameAtendido.Visible = False
End Sub

 
Private Sub Cmdatendido_Click()
FrameAtendido.Visible = True
Textorden.text = adodc1!tipoordencodigo
Textnumorden.text = adodc1!oc_cnumord
Textrazonsocial.text = adodc1!oc_crazsoc
End Sub

Private Sub cmdEdi2_Click()
On Error GoTo Err
Load frmRequerimientosDetalle
    With frmRequerimientosDetalle
        .activado = False
        .CtrAyu_articulo.xclave = Flex1.TextMatrix(Flex1.Row, 0)
        .TxtOrdfab.text = Flex1.TextMatrix(Flex1.Row, 1)
        If .CtrAyu_articulo.xclave <> "" Then
           .CtrAyu_articulo.Ejecutar
        End If
        .lblUni = Flex1.TextMatrix(Flex1.Row, 3)
        .txtCan = Flex1.TextMatrix(Flex1.Row, 4)
        .txtCan.Enabled = True
'        .tipo = Flex1.TextMatrix(Flex1.Row, 14)
'        If Flex1.TextMatrix(Flex1.Row, 3) <> Flex1.TextMatrix(Flex1.Row, 5) Then
'            .txtURe = Flex1.TextMatrix(Flex1.Row, 5)
'            .txtRef = Flex1.TextMatrix(Flex1.Row, 6)
'        Else
'            .txtURe = ""
'            .txtRef = ""
'        End If
        .CtrAyu_Ccosto.xclave = Flex1.TextMatrix(Flex1.Row, 5)
        .CtrAyu_Ccosto.Ejecutar
        .CtrAyu_Analitico.xclave = Flex1.TextMatrix(Flex1.Row, 6): .CtrAyu_Analitico.Ejecutar
        .Txtco1.text = Flex1.TextMatrix(Flex1.Row, 7)
        .Txtco2.text = Flex1.TextMatrix(Flex1.Row, 8)
        .Ctr_AyuFamilia.xclave = Flex1.TextMatrix(Flex1.Row, 9)
        .activado = True
        .Show 1
        
        If Not .cancelado Then
            If .tipo = "S" Then
              .txtCan = 1
            End If
            Flex1.TextMatrix(Flex1.Row, 0) = .CtrAyu_articulo.xclave
            Flex1.TextMatrix(Flex1.Row, 1) = .TxtOrdfab.text
            Flex1.TextMatrix(Flex1.Row, 2) = .CtrAyu_articulo.xnombre
            Flex1.TextMatrix(Flex1.Row, 5) = .CtrAyu_Ccosto.xclave
            Flex1.TextMatrix(Flex1.Row, 6) = .CtrAyu_Analitico.xclave
            If .txtURe = "" Then
                Flex1.TextMatrix(Flex1.Row, 3) = .lblUni
                Flex1.TextMatrix(Flex1.Row, 4) = .txtCan
            Else
                Flex1.TextMatrix(Flex1.Row, 3) = .txtURe
                Flex1.TextMatrix(Flex1.Row, 4) = .txtRef
            End If
            Flex1.TextMatrix(Flex1.Row, 7) = .Txtco1.text
            Flex1.TextMatrix(Flex1.Row, 8) = .Txtco2.text
            Flex1.TextMatrix(Flex1.Row, 9) = .Ctr_AyuFamilia.xclave
            
            
        End If
        Flex1.SetFocus
        cmdNue2.SetFocus
    End With
 Exit Sub
Err:
    MsgBox Err.Description
    Exit Sub
    Resume
End Sub
Private Sub CmdEli_Click()
    On Error GoTo EliErr
    
    If adodc1("oc_estadoorden") = 1 Or ESNULO(adodc1("estadooccodigo"), 0) <> "0" Then
        Mensaje = "Imposible anular la Orden de compra en su estado actual"
        MsgBox Mensaje, vbCritical, "Mensaje"
        TDBGrid1.SetFocus
        Exit Sub
    End If

    Dim strsql As String
    Dim voc As String
    Dim tipo As String
    Mensaje = "¿Está seguro que desea anular la Orden de compra?"
    If MsgBox(Mensaje, vbQuestion + vbYesNo + vbDefaultButton2, "Mensaje") = vbYes Then
        voc = adodc1("oc_cnumord")
        tipo = adodc1("tipoordencodigo")
        
        nTra = 1
        VGCNx.BeginTrans
        
        strsql = "UPDATE co_detordcompra SET oc_estadoorden=1  WHERE oc_cnumord='" & voc & "'"
        strsql = strsql & " and tipoordencodigo='" & tipo & "'"
        
        VGCNx.Execute strsql
        
        strsql = "UPDATE co_cabordcompra SET oc_estadoorden=1 WHERE oc_cnumord='" & voc & "'"
        strsql = strsql & " and tipoordencodigo='" & tipo & "'"
        
        VGCNx.Execute strsql

        VGCNx.CommitTrans
        nTra = 0
        
        If nTra = 0 Then
  '          adodc1.Requery
        End If
    End If
    Abre_Tabla_OCs
    TDBGrid1.Refresh
    If adodc1.RecordCount > 0 Then
        TDBGrid1.SetFocus
    Else
        cmdNue.SetFocus
    End If
    Exit Sub

EliErr:
    MsgBox Err.Description
    If nTra = 1 Then VGCNx.RollbackTrans
End Sub
Private Sub CmdEli2_Click()
    If Tiene_Entregas Then
        Mensaje = "El artículo tiene cantidad entregada"
        MsgBox Mensaje, vbExclamation, "Advertencia"
    End If
    
    Mensaje = "¿Desea quitar el artículo seleccionado?"
    If MsgBox(Mensaje, vbQuestion + vbYesNo + vbDefaultButton2, "Mensaje") = vbYes Then
        If Flex1.Rows - 1 = 1 Then
            Dim i As Integer
            
            For i = 0 To 9
                Flex1.TextMatrix(1, i) = ""
            Next
        Else
            Flex1.RemoveItem Flex1.Row
        End If
        Estado_Items
    End If
End Sub
Private Sub cmdGra_Click()
Dim SQLc As String
Dim SQLd As String
Dim Rs2 As New ADODB.Recordset
Dim i As Integer
Dim vFactor As Single, vCantid As Single
Dim vPreuni As Single, vDscpor As Single
Dim vDescto As Single, vIgv As Single
Dim vIgvpor As Single, vPrenet As Single
Dim vTotven As Single, vTotnet As Single
Dim vURef As String, txtMon As String
Dim txtEst As String, txtTip As Integer
Dim txtPro As String, txtSol As String
Dim lblPro As String, txtFor As String
On Error GoTo GrabErr
    
    txtTip = 0
    txtFor = ""
    
    If Trim(Ctrayu_tipoorden.xclave) = "" Then
       Mensaje = "Debe ingresar Código de Tipo de Orden"
       MsgBox Mensaje, vbExclamation, "Mensaje"
       Ctrayu_tipoorden.SetFocus
       Exit Sub
    End If
    
    txtPro = Trim(CtrAyu_Proveedor.xclave)
    
    If Not IsNumeric(Right(Trim(TxtHor.text), 2)) Or Not IsNumeric(Left(Trim(TxtHor.text), 2)) Then
        MsgBox "Ingreso de hora debe ser numerico, verifique", vbCritical, "Sistema"
        TxtHor.SetFocus
        Exit Sub
    ElseIf (Right(Trim(TxtHor.text), 2)) > 59 Or (Left(Trim(TxtHor.text), 2)) > 24 Then
        MsgBox "Ingreso de hora incorrecta, verifique", vbCritical, "Sistema"
        TxtHor.SetFocus
        Exit Sub
    End If
        
    If txtEmi > txtEnt Then
       MsgBox "Fecha de emision no debe ser mayor a la fecha de entrega", vbExclamation, "Error"
       Exit Sub
       txtEmi.SetFocus
    End If
      
    txtEst = ""
    txtSol = Trim(CtrAyu_solicitante.xclave)
    If txtSol = "" Then
        Mensaje = "Debe ingresar Solicitante"
        MsgBox Mensaje, vbExclamation, "Mensaje"
        CtrAyu_solicitante.SetFocus
        Exit Sub
    End If
      
    If Len(Trim(Ctr_almacenO.xclave)) = 0 Then
       MsgBox "Falta ingresar almacen de origen", vbCritical, "Sistema"
       Ctr_almacenO.SetFocus
       Exit Sub
    End If
    
    If Len(Trim(Ctr_almacenD.xclave)) = 0 Then
       MsgBox "Falta ingresar almacen de destino", vbCritical, "Sistema"
       Ctr_almacenD.SetFocus
       Exit Sub
    End If
    
    '----------------------------------------------------
    If nT = 1 Then
        Mensaje = "¿Desea guardar la Orden ?"
    Else
        Mensaje = "¿Desea guardar los cambios realizados?"
    End If
    
    If MsgBox(Mensaje, vbQuestion + vbYesNo, "Mensaje") = vbYes Then
       nTra = 1
       VGCNx.BeginTrans
       unum = Format(Val(lblNum), "00000000000")

       If nT = 1 Then      'Ingreso
         'unum = Format(Devolver_Dato(1, , " & trim(ctrayu_tipoordencodigo) & ", "tipoordencodigo", False,
         '      "ctnnumero"), "00000000000")
         SQLc = "select tipoordennumeracion from co_tipodeorden where tipoordencodigo='" & Trim(Ctrayu_tipoorden.xclave) & "' "
         Set Rs2 = New ADODB.Recordset
         Rs2.Open SQLc, VGCNx, adOpenKeyset, adLockReadOnly
         unum = Rs2!tipoordennumeracion + 1
          
          SQLc = "UPDATE co_tipodeorden SET tipoordennumeracion=" & unum & _
                " WHERE tipoordencodigo='" & Trim(Ctrayu_tipoorden.xclave) & "' "
            VGCNx.Execute SQLc
           unum = Format(Val(unum), "00000000000")
           lblNum = unum
            SQLc = "INSERT INTO co_cabordcompra (tipoordencodigo,oc_cnumord,oc_dfecdoc,oc_ccodpro," & _
                "oc_crazsoc,oc_ccotiza,oc_ccodmon,oc_cforpag,oc_dfecent," & _
                "oc_cobserv,oc_csolict,oc_centreg,oc_estadoorden,estadooccodigo,oc_nimport,oc_ndescue," & _
                "oc_nigv,oc_nventa,oc_dfecact,oc_chora,oc_cusuari,oc_cconver,almacenorigen,almacendestino,empresacodigo,PUNTOVTACODIGO , trasladofisico) VALUES ('" & _
                Ctrayu_tipoorden.xclave & "','" & lblNum & "','" & txtEmi & "','" & txtPro & "','" & _
                CtrAyu_Proveedor.xnombre & "','" & TxFcot.text & "','" & txtMon & "','" & txtFor & "','" & _
                txtEnt & "','" & _
                SupCadSQL(txtObs) & "','" & txtSol & "','" & txtEntE & "',' ','0'," & _
                CDbl(lblImp) & "," & CDbl(lblDes) & "," & CDbl(lblIgv) & "," & CDbl(lblCom) & _
                ",'" & txtEmi.Value & "','" & Format(TxtHor, "HH:mm") & "','" & VGUsuario & _
                "','" & txtEst & "','" & Ctr_almacenO.xclave & "','" & Ctr_almacenD.xclave & "','" & VGparametros.empresacodigo & "','" & VGparametros.puntovta & "','" & ChkTraslado.Value & "')"
            VGCNx.Execute SQLc
            
            For i = 1 To Flex1.Rows - 1
                vFactor = Val(Flex1.TextMatrix(i, 6))
                vCantid = Val(Flex1.TextMatrix(i, 4))
                If vCantid = 0 Then
                   vCantid = 1
                End If
                SQLd = "INSERT INTO co_detordcompra (tipoordencodigo,oc_cnumord,oc_ccodpro,oc_dfecdoc,oc_citem," & _
                  "oc_ccodigo,oc_cdesref,oc_cunidad,estadooccodigo," & _
                  "ord_fabnum,oc_ncantid,oc_nsaldo,oc_ncanten,centrocostocodigo,entidadcodigo," & _
                  "oc_ccomen1,oc_ccomen2,tipoarticulocodigo,fam_codigo,empresacodigo,puntovtacodigo)" & _
                  " VALUES ('" & Ctrayu_tipoorden.xclave & "','" & lblNum & "','" & txtPro & "','" & txtEmi _
                  & "','" & Format(i, "000") & "','" & Flex1.TextMatrix(i, 0) & "','" & _
                  Flex1.TextMatrix(i, 2) & "','" & Flex1.TextMatrix(i, 3) & "',0,'" & _
                  Flex1.TextMatrix(i, 1) & "'," & vCantid & "," & vCantid & ",0,'" & _
                  Trim(Flex1.TextMatrix(i, 5)) & "','" & Trim(Flex1.TextMatrix(i, 6)) & "','" & _
                  Flex1.TextMatrix(i, 7) & "','" & Flex1.TextMatrix(i, 8) & "',' ','" & Flex1.TextMatrix(i, 9) & "'," _
                  & " '" & VGparametros.empresacodigo & "','" & VGparametros.puntovta & "')"
                  
                VGCNx.Execute SQLd
                
                SQLd = "UPDATE maeart SET aprecom=" & Val(Flex1.TextMatrix(i, 8)) & _
                    ",acodpro='" & txtPro & "',afecven='" & txtEmi _
                    & "' WHERE acodigo='" & Flex1.TextMatrix(i, 0) & "'"
                VGCNx.Execute SQLd
            Next
        ElseIf nT = 2 Then     'Modificar
            SQLc = "UPDATE co_cabordcompra SET oc_dfecdoc='" & txtEmi & _
                "',oc_ccodpro='" & txtPro & "',oc_crazsoc='" & Trim(CtrAyu_Proveedor.xnombre) & _
                "',oc_ccotiza='" & TxFcot.text & "',oc_ccodmon='" & txtMon & "',oc_cforpag='" & _
                txtFor & "',oc_ntipcam=" & Val(txtTip) & ",oc_dfecent='" & _
                txtEnt & "',oc_cobserv='" & SupCadSQL(txtObs) & _
                "',oc_csolict='" & txtSol & "',oc_centreg='" & txtEntE & "',oc_nimport=" & _
                CDbl(lblImp) & ",oc_ndescue=" & CDbl(lblDes) & ",oc_nigv=" & CDbl(lblIgv) & _
                ",oc_nventa=" & CDbl(lblCom) & ",oc_dfecact='" & _
                txtEmi.Value & "',oc_chora='" & Format(TxtHor, "HH:mm") & "',oc_cusuari='" & _
                VGUsuario & "',oc_cconver='" & txtEst & "',almacenorigen='" & Ctr_almacenO.xclave & "', " _
                & " almacendestino='" & Ctr_almacenD.xclave & "',trasladofisico='" & ChkTraslado.Value & "' WHERE tipoordencodigo='" & Ctrayu_tipoorden.xclave & "' and oc_cnumord='" & lblNum & "' and empresacodigo='" & VGparametros.empresacodigo & "' and puntovtacodigo='" & VGparametros.puntovta & "'"
            VGCNx.Execute SQLc
            
            SQLd = "DELETE co_detordcompra WHERE tipoordencodigo='" & Ctrayu_tipoorden.xclave & "' and oc_cnumord='" & lblNum & "' and empresacodigo='" & VGparametros.empresacodigo & "' "
            'tipoordencodigo='" & Ctrayu_tipoorden.xclave & "' and oc_cnumord='" & lblNum & "'
            VGCNx.Execute SQLd
            
            For i = 1 To Flex1.Rows - 1
                vURef = ""
                vFactor = 0
                If Flex1.TextMatrix(i, 3) <> Flex1.TextMatrix(i, 5) Then
                    vURef = Flex1.TextMatrix(i, 5)
                    vFactor = Val(Flex1.TextMatrix(i, 6))
                End If
                vCantid = Val(Flex1.TextMatrix(i, 4))
                  
                SQLd = "INSERT INTO co_detordcompra (tipoordencodigo,oc_cnumord,oc_ccodpro,oc_dfecdoc,oc_citem," & _
                  "oc_ccodigo,oc_cdesref,oc_cunidad,estadooccodigo," & _
                  "ord_fabnum,oc_ncantid,oc_nsaldo,oc_ncanten,centrocostocodigo,entidadcodigo," & _
                  "oc_ccomen1,oc_ccomen2,tipoarticulocodigo,fam_codigo,empresacodigo,puntovtacodigo)" & _
                  " VALUES ('" & Ctrayu_tipoorden.xclave & "','" & lblNum & "','" & txtPro & "','" & txtEmi _
                  & "','" & Format(i, "000") & "','" & Flex1.TextMatrix(i, 0) & "','" & _
                  Flex1.TextMatrix(i, 2) & "','" & Flex1.TextMatrix(i, 3) & "',0,'" & _
                  Flex1.TextMatrix(i, 1) & "'," & vCantid & "," & vCantid & ",0,'" & _
                  Trim(Flex1.TextMatrix(i, 5)) & "','" & Trim(Flex1.TextMatrix(i, 6)) & "','" & _
                  Flex1.TextMatrix(i, 7) & "','" & Flex1.TextMatrix(i, 8) & "',' ','" & Flex1.TextMatrix(i, 9) & "'," _
                  & " '" & VGparametros.empresacodigo & "','" & VGparametros.puntovta & "')"
                  
                VGCNx.Execute SQLd
                
                SQLd = "UPDATE maeart SET aprecom=" & Val(Flex1.TextMatrix(i, 8)) & _
                    ",acodpro='" & txtPro & "',afecven='" & txtEmi _
                    & "' WHERE acodigo='" & Flex1.TextMatrix(i, 0) & "'"
                VGCNx.Execute SQLd
            Next
        End If
        
        VGCNx.CommitTrans
        nTra = 0
        adodc1.Find "oc_cnumord='" & lblNum & "'"
        
        If nT = 1 Then
            unum = Format(Val(unum) + 1, "00000000000")
            lblNum = unum
            Limpiar
            Vacia_FlexGrid
            Estado_Items
            txtEmi = Date
            txtEnt = Date
            txtTip = "0.000"
                        
        Else
            CmdSalir2_Click
        End If
    
End If
    
    OculObj02 False
    OculObj03 False
    OculObj04 True
    OculObj06 True
    FrameInicio.Visible = True

Call Abre_Tabla_OCs
Frame1.Visible = False
Exit Sub

GrabErr:
    MsgBox Err.Description

    If nTra = 1 Then
    VGCNx.RollbackTrans
    End If
    Exit Sub
    Resume
End Sub
Private Sub cmdImp_Click()
Dim formulas(3) As String
Dim tipoorden As String
unum = adodc1("oc_cnumord")
tipoorden = adodc1("tipoordencodigo")
CrystalReport1.Reset
CrystalReport1.WindowTitle = "al_impresionrequerimientos.rpt -- orden de compra"
   CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "al_impresionRequerimientos.rpt"
    CrystalReport1.DiscardSavedData = True
 
    'CrystalReport1.Connect = VGcadenareport2
    CrystalReport1.Connect = "Provider=SQLOLEDB;Password=" & VGParamSistem.PwdGEN & ";Persist Security Info=True;User ID=" & VGParamSistem.UsuarioGEN & ";Initial Catalog=" & VGParamSistem.BDEmpresaGEN & ";SERVER=" & VGParamSistem.ServidorGEN
    
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    Dim letras As String
    letras = NUMLET(adodc1("oc_nventa"))
    If adodc1("oc_ccodmon") = "01" Then
      letras = letras + " Nuevos Soles "
     Else
      letras = letras + " Dolares Americanos "
    End If
    CrystalReport1.formulas(0) = "@emp ='" & VGparametros.NomEmpresa & "'"
    CrystalReport1.formulas(1) = "@ruc ='" & VGparametros.RucEmpresa & "'"
    CrystalReport1.formulas(2) = "@letras ='" & letras & "'"
    CrystalReport1.StoredProcParam(0) = VGCNx.DefaultDatabase
    CrystalReport1.StoredProcParam(1) = tipoorden
   CrystalReport1.StoredProcParam(2) = unum
   If CrystalReport1.Status <> 2 Then
      CrystalReport1.Action = 1
   End If

End Sub
Private Sub cmdNue_Click()
 Dim cSqlM As String, cSelM As ADODB.Recordset
    nT = 1
    OculObj06 False
    OculObj04 False
    OculObj02 True
    OculObj03 True
    FrameInicio.Visible = False
    Proceso True
    lblImp = "0.00": lblTot = "0.00": lblIgv = "0.00"
    lblDes = "0.00": lblCom = "0.00"
    Frame1.Visible = True
    Fradatos.Visible = True
    Fradatos.Enabled = True
    cmdGra.Enabled = True
    CmdSalir2.Cancel = True
    
    Ctrayu_tipoorden.Enabled = True
    CtrAyu_Proveedor.Enabled = True
    lblRuc.Enabled = True
    CtrAyu_solicitante.Enabled = True
    Ctr_almacenO.Enabled = True
    Ctr_almacenD.Enabled = True
    ChkTraslado.Enabled = True
    TxFcot.Enabled = True
    txtObs.Enabled = True
    Flex1.Enabled = True
    cmdEdi2.Enabled = True
    
    Flex1.Rows = 1
    Ctrayu_tipoorden.SetFocus
End Sub
Private Sub cmdEdi_Click()
    If adodc1("oc_estadoorden") = 1 Then
        Mensaje = "La Orden de compra ha sido anulada, no se permitirá modificaciones"
        MsgBox Mensaje, vbExclamation, "Advertencia"
        cmdNue2.Enabled = False
        cmdEdi2.Enabled = False
        cmdEli2.Enabled = False
        cmdGra.Enabled = False
        OculObj06 False
        OculObj04 False
        OculObj02 True
        Mostrar adodc1("tipoordencodigo"), adodc1("oc_cnumord")
        OculObj03 True
        Proceso True
        Fradatos.Enabled = False
        
        Ctrayu_tipoorden.Enabled = False
        CtrAyu_Proveedor.Enabled = False
        lblRuc.Enabled = False
        CtrAyu_solicitante.Enabled = False
        Ctr_almacenO.Enabled = False
        Ctr_almacenD.Enabled = False
        ChkTraslado.Enabled = False
        TxFcot.Enabled = False
        txtObs.Enabled = False
        Flex1.Enabled = False
        cmdEdi2.Enabled = False
    Else
        nT = 2
        OculObj06 False
        OculObj04 False
        OculObj02 True
        Mostrar adodc1("tipoordencodigo"), adodc1("oc_cnumord")
        OculObj03 True
        Proceso True
        Fradatos.Enabled = True
        Frame1.Visible = True
        cmdEdi2.Enabled = True
        cmdEli2.Enabled = True
        cmdGra.Enabled = True
        
        Ctrayu_tipoorden.Enabled = False
        CtrAyu_Proveedor.Enabled = False
        lblRuc.Enabled = False
        CtrAyu_solicitante.Enabled = False
        Ctr_almacenO.Enabled = False
        Ctr_almacenD.Enabled = False
        ChkTraslado.Enabled = False
        TxFcot.Enabled = False
        txtObs.Enabled = False
        Flex1.Enabled = False
        cmdEdi2.Enabled = False
        
        txtEmi.SetFocus
        CmdSalir2.Cancel = True
    End If
    FrameInicio.Visible = False
End Sub
Private Sub cmdNue2_Click()
On Error GoTo errores

With frmRequerimientosDetalle
   .activado = False
'   .CtrAyu_articulo.xclave = ""
   .txtCan = "0.00"
   .TxtOrdfab.text = ""
   .lblFab.Caption = ""
   .Txtco1.text = ""
   .Txtco2.text = ""
   .activado = True
   .Show 1
   If .CtrAyu_articulo.xclave = "" Then .CtrAyu_articulo.xclave = "00"
   If .CtrAyu_articulo.xnombre = "" Then .CtrAyu_articulo.xnombre = "Ninguno"
   If .lblUnidad = "" Then .lblUnidad = "XX"
   If Not .cancelado Then
      If .tipo = "S" Then
         .txtCan = 1
      End If
      Flex1.AddItem Trim(.CtrAyu_articulo.xclave) & vbTab & .TxtOrdfab.text & vbTab & _
      Trim(.CtrAyu_articulo.xnombre) & vbTab & _
      .lblUni & vbTab & .txtCan.text & vbTab & .CtrAyu_Ccosto.xclave & vbTab & _
      .CtrAyu_Analitico.xclave & vbTab & .Txtco1.text & vbTab & .Txtco2.text & vbTab & .Ctr_AyuFamilia.xclave
      If Flex1.Rows - 1 > 0 Then
         Flex1.Row = Flex1.Rows - 1
      End If
      Estado_Items
      Flex1.SetFocus
      cmdNue2.SetFocus
  Else
      Flex1.SetFocus
      cmdNue2.SetFocus
  End If
End With

Exit Sub
errores:
MsgBox "Error " & Err.Description & Chr(13) & "", vbCritical, "Mensaje"

End Sub

Private Sub CmdSaliirAtendido_Click()
FrameAtendido.Visible = False
End Sub

Private Sub CmdSalir_Click()
     Unload Me
End Sub
Private Sub CmdSalir2_Click()
On Error GoTo Err
Limpiar
 '  Vacia_FlexGrid
    Estado_Items
    Estado_Botones
    OculObj02 False
    OculObj03 False
    OculObj04 True
    OculObj06 True
    Proceso False
    Frame1.Visible = False
    FrameInicio.Visible = True
    If adodc1.RecordCount > 0 Then
    '    TDBGrid1.SetFocus
    Else
        cmdNue.SetFocus
    End If
    CmdSalir.Cancel = True
    Exit Sub
Err:
    MsgBox Err.Description
    Exit Sub
    Resume

End Sub
Public Function SupCadSQL(s As String) As String
 Dim Aux As String
 If Not IsNull(s) Then
     Aux = Replace(s, "'", "''")
 End If
 SupCadSQL = Aux
 
End Function

Private Sub Ctr_almacenD_GotFocus()
'Ctr_almacenD.filtro = " taalma <>'" & Ctr_almacenO.xclave & "' "
'Ctr_almacenD.Ejecutar
End Sub

Private Sub Ctrayu_tipoorden_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Dim unum As String
    
Set VGvardllgen = New dllgeneral.dll_general
unum = VGvardllgen.ESNULO(ColecCampos("tipoordennumeracion").Value, "")
unum = Format(Val(unum) + 1, "00000000000")
lblNum = unum
    
    
    
End Sub


Private Sub CtrAyu_Proveedor_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    Set VGvardllgen = New dllgeneral.dll_general
    lblRuc.text = VGvardllgen.ESNULO(ColecCampos("clienteruc").Value, "")
End Sub
Private Sub CtrAyu_Proveedor_AlNoDevolverNada()
    lblRuc.text = ""
End Sub



Private Sub Form_Load()
FrameAtendido.Visible = False
Formato_FlexGrid
Call Ctrayu_tipoorden.conexion(VGCNx): Ctrayu_tipoorden.filtro = "(flagrequerimientosPedidos= 1) "
Call CtrAyu_Proveedor.conexion(VGCNx)
Call CtrAyu_solicitante.conexion(VGCNx)
Call Ctr_almacenO.conexion(VGCNx)

Call Ctr_almacenD.conexion(VGCNx)
Ctr_almacenD.filtro = " empresacodigo='" & VGparametros.empresacodigo & "' and puntovtacodigo='" & VGparametros.puntovta & "'"

FrameInicio.Visible = True
OculObj02 False
OculObj03 False
OculObj04 True
OculObj06 True
TDBGrid1.FetchRowStyle = True
txtEmi.Value = Date
txtEnt.Value = Date
unum = ""
Abre_Tabla_OCs
Estado_Botones
Frame1.Visible = False
    
CmdSalir.Picture = MDIPrincipal.ImageList2.ListImages.item("Retornar").Picture
cmdImp.Picture = MDIPrincipal.ImageList2.ListImages.item("Imprimir").Picture
CmdEli.Picture = MDIPrincipal.ImageList2.ListImages.item("Eliminar").Picture
cmdEdi.Picture = MDIPrincipal.ImageList2.ListImages.item("Modificar").Picture
cmdNue.Picture = MDIPrincipal.ImageList2.ListImages.item("Nuevo").Picture

CmdSalir2.Picture = MDIPrincipal.ImageList2.ListImages.item("Retornar").Picture
cmdGra.Picture = MDIPrincipal.ImageList2.ListImages.item("Grabar").Picture
cmdEli2.Picture = MDIPrincipal.ImageList2.ListImages.item("Eliminar").Picture
cmdEdi2.Picture = MDIPrincipal.ImageList2.ListImages.item("Modificar").Picture
cmdNue2.Picture = MDIPrincipal.ImageList2.ListImages.item("Nuevo").Picture

End Sub
Private Sub Reales_Positivos(k As Integer, t As TextBox)
Dim t1 As String
    k = Asc(UCase(Chr(k)))
    If k = 8 Then Exit Sub
    If k <> 45 And k <> 44 And k <> 32 And k <> 69 And k <> 43 Then
        t1 = Left(t, t.SelStart)
        t1 = t1 & Chr(k) & Right(t, Len(t) - Len(t1))
        If IsNumeric(t1) Then Exit Sub
    End If
    k = 0
    
End Sub
Public Function Existe(tipo As Integer, Cod As String, Tabla As String, Campo As String, Fecha As Boolean, Optional Cod2 As String, Optional cCampo2 As String, Optional Cod3 As String, Optional cCampo3 As String, Optional Cod4 As Boolean, Optional cCampo4 As String, Optional Cod5 As String, Optional cCampo5 As String) As Boolean
Dim cSel1 As ADODB.Recordset, cSL As String
Set cSel1 = New ADODB.Recordset

 If Fecha Then
        cSL = "Select * from " & Tabla & "  Where " & Campo & " =  '" & Cod & "'"
 Else
       If UCase(Tabla) = "PUNTO_VENTA" Then
                cSL = "Select * from " & Tabla & "  Where " & Campo & " =  '" & Cod & "'"
       Else
                cSL = "Select * from " & Tabla & "  Where " & Campo & " =  '" & Cod & "'"
       End If
       If Trim(Cod2) <> "" Then
            cSL = cSL & " And  " & cCampo2 & " =  '" & SupCadSQL(Cod2) & "'"
       End If
       If Trim(Cod3) <> "" Then
            cSL = cSL & " And  " & cCampo3 & " =  '" & SupCadSQL(Cod3) & "'"
       End If
       If Trim(cCampo4) <> "" Then
            If Cod4 = True Then
                cSL = cSL & " And  " & cCampo4
            Else
                cSL = cSL & " And  " & Not cCampo4
            End If
        End If
        If Trim(Cod5) <> "" Then
            cSL = cSL & " And  " & cCampo5 & " =  '" & Cod5 & "'"
        End If
 End If
 
Select Case tipo
Case 1 'Bd. Comun
            cSel1.Open cSL, VGCNx, adOpenStatic
Case 2 'Bd. Config
            cSel1.Open cSL, VGCNx, adOpenStatic
Case 3 'Bd. Contab
            cSel1.Open cSL, VGcnxCT, adOpenStatic
End Select

If cSel1.RecordCount > 0 Then
     Existe = True
Else
     Existe = False
End If
'csel1.Close
End Function
Sub Limpiar()
txtEntE = "": txtObs = ""
End Sub
Sub Mostrar(cC1 As String, CC2 As String)
    Dim cSqlM As String, cSelM As ADODB.Recordset
    Dim k As Integer, i As Integer, vd As String
    Dim vpu As Single, txtPro As String
    Dim txtSol As String
    
    lblNum = CC2
    If Escadena(adodc1("oc_ccodpro")) = "" Then
       CtrAyu_Proveedor.xclave = "00"
     Else
       CtrAyu_Proveedor.xclave = Escadena(adodc1("oc_ccodpro"))
    End If
    txtPro = CtrAyu_Proveedor.xclave
    CtrAyu_Proveedor.xnombre = Devolver_Dato(1, txtPro, "cp_proveedor", "clientecodigo", False, "clienterazonsocial")
    lblRuc.text = txtPro
    txtEmi = adodc1("oc_dfecdoc")
    txtEnt = adodc1("oc_dfecent")
    TxFcot.text = adodc1("oc_ccotiza")
    txtEntE = adodc1("oc_centreg")
    CtrAyu_solicitante.xclave = adodc1("oc_csolict")
    txtSol = CtrAyu_solicitante.xclave
    CtrAyu_solicitante.xnombre = Devolver_Dato(1, txtSol, "co_solicitantes", "solicitantecodigo", False, "solicitantenombre")
    txtObs = adodc1("oc_cobserv")
    Ctrayu_tipoorden.xclave = adodc1("tipoordencodigo")
    Ctrayu_tipoorden.xnombre = adodc1("tipoordendescripcion")
    'Ctrayu_tipoorden.Ejecutar
    Ctr_almacenO.xclave = adodc1("almacenorigen"): Ctr_almacenO.Ejecutar
    Ctr_almacenD.xclave = adodc1("almacendestino"):: Ctr_almacenD.Ejecutar
    ChkTraslado.Value = IIf(adodc1.Fields("trasladofisico") = True, 1, 0)
    TxtHor.text = adodc1("oc_chora")
    
    cSqlM = "SELECT * FROM co_detordcompra WHERE tipoordencodigo='" & cC1 & "' "
    cSqlM = cSqlM & " AND oc_cnumord='" & CC2 & "' ORDER BY oc_citem"
    Set cSelM = New ADODB.Recordset
    
    cSelM.Open cSqlM, VGCNx, adOpenStatic
    If cSelM.RecordCount() > 0 Then cSelM.MoveFirst
    
    k = 0
    Do While Not cSelM.EOF
        k = k + 1
        If cSelM("oc_ncantid") > 0 Then
           vpu = IIf(ESNULO(cSelM("oc_nfactor"), 0) = 0, 1, cSelM("oc_nfactor") / cSelM("oc_ncantid"))
         Else
           vpu = 1
        End If
        Flex1.AddItem cSelM("oc_ccodigo") & vbTab & cSelM("oc_ccodref") & vbTab & _
             cSelM("oc_cdesref") & vbTab & cSelM("oc_cunidad") & vbTab & _
             Format(cSelM("oc_ncantid"), "0.00") & vbTab & _
             cSelM("Centrocostocodigo") & vbTab & cSelM("entidadcodigo") & vbTab & _
             cSelM("oc_ccomen1") & vbTab & cSelM("oc_ccomen2") & vbTab & cSelM("fam_codigo"), 1
        If k = 1 Then
           Flex1.Rows = 2
        End If
        cSelM.MoveNext
    Loop
    cSelM.Close
End Sub

Sub Estado_Botones()
    If adodc1.RecordCount > 0 Then
        cmdImp.Enabled = True
    Else
        cmdImp.Enabled = False
    End If
End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
 With adodc1
    If .Sort = Empty Then
        .Sort = TDBGrid1.Columns.item(ColIndex).DataField & " asc"
    ElseIf Right(.Sort, 3) = "asc" Then
        .Sort = TDBGrid1.Columns.item(ColIndex).DataField & " desc"
    ElseIf Right(.Sort, 4) = "desc" Then
        .Sort = TDBGrid1.Columns.item(ColIndex).DataField & " asc"
    End If
    TDBGrid1.Refresh
 End With
End Sub


Private Sub txtEmi_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  SendKeys "{TAB}"
End If
End Sub
Private Sub txtEmi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not ValidFecha(txtEmi) Then
            Mensaje = "Fecha No Válida"
            MsgBox Mensaje, vbExclamation, "Error"
            txtEmi.SetFocus
        Else
            txtEnt.SetFocus
        End If
    End If
End Sub
Function ValidFecha(vText As String) As String
Dim cTxtNew As String, ncnt As Integer
Dim cTxt As String, cTxtDig As String
cTxtDig = "": cTxtNew = ""
For ncnt = 1 To Len(vText)
      cTxt = Mid(vText, ncnt, 1)
      If cTxt = "/" Then
         cTxtNew = cTxtNew & Str(Val(cTxtDig)) & "/"
         cTxtDig = ""
      Else
         If cTxt <> "_" Then cTxtDig = cTxtDig & cTxt
      End If
Next
If cTxtDig <> "" Then cTxtNew = cTxtNew & Str(Val(cTxtDig))

If IsDate(cTxtNew) Then
   ValidFecha = Format(CDate(cTxtNew), "dd/mm/yyyy")
End If
End Function


Private Sub txtEnt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  SendKeys "{TAB}"
End If
End Sub
Private Sub txtEnt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not ValidFecha(txtEnt) Then
            Mensaje = "Fecha No Válida"
            MsgBox Mensaje, vbExclamation, "Error"
            txtEnt.SetFocus
        End If
    End If
End Sub
Private Sub txtEntE_GotFocus()
    Enfoque txtEntE
End Sub

Private Sub TxtHor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub


Private Sub txtObs_GotFocus()
    Enfoque txtObs
End Sub
Private Sub txtObs_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmdEli2.Enabled Then
            Flex1.SetFocus
        Else
            cmdNue2.SetFocus
        End If
    End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub Enfoque(OBJ As Object)
  OBJ.SelStart = 0
  OBJ.SelLength = Len(OBJ)
End Sub
Sub Proceso(Estado As Boolean)
    Flex1.Visible = Estado
    cmdNue2.Visible = Estado
    cmdEdi2.Visible = Estado
    cmdEli2.Visible = Estado
End Sub
Sub Formato_FlexGrid()
    Flex1.ColWidth(0) = 1100
    Flex1.ColWidth(1) = 1000
    Flex1.ColWidth(2) = 2000
    Flex1.ColWidth(3) = 600
    Flex1.ColWidth(4) = 1000
    Flex1.ColWidth(5) = 800
    Flex1.ColWidth(6) = 800
    Flex1.ColWidth(7) = 1500
    Flex1.ColWidth(8) = 1500
    Flex1.ColWidth(9) = 500
End Sub
Sub Estado_Items()
    If Flex1.Rows - 1 = 1 Then
        If Flex1.TextMatrix(1, 0) = "" Then
            cmdEdi2.Enabled = False
            cmdEli2.Enabled = False
            cmdNue2.Enabled = True
            cmdNue2.SetFocus
        Else
            cmdEdi2.Enabled = True
            cmdEli2.Enabled = True
        End If
    Else
        cmdEdi2.Enabled = True
        cmdEli2.Enabled = True
    End If
End Sub
Sub Vacia_FlexGrid()
    Dim i As Integer
    
    If Flex1.Rows - 1 > 1 Then
       Do While Flex1.Rows - 1 > 1
          Flex1.RemoveItem 1
       Loop
        For i = 0 To 9
           Flex1.TextMatrix(1, i) = ""
       Next
    End If
End Sub
Function Tiene_Entregas() As Boolean
    Dim Adodc2 As ADODB.Recordset
    
    Set Adodc2 = New ADODB.Recordset
    
    Adodc2.Open "SELECT * FROM co_detordcompra WHERE oc_cnumord='" & lblNum & "' AND oc_ccodigo='" & _
        Flex1.TextMatrix(Flex1.Row, 0) & "' AND oc_ncanten>0", VGCNx, adOpenStatic
    Tiene_Entregas = False
    If Adodc2.RecordCount > 0 Then Tiene_Entregas = True
End Function


