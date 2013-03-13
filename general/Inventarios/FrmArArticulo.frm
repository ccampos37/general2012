VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Begin VB.Form FrmArArticulo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualiza Datos Generales de Articulos "
   ClientHeight    =   6270
   ClientLeft      =   1785
   ClientTop       =   2310
   ClientWidth     =   9975
   Icon            =   "FrmArArticulo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   9975
   Begin VB.Frame FrameNuevoGrupo 
      Caption         =   "Grupo"
      Height          =   1455
      Left            =   120
      TabIndex        =   63
      Top             =   0
      Width           =   9375
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayufamilia 
         Height          =   375
         Left            =   1440
         TabIndex        =   64
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   661
         XcodMaxLongitud =   10
         xcodwith        =   450
         NomTabla        =   "familia"
         ListaCampos     =   "FAM_CODIGO(1),FAM_NOMBRE(1)"
         XcodCampo       =   "FAM_CODIGO"
         XListCampo      =   "FAM_NOMBRE"
         ListaCamposDescrip=   "codigo, descripcion"
         ListaCamposText =   "FAM_CODIGO,FAM_NOMBRE"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayulinea 
         Height          =   375
         Left            =   1440
         TabIndex        =   65
         Top             =   600
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   661
         XcodMaxLongitud =   10
         xcodwith        =   450
         NomTabla        =   "lineas"
         ListaCampos     =   "lin_codigo(1), lin_nombre(1)"
         XcodCampo       =   "lin_codigo"
         XListCampo      =   "lin_nombre"
         ListaCamposDescrip=   "codigo, descripcion"
         ListaCamposText =   "lin_codigo, lin_nombre"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayugrupo 
         Height          =   375
         Left            =   1440
         TabIndex        =   66
         Top             =   960
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   661
         XcodMaxLongitud =   10
         xcodwith        =   450
         NomTabla        =   "grupo"
         ListaCampos     =   "gru_codigo(1),gru_nombre(1)"
         XcodCampo       =   "gru_codigo"
         XListCampo      =   "gru_nombre"
         ListaCamposDescrip=   "codigo, descripcion"
         ListaCamposText =   "gru_codigo,gru_nombre"
      End
      Begin VB.Label Label16 
         Caption         =   "Familia"
         Height          =   255
         Left            =   360
         TabIndex        =   69
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "Linea"
         Height          =   255
         Left            =   360
         TabIndex        =   68
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Grupo"
         Height          =   255
         Left            =   360
         TabIndex        =   67
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   0
      TabIndex        =   55
      Top             =   5040
      Width           =   9615
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&Grabar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   7650
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   240
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.CommandButton CmdModi 
         Caption         =   "&Modificar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   1185
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   240
         Width           =   870
      End
      Begin VB.CommandButton CmdEli 
         Caption         =   "&Eliminar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   2250
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   240
         Width           =   870
      End
      Begin VB.CommandButton CmdSalir2 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   8730
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   240
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.CommandButton CmdFicha 
         Caption         =   "&Ficha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   3315
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   240
         Width           =   870
      End
      Begin VB.CommandButton CmdIng 
         Caption         =   "&Crear"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   240
         Width           =   870
      End
      Begin VB.CommandButton CmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   4410
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   240
         Width           =   870
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5130
      Left            =   -15
      TabIndex        =   0
      Top             =   0
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   9049
      _Version        =   393216
      TabOrientation  =   3
      Style           =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Mant"
      TabPicture(0)   =   "FrmArArticulo.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrameNuevo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Busqueda"
      TabPicture(1)   =   "FrmArArticulo.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).Control(1)=   "DataGrid1"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "F.tecnica"
      TabPicture(2)   =   "FrmArArticulo.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).ControlCount=   1
      Begin VB.Frame FrameNuevo 
         Height          =   705
         Left            =   90
         TabIndex        =   47
         Top             =   240
         Width           =   9396
         Begin VB.TextBox Textfam 
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
            Left            =   990
            MaxLength       =   8
            TabIndex        =   48
            Top             =   270
            Width           =   975
         End
         Begin TextFer.TxFer TxFCorrelativo 
            Height          =   300
            Left            =   6960
            TabIndex        =   49
            Top             =   270
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
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
            Text            =   ""
            Valor           =   ""
         End
         Begin VB.Label Lblfam 
            Appearance      =   0  'Flat
            BackColor       =   &H00C47013&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   285
            Left            =   2070
            TabIndex        =   52
            Top             =   270
            Width           =   3495
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Familia :"
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
            TabIndex        =   51
            Top             =   300
            Width           =   690
         End
         Begin VB.Label Lblcorrelativo 
            AutoSize        =   -1  'True
            Caption         =   "Correlativo :"
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
            Left            =   5760
            TabIndex        =   50
            Top             =   300
            Width           =   1035
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3690
         Left            =   90
         TabIndex        =   9
         Top             =   1335
         Visible         =   0   'False
         Width           =   9396
         Begin VB.CheckBox CheckCom 
            Caption         =   "Comision"
            Height          =   285
            Left            =   6720
            TabIndex        =   54
            Top             =   720
            Value           =   1  'Checked
            Width           =   1905
         End
         Begin VB.CheckBox ChkSer 
            Caption         =   "Bien"
            Height          =   285
            Left            =   6780
            TabIndex        =   53
            Top             =   360
            Value           =   1  'Checked
            Width           =   1905
         End
         Begin VB.TextBox TxFamilia 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1935
            MaxLength       =   8
            TabIndex        =   44
            Text            =   "TxFamilia"
            Top             =   2160
            Width           =   975
         End
         Begin VB.TextBox TxTalla 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5535
            MaxLength       =   3
            TabIndex        =   43
            Text            =   "TxTalla"
            Top             =   2655
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.TextBox txPartAran 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7692
            MaxLength       =   30
            TabIndex        =   41
            Text            =   "txPartAran"
            Top             =   1890
            Width           =   1545
         End
         Begin VB.TextBox TxPeso 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5415
            MaxLength       =   9
            TabIndex        =   25
            Text            =   "TxPes"
            Top             =   3075
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.TextBox TxGrupo 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1935
            MaxLength       =   8
            TabIndex        =   24
            Text            =   "TxGrupo"
            Top             =   2955
            Width           =   975
         End
         Begin VB.TextBox TxLinea 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1935
            MaxLength       =   8
            TabIndex        =   23
            Text            =   "TxLinea"
            Top             =   2595
            Width           =   975
         End
         Begin VB.TextBox TxUnidad 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1935
            MaxLength       =   6
            TabIndex        =   22
            Text            =   "TxUnidad"
            Top             =   1770
            Width           =   975
         End
         Begin VB.TextBox TxAlterna 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1935
            MaxLength       =   50
            TabIndex        =   21
            Text            =   "TxAlterna"
            Top             =   1395
            Width           =   5175
         End
         Begin VB.TextBox TxDescripcion 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1935
            MaxLength       =   50
            TabIndex        =   20
            Text            =   "TxDescripcion"
            Top             =   1035
            Width           =   5175
         End
         Begin VB.TextBox TxFabricante 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1935
            MaxLength       =   40
            TabIndex        =   19
            Text            =   "TxFabric"
            Top             =   675
            Width           =   3750
         End
         Begin VB.TextBox TxCodigo 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1935
            MaxLength       =   20
            TabIndex        =   18
            Text            =   "TxCodigo"
            Top             =   315
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "FrmArArticulo.frx":091E
            Left            =   7500
            List            =   "FrmArArticulo.frx":092B
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   1380
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Frame Frame3 
            Height          =   615
            Left            =   5256
            TabIndex        =   13
            Top             =   2070
            Width           =   3975
            Begin VB.OptionButton OptLibre 
               Caption         =   "Libre"
               Height          =   255
               Left            =   3000
               TabIndex        =   16
               Top             =   240
               Width           =   855
            End
            Begin VB.OptionButton OptLote 
               Caption         =   "Stock x Lote"
               Height          =   255
               Left            =   1560
               TabIndex        =   15
               Top             =   240
               Width           =   1215
            End
            Begin VB.OptionButton OptSerie 
               Caption         =   "Stock x Serie"
               Height          =   255
               Left            =   120
               TabIndex        =   14
               Top             =   240
               Width           =   1455
            End
         End
         Begin VB.TextBox TxTipo 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1935
            MaxLength       =   2
            TabIndex        =   12
            Top             =   3330
            Width           =   495
         End
         Begin VB.TextBox TxMarca 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7692
            MaxLength       =   20
            TabIndex        =   11
            Text            =   "TxMarca"
            Top             =   2685
            Width           =   1545
         End
         Begin VB.TextBox TxtColor 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7692
            MaxLength       =   20
            TabIndex        =   10
            Text            =   "TxtColor"
            Top             =   3072
            Width           =   1545
         End
         Begin VB.Label Label10 
            Caption         =   "Familia                        :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   195
            TabIndex        =   46
            Top             =   2160
            Width           =   1770
         End
         Begin VB.Label LbFamilia 
            Caption         =   "LbFamilia"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3000
            TabIndex        =   45
            Top             =   2220
            Width           =   2925
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Talla :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4995
            TabIndex        =   42
            Top             =   2685
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Partida Arancelaria :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6165
            TabIndex        =   40
            Top             =   1920
            Width           =   1470
         End
         Begin VB.Label LbGrupo 
            Caption         =   "LbGrupo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3000
            TabIndex        =   39
            Top             =   2985
            Width           =   3105
         End
         Begin VB.Label LbLinea 
            Caption         =   "LbLinea"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3000
            TabIndex        =   38
            Top             =   2655
            Width           =   3720
         End
         Begin VB.Label LbUnidad 
            Caption         =   "LbUnidad"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3015
            TabIndex        =   37
            Top             =   1830
            Width           =   2895
         End
         Begin VB.Label Label26 
            Caption         =   "Tipo Articulo              :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   195
            TabIndex        =   36
            Top             =   3315
            Width           =   1785
         End
         Begin VB.Label Label24 
            Caption         =   "Peso Articulo             :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4395
            TabIndex        =   35
            Top             =   3195
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label12 
            Caption         =   "Grupo                        :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   195
            TabIndex        =   34
            Top             =   2955
            Width           =   1665
         End
         Begin VB.Label Label11 
            Caption         =   "Linea                          :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   195
            TabIndex        =   33
            Top             =   2595
            Width           =   1770
         End
         Begin VB.Label Label9 
            Caption         =   "Unidad Medida           :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   195
            TabIndex        =   32
            Top             =   1755
            Width           =   1695
         End
         Begin VB.Label Label4 
            Caption         =   "Descripción Alterna    :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   195
            TabIndex        =   31
            Top             =   1395
            Width           =   1815
         End
         Begin VB.Label Label3 
            Caption         =   "Descripción                 :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   195
            TabIndex        =   30
            Top             =   1050
            Width           =   1770
         End
         Begin VB.Label Label2 
            Caption         =   "Código Fabricante      :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   195
            TabIndex        =   29
            Top             =   690
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Código                        :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   195
            TabIndex        =   28
            Top             =   360
            Width           =   1635
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Clase :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7110
            TabIndex        =   27
            Top             =   2715
            Width           =   495
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Cod. Conversion :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6330
            TabIndex        =   26
            Top             =   3090
            Width           =   1305
         End
      End
      Begin VB.Frame Frame5 
         Height          =   615
         Left            =   -74850
         TabIndex        =   4
         Top             =   30
         Width           =   9288
         Begin VB.ComboBox CmbOrden 
            Height          =   315
            ItemData        =   "FrmArArticulo.frx":094E
            Left            =   4560
            List            =   "FrmArArticulo.frx":0967
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   180
            Width           =   2265
         End
         Begin VB.TextBox TxFiltro 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   990
            TabIndex        =   5
            Text            =   "TxFiltro"
            Top             =   180
            Width           =   2655
         End
         Begin VB.Label Label33 
            Caption         =   "Orden :"
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
            Left            =   3870
            TabIndex        =   8
            Top             =   210
            Width           =   855
         End
         Begin VB.Label Label32 
            Caption         =   "Buscar  :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   210
            TabIndex        =   7
            Top             =   225
            Width           =   930
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datos Técnicos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4815
         Left            =   -74940
         TabIndex        =   1
         Top             =   45
         Visible         =   0   'False
         Width           =   9345
         Begin VB.TextBox TxTecnica 
            Height          =   3450
            Left            =   210
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Text            =   "FrmArArticulo.frx":09C1
            Top             =   615
            Visible         =   0   'False
            Width           =   8970
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "FrmArArticulo.frx":09C7
         Height          =   3945
         Left            =   -74850
         TabIndex        =   3
         Top             =   600
         Width           =   9270
         _ExtentX        =   16351
         _ExtentY        =   6959
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   2
         WrapCellPointer =   -1  'True
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "ACODIGO"
            Caption         =   "  CODIGO"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "ADESCRI"
            Caption         =   "                      DESCRIPCION"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "AUNIDAD"
            Caption         =   "  UNIDAD"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "ACODIGO2"
            Caption         =   "COD. FABRI. "
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "#.##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "AFAMILIA"
            Caption         =   " FAMILIA"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "AMODELO"
            Caption         =   "  LINEA"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "AGRUPO"
            Caption         =   "   GRUPO"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "ATIPO"
            Caption         =   "      TIPO"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "ACUENTA"
            Caption         =   " CTA. CONT."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "AFSTOCK"
            Caption         =   "AFSTOCK"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            ScrollBars      =   3
            BeginProperty Column00 
               Alignment       =   2
               Locked          =   -1  'True
               WrapText        =   -1  'True
               ColumnWidth     =   2190.047
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               WrapText        =   -1  'True
               ColumnWidth     =   5760
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               Locked          =   -1  'True
               WrapText        =   -1  'True
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               Locked          =   -1  'True
               WrapText        =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               Locked          =   -1  'True
               WrapText        =   -1  'True
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               Locked          =   -1  'True
               WrapText        =   -1  'True
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               Locked          =   -1  'True
               WrapText        =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
               Locked          =   -1  'True
               WrapText        =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               Locked          =   -1  'True
               WrapText        =   -1  'True
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column09 
               Locked          =   -1  'True
               WrapText        =   -1  'True
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmArArticulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adodc1 As ADODB.Recordset
Dim Adodc2 As ADODB.Recordset
Dim Adodc3 As ADODB.Recordset
Dim cSel1 As ADODB.Recordset
Dim cSql1 As String, CSQL2 As String
Dim nT As Integer       'Ingreso,Modificación,Ficha Tecnica
Dim cCod As String, cDes As String
Dim nCom As Integer, nUni As Integer
Dim cEstado As String
Dim cUsuario As String
Dim nTra As Integer, nCursor As Integer
Dim EstadoAnt As Integer
Private Sub OculObj01(ntipo As Boolean) ' Ficha Tecnica
Frame2.Visible = ntipo
TxTecnica.Visible = ntipo
End Sub
Private Sub OculObj02(ntipo As Boolean)  'Grabar y salir
Cmdgrabar.Visible = ntipo
CmdSalir2.Visible = ntipo
End Sub
Private Sub OculObj03(ntipo As Boolean) ' Todos los datos
Frame1.Visible = ntipo
End Sub
Private Sub OculObj04(ntipo As Boolean) ' Botones principales
CmdIng.Visible = ntipo
CmdModi.Visible = ntipo
CmdEli.Visible = ntipo
CmdFicha.Visible = ntipo
CmdSalir.Visible = ntipo
End Sub
Private Sub OculObj05(ntipo As Boolean)  'Orden y Filtro
Frame5.Visible = ntipo
Label32.Visible = ntipo
TxFiltro.Visible = ntipo
Label33.Visible = ntipo
cmbOrden.Visible = ntipo
End Sub
Private Sub OculObj06(ntipo As Boolean)  'Datagrid
DataGrid1.Visible = ntipo
End Sub
Private Sub CmbOrden_Click()             ' Ordenar por
Dim cD As String
On Error GoTo Err
nCom = cmbOrden.ListIndex

Set adodc1 = New ADODB.Recordset
cD = "Select ACODIGO,ADESCRI,AUNIDAD,ACODIGO2,AFAMILIA,ALINEA,AGRUPO,ATIPO,ACUENTA,AMARCA,AFSTOCK FROM MAEART "

Select Case nCom
Case 0
            cD = cD & " ORDER BY ACODIGO"
Case 1
            cD = cD & " ORDER BY ADESCRI"
Case 2
            cD = cD & " ORDER BY AGRUPO"
Case 3
            cD = cD & " ORDER BY AFAMILIA"
Case 4
            cD = cD & " ORDER BY ALINEA"
Case 5
            cD = cD & " ORDER BY ACODIGO2"
Case 6
            cD = cD & " ORDER BY ATIPO"
End Select
adodc1.Open cD, VGCNx, adOpenStatic
TxFiltro = ""
Set DataGrid1.DataSource = adodc1
If DataGrid1.Visible Then DataGrid1.SetFocus
Exit Sub
Err:
  MsgBox Err.Description & Chr(13) & "Salir del Formulario", vbInformation, "Aviso"
End Sub

Private Sub CmdEli_Click()              ' Elimina
Dim ACMD As New ADODB.Command
Dim CSQL2 As String
Dim nTra1 As Integer
Dim nN As Integer

On Error GoTo EliErr

If adodc1.RecordCount > 0 Then
    cCod = adodc1("ACODIGO")
    'SE HACIA REFERENCIA A STKART .......
    'RMM*******************************************************************
    cSql1 = " SELECT TOP 1 MOVALMDET.DECODIGO FROM MOVALMCAB INNER JOIN MOVALMDET ON (MOVALMCAB.CANUMDOC = MOVALMDET.DENUMDOC) AND (MOVALMCAB.CATD = MOVALMDET.DETD) AND (MOVALMCAB.CAALMA = MOVALMDET.DEALMA) WHERE (((MOVALMDET.DECODIGO)='" & cCod & "')) AND  CASITGUI <> 'A'"
    'RMM*******************************************************************
    Set cSel1 = New ADODB.Recordset
    cSel1.Open cSql1, VGCNx, adOpenStatic
    If cSel1.RecordCount > 0 Then          ' vGAlmacen
        MsgBox "El artículo tiene Movimientos de Almacén, no se puede Eliminar", vbInformation, "Mensaje"
        cSel1.Close
        Exit Sub
    End If
    cSel1.Close
    
    cSql1 = " SELECT *  FROM KITS WHERE KITS.CODKIT='" & cCod & "'"
    'RMM*******************************************************************
    Set cSel1 = New ADODB.Recordset
    cSel1.Open cSql1, VGCNx, adOpenStatic
    If cSel1.RecordCount > 0 Then          ' vGAlmacen
        MsgBox "El artículo es parte de un codigo KIT, no se puede Eliminar", vbInformation, "Mensaje"
        cSel1.Close
        Exit Sub
    End If
    cSel1.Close
        
    If MsgBox("   Desea Eliminar " & Chr(10) & "" & Mid(adodc1("ADESCRI"), 1, 25) & "", vbQuestion + vbOKCancel, "Mensaje") = vbOK Then
   '      nN = Pos_Dato(adodc1)
        cSql1 = "Delete from MAEART where ACODIGO = '" & cCod & "'"
        CSQL2 = "Delete from STKART where STCODIGO = '" & cCod & "'"
        nTra = 1
        VGCNx.BeginTrans
        VGCNx.Execute cSql1
        VGCNx.Execute CSQL2
        VGCNx.CommitTrans
        
        cSql1 = "Delete from STKSERI where STSCODIGO= '" & cCod & "'"
        CSQL2 = "Delete from STKLOTE where STSCODIGO= '" & cCod & "'"
        VGCNx.Execute cSql1
        VGCNx.Execute CSQL2
        
        nTra = 0
        
     '**** grabamos en facturacion
        Set ACMD.ActiveConnection = VGCNx
        ACMD.CommandText = "al_actualizaproducto_pro"
        ACMD.CommandType = adCmdStoredProc
        ACMD.Prepared = True
        With ACMD
            .Parameters("@baseini") = VGCNx.DefaultDatabase
            .Parameters("@basefin") = VGCNx.DefaultDatabase
            .Parameters("@almacen") = "01"
            .Parameters("@articulo") = cCod
            .Parameters("@tipo") = "3"
        End With
        ACMD.Execute
        
        Set ACMD = Nothing
    '*********
         
        adodc1.Requery
        adodc1.AbsolutePosition = nN
    End If
Else
    MsgBox "No existen registros", vbInformation, "Mensaje"
End If
DataGrid1.SetFocus
Exit Sub
EliErr:
    MsgBox Err.Description
    If nTra = 1 Then VGCNx.RollbackTrans
End Sub

Private Sub CmdFicha_Click()            ' Ficha Tecnica
If adodc1.RecordCount > 0 Then
    SSTab1.Tab = 2
    nT = 3
    TxTecnica = ""
    OculObj04 (False)
    OculObj05 (False)
    OculObj06 (False)
    OculObj01 (True)
    OculObj02 (True)
    Cmdgrabar.Visible = True
    cCod = adodc1("ACODIGO")
    cDes = adodc1("ADESCRI")
    
    cSql1 = "Select ACODIGO,Acomenta from MaeART where ACODIGO = '" & cCod & "'"
    Set cSel1 = New ADODB.Recordset
    cSel1.Open cSql1, VGCNx, adOpenStatic
    If cSel1.RecordCount > 0 Then
        If Not IsNull(cSel1("Acomenta")) Then TxTecnica = cSel1("Acomenta")
        Frame2.Caption = "FICHA TECNICA :" & cDes
    Else
        MsgBox "El registro ha sido eliminado", vbInformation, "Mensaje"
        cSel1.Close: CmdSalir2_Click
        Exit Sub
    End If
    cSel1.Close
    TxTecnica.SetFocus
Else
    MsgBox "No existe Registros", vbInformation, "Mensaje"
    CmdSalir2_Click
End If
End Sub

Private Sub CmdGrabar_Click()           ' Grabar
'Dim acmd As New ADODB.Command
On Error GoTo GrabErr

If nT <> 3 Then
    If Trim(TxCodigo) = "" Then
        MsgBox "Ingrese Código", vbInformation, "Mensaje"
        TxCodigo.SetFocus: Exit Sub
    End If
    If Trim(TxDescripcion) = "" Then
        MsgBox "Ingrese Descripcion", vbInformation, "Mensaje"
        TxDescripcion.SetFocus: Exit Sub
    End If
    If Trim(TxUnidad) <> "" Then
        If Existe(1, TxUnidad, "TABUNIMED", "UM_ABREV", False) = False Then
            MsgBox "Ingrese Unidad de Medida", vbInformation, "Mensaje"
            TxUnidad.SetFocus: Exit Sub
        End If
    Else
        MsgBox "Ingrese Unidad de Medida", vbInformation, "Mensaje"
        TxUnidad.SetFocus: Exit Sub
    End If
    'Para que exista grupo tiene que haber lineas y para que haya lineas tiene que haber
    'familias
    
    If Trim(TxGrupo) <> "" And Trim(TxLinea) = "" Then
       MsgBox "Ud. tiene que registrar Linea para asignar Grupo", vbInformation, "Mensaje"
       TxLinea.SetFocus: Exit Sub
    End If
    If Trim(TxLinea) <> "" And Trim(TxFamilia) = "" Then
       MsgBox "Ud. tiene que registrar Familia para asignar Linea", vbInformation, "Mensaje"
       TxFamilia.SetFocus: Exit Sub
    End If
    
    If Trim(TxFamilia) <> "" Then
        If Existe(1, TxFamilia, "FAMILIA", "FAM_CODIGO", False) = False Then
            MsgBox "Codigo de Familia no existe", vbInformation, "Mensaje"
            TxFamilia.SetFocus: Exit Sub
        End If
        If Trim(TxLinea) <> "" Then
             If Existe(1, TxLinea, "LINEAS", "LIN_CODIGO", False, TxFamilia, "FAM_CODIGO") = False Then
                 MsgBox "Codigo de Linea con la Familia señalada no existe", vbInformation, "Mensaje"
                 TxLinea.SetFocus: Exit Sub
              End If
                 If Trim(TxGrupo) <> "" Then
                    If Existe(1, TxGrupo, "GRUPO", "GRU_CODIGO", False, TxFamilia, "FAM_CODIGO", TxLinea, "LIN_CODIGO") = False Then
                       MsgBox "Codigo de Grupo con Familia y Linea señalada no existe ", vbInformation, "Mensaje"
                       TxGrupo.SetFocus: Exit Sub
                    End If
                 End If
         End If
    End If
'    If EstadoAnt = 3 And Not OptSerie.Value Then
'        If existe_serie Then OptSerie.Value = True
'    ElseIf EstadoAnt = 2 And Not OptLote.Value Then
'        If existe_serie Then OptLote.Value = True
'    End If
    
    
End If

If MsgBox("Es correcta la Información", vbQuestion + vbOKCancel, "Mensaje") = vbOK Then
    If Trim(TxPeso) = "" Then TxPeso = 0
    If nT = 1 Then      'Ingreso
        If codigo(TxCodigo) = False Then
            MsgBox "Código de Artículo ya existe", vbInformation, "Mensaje"
            TxCodigo.SetFocus: Exit Sub
        End If
        CSQL2 = "Insert Into MaeArt (ACODIGO,ACODIGO2,ADESCRI,ADESCRI2,AUNIDAD,AFAMILIA,"
        CSQL2 = CSQL2 & "Alinea,AGRUPO,APESO,ATIPO,AUSER,AESTADO,AFSERIE,AFLOTE,ACODMON,AIGVPOR,AFLAGIGV,AMARCA,afecha,ACOLOR,PA,TALLA,AFSTOCK,apcom) VALUES  "
        CSQL2 = CSQL2 & "('" & SupCadSQL(TxCodigo) & "','" & SupCadSQL(TxFabricante) & "','" & SupCadSQL(TxDescripcion) & "','" & SupCadSQL(TxAlterna) & "',"
        CSQL2 = CSQL2 & "'" & SupCadSQL(TxUnidad) & "','" & SupCadSQL(TxFamilia) & "','" & SupCadSQL(TxLinea) & "','" & TxGrupo & "',"
        CSQL2 = CSQL2 & "" & TxPeso & ",'" & TxTipo & "',"
        CSQL2 = CSQL2 & "'" & SupCadSQL(VGUsuario) & "','V','" & IIf(OptSerie.Value, "S", "N") & "', '" & IIf(OptLote.Value, "S", "N") & "',"
        CSQL2 = CSQL2 & "'01',19,0,'" & SupCadSQL(TxMarca) & "',getdate(),'" & SupCadSQL(TxtColor) & "','" & Trim(txPartAran) & "','" & Trim(TxTalla) & "','" & ChkSer.Value & "',"
        CSQL2 = CSQL2 & CheckCom.Value & ")"
        cCod = TxCodigo
    ElseIf nT = 2 Then     'Modificar             Trim(Mid(Combo1.text, 1, 1))
        CSQL2 = "Update MaeArt Set ACODIGO = '" & SupCadSQL(TxCodigo) & "',ACODIGO2 = '" & SupCadSQL(TxFabricante) & "',"
        CSQL2 = CSQL2 & "ADESCRI = '" & SupCadSQL(TxDescripcion) & "',ADESCRI2 = '" & SupCadSQL(TxAlterna) & "',"
        CSQL2 = CSQL2 & "AUNIDAD = '" & SupCadSQL(TxUnidad) & "',AFAMILIA = '" & SupCadSQL(TxFamilia) & "',"
        CSQL2 = CSQL2 & "ALINEA = '" & TxLinea & "',AGRUPO = '" & TxGrupo & "',"
        CSQL2 = CSQL2 & "APESO = " & TxPeso & ","
        CSQL2 = CSQL2 & "ATIPO = '" & SupCadSQL(TxTipo) & "',AUSER = '" & SupCadSQL(VGUsuario) & "',"
        CSQL2 = CSQL2 & "AFSERIE = '" & IIf(OptSerie.Value, "S", "N") & "',"
        CSQL2 = CSQL2 & "AFLOTE = '" & IIf(OptLote.Value, "S", "N") & "',"
        CSQL2 = CSQL2 & "AESTADO = '" & cEstado & "',"
        CSQL2 = CSQL2 & "ACOLOR = '" & SupCadSQL(TxtColor) & "',"
        CSQL2 = CSQL2 & "AMARCA = '" & SupCadSQL(TxMarca) & "',"
        CSQL2 = CSQL2 & "PA= '" & SupCadSQL(txPartAran) & "',"
        CSQL2 = CSQL2 & "TALLA= '" & SupCadSQL(TxTalla) & "',AFSTOCK='" & ChkSer.Value & "',"
        CSQL2 = CSQL2 & "apcom= " & CheckCom.Value & ""
        CSQL2 = CSQL2 & " Where ACODIGO = '" & SupCadSQL(TxCodigo) & "'"
        cCod = SupCadSQL(TxCodigo)
    ElseIf nT = 3 Then      'Ficha Tecnica
        CSQL2 = "Update MaeART set AComenta = '" & SupCadSQL(TxTecnica) & "' "
        CSQL2 = CSQL2 & "Where ACODIGO = '" & SupCadSQL(cCod) & "'"
    End If
    nTra = 1
    'CSQL2 = "Insert Into MaeArt (ACODIGO,ACODIGO2,ADESCRI,ADESCRI2,AUNIDAD,AFAMILIA,ALINEA,AGRUPO,APESO,ATIPO,AUSER,AESTADO,AFECHA,AFSERIE,AFLOTE,ACODMON,AIGVPOR,AFLAGIGV,AMARCA,ACOLOR,PA,TALLA) VALUES  ('001','','JERSEY 20/1 AZUL MARINO','','KGS','00','00','00',50,'','STAR','V','19/12/2002','N', 'S','MN',18,0,'','','','')"
    VGCNx.BeginTrans
    VGCNx.Execute CSQL2
    VGCNx.CommitTrans
    nTra = 0
    adodc1.Requery
    FrameNuevoGrupo.Visible = False
    adodc1.Find "ACODIGO = '" & cCod & "'"
End If

If nT = 1 Then
    Limpiar
    TxCodigo.SetFocus
ElseIf nT = 2 Or nT = 3 Then
    CmdSalir2_Click
End If

Exit Sub
GrabErr:
    MsgBox Err.Description
    If nTra = 1 Then VGCNx.RollbackTrans
    Exit Sub
    Resume
End Sub

Private Sub CmdIng_Click()      'Ingresar
nT = 1
EstadoAnt = 0
Me.Caption = "Ingreso de Articulos"
SSTab1.Tab = 0
OculObj04 (False)
OculObj05 (False)
OculObj06 (False)
OculObj02 (True)
OculObj03 (True)
OculObj01 (False)
Limpiar
'RMM*************************
OptLote.Enabled = True
OptSerie.Enabled = True
OptSerie.Enabled = True
'RMM*************************
muestracrear
ChkSer.Value = 1
TxCodigo.Enabled = True
TxCodigo.SetFocus
End Sub
Private Sub muestracrear()
If VGparametros.tipogeneracioncodigo = 2 And VGparametros.tipocreacioncodigo <> "L" Then
        FrameNuevo.Visible = True
 ElseIf VGparametros.tipogeneracioncodigo = 4 Then
        FrameNuevoGrupo.Visible = True
  Else
 End If
End Sub

Private Sub CmdModi_Click()     'Modificar
If adodc1.RecordCount > 0 Then
    nT = 2
    Me.Caption = "Modificación de Articulos"
    SSTab1.Tab = 0
    OculObj04 (False)
    OculObj05 (False)
    OculObj06 (False)
    OculObj02 (True)
    OculObj03 (True)
    Limpiar
    cCod = adodc1("ACODIGO")
    TxCodigo.Enabled = False
    Cmdgrabar.Visible = True
    Mostrar (cCod)
    If Trim(TxCodigo) <> "" And TxCodigo.Visible Then TxFabricante.SetFocus
Else
    MsgBox "No existen registros", vbInformation, "Mensaje"
End If
End Sub

Private Sub CmdSalir_Click()    'Salida principal del formulario
Unload Me
End Sub

Private Sub CmdSalir2_Click()   'Salida de la segunda pantalla
Me.Caption = "Actualiza Datos Generales de Articulos"
SSTab1.Tab = 1
OculObj01 (False)
OculObj02 (False)
OculObj03 (False)
OculObj04 (True)
OculObj05 (True)
OculObj06 (True)
InhabObj (True)
Cmdgrabar.Enabled = True
Cmdgrabar.Visible = False
DataGrid1.SetFocus
FrameNuevoGrupo.Visible = False
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then TxPeso.SetFocus
End Sub


Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyBack Then
 If Len(TxFiltro) - 1 > 0 Then
  TxFiltro = Left(TxFiltro, Len(TxFiltro) - 1)
 Else
  TxFiltro = ""
 End If
 KeyAscii = 0
ElseIf KeyAscii <> 13 Then
 TxFiltro = TxFiltro & Chr(KeyAscii)
End If
End Sub

Private Sub Form_Activate()
TxFiltro = ""
cmbOrden.ListIndex = 0
If DataGrid1.Visible Then DataGrid1.SetFocus
End Sub

Private Sub Form_Load()
SSTab1.Tab = 1
central Me         'Centra Formulario
'Init_ControlDataGrid DataGrid1
Limpiar
FrameNuevo.Visible = False
FrameNuevoGrupo.Visible = False
OculObj01 (False)
OculObj02 (False)
OculObj03 (False)
OculObj04 (True)
OculObj05 (True)
OculObj06 (True)
Set adodc1 = New ADODB.Recordset
adodc1.Open "Select ACODIGO,ADESCRI,AUNIDAD,ACODIGO2,AFAMILIA,ALINEA,AGRUPO,ATIPO,ACUENTA,AMARCA,AFSTOCK FROM MAEART ORDER BY ACODIGO", VGCNx, adOpenStatic, adLockReadOnly
adodc1.Requery
Set DataGrid1.DataSource = adodc1
DataGrid1.Refresh

Call Ctr_AyuFamilia.conexion(VGCNx)
Call Ctr_ayulinea.conexion(VGCNx)
Call Ctr_Ayugrupo.conexion(VGCNx)

CmdIng.Picture = MDIPrincipal.ImageList2.ListImages.item("Insertar").Picture
CmdModi.Picture = MDIPrincipal.ImageList2.ListImages.item("Modificar").Picture
CmdEli.Picture = MDIPrincipal.ImageList2.ListImages.item("Eliminar").Picture
CmdFicha.Picture = MDIPrincipal.ImageList2.ListImages.item("Nuevo").Picture
Cmdgrabar.Picture = MDIPrincipal.ImageList2.ListImages.item("Grabar").Picture
CmdSalir.Picture = MDIPrincipal.ImageList2.ListImages.item("Retornar").Picture
CmdSalir2.Picture = MDIPrincipal.ImageList2.ListImages.item("Retornar").Picture

End Sub
Private Sub Limpiar()       'Limpia variables
TxCodigo = "": TxFabricante = "": TxDescripcion = ""
TxAlterna = "": TxUnidad = "": LbUnidad = ""
TxFamilia = "": LbFamilia = "": TxLinea = ""
LbLinea = "": TxGrupo = "": LbGrupo = ""
TxPeso = ""
TxMarca = ""
TxtColor = ""
cUsuario = "": Combo1.ListIndex = 0
txPartAran = "": TxTalla = ""
OptLibre = True
TxTipo = ""
txPartAran = ""
CheckCom.Value = 0
ChkSer.Value = 0

End Sub

Private Sub TextFam_DblClick()
Set Adodc2 = New ADODB.Recordset
Adodc2.Open "SELECT FAM_CODIGO,FAM_NOMBRE FROM FAMILIA", VGCNx, adOpenStatic, adLockOptimistic
frmReferencia.Conectar Adodc2, "SELECT FAM_CODIGO,FAM_NOMBRE FROM FAMILIA"
frmReferencia.Label1.Caption = "Familias de Artículos"
frmReferencia.Show vbModal
Adodc2.Close
If vGUtil(1) <> "" Then
   TxFamilia = (vGUtil(1))
   LbFamilia = vGUtil(2)
   Textfam = TxFamilia
   Adodc2.Open "SELECT * FROM FAMILIA where fam_codigo='" & Textfam & "'", VGCNx, adOpenStatic, adLockOptimistic
   TxFcorrelativo.text = IIf(IsNull(Adodc2!correlativocodigo), 1, Adodc2!correlativocodigo)
   TxCodigo.text = Trim(TxFamilia.text) + Right("0000000000" + RTrim(TxFcorrelativo.text), VGparametros.VGLongCodigo - Len(Trim(TxFamilia)))
   Adodc2.Close
   Adodc2.Open "update FAMILIA set correlativocodigo=" & TxFcorrelativo.text + 1 & " where fam_codigo='" & Textfam & "'", VGCNx, adOpenStatic, adLockOptimistic
   FrameNuevo.Visible = False
   SendKeys "{tab}"
End If
End Sub


Public Sub TxFabricante_KeyPress(KeyAscii As Integer) 'CODIGO DEL FABRICANTE
If KeyAscii = 13 Then
    TxDescripcion.SetFocus
Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub

Private Sub TxFiltro_Change()
If adodc1.RecordCount > 0 Then
    If Trim(TxFiltro) <> "" Then
        nCursor = adodc1.Bookmark
        adodc1.AbsolutePosition = 1
        adodc1.MoveFirst
        
        If cmbOrden.ListIndex = 0 Then
            adodc1.Find "ACODIGO like '" & Trim(UCase(TxFiltro)) & "*'"
        ElseIf cmbOrden.ListIndex = 1 Then
            adodc1.Find "ADESCRI like '*" & Trim(UCase(TxFiltro)) & "*'"
        ElseIf cmbOrden.ListIndex = 2 Then
            adodc1.Find "AGRUPO like '" & Trim(UCase(TxFiltro)) & "*' "
        ElseIf cmbOrden.ListIndex = 3 Then
            adodc1.Find "AFAMILIA like '" & Trim(UCase(TxFiltro)) & "*' "
        ElseIf cmbOrden.ListIndex = 4 Then
            adodc1.Find "Alinea like '" & Trim(UCase(TxFiltro)) & "*' "
        ElseIf cmbOrden.ListIndex = 5 Then
            adodc1.Find "ACODIGO2 like '" & Trim(UCase(TxFiltro)) & "*' "
        ElseIf cmbOrden.ListIndex = 6 Then
            adodc1.Find "ATIPO like '" & Trim(UCase(TxFiltro)) & "*' "
        ElseIf cmbOrden.ListIndex = 6 Then
            adodc1.Find "AFSTOCK like '" & Trim(UCase(TxFiltro)) & "*' "
        End If
        If adodc1.EOF Then adodc1.AbsolutePosition = nCursor
    End If
End If
End Sub

Private Sub TxMarca_DblClick()
Set Adodc2 = New ADODB.Recordset
Adodc2.Open "SELECT COD_MARCA,DESCRI_MARCA FROM MAEMARCA", VGCNx, adOpenStatic, adLockOptimistic
frmReferencia.Conectar Adodc2, "SELECT COD_MARCA,DESCRI_MARCA  FROM MAEMARCA"
frmReferencia.Label1.Caption = "Tipo de Clase de Articulo"
frmReferencia.Show vbModal

If vGUtil(1) <> "" Then
  TxMarca = vGUtil(1)
  SendKeys "{tab}"
End If
End Sub

Private Sub TxMarca_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then TxMarca_DblClick
End Sub

Private Sub TxMarca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     If TxMarca <> "" Then
            If Existe(1, TxMarca, "MAEMARCA", "COD_MARCA", False) Then
                    SendKeys "{TAB}"
            Else
                    MsgBox "El Tipo de Clase  Articulo no existe ", vbInformation, "Información"
                    TxMarca.SetFocus
            End If
      Else
            SendKeys "{TAB}"
      End If
End If
End Sub

Private Sub TxTalla_DblClick()
    Set Adodc2 = New ADODB.Recordset
    Adodc2.Open "SELECT CODIGO,DESCRIP FROM TALLA", VGCNx, adOpenStatic, adLockOptimistic
    frmReferencia.Conectar Adodc2, "SELECT CODIGO,DESCRIP FROM TALLA"
    frmReferencia.Label1.Caption = "Lista de Tallas"
    frmReferencia.Show vbModal
    Adodc2.Close
    If vGUtil(1) <> "" Then
      TxTalla = vGUtil(1)
      SendKeys "{tab}"
    End If
End Sub

Private Sub TxTalla_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then TxTalla_DblClick
End Sub

Private Sub TxTalla_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     If TxTalla <> "" Then
            If Existe(1, TxTalla, "TALLA", "CODIGO", False) Then
                    SendKeys "{TAB}"
            Else
                    MsgBox "No se encuentra esta talla ", vbInformation, "Información"
                    TxTalla.SetFocus
            End If
      Else
            SendKeys "{TAB}"
      End If
End If
End Sub

Private Sub TxtColor_DblClick()
Set Adodc2 = New ADODB.Recordset
Adodc2.Open "SELECT COD_COLOR,DESCRI_COLOR FROM MAECOLOR", VGCNx, adOpenStatic, adLockOptimistic
frmReferencia.Conectar Adodc2, "SELECT COD_COLOR,DESCRI_COLOR  FROM MAECOLOR"
frmReferencia.Label1.Caption = "Tipo de Color"
frmReferencia.Show vbModal
Adodc2.Close
If vGUtil(1) <> "" Then
  TxtColor = vGUtil(1)
  SendKeys "{tab}"
End If
End Sub

Private Sub TxtColor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then TxtColor_DblClick
End Sub

Private Sub TxtColor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     If TxtColor <> "" Then
            If Existe(1, TxtColor, "MAECOLOR", "COD_COLOR", False) Then
                    SendKeys "{TAB}"
            Else
                    MsgBox "El Tipo de color no existe ", vbInformation, "Información"
                    TxtColor.SetFocus
            End If
      Else
            SendKeys "{TAB}"
      End If
End If
End Sub

Private Sub TxTipo_DblClick()
Set Adodc2 = New ADODB.Recordset
Adodc2.Open "SELECT COD_TIPO,DES_TIPO FROM TIPO_ARTICULO", VGCNx, adOpenStatic, adLockOptimistic
frmReferencia.Conectar Adodc2, "SELECT COD_TIPO,DES_TIPO  FROM TIPO_ARTICULO"
frmReferencia.Label1.Caption = "Tipo de Artículo"
frmReferencia.Show vbModal
If vGUtil(1) <> "" Then
  TxTipo = vGUtil(1)
  SendKeys "{tab}"
End If
End Sub

Private Sub TxTipo_GotFocus()
Enfoque TxTipo
End Sub

Private Sub TxTipo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then TxTipo_DblClick
End Sub

Private Sub TxTipo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     If TxTipo <> "" Then
            If Existe(1, TxTipo, "TIPO_ARTICULO", "COD_TIPO", False) Then
                    SendKeys "{TAB}"
            Else
                    MsgBox "El Tipo de Articulo no existe ", vbInformation, "Información"
                    TxTipo.SetFocus
            End If
      Else
            SendKeys "{TAB}"
      End If
End If
End Sub

Private Sub Txunidad_DblClick()
Set Adodc2 = New ADODB.Recordset

Adodc2.Open "SELECT UM_ABREV,UM_NOMBRE FROM TABUNIMED", VGCNx, adOpenStatic, adLockOptimistic
frmReferencia.Conectar Adodc2, "SELECT UM_ABREV,UM_NOMBRE FROM TABUNIMED"
frmReferencia.Label1.Caption = "Unidades de Medida"
frmReferencia.Show vbModal
If vGUtil(1) <> "" Then
  TxUnidad = vGUtil(1)
  LbUnidad = vGUtil(2)
  SendKeys "{tab}"
End If
End Sub
Private Sub TxUnidad_GotFocus()
Enfoque TxUnidad
End Sub

Private Sub TxUnidad_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Txunidad_DblClick
ElseIf KeyCode = 46 Or KeyCode = 8 Then
    LbUnidad = ""
End If
End Sub

Private Sub TxUnidad_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   If Trim(TxUnidad) <> "" Then
      If Existe(1, TxUnidad, "TABUNIMED", "UM_ABREV", False) = False Then
         MsgBox "El código de Unidad no existe", vbInformation, "Mensaje"
         TxUnidad.SetFocus
      Else
          LbUnidad = Devolver_Dato(1, TxUnidad, "TABUNIMED", "UM_ABREV", False, "UM_NOMBRE")
          TxFamilia.SetFocus
      End If
   Else
      MsgBox "Ingrese Código de Unidad de Medida", vbInformation, mensaje1
      TxUnidad.SetFocus: Exit Sub
   End If
    
Else
    If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub
Private Sub TxFamilia_DblClick()
Set Adodc2 = New ADODB.Recordset
Adodc2.Open "SELECT FAM_CODIGO,FAM_NOMBRE FROM FAMILIA", VGCNx, adOpenStatic, adLockOptimistic
frmReferencia.Conectar Adodc2, "SELECT FAM_CODIGO,FAM_NOMBRE FROM FAMILIA"
frmReferencia.Label1.Caption = "Familias de Artículos"
frmReferencia.Show vbModal
If vGUtil(1) <> "" Then
  If TxFamilia <> vGUtil(1) Then
    TxLinea = ""
    LbLinea = ""
    TxGrupo = ""
    LbGrupo = ""
  End If
  TxFamilia = (vGUtil(1))
  LbFamilia = vGUtil(2)
  SendKeys "{tab}"
End If
End Sub

Private Sub TxFamilia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    TxFamilia_DblClick
ElseIf KeyCode = 46 Or KeyCode = 8 Then
    LbFamilia = ""
End If
End Sub

Private Sub TxFamilia_GotFocus()
Enfoque TxFamilia
End Sub
Private Sub TxFamilia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(TxFamilia) <> "" Then
       If Existe(1, TxFamilia, "FAMILIA", "FAM_CODIGO", False) = False Then
          MsgBox "El código de Familia no existe", vbInformation, "Mensaje"
          TxFamilia.SetFocus: Exit Sub
       Else
            LbFamilia = Devolver_Dato(1, TxFamilia, "FAMILIA", "FAM_CODIGO", False, "FAM_NOMBRE")
       End If
    Else
        LbFamilia = ""
         MsgBox "Ingrese el código de Familia ", vbInformation, "Mensaje"
         TxFamilia.SetFocus: Exit Sub
    End If
    TxLinea.SetFocus
Else
    If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub
Private Sub TxLinea_DblClick()
Set Adodc2 = New ADODB.Recordset
If TxFamilia <> "" Then
    Adodc2.Open "SELECT LIN_CODIGO,LIN_NOMBRE FROM LINEAS WHERE FAM_CODIGO='" & TxFamilia & "'", VGCNx, adOpenStatic, adLockOptimistic
    frmReferencia.Conectar Adodc2, "SELECT LIN_CODIGO,LIN_NOMBRE FROM LINEAS WHERE FAM_CODIGO='" & TxFamilia & "'"
    frmReferencia.Label1.Caption = "Lineas de Artículos"
    frmReferencia.Show vbModal
    If vGUtil(1) <> "" Then
       If TxLinea <> vGUtil(1) Then
          TxGrupo = ""
          LbGrupo = ""
       End If
       TxLinea = (vGUtil(1))
       LbLinea = vGUtil(2)
       SendKeys "{tab}"
    End If
Else
   MsgBox "Para asignar Linea, Ud. debe asignar Familia", vbInformation, "Mensaje"
   TxFamilia.SetFocus
End If
End Sub
Private Sub TxLinea_GotFocus()
Enfoque TxLinea
End Sub
Private Sub TxLinea_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    TxLinea_DblClick
ElseIf KeyCode = 46 Or KeyCode = 8 Then
    LbLinea = ""
End If
End Sub

Private Sub TxLinea_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(TxLinea) <> "" Then
       If TxFamilia <> "" Then
          If Existe(1, TxLinea, "LINEAS", "LIN_CODIGO", False, TxFamilia, "FAM_CODIGO") = False Then
             MsgBox "El código de Linea con la Familia asignada, no existe", vbInformation, "Mensaje"
             TxLinea.SetFocus: Exit Sub
          End If
       Else
          MsgBox "Para asignar Linea, Ud. debe asignar Familia", vbInformation, "Mensaje"
          TxFamilia.SetFocus: Exit Sub
       End If
       LbLinea = Devolver_Dato(1, TxLinea, "LINEAS", "LIN_CODIGO", False, "LIN_NOMBRE", TxFamilia, "fam_codigo")
    Else
        LbLinea = ""
    End If
    TxGrupo.SetFocus
Else
    If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub

Private Sub TxCodigo_GotFocus()
Enfoque TxCodigo
End Sub
Private Sub TxCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If codigo(TxCodigo) Then
        TxFabricante.SetFocus
    Else
        If Trim(TxCodigo) = "" Then
            MsgBox "Ingrese el Código del articulo", vbInformation, "Mensaje"
        Else
            MsgBox "El Código ya existe", vbInformation, "Mensaje"
        End If
        TxCodigo.SetFocus
    End If
Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub

Private Sub TxGrupo_DblClick()
Set Adodc2 = New ADODB.Recordset

If TxLinea <> "" Then
    If TxFamilia <> "" Then
        Adodc2.Open "SELECT GRU_CODIGO,GRU_NOMBRE FROM GRUPO WHERE FAM_CODIGO='" & TxFamilia & "' AND LIN_CODIGO='" & TxLinea & "'", VGCNx, adOpenStatic, adLockOptimistic
        frmReferencia.Conectar Adodc2, "SELECT GRU_CODIGO,GRU_NOMBRE FROM GRUPO WHERE FAM_CODIGO='" & TxFamilia & "' AND LIN_CODIGO='" & TxLinea & "'"
        frmReferencia.Label1.Caption = "Grupos de Artículos"
        frmReferencia.Show vbModal
        If vGUtil(1) <> "" Then
           TxGrupo = (vGUtil(1))
           LbGrupo = vGUtil(2)
           SendKeys "{tab}"
        End If
    Else
       MsgBox "Para asignar Grupo, Ud. debe asignar Familia y Linea", vbInformation, "Mensaje"
       TxFamilia.SetFocus
   End If
Else
   MsgBox "Para asignar Grupo, Ud. debe asignar Linea", vbInformation, "Mensaje"
   TxLinea.SetFocus
End If
End Sub
Private Sub TxGrupo_GotFocus()
Enfoque TxGrupo
End Sub
Private Sub TxGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    TxGrupo_DblClick
ElseIf KeyCode = 46 Or KeyCode = 8 Then
    LbGrupo = ""
End If

End Sub

Private Sub TxGrupo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(TxGrupo) <> "" Then
       If Trim(TxLinea) <> "" Then
            If TxFamilia <> "" Then
               If Existe(1, TxGrupo, "GRUPO", "GRU_CODIGO", False, TxFamilia, "FAM_CODIGO", TxLinea, "LIN_CODIGO") = False Then
                  MsgBox "El código de Grupo con la Familia y Linea asignados, no existe", vbInformation, "Mensaje"
                  TxGrupo = "": TxGrupo.SetFocus: Exit Sub
               End If
            Else
               MsgBox "Para asignar Grupo, Ud. debe asignar Familia y Linea", vbInformation, "Mensaje"
               TxFamilia.SetFocus: Exit Sub
            End If
        Else
            MsgBox "Para asignar Grupo, Ud. debe asignar Linea", vbInformation, "Mensaje"
            TxLinea.SetFocus: Exit Sub
        End If
        LbGrupo = Devolver_Dato(1, TxGrupo, "GRUPO", "GRU_CODIGO", False, "GRU_NOMBRE", TxFamilia, "Fam_codigo", TxLinea, "Lin_Codigo")
    Else
        LbGrupo = ""
    End If
    SendKeys "{tab}"
    'Combo1.SetFocus
Else
    If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub

Private Sub TxPeso_GotFocus()
Enfoque TxPeso
End Sub
Private Sub TxPeso_KeyPress(KeyAscii As Integer)
Dim I As Integer

If KeyAscii = 13 Then
   Cmdgrabar.SetFocus
Else
  If (Chr$(KeyAscii) > "9" Or Chr$(KeyAscii) < "0") And Chr$(KeyAscii) <> "." And KeyAscii <> 8 Then
     KeyAscii = 0
  Else
     If Chr$(KeyAscii) = "." Then
        For I = 1 To Len(TxPeso)
            If Mid(TxPeso, I, 1) = "." Then KeyAscii = 0: Exit Sub
        Next
     End If
  End If
End If
End Sub
Private Sub TxAlterna_GotFocus()
Enfoque TxAlterna
End Sub
Private Sub TxAlterna_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxUnidad.SetFocus
Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub

Private Sub TxDescripcion_GotFocus()
Enfoque TxDescripcion
End Sub
Private Sub TxDescripcion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxAlterna.SetFocus
Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub

Private Sub Mostrar(cC1 As String) 'Muestra los datos
Dim cSqlM As String, cSelM As ADODB.Recordset
If Trim(cC1) = "" Then
    MsgBox "No hay registros para mostrar", vbInformation, "Mensaje"
    Exit Sub
End If
cSqlM = "Select * From MaeART Where ACODIGO = '" & SupCadSQL(cC1) & "'"
Set cSelM = New ADODB.Recordset
cSelM.Open cSqlM, VGCNx, adOpenStatic
If cSelM.RecordCount > 0 Then
    TxCodigo = cSelM("ACODIGO")
    If Not IsNull(cSelM("ACODIGO2")) Then TxFabricante = cSelM("ACODIGO2")
    If Not IsNull(cSelM("ADESCRI")) Then TxDescripcion = cSelM("ADESCRI")
    If Not IsNull(cSelM("ADESCRI2")) Then TxAlterna = cSelM("ADESCRI2")
    If Not IsNull(cSelM("AMARCA")) Then TxMarca = cSelM("AMARCA")
    
    'Fernando: 31/08/2001:
    If Not IsNull(cSelM("PA")) Then txPartAran = cSelM("PA")
    If Not IsNull(cSelM("TALLA")) Then TxTalla = cSelM("TALLA")
    '***
    If Not IsNull(cSelM("AUNIDAD")) Then TxUnidad = cSelM("AUNIDAD")
    If Not IsNull(cSelM("AFAMILIA")) Then TxFamilia = cSelM("AFAMILIA")
    If Not IsNull(cSelM("ALINEA")) Then TxLinea = cSelM("ALINEA")
    If Not IsNull(cSelM("AGRUPO")) Then TxGrupo = cSelM("AGRUPO")
    If Not IsNull(cSelM("APESO")) Then TxPeso = cSelM("APESO")
    If Not IsNull(cSelM("ACOLOR")) Then TxtColor = cSelM("ACOLOR")
    If Not IsNull(cSelM("ATIPO")) Then TxTipo = cSelM("ATIPO")
    If Not IsNull(cSelM("APCOM")) Then CheckCom.Value = cSelM("APCOM")
    If Not IsNull(cSelM("AFSTOCK")) Then ChkSer.Value = cSelM("AFSTOCK")
    OptLibre = True
    
    'Si el Articulo Tiene Movimientos no debe ser modificado el manejo del Stock '
    '*RMM***********************************************************************
    If ClsTock.ArticuloConMovimiento(TxCodigo, VGCNx) Then
       OptLote.Enabled = False
       OptSerie.Enabled = False
       OptLibre.Enabled = False
    Else
       OptLote.Enabled = True
       OptSerie.Enabled = True
       OptLibre.Enabled = True
    End If
    
    EstadoAnt = 1
    If Not IsNull(cSelM("AFLOTE")) Then
          OptLote.Value = IIf(cSelM("AFLOTE") = "S", True, False)
          EstadoAnt = 2
    End If
     If Not IsNull(cSelM("AFSERIE")) Then
          OptSerie.Value = IIf(cSelM("AFSERIE") = "S", True, False)
          EstadoAnt = 3
    End If
    
    If Not IsNull(cSelM("AUSER")) Then cUsuario = cSelM("AUSER")
    If Not IsNull(cSelM("AESTADO")) Then cEstado = cSelM("AESTADO")
    LbUnidad = Devolver_Dato(1, TxUnidad, "TABUNIMED", "UM_ABREV", False, "UM_NOMBRE")
    LbFamilia = Devolver_Dato(1, TxFamilia, "FAMILIA", "FAM_CODIGO", False, "FAM_NOMBRE")
    LbLinea = Devolver_Dato(1, TxLinea, "LINEAS", "LIN_CODIGO", False, "LIN_NOMBRE", TxFamilia, "FAM_CODIGO")
    LbGrupo = Devolver_Dato(1, TxGrupo, "GRUPO", "GRU_CODIGO", False, "GRU_NOMBRE", TxLinea, "LIN_CODIGO", TxFamilia, "FAM_CODIGO")
Else
    MsgBox "No existe registro", vbInformation, "Mensaje"
    CmdSalir2_Click
End If
cSelM.Close
End Sub

Private Sub InhabObj(ntipo As Boolean) ' Habilita e Inhabilita los objetos
TxCodigo.Enabled = ntipo
TxFabricante.Enabled = ntipo
TxDescripcion.Enabled = ntipo
TxAlterna.Enabled = ntipo
TxUnidad.Enabled = ntipo
TxFamilia.Enabled = ntipo
TxLinea.Enabled = ntipo
TxGrupo.Enabled = ntipo
'Combo1.Enabled = nTipo
TxTipo = ntipo
TxPeso.Enabled = ntipo
End Sub

Private Sub TxFabricante_GotFocus()
Enfoque TxFabricante
End Sub

Function existe_lote() As Boolean
Dim RSQL As String
   existe_lote = False
   RSQL = "select STSLOTE from STKLOTE where   STSCODIGO = '" & TxCodigo & "'"
   Set Adodc3 = New ADODB.Recordset
   Adodc3.Open RSQL, VGCNx, adOpenStatic
   If Not Adodc3.EOF Then
        MsgBox "Lote Registrado en Almacen,   ", vbInformation, "Aviso"
        existe_lote = True
   End If
   Adodc3.Close
End Function


Function existe_serie() As Boolean
Dim RSQL As String
   existe_serie = False
   RSQL = "select STSSERIE from STKSERI where   STSCODIGO = '" & TxCodigo & "'"
   Set Adodc3 = New ADODB.Recordset
   Adodc3.Open RSQL, VGCNx, adOpenStatic
   If Not Adodc3.EOF Then
            MsgBox "Serie Registrada en Almacen,   ", vbInformation, "Aviso"
            existe_serie = True
   End If
   Adodc3.Close
End Function
Private Sub Ctr_AyuFamilia_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Ctr_ayulinea.filtro = " FAM_CODIGO='" & Ctr_AyuFamilia.xclave & "'"
End Sub

Private Sub Ctr_Ayugrupo_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Dim xx As String
xx = ""
TxCodigo = Ctr_AyuFamilia.xclave & Ctr_ayulinea.xclave & Ctr_Ayugrupo.xclave
TxFamilia = Ctr_AyuFamilia.xclave
TxLinea = Ctr_ayulinea.xclave
TxGrupo = Ctr_Ayugrupo.xclave
Lblfam = Ctr_AyuFamilia.xnombre
LbLinea = Ctr_ayulinea.xnombre
LbGrupo = Ctr_Ayugrupo.xnombre

End Sub

Private Sub Ctr_ayulinea_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Ctr_Ayugrupo.filtro = " FAM_CODIGO='" & Ctr_AyuFamilia.xclave & "' and lin_CODIGO='" & Ctr_ayulinea.xclave & "'"
End Sub
