VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Begin VB.Form FrmmntNotas 
   Caption         =   "Form1"
   ClientHeight    =   9465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9465
   ScaleWidth      =   12105
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   975
      Left            =   1800
      TabIndex        =   88
      Top             =   3240
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   1720
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2175
      Left            =   360
      TabIndex        =   87
      Top             =   3120
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   3836
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame4 
      Height          =   2535
      Left            =   120
      TabIndex        =   86
      Top             =   3000
      Width           =   11895
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comentarios"
      Height          =   2550
      Left            =   1560
      TabIndex        =   30
      Top             =   3000
      Visible         =   0   'False
      Width           =   8265
      Begin VB.TextBox Text12 
         Height          =   1935
         Left            =   600
         MultiLine       =   -1  'True
         TabIndex        =   33
         Top             =   360
         Width           =   5655
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Grabar"
         Height          =   495
         Left            =   6600
         TabIndex        =   32
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Cancelar"
         Height          =   495
         Left            =   6600
         TabIndex        =   31
         Top             =   1080
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Height          =   675
      Left            =   210
      TabIndex        =   25
      Top             =   5490
      Width           =   11745
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   10020
         TabIndex        =   27
         Top             =   240
         Width           =   1425
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1410
         TabIndex        =   26
         Top             =   270
         Width           =   1395
      End
      Begin VB.Label Label16 
         Caption         =   "Total  Cantidad"
         Height          =   195
         Index           =   0
         Left            =   8760
         TabIndex        =   29
         Top             =   270
         Width           =   1395
      End
      Begin VB.Label Label16 
         Caption         =   "Total  Items"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   28
         Top             =   300
         Width           =   1395
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   195
      Top             =   6075
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Framedet 
      Height          =   3240
      Left            =   120
      TabIndex        =   42
      Top             =   6120
      Width           =   11895
      Begin VB.TextBox Fechavcto 
         Height          =   285
         Left            =   7335
         TabIndex        =   56
         Top             =   570
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1830
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   570
         Visible         =   0   'False
         Width           =   2100
      End
      Begin VB.TextBox Text6 
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   54
         Top             =   600
         Width           =   2070
      End
      Begin VB.TextBox TxtUniRef 
         Height          =   285
         Left            =   8610
         TabIndex        =   53
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox TxtCantidad 
         Height          =   375
         Left            =   1800
         TabIndex        =   52
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox TxtArticulo 
         Height          =   285
         Left            =   960
         MaxLength       =   20
         TabIndex        =   51
         Top             =   2640
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.CheckBox chkserie 
         Caption         =   "Por Cantidades"
         Height          =   255
         Left            =   4080
         TabIndex        =   50
         Top             =   570
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txEquip 
         Height          =   285
         Left            =   8520
         MaxLength       =   10
         TabIndex        =   49
         Top             =   1920
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.TextBox TxordFab 
         Height          =   285
         Left            =   7530
         MaxLength       =   10
         TabIndex        =   48
         Top             =   1350
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.TextBox txccosto 
         Height          =   375
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   47
         Top             =   1800
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.TextBox txtcanref 
         Height          =   285
         Left            =   4920
         MaxLength       =   10
         TabIndex        =   46
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton CmdDetEnvio 
         Caption         =   "&Enviar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton CmdDetKimpia 
         Caption         =   "&Limpiar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton CmdDetSalir 
         Cancel          =   -1  'True
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
         Height          =   1095
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   2040
         Width           =   1215
      End
      Begin MSMask.MaskEdBox FechaFabric 
         Height          =   285
         Left            =   10380
         TabIndex        =   57
         Top             =   570
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   -2147483634
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   285
         Left            =   7350
         TabIndex        =   58
         Top             =   570
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   -2147483634
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuart 
         Height          =   375
         Left            =   1800
         TabIndex        =   59
         Top             =   240
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   661
         XcodMaxLongitud =   20
         xcodwith        =   1500
         NomTabla        =   "maeart"
         ListaCampos     =   "acodigo(1),adescri(1),acodigo2(2),aunidad(2)"
         XcodCampo       =   "acodigo"
         XListCampo      =   "adescri"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "acodigo,adescri,acodigo2,aunidad"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuAnalitico 
         Height          =   315
         Left            =   4200
         TabIndex        =   60
         Top             =   1680
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   556
         XcodMaxLongitud =   11
         xcodwith        =   900
         NomTabla        =   "v_analiticoentidad"
         TituloAyuda     =   "Busqueda de Centro de Costos"
         ListaCampos     =   "entidadcodigo(1),entidadrazonsocial(1),entidaddireccion(1)"
         XcodCampo       =   "entidadcodigo"
         XListCampo      =   "entidadrazonsocial"
         ListaCamposDescrip=   "Código,Descripción,cliente"
         ListaCamposText =   "entidadcodigo,entidadrazonsocial,entidaddireccion"
         Requerido       =   0   'False
      End
      Begin VB.Label lbcantstk 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbcantstk"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   82
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Nro Serie \ Lote"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   81
         Top             =   630
         Width           =   1140
      End
      Begin VB.Label lbEtiNum 
         Caption         =   "Num de Item:"
         Height          =   255
         Left            =   9270
         TabIndex        =   80
         Top             =   270
         Width           =   975
      End
      Begin VB.Label LblNroReg 
         Caption         =   "LblNroReg"
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
         Left            =   10440
         TabIndex        =   79
         Top             =   270
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label26 
         Caption         =   "Label9"
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
         Left            =   8970
         TabIndex        =   78
         Top             =   2280
         Width           =   1980
      End
      Begin VB.Label LblCantidad 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label8"
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
         Left            =   7380
         TabIndex        =   77
         Top             =   2250
         Width           =   1335
      End
      Begin VB.Label Label25 
         Caption         =   "Label7"
         Height          =   195
         Left            =   6720
         TabIndex        =   76
         Top             =   2310
         Width           =   675
      End
      Begin VB.Label Label24 
         Caption         =   "Unidad referencial"
         Height          =   255
         Left            =   7200
         TabIndex        =   75
         Top             =   990
         Width           =   1335
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad en Stock"
         Height          =   195
         Left            =   240
         TabIndex        =   74
         Top             =   1050
         Width           =   1320
      End
      Begin VB.Label Label22 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label14"
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
         Left            =   6120
         TabIndex        =   73
         Top             =   960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label21 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label13"
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
         Left            =   7440
         TabIndex        =   72
         Top             =   2640
         Visible         =   0   'False
         Width           =   4485
      End
      Begin VB.Label Label20 
         Caption         =   "Fecha Fab."
         Height          =   255
         Left            =   9420
         TabIndex        =   71
         Top             =   600
         Width           =   945
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad"
         Height          =   195
         Left            =   240
         TabIndex        =   70
         Top             =   1350
         Width           =   630
      End
      Begin VB.Label Label18 
         Caption         =   "Fecha Vcto."
         Height          =   255
         Left            =   6300
         TabIndex        =   69
         Top             =   600
         Width           =   1245
      End
      Begin VB.Label Label17 
         Caption         =   "Unidad Estandar"
         Height          =   195
         Left            =   3480
         TabIndex        =   68
         Top             =   1020
         Width           =   1215
      End
      Begin VB.Label LblCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Codigo"
         Height          =   195
         Left            =   240
         TabIndex        =   67
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblUniEst 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblUniEst"
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
         Height          =   255
         Left            =   4890
         TabIndex        =   66
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblordfab 
         Caption         =   "Orden Fabricación"
         Height          =   255
         Left            =   6120
         TabIndex        =   65
         Top             =   1380
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label lblccosto3 
         AutoSize        =   -1  'True
         Caption         =   "Merma"
         Height          =   195
         Left            =   3480
         TabIndex        =   64
         Top             =   1320
         Width           =   480
      End
      Begin VB.Label lblccosto 
         AutoSize        =   -1  'True
         Caption         =   "Centro de Costo"
         Height          =   195
         Left            =   240
         TabIndex        =   63
         Top             =   1890
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Label LblPrecio 
         Height          =   345
         Left            =   1830
         TabIndex        =   62
         Top             =   2550
         Width           =   1725
      End
      Begin VB.Label Lblanalitico 
         AutoSize        =   -1  'True
         Caption         =   "Analitico"
         Height          =   195
         Left            =   3480
         TabIndex        =   61
         Top             =   1725
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Generales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   11850
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5505
         MaxLength       =   11
         TabIndex        =   8
         Top             =   195
         Width           =   1275
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6840
         MaxLength       =   20
         TabIndex        =   7
         Top             =   960
         Width           =   1515
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   6
         Top             =   2160
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Valorizado"
         ForeColor       =   &H80000006&
         Height          =   255
         Left            =   2955
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Cmddetalle 
         Caption         =   "<<      Insertar producto(s)     >>"
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
         Height          =   390
         Left            =   5115
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2220
         Width           =   5295
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   9240
         TabIndex        =   2
         Top             =   960
         Width           =   405
      End
      Begin VB.CheckBox ChkTalla 
         Alignment       =   1  'Right Justify
         Caption         =   "Ingresos por Tallas"
         Height          =   225
         Left            =   8370
         TabIndex        =   1
         Top             =   225
         Visible         =   0   'False
         Width           =   1725
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1365
         TabIndex        =   4
         Top             =   225
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   99614721
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   36161
      End
      Begin TextFer.TxFer TxNdoc 
         Height          =   375
         Left            =   7320
         TabIndex        =   9
         Top             =   1680
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
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
         MaxLength       =   10
         Text            =   ""
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
      End
      Begin TextFer.TxFer TxSerie 
         Height          =   375
         Left            =   6480
         TabIndex        =   10
         Top             =   1680
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
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
         MaxLength       =   4
         Text            =   ""
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuCliente 
         Height          =   375
         Left            =   1320
         TabIndex        =   34
         Top             =   1320
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   661
         Enabled         =   0   'False
         XcodMaxLongitud =   11
         xcodwith        =   1200
         NomTabla        =   "vt_cliente"
         ListaCampos     =   "clientecodigo(1),clienterazonsocial(1)"
         XcodCampo       =   "clientecodigo"
         XListCampo      =   "clienterazonsocial"
         ListaCamposDescrip=   "cliente,Razon social"
         ListaCamposText =   "clientecodigo,clienterazonsocial"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuDocref 
         Height          =   255
         Left            =   1320
         TabIndex        =   35
         Top             =   1800
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   450
         Enabled         =   0   'False
         XcodMaxLongitud =   2
         xcodwith        =   200
         NomTabla        =   "tipo_docu"
         ListaCampos     =   "TDO_TIPDOC(1),TDO_DESCRI(1)"
         XcodCampo       =   "TDO_TIPDOC"
         XListCampo      =   "TDO_DESCRI"
         ListaCamposDescrip=   "Tipo, Descripcion"
         ListaCamposText =   "TDO_TIPDOC,TDO_DESCRI"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuProveedor 
         Height          =   315
         Left            =   1320
         TabIndex        =   83
         Top             =   960
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
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuTransa 
         Height          =   375
         Left            =   1320
         TabIndex        =   84
         Top             =   600
         Width           =   3720
         _ExtentX        =   6562
         _ExtentY        =   661
         XcodMaxLongitud =   0
         xcodwith        =   100
         NomTabla        =   "tabtransa"
         TituloAyuda     =   "Transaciones"
         ListaCampos     =   "tt_codmov(1),tt_descri(1),tt_dr(1),tt_codtrans_auto(1),tt_clie(2),tt_dr(2),intercompanias(1),tt_equip(1)"
         XcodCampo       =   "tt_codmov"
         XListCampo      =   "tt_descri"
         ListaCamposDescrip=   "Codigo,Descripcion,doc.ref.,trans.auto,Ctrl.Cliente,Doc.ref.Proyectos"
         ListaCamposText =   "tt_codmov,tt_descri,tt_dr,tt_codtrans_auto,tt_clie,tt_dr,intercompanias,tt_equip"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuAlmacen 
         Height          =   375
         Left            =   6120
         TabIndex        =   85
         Top             =   600
         Width           =   4080
         _ExtentX        =   7197
         _ExtentY        =   661
         XcodMaxLongitud =   2
         xcodwith        =   200
         NomTabla        =   "tabalm"
         TituloAyuda     =   "Almacenes"
         ListaCampos     =   "TAALMA(1),TADESCRI(1),empresacodigo(1)"
         XcodCampo       =   "TAALMA"
         XListCampo      =   "TADESCRI"
         ListaCamposDescrip=   "Codigo,Descripcion,empresa"
         ListaCamposText =   "TAALMA,TADESCRI,empresacodigo"
      End
      Begin VB.Label LblCC 
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
         Left            =   8610
         TabIndex        =   24
         Top             =   2310
         Width           =   1890
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Doc. :"
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   150
         TabIndex        =   23
         Top             =   285
         Width           =   930
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Transaccion :"
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   150
         TabIndex        =   22
         Top             =   645
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Num. Doc :"
         ForeColor       =   &H80000006&
         Height          =   255
         Left            =   4590
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor :"
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   1035
         Width           =   825
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tip Doc Ref :"
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   150
         TabIndex        =   19
         Top             =   1860
         Width           =   960
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Cod. Cliente :"
         Height          =   195
         Left            =   150
         TabIndex        =   18
         Top             =   1410
         Width           =   945
      End
      Begin VB.Label Label8 
         Caption         =   "Orden Compra"
         ForeColor       =   &H80000006&
         Height          =   210
         Left            =   5595
         TabIndex        =   17
         Top             =   1050
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Autorizacion"
         ForeColor       =   &H80000006&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2280
         Width           =   870
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Almacen :"
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   5220
         TabIndex        =   15
         Top             =   705
         Width           =   705
      End
      Begin VB.Label Label14 
         Caption         =   "Num. Ref"
         ForeColor       =   &H80000006&
         Height          =   255
         Left            =   5610
         TabIndex        =   14
         Top             =   1785
         Width           =   810
      End
      Begin VB.Label lblauto 
         Caption         =   "lblauto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1980
         TabIndex        =   13
         Top             =   2280
         Width           =   1965
      End
      Begin VB.Label Label1 
         Caption         =   "Moneda"
         Height          =   210
         Left            =   8550
         TabIndex        =   12
         Top             =   1035
         Width           =   900
      End
      Begin VB.Label LbltComp 
         Height          =   255
         Left            =   7335
         TabIndex        =   11
         Top             =   225
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.Frame FrameOpccab 
      Height          =   1575
      Left            =   1680
      TabIndex        =   36
      Top             =   6240
      Width           =   8175
      Begin VB.CommandButton Command8 
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
         Height          =   1005
         Left            =   6510
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   270
         Width           =   1155
      End
      Begin VB.CommandButton Command3 
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
         Height          =   1005
         Left            =   3435
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   270
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.CommandButton Command7 
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
         Height          =   1005
         Left            =   4980
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   240
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Modificar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   1890
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   270
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Adicionar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   270
         Visible         =   0   'False
         Width           =   1155
      End
   End
End
Attribute VB_Name = "FrmmntNotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'    vgRegEnt = 0 significa salida
'    vgregent = 1 significa ingreso
'    VGSeleccion = 1 Significa que es seleccion con frame de tipo de cambio
'    VGSeleccion = 2 Significa que es seleccion sin frame de tipo de cambio para modificar el contenido
'    VGSeleccion = 3 Significa que es seleccion sin frame de tipo de cambio para agregar item
'    VGform significa con formulario esta trabajando
'     text9    autorizado
'     text10  cencos
'     Ctr_AyuAlmacen.xclave  almacen
Option Explicit
'Dim db As Database
Dim VGDllGeneral As New dllgeneral.dll_general
Dim nument As Long
Dim precioprom As Double
Dim CANTIDAD As Double
Dim canttemp As Double
Dim Campo As String * 2
Dim contador As Integer
Dim auxdisp As Integer
Dim num As Integer
Dim TT_CONTADOR As Integer
Dim estadocosto As Integer
Dim cadena As String
Dim alma As String
Dim tipo As String * 2
Dim dato As String
Dim empresaorigen As String
Dim NumDoc As String
Dim Codigo2 As String
Dim Comenta  As Boolean
Dim WithEvents Conex As ADODB.Connection
Attribute Conex.VB_VarHelpID = -1
Dim Completo As Boolean
Dim Nimprimir As Integer
Public CENTROCOSTO As Integer
Dim analitico As Integer
Dim WithEvents rsmantenimiento As ADODB.Recordset
Attribute rsmantenimiento.VB_VarHelpID = -1

'***********************************
Dim flaglote As String
Dim flagserie As String
Dim FACTOR As Integer
Dim xserie As String
Dim I As Integer
Dim fin As Integer
Dim graba As Boolean
Dim cant As Double
Dim dato_invalido As Boolean
Dim hubo_error  As Boolean
Dim serie_lote  As String
'*************
Dim rsSTKART As New ADODB.Recordset

Private Sub Cmddetalle_Click()
 Dim rf As New ADODB.Recordset
 Dim contitem As Integer
 contitem = 0
 If CmddetalleOk = 0 Then
    Exit Sub
 End If
     contitem = contitem + 1
     Cmddetalle.Enabled = False
     Check1.Enabled = False
     Text2 = "01"
     muestra
Dim criterio As String
Set rsSTKART = VGCNx.Execute("Select * from STKART WHERE STALMA='" & VGAlma & "'")
Call Ctr_AyuAnalitico.conexion(VGCNx)
Ctr_AyuAnalitico.filtro = " tipoanaliticocodigo='" & VGParamSistem.tipoanaliticocodigo & "' and  isnull(proyectocierre,0)=0 "
Call Ctr_Ayuart.conexion(VGCNx)
CmdDetEnvio.Enabled = False
deshabilitartx5_tx3 (False)
Text6.Enabled = False
LblNroReg.Visible = False
lbEtiNum.Visible = False
VGForm1 = 2
limpia
  
  'revisar cuando viene de modificar
   'VGRegEnt = 1  en cualquier formulario  significa entrada
If VGRegEnt = 1 Then
     Label7.Caption = "Cantidad a Entrar "
Else
     Label7.Caption = "Cantidad a Salir"
     FechaFabric.Visible = False
     Label6.Visible = False
     TxtUniRef.Visible = False
End If

CmdDetEnvio.Picture = MDIPrincipal.ImageList2.ListImages.item("Insertar").Picture
CmdDetKimpia.Picture = MDIPrincipal.ImageList2.ListImages.item("Sacar").Picture
CmdDetSalir.Picture = MDIPrincipal.ImageList2.ListImages.item("Retornar").Picture



 
End Sub
Private Function CmddetalleOk()
CmddetalleOk = 0
 If Ctr_AyuProveedor.Enabled And Ctr_AyuProveedor.Visible And Ctr_AyuProveedor.xclave = "" Then
         MsgBox "falta llenar el Codigo del proveedor", vbExclamation, mensaje1
         Ctr_AyuProveedor.SetFocus
         Exit Function
  End If
  If Ctr_AyuCliente.Enabled And Ctr_AyuCliente.Visible And Ctr_AyuCliente.xclave = "" Then
         MsgBox "falta llenar el Codigo del Cliente", vbExclamation, mensaje1
         Ctr_AyuCliente.SetFocus
         Exit Function
  End If
 
 If Ctr_AyuProveedor.Enabled And Ctr_AyuProveedor.Visible Then
    If TxSerie.text = "" Then
             MsgBox "Digite Numero de serie", vbInformation, "Información"
              TxSerie.SetFocus
              Exit Function
           ElseIf TxNdoc.text = "" Then
             MsgBox "Digite Numero de Numero de Documento", vbInformation, "Información"
            TxNdoc.SetFocus: Exit Function
     End If
 End If
 If Trim(Text9) <> "" Then
     If Trim(validarautorizado(Text9)) = "" Then
        MsgBox "El Autorizado no existe", vbInformation, "Información"
        Text9.SetFocus: Exit Function
     End If
 Else
     If Text9.Enabled And Text9.Visible Then
         MsgBox "Falta llenar el Codigo del Autorizado", vbExclamation, mensaje1
         Text9.SetFocus
         Exit Function
     End If
 End If
 If Ctr_AyuAlmacen.Enabled And Ctr_AyuAlmacen.Visible And Ctr_AyuAlmacen.xclave = "" Then
         MsgBox "falta llenar el Codigo del almacen", vbExclamation, mensaje1
         Ctr_AyuAlmacen.SetFocus
         Exit Function
 End If
 If Ctr_AyuDocref.Enabled And Ctr_AyuDocref.Visible And Ctr_AyuDocref.xclave = "" Then
         MsgBox "falta llenar Tipo de documento de referencia", vbExclamation, mensaje1
         Ctr_AyuDocref.SetFocus
         Exit Function
 End If
 CmddetalleOk = 1
 End Function

'Ingreso
Private Sub Command1_Click()
If Check1.Value = 0 Then
   VGSeleccion = 1
   buscar_trans
Else
   If DataGrid.RecordCount = 1 Then
      VGValnuevo = True
      VGSeleccion = 1
   Else
      VGSeleccion = 3
   End If
   FormCreacion.Caption = "Ingreso del Articulo"
   FormCreacion.Show 1
End If
End Sub

Private Sub Command2_Click()
If DataGrid.RecordCount = 1 Then
    MsgBox "No hay registros para Modificar", vbInformation, "Información"
    Exit Sub
End If
VGSeleccion = 2
If Check1.Value = 0 Then
    buscar_trans
    FrmCreacionSin.Caption = "Modificación del Detalle"
    FrmCreacionSin.Show 1
Else
    FormCreacion.Caption = "Modificación del Detalle"
    FormCreacion.Show 1
End If
End Sub
'Eliminar
Private Sub Command3_Click()
Dim I As Integer

If DataGrid.RecordCount = 1 Then
    MsgBox "No hay registros para Eliminar", vbInformation, "Información"
    Exit Sub
End If
If MsgBox("Desea Eliminar el Registro?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
    
    I = DataGrid.RecordCountel
    If DataGrid.RecordCount > 2 Then
        MSFlexGrid1.RemoveItem I
    Else
        MSFlexGrid1.Clear
        DataGrid.RecordCount = 1
        MSFlexGrid1.Row = 0
        inicializaFG
        Command7.SetFocus
    End If
End If
End Sub

Private Function ValidaCmddetalle()
 Dim contitem As Integer
 contitem = 0
 ValidaCmddetalle = False
 If Ctr_AyuProveedor.xclave.Enabled And Ctr_AyuProveedor.xclave.Visible Then
         MsgBox "falta llenar el Codigo del proveedor", vbExclamation, mensaje1
         Enfoque Ctr_AyuProveedor.xclave
         Exit Function
     End If
 End If
If Ctr_AyuCliente.Enabled And Ctr_AyuCliente.Visible Then
         MsgBox "falta llenar el Codigo del Cliente", vbExclamation, mensaje1
         Ctr_AyuCliente.SelStart = 0: Ctr_AyuCliente.SelLength = Len(Ctr_AyuCliente)
         Ctr_AyuCliente.SetFocus
         Exit Function
     End If
 End If
 If Ctr_AyuProveedor.xclave.Enabled And Ctr_AyuProveedor.xclave.Visible Then
   If Ctr_AyuDocref.xclave = "" And VGRegEnt = 1 Then
     If Ctr_AyuDocref.xclave = "" Then
        MsgBox "Digite Tipo de Documento", vbInformation, "Información"
        If Ctr_AyuDocref.Enabled = True Then Ctr_AyuDocref.SetFocus: Exit Function
      ElseIf TxSerie.text = "" Then
             MsgBox "Digite Numero de serie", vbInformation, "Información"
              TxSerie.SetFocus: Exit Function
           ElseIf TxNdoc.text = "" Then
             MsgBox "Digite Numero de Numero de Documento", vbInformation, "Información"
            TxNdoc.SetFocus: Exit Function
     End If
   End If
 End If
 If Text9.Enabled And Text9.Visible Then
         MsgBox "Falta llenar el Codigo del Autorizado", vbExclamation, mensaje1
         Text9.SetFocus
         Exit Function
     End If
 End If
If Ctr_AyuAlmacen.xclave.Enabled And Ctr_AyuAlmacen.xclave.Visible Then
         MsgBox "falta llenar el Codigo del almacen", vbExclamation, mensaje1
         Ctr_AyuAlmacen.xclave.SetFocus
     End If
 End If
 contitem = contitem + 1
 Cmddetalle.Enabled = truee
     Check1.Enabled = False
     Text2 = "01"
     muestra
Text1.text = Format(TxSerie.text, "0000") + Format(TxNdoc.text, "0000000000")
End Function

Private Sub Command4_Click()
' GRABA EL COMENTARIO DE LA GUIA
 Dim RSQL As String
 Dim rpta As String
 On Error GoTo Err
 RSQL = "Update MovAlmCab set CAGLOSA = '" & Text12 & "' "
 RSQL = RSQL & "Where CAALMA = '" & VGAlma & "'AND  CATD= '" & tipo & "' AND CANUMDOC = '" & Trim(Text4) & "'" '
 VGCNx.Execute RSQL
 Frame2.Visible = False
 crtlvisible (True)
' inicializar
 rpta = MsgBox("¿Desea Imprimir?", vbYesNo + vbQuestion, "Aviso")
 If rpta = vbYes Then
    imprimir
 End If
 inicializar
 inicializaFG
 Exit Sub
Err:
   MsgBox Err.Description
End Sub

Private Sub command5_Click()
' CANCELA EL COMENTARIO
  Dim rpta As Integer
  Frame2.Visible = False
  crtlvisible (True)
  inicializar
  rpta = MsgBox("¿Desea Imprimir?", vbYesNo + vbQuestion, "Aviso")
  If rpta = vbYes Then
     imprimir
  End If
End Sub
'****************************** Graba la NI ,NS ****************
Private Sub Command7_Click()
 ' Dim adodc2 As ADODB.Recordset
  Dim Data2 As New ADODB.Recordset
  Dim criterio As String
  Dim cadena As String
  Dim cadena1 As String
  Dim cadena2 As String
  Dim rpta As Integer
  Dim merma As Integer
    Dim FACTOR As Double
  Dim uSql As String
  On Error GoTo GrabErr
  
    
   CANTIDAD = 0
   If DataGrid.RecordCount = 1 Then
     MsgBox "No se puede grabar,debe adicionar registro", vbInformation, mensaje1
     Exit Sub
   End If
   If Not IsNumeric(Text4) Then
     MsgBox "Numero de Documento no consecutivo", vbExclamation, "Aviso"
     Exit Sub
   End If
   Text4 = Format(Text4, String(11, "0"))
' cambio en el control de numero de documentos
   Dim J As Integer
   Dim vgxregent As Integer
   Dim xdato As String
   Dim xtipo As String
   Dim X As Boolean
Set Data2 = Nothing
Set Data2 = Nothing
Data2.Open "movalmdet", VGCNx, adOpenDynamic, adLockOptimistic
  J = 0
  vgxregent = VGRegEnt
  xdato = dato
  xtipo = tipo
  For J = 1 To 2
  Nimprimir = 0
  If J = 2 Then
    VGRegEnt = 2
    dato = "S"
    tipo = "NS"
    Ctr_AyuTransa.xclave = "90"
'   Else
'    Exit For
  End If
   If J = 1 Then X = existe_numdoc(Text4, tipo)
'   Screen.MousePointer = 11
   If J = 1 Then grabacabecera
   FACTOR = 1    ' factor de conversion
   contador = 1  ' Contador de item
   'graba detalle
   NumDoc = Text4
   merma = 0
   While DataGrid.RecordCount > contador
     If (IIf(VGRegEnt = 1, True, True)) Then      'verificastk
       cadena = MSFlexGrid1.TextMatrix(contador, 0)
       CANTIDAD = 0
       If Not VGActualizar Then
              If J = 1 Or J = 2 And Val(MSFlexGrid1.TextMatrix(contador, 14)) <> 0 Then
                 Data2.AddNew
                 If J = 2 Then X = existe_numdoc(Text4, tipo)
                    NumDoc = Text4
                 If J = 2 Then grabacabecera
              End If
       Else
              criterio = "DECODIGO = '" & UCase(cadena) & "'"
              criterio = criterio + " and  DEALMA = '" & Ctr_AyuAlmacen.xclave & "'"  ' VGAlma & "'"
              Data2.Find criterio
              If Data2.RecordCount = 0 Then
                MsgBox " No encontrado...!!!", vbInformation, "AVISO"
                Data2.Close
                Set Data2 = Nothing
                Exit Sub
              End If
       End If
      If J = 1 Or J = 2 And Val(MSFlexGrid1.TextMatrix(contador, 14)) <> 0 Then
         Data2("DEALMA") = Ctr_AyuAlmacen.xclave    'VGAlma
         Data2("DETD") = tipo ' "NS ,NI"
         Data2("DENUMDOC") = Text4.text
         Data2("DEITEM") = contador
         Data2("DECODIGO") = UCase(MSFlexGrid1.TextMatrix(contador, 0))   ' Format(MSFlexGrid1.TextMatrix(contador, 0), "00000000")
         Data2("DEDESCRI") = MSFlexGrid1.TextMatrix(contador, 1) 'Antes no se debe grababa se consulta a MAEART
         If J = 1 Then
            CANTIDAD = Val(MSFlexGrid1.TextMatrix(contador, 6))
          Else
            CANTIDAD = Val(MSFlexGrid1.TextMatrix(contador, 14))
         End If
         Data2("DECANTID") = CANTIDAD
         Data2("DECODMON") = Text2  'antes no se graba en detalle se consultaba a la cabecera
         Data2("DEUNIDAD") = MSFlexGrid1.TextMatrix(contador, 4) 'Antes no se debe grababa se consulta a MAEART
         Data2("DECANREF1") = "" & IIf(MSFlexGrid1.TextMatrix(contador, 14) = "", 0, MSFlexGrid1.TextMatrix(contador, 14))
         If MSFlexGrid1.TextMatrix(contador, 3) <> "" Then
            grabastk
            If MSFlexGrid1.TextMatrix(contador, 5) <> "" Then    'si tiene precio de costo
                Data2("DEPRECIO") = Val(MSFlexGrid1.TextMatrix(contador, 7)) ' * VGTipCamb '******el precio
                Data2("DETIPCAM") = MSFlexGrid1.TextMatrix(contador, 15) 'DevolverTCambio(DTPicker1.Value)
            ElseIf (estadocosto = 1 And VGRegEnt = 0) Or Text10.Visible Then  'SALIDA VALORIZADA  0 - SALIDA,1 - ENTRADA, text10 indica salida x C
                Data2("DEPRECIO") = precioprom  '******'valorizacion de precio prom
            Else
                Data2("DEPRECIO") = 0
            End If
            Data2("DECENCOS") = MSFlexGrid1.TextMatrix(contador, 11)
            Data2("DEORDFAB") = MSFlexGrid1.TextMatrix(contador, 12)
            Data2("DEQUIPO") = MSFlexGrid1.TextMatrix(contador, 13)
            alma = Ctr_AyuAlmacen.xclave '' VGAlma  'indica el almacen
            'mejorar a una funcion
            If MSFlexGrid1.TextMatrix(contador, 10) = "S" Then
                grabaserie alma, cadena
                Data2("DESERIE") = MSFlexGrid1.TextMatrix(contador, 2)
            End If
            If MSFlexGrid1.TextMatrix(contador, 10) = "N" Then
                grabalote alma, cadena
                Data2("DELOTE") = MSFlexGrid1.TextMatrix(contador, 2)
            End If
         End If
         Data2.Update
       End If
     End If
     contador = contador + 1
   Wend
   'data2.Refresh
   
   Dim cad As String
   If Ctr_AyuAlmacen.xclave <> "" And (TxTransa = "TD" Or TxTransa = "SD") Then
     contador = 1
     While DataGrid.RecordCount > contador
        CANTIDAD = Val(MSFlexGrid1.TextMatrix(contador, 6))
        cad = insertar1
        Completo = False
        Conex.BeginTrans
        Conex.Execute cad
        Conex.CommitTrans
        Do
          DoEvents
        Loop Until Completo
        
        grabastk1                'graba en la tabla stk del otro almacen
        alma = Ctr_AyuAlmacen.xclave          'codigo del almacen
        tipo = "NI"                'cuando se realiza otra traansaccion
        If MSFlexGrid1.TextMatrix(contador, 10) = "S" Then grabaserie Ctr_AyuAlmacen.xclave, cadena
        If MSFlexGrid1.TextMatrix(contador, 10) = "N" Then grabalote Ctr_AyuAlmacen.xclave, cadena
        tipo = "NS"
        contador = contador + 1
     Wend
   End If
   
  'Activa el menu en las opciones reporte y consulta
  If Comenta And Nimprimir = 1 Then
     rpta = MsgBox("Desea Agregar Comentarios", vbYesNo + vbQuestion, "Aviso")
  Else
     rpta = vbNo
  End If
  If rpta = vbYes Then
     crtlvisible (False)
     Frame2.Visible = True
     Text12.SetFocus
  Else
   '  TxTransa.Enabled = True
     If Nimprimir = 1 Then
        rpta = MsgBox("¿Desea Imprimir?", vbYesNo + vbQuestion, "Aviso")
        If rpta = vbYes Then
           imprimir
        End If
     End If
 End If
Next
VGRegEnt = vgxregent
dato = xdato
tipo = xtipo
inicializar
inicializaFG
 VGSoles = True
 VGTipCamb = 1
 Screen.MousePointer = 1
 Exit Sub
GrabErr:
 'Resume
 MsgBox Err.Description, vbExclamation, "Error"
'Resume
 Screen.MousePointer = 1
 Exit Sub
 Resume
End Sub

Private Sub Command8_Click()
'*********************************** SALIR
Dim rpta As Integer
   If DataGrid.RecordCount > 1 Then
     rpta = MsgBox("Desea Grabar", vbYesNo + vbQuestion, "Aviso")
     If rpta = vbYes Then
       Command7_Click
     End If
   End If
   VGval = False
   Ctr_AyuProveedor.Enabled = True
   Text8.Visible = True
   Label8.Visible = True
   Text8.Enabled = True
   Check1.Enabled = True
   VGForm = 5
   Unload Me
End Sub

Private Sub Check1_Click()
   VGval = True   'Para toda la valorizacion'
   VGValnuevo = True   'Para la pantalla de inicio'
   VGForm = 1
   SendKeys "{tab}"
End Sub

Private Sub Conex_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
  Completo = True
End Sub

Private Sub Ctr_ayuAlmacen_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Dim Adodc3 As New ADODB.Recordset
VGAlma = Trim(Ctr_AyuAlmacen.xclave)
Set Adodc3 = Nothing
Set Adodc3 = VGCNx.Execute("select * from tabalm where taalma='" & VGAlma & "'")
  If Adodc3.RecordCount > 0 Then
    If VGRegEnt = 1 Then
      Text4 = Format(Adodc3!tanument, "00000000000")
    Else
      Text4 = Format(Adodc3!tanumsal, "00000000000")
    End If
    empresaorigen = Adodc3!empresacodigo
  End If
End Sub

Private Sub Ctr_AyuCliente_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
If analitico = 1 Then
   SQL = " clientecodigo='" & Ctr_AyuCliente & "' and proyectocierre=0 and tipoanaliticocodigo='" & VGParamSistem.tipoanaliticocodigo & "'"
   Set acliente = VGCNx.Execute(" select * from gr_proyectos where " & SQL)
   If acliente.RecordCount = 0 Then
      MsgBox ("No existe proyectos activos para este cliente ")
      Text5.SetFocus
      FrmCreacionSin.Ctr_AyuAnalitico.Visible = False
      Exit Sub
    Else
      FrmCreacionSin.Ctr_AyuAnalitico.filtro = SQL

    End If
End If

End Sub

Private Sub Ctr_AyuTransa_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    buscar_trans

End Sub

Private Sub DTPicker1_Change()
If DTPicker1.Value > VGParamSistem.fechatrabajo Then
      MsgBox "Fecha de documento mayor a fecha de trabajo", vbInformation, "Mensaje"
      DTPicker1.SetFocus
End If
DTPicker1.Value = UltimoCierreFech(DTPicker1.Value)
VGTipCamb = DevolverTCambio(DTPicker1.Value)
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   SendKeys "{tab}"
End If
End Sub

Private Sub Form_Activate()
   Dim J, kTotal As Double
   VGtipocreacion = 1
   If DataGrid.RecordCount > 1 Then
      Text5 = Format(DataGrid.RecordCount, "##,###,##0.00")
      kTotal = 0
      For J = 1 To DataGrid.RecordCount
        kTotal = kTotal + DataGrid1.Columns(3) '  CDbl(MSFlexGrid1.TextMatrix(J, 3))
      Next
      Text3 = Format(kTotal, "##,###,##0.00")
   Else
      Text5 = Format(0, "##,###,##0.00")
      Text3 = Format(0, "##,###,##0.00")
   End If
   If VGAutomatico Then
     Text4.Enabled = False
   End If
End Sub

Private Sub Form_Load()
   Dim rs As New ADODB.Recordset
   Dim clsmovimientos As New ClasMovimientos
   Dim RSQL As String
   Dim numsal As String
   DoEvents
   Call ctr_ayudas
   FrameOpccab.Visible = True
   Framedet.Visible = False
   
    VGSeleccion = 1               'Indica el modo de apertura = 1 y modificacion=2
    VGtipocreacion = 1
    VGActualizar = False
    VGSoles = True
    VGForm = 5
    LIMPIACABECERA
    DTPicker1.MaxDate = VGParamSistem.fechatrabajo
    DTPicker1.Value = UltimoCierreFech(CDate(Format(VGParamSistem.fechatrabajo, "dd/MM/yyyy")))
    VGTipCamb = DevolverTCambio(DTPicker1.Value)
    
    RSQL = "select  TANUMENT, TANUMSAL from TabAlm  "
    Set rs = VGCNx.Execute(RSQL)
    If rs.RecordCount() = 0 Then
       MsgBox ("No existe registro del almacen en tabla de almacenes")
       GoTo salir
    End If
    nument = IIf(IsNull(rs(0)), 1, rs(0))
    numsal = IIf(IsNull(rs(1)), 1, rs(1))
    VGCNx.Execute RSQL
    
    If VGRegEnt = 1 Then
      Text4.text = Format(Val(nument), "00000000000")
      Me.Caption = "Registro de Entrada"
      dato = "I"
      tipo = "NI"
      Codigo2 = "NOTA DE INGRESO"
      Text2.Visible = True
      ChkTalla.Caption = "Ingreso por Tallas"
      ocultarlabel
      
    Else
       ChkTalla.Caption = "Salida por Tallas"
       Me.Caption = "Registro de Salida"
       dato = "S"
       tipo = "NS"
       Text2.Visible = False
       Label1.Visible = False
       Codigo2 = "NOTA DE SALIDA"
       Check1.Visible = False
       Text4.text = Format(Val(numsal) + 1, "00000000000")
    End If
    VGval = False
    habilitado (False)
    Set rsmantenimiento = New ADODB.Recordset
    Call clsmovimientos.CreaRsTempDetalle(rsmantenimiento)
    rsmantenimiento.Open
    Text4.Enabled = False
    
    Command1.Picture = MDIPrincipal.ImageList2.ListImages("Adicionar").Picture
    Command2.Picture = MDIPrincipal.ImageList2.ListImages("Modificar").Picture
    Command3.Picture = MDIPrincipal.ImageList2.ListImages("Eliminar").Picture
    Command7.Picture = MDIPrincipal.ImageList2.ListImages("Grabar").Picture
    Command8.Picture = MDIPrincipal.ImageList2.ListImages("Retornar").Picture
    
    Exit Sub
salir:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        VGTipCamb = DevolverTCambio(VG_FecTrab)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 '************************** NUM REF
  If Ctr_AyuCliente = "" And KeyAscii = 13 Then
     SendKeys "{tab}"
     KeyAscii = 0
     Exit Sub
  End If
 
 If KeyAscii = 13 And Text1.text <> "" Then
    If Not IsNumeric(Text1) And (Ctr_AyuDocref.xclave = "BV") Then
       MsgBox "Ingrese el Numero de  la Boleta", vbOKOnly, "Aviso"
       Exit Sub
    End If
    If Ctr_AyuDocref.xclave = "FT" And Check1.Value = 1 Then
       FormCreacion.Ctr_AyuDocref.xclave = Text1
       FormCreacion.Ctr_AyuDocref.Enabled = False
    End If
       If Text8.Enabled Then
             Text8.SetFocus
       ElseIf Ctr_AyuCliente.Enabled Then
             Ctr_AyuCliente.SetFocus
       ElseIf Text9.Enabled Then
             Text9.SetFocus
       ElseIf Text10.Enabled Then
             Text10.SetFocus
       ElseIf Ctr_AyuAlmacen.xclave.Enabled Then
             Ctr_AyuAlmacen.xclave.SetFocus
       Else
             Cmddetalle_Click
       End If
 Else
    If Ctr_AyuDocref.xclave = "BV" Then
        If (Chr$(KeyAscii) > "9" Or Chr$(KeyAscii) < "0") And KeyAscii <> 8 Then KeyAscii = 0
    End If
 End If
 Set Conex = New ADODB.Connection
 
End Sub

Private Sub Text10_DblClick()
  Dim Adodc3 As ADODB.Recordset   'Centro de Costos
  Set Adodc3 = New ADODB.Recordset
  If UCase(Dir$(cRuta4)) = UCase(cNomBd4) Then
        Adodc3.Open "SELECT cencost_codigo,cencost_descripcion FROM centro_costos where  len(cencost_codigo) = '6' ", VGcnxCT, adOpenStatic, adLockOptimistic
  Else
        Adodc3.Open "SELECT cencost_codigo,cencost_descripcion FROM centro_costos ", VGCNx, adOpenStatic, adLockOptimistic
  End If
  
        frmReferencia.Conectar Adodc3, "SELECT cencost_codigo,cencost_descripcion FROM centro_costos  "
        frmReferencia.Label1.Caption = "Centro de Costos"
        frmReferencia.Show vbModal
        Adodc3.Close
        If vGUtil(1) <> "" Then
                 Text10 = vGUtil(1)
                 LblCC = vGUtil(2)
        End If
        If Text10 <> "" Then Text10_KeyPress (13)
 
End Sub

Private Sub Text10_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
   Text10_DblClick
End If
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
'**********************CENTRO COSTO
If KeyAscii = 13 And Text10.text <> "" Then
  If Trim(Text10.text) <> "" Then
     If Existe(1, Text10, "CENTRO_COSTOS", "cencost_codigo", False) = False Then
              MsgBox "Centro de Costo no existe", vbInformation, "Mensaje"
             Text10.SetFocus: Exit Sub
     End If
     If Ctr_AyuAlmacen.xclave.Enabled Then
          Ctr_AyuAlmacen.xclave.SetFocus
      Else
          Tabula (KeyAscii)
          'Cmddetalle_Click
      End If
   Else
      MsgBox "Ingrese el numero de Centro de Costo", vbInformation, mensaje1
      Text10.SetFocus
   End If
End If
End Sub



Private Sub Text9_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 112 Then
       Text9_DblClick
   End If
End Sub

Private Sub TxNdoc_GotFocus()
Call TxNdoc_KeyPress(13)
End Sub

Private Sub TxNdoc_KeyPress(KeyAscii As Integer)

Dim RSQL As New ADODB.Recordset
If KeyAscii = 13 And Len(TxNdoc.text) > 0 And VGRegEnt = 1 Then
   If IsNumeric(RTrim(TxNdoc.text)) Then TxNdoc.text = Right("0000000000" & RTrim(TxNdoc.text), TxNdoc.MaxLength)
   SQL = " select * from movalmcab where empresacodigo='" & VGParametros.empresacodigo & "' and CACODPRO ='" & RTrim(Ctr_AyuProveedor.xclave) & " '"
   SQL = SQL & " and carftdoc='" & Ctr_AyuDocref.xclave & "' and carfndoc='" & Format(TxSerie.text, "0000") & Format(TxNdoc.text, "0000000000") & "'"
   Set RSQL = VGCNx.Execute(SQL)
   If RSQL.RecordCount > 0 Then
      MsgBox (" Ya existe ingresado Numero de documento ")
      TxNdoc.SetFocus
      Exit Sub
   End If
End If
End Sub

Private Sub TxSerie_LostFocus()
If IsNumeric(RTrim(TxSerie.text)) Then TxSerie.text = Format(TxSerie.text, "0000")
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If TxTransa.Enabled = False And KeyAscii = 13 And Text4 <> "" Then
   Text4 = Format(Text4, "00000000000")
   If Command7.Visible = True Then
     Command7.SetFocus
   End If
End If
End Sub

Private Sub siguiente_tx5()
          If Ctr_AyuDocref.Enabled Then
             Ctr_AyuDocref.SetFocus
          ElseIf Text8.Enabled Then
             Text8.SetFocus
          ElseIf Ctr_AyuCliente.Enabled Then
             Ctr_AyuCliente.SetFocus
          ElseIf Text9.Enabled Then
             Text9.SetFocus
          ElseIf Text10.Enabled Then
             Text10.SetFocus
          ElseIf Ctr_AyuAlmacen.xclave.Enabled Then
             Ctr_AyuAlmacen.xclave.SetFocus
          Else
              Cmddetalle_Click
          End If
End Sub

Private Sub siguiente_tx7()
   'lblClie = Mid(lblClie, 1, 10)
   lblClie = lblClie.Caption
   If Ctr_AyuCliente <> "" Then
          If Text9.Enabled And Text9.Visible And Trim(Text9) = "" Then
             Text9.SetFocus
          ElseIf Text10.Enabled And Text10.Visible And Trim(10) = "" Then
             Text10.SetFocus
          ElseIf Ctr_AyuAlmacen.xclave.Enabled And Ctr_AyuAlmacen.xclave.Visible Then
             Ctr_AyuAlmacen.xclave.SetFocus
          Else
              Cmddetalle_Click
          End If
   End If
End Sub
 '***** Orden de compra
Private Sub Text8_KeyPress(KeyAscii As Integer)
  Dim criterio As String
  If KeyAscii = 13 Then
        Text8 = Trim(Text8)
        If Text8 <> "" Then
            criterio = "CANUMORD = '" & Text8.text & "' AND  CACODPRO ='" & Ctr_AyuProveedor.xclave & "'"
'            Data1.Recordset.FindFirst criterio
            If VGDllGeneral.VerificaDatoExistente(VGCNx, "select * from movalmcab where " & criterio) = 1 Then
              MsgBox "La Orden de Compra ya fue registrada !", vbExclamation, mensaje1
              Exit Sub
            End If
        End If
        If Ctr_AyuCliente.Enabled And Ctr_AyuCliente.Visible Then
           Ctr_AyuCliente.SetFocus
        ElseIf Text9.Enabled And Text9.Visible Then
           Text9.SetFocus
        End If
End If
  
End Sub

Private Sub ocultarlabel()
    Label7.Visible = False
    Ctr_AyuCliente.Visible = False
    Label9.Visible = False
    Text9.Visible = False
    Label11.Visible = False
    Ctr_AyuAlmacen.Visible = False
End Sub

Private Sub Text9_DblClick()
  FormAyuda.Show 1
  If Text10.Enabled And Text10 <> "" Then
        Text10.SetFocus
  ElseIf Ctr_AyuAlmacen.xclave.Enabled And Ctr_AyuAlmacen.xclave <> "" Then
        Ctr_AyuAlmacen.xclave.SetFocus
  ElseIf TxSerie.text <> "" Then
         TxNdoc.SetFocus
  Else
        SendKeys "{tab}"
        
  End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
      If KeyAscii = 13 Then          'Autorizado
            If Trim(Text9) <> "" Then
                    If Trim(validarautorizado(Text9)) = "" Then
                            MsgBox "No existe el Autorizado", vbInformation, "Mensaje"
                            If Text9.Enabled And Text9.Visible Then Text9.SetFocus
                            Exit Sub
                    End If
                    lblauto = Mid(validarautorizado(Text9), 1, 10)
                    SendKeys "{tab}"
            ElseIf Ctr_AyuAlmacen.xclave.Enabled And Ctr_AyuAlmacen.xclave.Visible Then
                    Ctr_AyuAlmacen.xclave.SetFocus
            Else
                    Cmddetalle.SetFocus
            End If
       End If
End Sub

Private Sub muestra()
     Dim numfil As Integer
    ' Dim nument As Long
     Dim numsal As String
     Dim rs As New ADODB.Recordset
     Dim RSQL As String
    
     If Trim(Ctr_AyuAlmacen.xclave) <> "" Then
        VGAlma = Ctr_AyuAlmacen.xclave
        RSQL = "select  TANUMENT, TANUMSAL from TabAlm  WHERE TAALMA='" & Ctr_AyuAlmacen.xclave & "'"  ' VGAlma & "' "
        'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
        Set rs = VGCNx.Execute(RSQL)
        
        Command1.Visible = True
        Command2.Visible = True
        Command3.Visible = True
        Command7.Visible = True
        If Check1.Value = 0 Then
            VGSeleccion = 1
            buscar_trans
            'Fernando: 06/09/2001:
            '***
         Else
            VGSeleccion = 1
            FormCreacion.Caption = "Ingreso del Detalle"
            FormCreacion.Show 1
        End If
     Else
        MsgBox "No ningún Almacen Activo", vbInformation, "Información"
     End If
     Framedet.Visible = True
     FrameOpccab.Visible = False
End Sub

Public Function insertar1() As String

  Dim cad As String
  If MSFlexGrid1.TextMatrix(contador, 10) = "S" Then
          cad = "insert into MovAlmDet (DEALMA,DETD,DENUMDOC,DECODIGO,DECANTID,DEPRECIO,DEITEM,DESERIE,DECODMON,DECENCOS,DEORDFAB,DEQUIPO) values ('" & Ctr_AyuAlmacen.xclave & "','" & Campo & "','" & Format(nument, "0000000000") & "','" & UCase(MSFlexGrid1.TextMatrix(contador, 0)) & "'," & CANTIDAD & "," & Val(precioprom) & "," & contador & ",'" & MSFlexGrid1.TextMatrix(contador, 2) & "','" & Text2 & "','" & MSFlexGrid1.TextMatrix(contador, 11) & "','" & MSFlexGrid1.TextMatrix(contador, 12) & "','" & MSFlexGrid1.TextMatrix(contador, 13) & "') "
  ElseIf MSFlexGrid1.TextMatrix(contador, 10) = "N" Then
          cad = "insert into MovAlmDet (DEALMA,DETD,DENUMDOC,DECODIGO,DECANTID,DEPRECIO,DEITEM,DELOTE,DECODMON,DECENCOS,DEORDFAB,DEQUIPO) values ('" & Ctr_AyuAlmacen.xclave & "','" & Campo & "','" & Format(nument, "0000000000") & "','" & UCase(MSFlexGrid1.TextMatrix(contador, 0)) & "'," & CANTIDAD & "," & Val(precioprom) & "," & contador & ",'" & MSFlexGrid1.TextMatrix(contador, 2) & "','" & Text2 & "','" & MSFlexGrid1.TextMatrix(contador, 11) & "','" & MSFlexGrid1.TextMatrix(contador, 12) & "','" & MSFlexGrid1.TextMatrix(contador, 13) & "')"
  Else
          cad = "insert into MovAlmDet (DEALMA,DETD,DENUMDOC,DECODIGO,DECANTID,DEPRECIO,DEITEM,DECODMON,DECENCOS,DEORDFAB,DEQUIPO) values ('" & Ctr_AyuAlmacen.xclave & "','" & Campo & "','" & Format(nument, "0000000000") & "','" & UCase(MSFlexGrid1.TextMatrix(contador, 0)) & "'," & CANTIDAD & "," & Val(precioprom) & "," & contador & ",'" & Text2 & "','" & MSFlexGrid1.TextMatrix(contador, 11) & "','" & MSFlexGrid1.TextMatrix(contador, 12) & "','" & MSFlexGrid1.TextMatrix(contador, 13) & "') "
  End If
  insertar1 = cad
End Function

Public Sub grabaalmacen()
 'proceso para una transferencia
  Dim uSql As String
  Dim insertar1 As String
  Dim Adodc3 As ADODB.Recordset
  Set Adodc3 = New ADODB.Recordset
  
  Adodc3.Open "select  TANUMENT from tabAlm where TAALMA =  '" & Ctr_AyuAlmacen.xclave & " '", VGCNx, adOpenStatic, adLockOptimistic
  'Set rS = db.OpenRecordset(rSql, dbOpenSnapshot)
  If Adodc3.EOF Then
     MsgBox "No se ha declarado la numeracion para el almacen destino", vbInformation, "Aviso"
     Adodc3.Close
     Exit Sub
  End If
  nument = Adodc3(0) + 1
  Campo = "NI" 'verifica que el numero sea consecutivo
     
     Set Adodc3 = New ADODB.Recordset
     Adodc3.Open "SELECT  CANUMDOC from MOVALMCAB where CAALMA ='" & Ctr_AyuAlmacen.xclave & "' AND  CATD = '" & Campo & "' and CANUMDOC =  '" & Format(nument, "0000000000") & "' ", VGCNx, adOpenStatic, adLockOptimistic
     If Not Adodc3.EOF Then
       Set Adodc3 = New ADODB.Recordset
       Adodc3.Open "SELECT MAX (CANUMDOC) from MOVALMCAB where CAALMA ='" & Ctr_AyuAlmacen.xclave & "' AND  CATD = '" & Campo & "' ", VGCNx, adOpenStatic, adLockOptimistic
       nument = Adodc3(0) + 1
     End If
     Adodc3.Close
    
  insertar1 = "insert into MovAlmCab (CAALMA,CATD,CANUMDOC,CACODMOV,CAFECDOC,CATIPMOV,CASITGUI,CARFTDOC,CARFNDOC,CARFALMA,CAHORA,CACODPRO,CANOMPRO,CACODCLI,CANOMCLI,CACODMON) "
  insertar1 = insertar1 & " values ('" & Ctr_AyuAlmacen.xclave & "','" & Campo & "','" & Format(nument, "0000000000") & "','03','" & DTPicker1 & "','I','V','NS','" & Text4 & "','01','" & Time & "','" & SupCadSQL(Trim(UCase$(CtrAyu_Proveedor.xclave))) & "','" & SupCadSQL(LTrim(Ctr_AyuProveedor.xclave)) & "','" & SupCadSQL(Mid$(UCase$(Ctr_AyuCliente.xclave), 1, 11)) & "','" & SupCadSQL(LTrim(lblClie.Caption)) & "','" & Text2 & "')"
  VGCNx.Execute insertar1
  uSql = "Update TabAlm set TANUMENT = " & nument & " where TAALMA='" & Ctr_AyuAlmacen.xclave & "' "
  VGCNx.Execute uSql
 
    
End Sub

Public Sub grabastk()
  Dim ACMD As New ADODB.Command
  Dim cadena As String
  Dim criterio As String
  Dim entrada As Boolean
  On Error GoTo GrabErr
   
cadena = MSFlexGrid1.TextMatrix(contador, 0)
Set rsSTKART = New ADODB.Recordset
rsSTKART.Open "Select * from STKART ", VGCNx, adOpenDynamic, adLockOptimistic
criterio = " STCODIGO = '" & cadena & "' and  STALMA ='" & Ctr_AyuAlmacen.xclave & "'"
rsSTKART.Filter = criterio

If Not rsSTKART.EOF Then      'si existe el articulo
  
                canttemp = IIf(IsNull(rsSTKART("STSKDIS")), 0, rsSTKART("STSKDIS"))  ' revisar si validar en creacion
                rsSTKART("STKFECULT") = DTPicker1.Value
                If VGRegEnt = 1 Then
                    If LbltComp.Caption = 1 Then
                        rsSTKART("STSKCOM") = rsSTKART("STSKCOM") - CANTIDAD
                    Else
                        rsSTKART("STSKDIS") = rsSTKART("STSKDIS") + CANTIDAD
                    End If
                   'aqui actualiza
                   If Not IsNull(rsSTKART("STKPREPRO")) Then
                      precioprom = rsSTKART("STKPREPRO")
                      If MSFlexGrid1.TextMatrix(contador, 5) <> "" Then
                         rsSTKART("STKPREULT") = Val(MSFlexGrid1.TextMatrix(contador, 7)) * VGTipCamb 'el precio
                         If VGval And (canttemp + CANTIDAD) <> 0 Then
                          'valorizaAnte                          'valorizaActual                                                  saldoActu
                            rsSTKART("STKPREPRO") = Round(((precioprom * canttemp) + CANTIDAD * Val(Val(MSFlexGrid1.TextMatrix(contador, 7)) * VGTipCamb)) / (canttemp + CANTIDAD), 6)
                         End If
                      End If
                    Else
                      precioprom = 0
                      If MSFlexGrid1.TextMatrix(contador, 5) <> "" Then
                         rsSTKART("STKPREPRO") = Round(Val(MSFlexGrid1.TextMatrix(contador, 7)) * VGTipCamb, 6) 'el precio
                         If VGval Then
                            rsSTKART("STKPREULT") = Round(Val(MSFlexGrid1.TextMatrix(contador, 7)) * VGTipCamb) 'el precio
                            rsSTKART("STKFECULT") = DTPicker1.Value
                         End If
                      End If
                   End If
                Else
                  'para la salida
                   rsSTKART("STSKDIS") = rsSTKART("STSKDIS") - CANTIDAD
                   'aqui actualiza
                   If Not IsNull(rsSTKART("STKPREPRO")) Then
                      precioprom = Round(rsSTKART("STKPREPRO"), 6)
                    Else
                      precioprom = 0
                   End If
               End If
       Else
            rsSTKART.AddNew                   'existe
            rsSTKART("STALMA") = Ctr_AyuAlmacen.xclave    'VGAlma   '"01"
            rsSTKART("STCODIGO") = MSFlexGrid1.TextMatrix(contador, 0)
            rsSTKART("STKFECULT") = DTPicker1.Value
            If VGRegEnt Then
                rsSTKART("STSKDIS") = CANTIDAD
                rsSTKART("STKPREULT") = Val(MSFlexGrid1.TextMatrix(contador, 7)) * VGTipCamb    'el costo de ingreso
                If MSFlexGrid1.TextMatrix(contador, 5) <> "" Then
                      rsSTKART("STKPREPRO") = Round(Val(MSFlexGrid1.TextMatrix(contador, 7)) * VGTipCamb, 6) '******el  costo = costo prom
               End If
            End If
          'Grabamos en Facturacion
          Set ACMD.ActiveConnection = VGGeneral
          ACMD.CommandText = "al_actualizaproducto_pro"
          ACMD.CommandType = adCmdStoredProc
          ACMD.Prepared = True
          With ACMD
            .Parameters("@baseini") = VGCNx.DefaultDatabase
            .Parameters("@basefin") = VGCNx.DefaultDatabase
            .Parameters("@almacen") = VGAlma
            .Parameters("@articulo") = MSFlexGrid1.TextMatrix(contador, 0)
            .Parameters("@tipo") = "1"
         End With
         ACMD.Execute
         Set ACMD = Nothing
         entrada = IIf(VGRegEnt = 1, True, False)
         Call ValMes(VGAlma, entrada) 'para la valorizacion
 End If
 rsSTKART.Update
 rsSTKART.Close
 Exit Sub
GrabErr:
 MsgBox Err.Description
 Exit Sub
 Resume
End Sub

Public Sub grabastk1()
   Dim criterio As String
   Dim cadena As String
   Dim ACMD As New ADODB.Command
   
   On Error GoTo GrabErr
   cadena = MSFlexGrid1.TextMatrix(contador, 0)
   criterio = " STCODIGO ='" & cadena & "' and  STALMA ='" & Ctr_AyuAlmacen.xclave & "'"
   rsSTKART.Filter = criterio
   If rsSTKART.EOF Then
     rsSTKART.AddNew
     rsSTKART("STSKDIS") = CANTIDAD
     rsSTKART("STKPREPRO") = Round(precioprom, 6)
     rsSTKART("STALMA") = Ctr_AyuAlmacen.xclave  '"01"
     rsSTKART("STCODIGO") = MSFlexGrid1.TextMatrix(contador, 0)
     
      Set ACMD.ActiveConnection = VGCNx
       ACMD.CommandText = "al_actualizaproducto_pro"
        ACMD.CommandType = adCmdStoredProc
        ACMD.Prepared = True
        With ACMD
            .Parameters("@baseini") = VGCNx.DefaultDatabase
            .Parameters("@basefin") = VGBase2
            .Parameters("@almacen") = Ctr_AyuAlmacen.xclave
            .Parameters("@articulo") = MSFlexGrid1.TextMatrix(contador, 0)
            .Parameters("@tipo") = "1"
        End With
        ACMD.Execute
        Set ACMD = Nothing
   Else
     
     auxdisp = rsSTKART("STSKDIS")
     If rsSTKART("STKPREPRO") <> 0 And (canttemp + auxdisp) <> 0 Then 'no se registrado algun precio
       rsSTKART("STKPREPRO") = Round((precioprom * canttemp + auxdisp * rsSTKART("STKPREPRO")) / (canttemp + auxdisp), 6)
       rsSTKART("STKFECULT") = DTPicker1.Value
       rsSTKART("stkultfechacompra") = DTPicker1.Value
     End If
      rsSTKART("STSKDIS") = rsSTKART("STSKDIS") + CANTIDAD
   End If
   rsSTKART.Update
'   Data3.Refresh
   Call ValMes(Ctr_AyuAlmacen.xclave, True)  'para la valorizacion
   Exit Sub
GrabErr:
    MsgBox Err.Description
    'Resume
End Sub

Public Sub buscar_trans()
  Dim criterio As String
  Dim rs As New ADODB.Recordset
  Dim RSQL As String
  analitico = 0
   On Error GoTo GrabErrR

    'Busco la transaccion
    RSQL = "select * from TabTransa where TT_CODMOV ='" & Ctr_AyuTransa.xclave & "' and TT_TIPMOV ='" & dato & "'"
    Set rs = VGCNx.Execute(RSQL)
    If rs.RecordCount = 0 Then
       MsgBox "El tipo de transaccion no existe !", vbOKOnly, "Error"
       LIMPIACABECERA
       habilitado (False)
       Exit Sub
    End If
    habilitado (True)
    If Not IsNull(rs("TT_CONT")) Then
       TT_CONTADOR = rs("TT_CONT")
    Else
       MsgBox "El tipo de transacción no esta inicializara !" & Chr(13) & "Para inicializarla ir a la tabla de Transacción", vbOKOnly + vbExclamation, "Error"
       habilitado (False)
       Exit Sub
    End If
    estadocosto = ESNULO(rs("estadocosto"), 0)
    If rs("TT_PRV") = "N" Then
       Ctr_AyuProveedor.Visible = False
     Else
      Ctr_AyuProveedor.Visible = True
    End If
    If rs("tt_alma") = "S" Then
    Ctr_AyuAlmacen.Visible = True
      Label11.Visible = True
    End If
    If rs("TT_DR") = "N" Then
       Ctr_AyuDocref.Visible = False
     Else
       Ctr_AyuDocref.Enabled = True
    End If
    
    If rs("TT_AT") = "N" Then
       Text9.Enabled = False
       Label9.Visible = False
       Text9.Visible = False
    Else
       Label9.Visible = True
       Text9.Visible = True
       Text9.Enabled = True
    End If
    CENTROCOSTO = 0
    If rs("TT_CC") = "N" Then
       Check1.Enabled = True
    Else
       CENTROCOSTO = 1

       Check1.Enabled = False
       Check1.Value = 0
    End If
        
    If rs("TT_OC") = "N" Then
       Text8.Visible = False
       Label8.Visible = False
     Else
       Text8.Visible = True
       Label8.Visible = True
    End If
    If rs("TT_CLIE") = "S" Then
        Label8.Visible = False
        Text8.Visible = False
        Label7.Visible = True
        Ctr_AyuCliente.Visible = True
        Ctr_AyuCliente.Enabled = True
   Else
        Label8.Visible = True
        Text8.Visible = True
        Ctr_AyuCliente.Enabled = False
        Label7.Visible = False
        Ctr_AyuCliente.Visible = False
   End If
'*RMM*************************
   If rs("TT_ORDFAB") = "S" Then
      lblordfab.Visible = True
      TxordFab.Visible = True
   Else
      lblordfab.Visible = False
      TxordFab.Visible = False
   End If
   
   If rs("TT_EQUIP") = "S" Then
      analitico = 1
      Ctr_AyuAnalitico.Visible = True
   Else
      Ctr_AyuAnalitico.Visible = False
  End If
  If rs("ingresosfuturos") = "S" Then
            LbltComp.Caption = 1
   Else
      LbltComp.Caption = 0
  End If
     
'*RMM*************************
   Comenta = IIf(rs("TT_CO") = "S", True, False)
   Cmddetalle.Enabled = True
   Exit Sub
GrabErrR:
 MsgBox Err.Description, vbInformation, "Aviso"
 Exit Sub
 Resume

End Sub

Private Sub grabacabecera()
  Dim criterio As String
  Dim cadena As String
  Dim FACTOR As Double
  Dim uSql As String
  Dim Data1 As New ADODB.Recordset
  VGCNx.BeginTrans
  Set Data1 = Nothing
  Data1.Open "movalmcab", VGCNx, adOpenDynamic, adLockOptimistic
   On Error GoTo GrabErr
  'Desea grabar el registro
   If Text4.text <> "" Then
      VGAlma = "" & Trim(Ctr_AyuAlmacen.xclave)
      If Not VGActualizar Then
         Data1.AddNew
         Data1("empresacodigo") = VGParametros.empresacodigo
         Data1("CAALMA") = VGAlma
         Data1("CANUMDOC") = Mid$(UCase$(Text4.text), 1, 12)
      Else
         criterio = " CANUMDOC ='" & Text4 & "'"
         criterio = criterio + " and  CAALMA ='" & VGAlma & "'"
         Data1.Find criterio
      End If
      Data1("CATIPMOV") = dato
      Data1("CATD") = tipo
      Data1("CAHORA") = Format(Time, "hh:mm:ss")
      Data1("CAFECDOC") = DTPicker1.Value            ' CDate(Text2.text)
      Data1("CACOTIZA") = IIf(Len(Trim(tx_ordfab)) = 0, " ", tx_ordfab)
      
      If Trim(Text1.text) <> "" Then
         Data1("CARFNDOC") = SupCadSQL(Trim(Text1.text))
      Else
         Data1("CARFNDOC") = " "
      End If
      If Ctr_AyuTransa.xclave <> "" Then
         Data1("CACODMOV") = SupCadSQL(Mid$(UCase$(Ctr_AyuTransa.xclave), 1, 2))
      Else
         Data1("CACODMOV") = " "
      End If
      Text4 = Trim(UCase$(Text4.text))
      Data1("CANUMDOC") = Text4
      If Trim(CtrAyu_Proveedor.xclave) <> "" Then
         Data1("CACODPRO") = SupCadSQL(Trim(UCase$(Ctr_AyuProveedor.xclave)))
         Data1("CANOMPRO") = SupCadSQL(LTrim(Ctr_AyuProveedor.xnombre))
      Else
         Data1("CACODPRO") = " "
      End If
      Data1("CAFECACT") = Now
      If Trim(Ctr_AyuDocref.xclave) <> "" Then
         Data1("CARFTDOC") = SupCadSQL(Mid$(UCase$(Ctr_AyuDocref.xclave), 1, 2))
      Else
         Data1("CARFTDOC") = " "
      End If
      If Ctr_AyuCliente.Visible And Ctr_AyuCliente.xclave <> "" Then
         Data1("CACODCLI") = SupCadSQL(Mid$(UCase$(Ctr_AyuCliente.xclave), 1, 11))
      Else
         Data1("CACODCLI") = " "
      End If
      
     If Trim(Text8.text) <> "" And VGRegEnt = 1 Then
         Data1("CANUMORD") = SupCadSQL(Trim(UCase$(Text8.text)))
      Else
         Data1("CANUMORD") = " "
      End If
      If Text9.Visible And Trim(Text9) <> "" Then
         Data1("CASOLI") = Mid$(UCase$(Text9.text), 1, 3)
      Else
         Data1("CASOLI") = " "
      End If
      Data1("CAUSUARI") = UCase(VGUsuario)
      If Text10.Visible And Trim(Text10.text) <> "" Then
         Data1("CACENCOS") = Text10.text
      Else
         Data1("CACENCOS") = " "
      End If
      If Ctr_AyuAlmacen.xclave.Visible And Trim(Ctr_AyuAlmacen.xclave) <> "" Then
         Data1("CARFALMA") = Mid$(UCase$(Ctr_AyuAlmacen.xclave), 1, 2)
      Else
         Data1("CARFALMA") = " "
      End If
      Data1("CACODMON") = Text2
      'Data1.Recordset("CATIPCAM") = VGTipCamb
      Data1("CATIPCAM") = DevolverTCambio(DTPicker1.Value)
      VGCodMon = Text2
      Data1("CASITGUI") = "V"
      'Data1.Recordset("CASITUA") = "V"
      Data1("CAESTIMP") = "V"
      Data1("empresacodigo") = empresaorigen
      Data1.Update
   End If
   Data1.Close
   VGCNx.CommitTrans
   Nimprimir = 1
   Exit Sub
GrabErr:
       MsgBox Err.Description
       VGCNx.RollbackTrans
       Exit Sub
       Resume
End Sub
Function ValidarDoc(txt As TextBox) As String
  
  Dim rs As New ADODB.Recordset
  Dim RSQL As String
  
    RSQL = "select TDO_DESCRI  from TIPO_DOCU  where TDO_TIPDOC='" & SupCadSQL(txt.text) & "'"
  '  Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
    Set rs = VGCNx.Execute(RSQL)
    If rs.EOF Then
       MsgBox "El tipo de codigo no existe !", vbOKOnly, "Error"
       ValidarDoc = ""
       txt.SetFocus
       Exit Function
    End If
    ValidarDoc = rs(0)
    rs.Close

End Function

Function transa(text As TextBox) As String
 Dim rs As Recordset
 Dim RSQL As String
  RSQL = "select  TT_DESCRI FROM TabTransa where TT_CODMOV= '" & text & "' and TT_TIPMOV ='" & dato & "'" '

  Set rs = VGCNx.Execute(RSQL)
  If Not rs.EOF Then
    transa = rs(0)
  Else
    MsgBox "El tipo de codigo no existe !", vbOKOnly, "Error"
    transa = ""
  End If
   rs.Close
End Function
Function tipref(text As TextBox) As String
 Dim rs As Recordset
 Dim RSQL As String
  RSQL = "select  TDO_DESCRI FROM TIPO_DOCU where TDO_TIPDOC= '" & text & "'" '
  'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
  Set rs = VGCNx.Execute(RSQL)
  If Not rs.EOF Then
    tipref = rs(0)
  Else
    MsgBox "El tipo de codigo no existe !", vbOKOnly, "Error"
    tipref = ""
  End If
  rs.Close
End Function

Function prove(txt As TextBox) As String
 Dim rs As New ADODB.Recordset
 Dim RSQL As String
   RSQL = "select clienterazonsocial as PRVCNOMBRE FROM cp_proveedor where clientecodigo= '" & SupCadSQL(txt.text) & "'" '

   Set rs = VGCNx.Execute(RSQL)
   If Not rs.EOF Then
     prove = rs(0)
   Else
     MsgBox "El codigo del proveedor no existe !", vbExclamation, "Error"
     prove = ""
  End If
  rs.Close
End Function

Private Sub LIMPIACABECERA()
   Ctr_AyuProveedor.xclave = ""
   Ctr_AyuDocref.xclave = ""
   Ctr_AyuCliente.xclave = ""
   Text8 = ""
   Text9 = ""
   Ctr_AyuAlmacen.xclave = ""
   lblauto = ""
   LblCC = ""
   Text2 = ""
End Sub

Private Sub habilitado(bol As Boolean)
   Ctr_AyuProveedor.Enabled = bol
   Ctr_AyuDocref.Enabled = bol
    
   Text8.Enabled = bol
   Ctr_AyuCliente.Enabled = bol
   Text9.Enabled = bol

   Ctr_AyuAlmacen.Enabled = bol

End Sub
Private Sub inicializar()

  Ctr_AyuTransa.xclave = ""
  Text4.text = ""
  Check1.Value = 0
'  TxTransa.Enabled = True
  ocultarlabel
  Text12 = ""
  MSFlexGrid1.Clear
  DataGrid.RecordCount = 1
 ' inicializaFG
  Command1.Visible = False
  Command2.Visible = False
  Command3.Visible = False
  Command7.Visible = False
  'inicializar
  If Ctr_AyuDocref.text = "F" Then
    FormCreacion.Ctr_AyuDocref.Enabled = True
    FormCreacion.Ctr_AyuDocref.text = ""
  End If
  habilitado (True)
  LIMPIACABECERA
  habilitado (False)
  VGval = False
  Check1.Enabled = True
  Cmddetalle.Enabled = True
 
End Sub


Private Sub ValMes(almacen As String, entrada As Boolean)
  Dim cadena As String
  Dim criterio As String
  Dim adoreg As ADODB.Recordset
  Dim RSQL As String
  Dim Cantent As Double
  Dim Cantsal As Double
  Dim mespro As String
  Dim uSql As String
  On Error GoTo Err
   mespro = Year(DTPicker1) & Format(Month(DTPicker1), "00")
   cadena = MSFlexGrid1.TextMatrix(contador, 0) 'codigo del art
   RSQL = "select SMCANENT,SMCANSAL FROM MoResMes where  SMALMA = '" & almacen & "' and SMMESPRO = '" & mespro & "' AND  SMCODIGO= '" & cadena & "'"  '
   Set adoreg = New ADODB.Recordset
   adoreg.Open RSQL, VGCNx, adOpenDynamic, adLockOptimistic
    If Not adoreg.EOF Then 'existe
      If entrada Then
        Cantent = adoreg(0) + CANTIDAD
        uSql = "Update MoResMes set SMCANENT = " & Cantent & " where SMALMA='" & almacen & "'  and  SMCODIGO ='" & cadena & "' AND SMMESPRO ='" & mespro & "' "
       Else
        Cantsal = adoreg(1) + CANTIDAD
        uSql = "Update MoResMes set SMCANSAL = " & Cantsal & " where SMALMA='" & almacen & "' and   SMCODIGO ='" & cadena & "' AND SMMESPRO ='" & mespro & "' "
       End If
   Else
      If entrada Then
        Cantent = CANTIDAD
        Cantsal = 0
      Else
        Cantsal = CANTIDAD
        Cantent = 0
      End If
       uSql = "insert into MoResMes (SMALMA,SMCODIGO,SMMESPRO,SMCANENT,SMCANSAL,SMMNPREUNI,SMUSPREUNI,SMSALDOINI) VALUES ('" & almacen & "','" & cadena & "','" & mespro & "' ," & Cantent & "," & Cantsal & "," & Val(cNull(rsSTKART("STKPREPRO"))) & ",0,0) "

   End If
   VGCNx.Execute uSql
  Exit Sub
Err:
   MsgBox Err.Description
   
End Sub

Private Sub crtlvisible(dato As Boolean)
   MSFlexGrid1.Visible = dato
   Command1.Visible = dato
   Command2.Visible = dato
   Command3.Visible = dato
   Command7.Visible = dato
   Command8.Visible = dato

End Sub

Private Sub grabaserie(alma As String, codigo As String)
Dim uSql As String
Dim Serie As String
Dim valor As Integer
Dim rs As Recordset
Dim RSQL As String
Dim fecfab As Date
Dim fecven As Date
    Serie = MSFlexGrid1.TextMatrix(contador, 2)
    RSQL = "select STSSKDIS FROM STKSERI where   STSALMA= '" & alma & "' and  STSCODIGO= '" & codigo & "' and STSSERIE= '" & Serie & "'" '
    
    Set rs = VGCNx.Execute(RSQL)
    If Not rs.EOF Then
       valor = IIf(tipo = "NI", 1, 0)
       uSql = "update STKSERI set STSSKDIS = " & valor & " WHERE  STSALMA='" & alma & "' and STSCODIGO='" & codigo & "'AND STSSERIE='" & Serie & "'"
    Else
       uSql = "insert into STKSERI (STSALMA,STSCODIGO,STSSERIE,STSSKDIS)   VALUES ('" & alma & "','" & codigo & "','" & Serie & "',1) "
    End If
    rs.Close
    
    Set rs = Nothing
    VGCNx.Execute uSql
       
End Sub
Function existe_numdoc(text As TextBox, stipo As String) As Boolean
Dim numsal As String
Dim rs As New ADODB.Recordset
Dim RSQL As String
VGAlma = Ctr_AyuAlmacen.xclave
If Trim(Ctr_AyuAlmacen.xclave) <> "" Then
VGCNx.BeginTrans

   rs.Open "select  TANUMENT, TANUMSAL from TabAlm  WHERE TAALMA='" & Ctr_AyuAlmacen.xclave & "'", VGCNx, adOpenDynamic, adLockOptimistic

   nument = IIf(IsNull(rs(0)), 1, rs(0))
   numsal = IIf(IsNull(rs(1)), 1, rs(1))
   If VGRegEnt = 1 Then
      Text4.text = Format(nument, "00000000000")
      rs("tanument") = rs("tanument") + 1
      nument = Text4.text
    Else
      Text4.text = Format(numsal, "00000000000")
      rs("tanumsal") = rs("tanumsal") + 1
      numsal = Text4.text
   End If
   rs.Update
   rs.Close
VGCNx.CommitTrans
End If
existe_numdoc = False
End Function
Function existe_ordcom(text As TextBox) As Boolean
Dim criterio As String
Dim RSQ As New ADODB.Recordset

 If Text8 <> "" And Ctr_AyuProveedor.xclave <> "" Then
    criterio = "CANUMORD = '" & Text8.text & "' AND  CACODPROV ='" & CtrAyu_Proveedor.xclave & "'"

    Set RSQ = VGCNx.Execute("select * from movalmcab where " & criterio)
    If RSQ.RecordCount > 0 Then
        MsgBox "El Numero documento ya ha sido registrado !", vbExclamation, "Error"
        existe_ordcom = True
        Exit Function
    End If
  End If
  existe_ordcom = False
End Function
Function existe_almacen(text As TextBox) As String
  Dim RSQL As String
  Dim rs As New ADODB.Recordset
  
   RSQL = "SELECT TADESCRI FROM TabAlm where  TAALMA= '" & text & "' and empresacodigo='" & VGParametros.empresacodigo & "' and puntovtacodigo='" & VGParametros.puntovta & "'"
   'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
   Set rs = VGCNx.Execute(RSQL)
   If Not rs.EOF Then 'existe
     existe_almacen = rs(0)
   Else
     MsgBox "El codigo del almacen no existe !", vbOKOnly + vbInformation, "Error"
     existe_almacen = ""
   End If
   rs.Close
End Function

Function existe_clie(text As TextBox) As String
  Dim RSQL As String
  Dim rs As New ADODB.Recordset
  RSQL = "SELECT CNOMCLI FROM maecli where CCODCLI= '" & Trim(text) & "'"
  Set rs = VGCNx.Execute(RSQL)
  If rs.RecordCount > 0 Then 'existe
     existe_clie = rs(0)
  Else
     existe_clie = ""
  End If
  rs.Close
End Function

Function validarautorizado(text As TextBox) As String
  Dim RSQL As String
  Dim rs As Recordset
  Dim codayu As String
  codayu = 12
  RSQL = "Select TCLAVE,TDESCRI from TABAYU  where TCOD= '" & codayu & "' and  Tclave ='" & Trim(text) & "'"
  'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
   Set rs = VGCNx.Execute(RSQL)
   If Not rs.EOF Then 'existe
     validarautorizado = rs(1)
   Else
     validarautorizado = ""
  End If
  rs.Close
End Function

'******************************************************
'Procedimiento que permite verificar antes de grabar
Function verificastk() As Boolean
  Dim cadena As String
   cadena = MSFlexGrid1.TextMatrix(contador, 0)
   If MSFlexGrid1.TextMatrix(contador, 10) = "S" Then
     verificastk = IIf(existe_serie(cadena), True, False)
   ElseIf MSFlexGrid1.TextMatrix(contador, 10) = "N" Then
      verificastk = IIf(SaldoLote(cadena), True, False)
   ElseIf consulta_stk Then
     verificastk = True
   Else
     verificastk = False
  End If
End Function

'Las siguientes consultas verifican si existe stock antes de grabar
'solo si esta saliendo mercaderia se hace la consulta
Function consulta_stk() As Boolean
Dim RSQL As String
Dim rs As Recordset
Dim cadena As String
   cadena = MSFlexGrid1.TextMatrix(contador, 0)
   RSQL = "select  stskdis from stkart  WHERE STALMA='" & VGAlma & "'  and stcodigo ='" & cadena & "'"
   'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
   Set rs = VGCNx.Execute(RSQL)
   If Not rs.EOF Then
     If CANTIDAD > rs(0) Then
       consulta_stk = False
     Else
       consulta_stk = True
     End If
   End If
   rs.Close
End Function

Function SaldoLote(text As String) As Boolean
Dim rs As Recordset
Dim RSQL As String
Dim Lote As String

   Lote = MSFlexGrid1.TextMatrix(contador, 2)
   RSQL = "select  STSLKDIS from STKLOTE where  STSALMA ='" & VGAlma & "' and STSCODIGO = '" & text & "' and STSLOTE = '" & Lote & "'"
'   Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
   Set rs = VGCNx.Execute(RSQL)
   If Not rs.EOF Then
     If CANTIDAD > rs(0) Then
       MsgBox "No hay stock del" & text & "lote:" & Lote, vbInformation, "Aviso"
       SaldoLote = False
     Else
       SaldoLote = True
     End If
   End If
   rs.Close
End Function

Function SaldoSerie(text As String) As Boolean
Dim rs As Recordset
Dim RSQL As String
Dim Serie As String
   Serie = MSFlexGrid1.TextMatrix(contador, 2)
   RSQL = "select STSSKDIS from STKSERI where  STSALMA ='" & VGAlma & "' and STSCODIGO = '" & text & "' and STSSERIE = '" & Serie & "'"
   'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
   Set rs = VGCNx.Execute(RSQL)
   If Not rs.EOF Then
     If CANTIDAD > rs(0) Then
       MsgBox "No hay stock " & text & " serie: " & Serie, vbInformation, "Aviso"
       existe_serie = False
     Else
       existe_serie = True
     End If
   End If
   rs.Close
End Function
Private Sub imprimir()
    Dim cadena As String
    Dim cFormato As String
    Dim cDireccion As String
    Dim cRuc As String
    Dim cNomRepor  As String
    Dim aBusca As New ADODB.Recordset
    
                           CrystalReport1.Reset
                            cNomRepor = "REPNOTAING.rpt"
                            CrystalReport1.ReportFileName = VGParamSistem.RutaReport & cNomRepor
               
                            CrystalReport1.Connect = VGCadenaReport2
                            CrystalReport1.StoredProcParam(0) = VGParamSistem.BDEmpresa
                            CrystalReport1.StoredProcParam(1) = VGAlma
                            CrystalReport1.StoredProcParam(2) = tipo
                            CrystalReport1.StoredProcParam(3) = Text4.text
                            
                            CrystalReport1.DiscardSavedData = True
                            CrystalReport1.Destination = crptToWindow
                            ''CrystalReport1.SelectionFormula = cadena
                            ''CrystalReport1.Formulas(0) = "Empresa = '" & VGparametros.RucEmpresa & "'"
                            ''CrystalReport1.Formulas(1) = "Direccion = '" & cDireccion & "' "
                            ''CrystalReport1.Formulas(2) = "Ruc = '" & cRuc & "' "
                            CrystalReport1.formulas(0) = "fecha='" & DTPicker1.Value & "'"
                            
                            
                            CrystalReport1.formulas(1) = "xtrans = '" & lbltrans.Caption & "' "
                            CrystalReport1.formulas(2) = "xtd = '" & Trim(tipo) & "' "
                            CrystalReport1.formulas(3) = "xndoc = '" & Text4.text & "' "
                            
                            
                            If tipo = "NI" Then
                                CrystalReport1.WindowTitle = "RepNotaIng -- Impresion de Notas de Ingreso"
                                CrystalReport1.formulas(4) = "Xnalma = '" & Text10.text & "' "
                                CrystalReport1.formulas(5) = "Dalma = '" & LblCC.Caption & "' "
                                CrystalReport1.formulas(6) = "AlmaDes = '" & VGAlma & "' "
                                CrystalReport1.formulas(7) = "Dalmades = '" & lblalmacen.Caption & "' "
                            
                            ElseIf tipo = "NS" Then
                                CrystalReport1.WindowTitle = "RepNotaIng -- Impresion de Notas de Salida"
                                CrystalReport1.formulas(4) = "Xnalma = '" & VGAlma & "' "
                                CrystalReport1.formulas(5) = "Dalma = '" & lblalmacen.Caption & "' "
                                CrystalReport1.formulas(6) = "AlmaDes = '" & Text10.text & "' "
                                CrystalReport1.formulas(7) = "Dalmades = '" & LblCC.Caption & "' "
                        
                            End If
                            
                            CrystalReport1.formulas(8) = "NRef = '" & Text1.text & "' "
                            CrystalReport1.formulas(9) = "DocRef = '" & Ctr_AyuDocref.text & "' "
                            CrystalReport1.formulas(10) = "TTrans = '" & Ctr_AyuTransa.xclave & "' "
                            CrystalReport1.formulas(11) = "emp = '" & VGParametros.RucEmpresa & "'"
                            CrystalReport1.WindowShowPrintBtn = True
                            CrystalReport1.WindowShowRefreshBtn = True
                            CrystalReport1.WindowShowSearchBtn = True
                            CrystalReport1.WindowShowPrintSetupBtn = True
                            CrystalReport1.WindowState = crptMaximized
                            
                            
                            If CrystalReport1.Status <> 2 Then
                                CrystalReport1.Action = 1
                                VGCNx.Execute "Update MovAlmCab Set CaEstImp = 'I' Where CATD = '" & tipo & "' and CANUMDOC = '" & Text4.text & "'"
                            End If
        Exit Sub
ErrImp:
     MsgBox Err.Description
     Resume Next
End Sub



Private Sub imprimirBK()
Dim cadena As String
If TxTransa = "DP" Then
   CrystalReport1.WindowTitle = "Inv520 -- Control de Inventarios"
   CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "\inv520.rpt"
Else
   CrystalReport1.WindowTitle = "Inv043 -- Control de Inventarios"
   CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "\inv043.rpt"
End If
Ubi_Tab CrystalReport1
cadena = "{MOVALMCAB.CAALMA} = '" & VGAlma & "'  and {MOVALMCAB.CATD} = '" & tipo & "' and {MOVALMCAB.CANUMDOC} = '" & NumDoc & "'"
CrystalReport1.DiscardSavedData = True
CrystalReport1.Destination = crptToWindow
CrystalReport1.WindowTitle = " Control de Inventarios"
CrystalReport1.ReplaceSelectionFormula (cadena)
CrystalReport1.WindowShowPrintBtn = True
CrystalReport1.WindowShowRefreshBtn = True
CrystalReport1.WindowShowSearchBtn = True
CrystalReport1.WindowShowPrintSetupBtn = True
CrystalReport1.formulas(0) = "empresa ='" & VGParametros.RucEmpresa & "'"
CrystalReport1.formulas(1) = "nota ='" & Codigo2 & "'"
CrystalReport1.formulas(2) = "hora ='" & Time & "'"
If VGRegEnt = 0 Then
    CrystalReport1.formulas(3) = "Tipo = 'S'"
Else
    CrystalReport1.formulas(3) = "Tipo = 'I'"
End If
CrystalReport1.Action = 1

If VGRegEnt <> 1 And TxTransa = "TD" Then
    If vbOK = MsgBox(" Desea imprimir la nota de Ingreso", vbInformation + vbOKCancel, "Aviso") Then
        CrystalReport1.WindowTitle = "Inv043 -- Control de Inventarios"
        CrystalReport1.ReportFileName = RUTA & "reportes\inv043.rpt"
        Ubi_Tab CrystalReport1
        cadena = "{MOVALMCAB.CAALMA} = '" & Ctr_AyuAlmacen.xclave & "'  and {MOVALMCAB.CATD} = '" & Campo & "' and {MOVALMCAB.CANUMDOC} = '" & Format(nument, "0000000000") & "'"
        CrystalReport1.DiscardSavedData = True
        CrystalReport1.Destination = crptToWindow
        CrystalReport1.WindowTitle = " Control de Inventarios"
        CrystalReport1.ReplaceSelectionFormula (cadena)
        CrystalReport1.WindowShowPrintBtn = True
        CrystalReport1.WindowShowRefreshBtn = True
        CrystalReport1.WindowShowSearchBtn = True
        CrystalReport1.WindowShowPrintSetupBtn = True
        CrystalReport1.formulas(0) = "empresa ='" & VGParametros.RucEmpresa & "'"
        CrystalReport1.formulas(1) = "nota ='NOTA DE INGRESO'"
        CrystalReport1.formulas(2) = "hora ='" & Time & "'"
        CrystalReport1.formulas(3) = "Tipo = 'S'"
        CrystalReport1.Action = 1
   End If
End If
End Sub



Private Sub ctr_ayudas()
Call Ctr_AyuAlmacen.conexion(VGCNx)
Call Ctr_AyuCliente.conexion(VGCNx)
Call Ctr_AyuDocref.conexion(VGCNx)
Call Ctr_AyuTransa.conexion(VGCNx)
Call Ctr_AyuProveedor.conexion(VGCNx)
If VGRegEnt = 1 Then
   Ctr_AyuTransa.filtro = "tt_tipmov='I' and rtrim(tt_codtrans_auto)=''"
Else
Ctr_AyuTransa.filtro = "tt_tipmov='S' and rtrim(tt_codtrans_auto)=''"
End If
Ctr_AyuAlmacen.filtro = "empresacodigo='" & VGParametros.empresacodigo & "' and puntovtacodigo='" & VGParametros.puntovta & "'"
End Sub


Private Sub Combo1_Click()
   CmdDetEnvio.Enabled = True
   CmdDetEnvio.SetFocus
End Sub
'Enviar

Private Sub Combo1_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
   CmdDetEnvio.Enabled = True
   CmdDetEnvio.SetFocus
 End If
End Sub

Public Sub CmdDetEnvio_Click()
Dim criterio As String
Dim dato1 As String
Dim ncombo As Integer
Dim kflag, J As Integer
kflag = 0
For J = 1 To FrmmntNotas.DataGrid.RecordCount
    If FrmmntNotas.CENTROCOSTO = 1 Then
       criterio = Trim(FrmmntNotas.DataGrid1.Columns(0)) + Trim(FrmmntNotas.DataGrid1.Columns(2)) + Trim(FrmmntNotas.DataGrid1.Columns(11))
'       criterio = Trim(FrmmntNotas.MSFlexGrid1.TextMatrix(J, 0)) + Trim(FrmmntNotas.MSFlexGrid1.TextMatrix(J, 2)) + Trim(FrmmntNotas.MSFlexGrid1.TextMatrix(J, 11))
       dato1 = Trim(TxtArticulo) + Trim(Text6) + Trim(txccosto)
     Else
       criterio = Trim(FrmmntNotas.DataGrid.Columns(0)) + Trim(FrmmntNotas.DataGrid.Columns(2))
'       criterio = Trim(FrmmntNotas.MSFlexGrid1.TextMatrix(J, 0)) + Trim(FrmmntNotas.MSFlexGrid1.TextMatrix(J, 2))
       dato1 = Trim(TxtArticulo) + Trim(Text6)
    End If
    If criterio = dato1 Then
       kflag = 1
       Exit For
    End If
Next
If kflag = 1 Then
   If Trim(Text6) <> "" Then
      MsgBox "Ya existe el lote para el articulo...Verifique!!!", vbInformation, "AVISO"
    ElseIf Trim(txccosto) <> "" Then
            MsgBox "Ya existe el articulo + centro de costos ...Verifique!!!", vbInformation, "AVISO"
         Else
           MsgBox "Ya existe el articulo...Verifique!!!", vbInformation, "AVISO"
   End If
   Exit Sub
Else
  TxtArticulo = Trim(TxtArticulo)
End If

If Not IsNumeric(TxtCantidad.text) Then
       MsgBox "Ingrese cantidad respectiva", vbOKOnly + vbExclamation, "Error"
       TxtCantidad.SetFocus
       TxtCantidad.SelStart = 0: TxtCantidad.SelLength = Len(TxtCantidad)
       Exit Sub
End If
If Val(lbcantstk) < Val(TxtCantidad) And (VGRegEnt <> 1) Then
    MsgBox "La cantidad no puede ser mayor al stock", vbOKOnly + vbExclamation, "Error"
    If TxtCantidad.Enabled Then TxtCantidad.SetFocus
    Exit Sub
End If
If flagserie = "S" And (VGRegEnt = 1) And Text6 = "" Then 'And Not Combo1.Enabled
     MsgBox "Ingrese el Número de serie", vbOKOnly + vbExclamation, "Error"
     Text6.SetFocus
     Exit Sub
End If
If flaglote = "S" And (Text6 = "") Then 'And Not Combo1.Enabled
     MsgBox "Ingrese el Número de Lote", vbOKOnly + vbExclamation, "Error"
     Text6.SetFocus
     Exit Sub
  End If
If (flagserie = "S") Then
    If FrmmntNotas.DataGrid.RecordCount <> 1 Then
        For ncombo = 1 To FrmmntNotas.DataGrid.RecordCount - 1
          If Combo1.Visible Then
            If UCase(Combo1.text) = UCase(FrmmntNotas.MSFlexGrid1.TextMatrix(ncombo, 2)) Then
              MsgBox "Ya se ingreso la serie", vbInformation, "Aviso"
              Combo1.SetFocus
              Exit Sub
            End If
          ElseIf Text6 <> "" Then
            If UCase(Text6.text) = UCase(FrmmntNotas.MSFlexGrid1.TextMatrix(ncombo, 2)) Then
              MsgBox "Ya se ingreso la serie", vbInformation, "Aviso"
              Text6.SetFocus
              Exit Sub
            End If
          End If
        Next ncombo
    End If
End If
If (flagserie = "S") And (VGRegEnt = 1) Then
   If existe_serie(Text6) Then Exit Sub
End If
If flaglote = "S" And (VGRegEnt <> 1) Then
   If Not SaldoLote(Text6) Then
        Text6.SetFocus
        Exit Sub
   End If
End If
If flagserie = "N" And VGSeleccion <> 2 Then Carga ' revisar  verifica la conversion de unidades
If VGForm <> 6 Or VGEstadomodi Or VGtipocreacion = 2 Then
    'verifico si actualizo
                    'GuiaSalida
    ' Else
                   'ingreso o salida
     CANTIDAD = Val(TxtCantidad)
     ingreso_salida
     If VGEstadomodi Then
          Unload Me
          Exit Sub
     End If
End If
limpia
'*************************
'FrmRegistro.buscar_trans
'*************************
Combo1.Visible = False
Text6.Visible = True
TxtArticulo.Enabled = True
'Entra a las multiples opciones

 
 If I <> 0 Then         'solo entra cuando hay  dato en temporal de salida
   I = I + 1             'contador de item
   LblNroReg.Visible = True
   LblNroReg = I
   
   If I < fin Then
                DisplayDisp         'funcion de llenar los datos
                If flagserie = "S" Or flaglote = "S" Then
                       If (VGRegEnt <> 1) And (flagserie = "S") Then
                          Combo1.Visible = True
                          Combo1.SetFocus
                          CmdDetEnvio.Enabled = True
                          TxtCantidad.Enabled = False
                          Text6.Visible = False
                       ElseIf flagserie = "S" Then
                           MaskEdBox1.BackColor = &H8000000F
                           FechaFabric.BackColor = &H80000009
                           MaskEdBox1.Enabled = False
                           FechaFabric.Enabled = True
                           Text6.SetFocus
                       Else
                           MaskEdBox1.BackColor = &H80000009
                           FechaFabric.BackColor = &H80000009
                           MaskEdBox1.Enabled = True
                           FechaFabric.Enabled = True
                           Text6.SetFocus
                       End If
                 Else
                        MaskEdBox1.BackColor = &H8000000F
                        FechaFabric.BackColor = &H8000000F
                       TxtCantidad.Enabled = True
                       TxtCantidad.SetFocus
                End If
      Else
                CmdDetSalir.SetFocus
      End If
End If
CmdDetEnvio.Enabled = False
Ctr_Ayuart.SetFocus
Ctr_Ayuart.xclave = "": Ctr_Ayuart.Ejecutar
graba = False
If VGSeleccion = 2 Or VGSeleccion = 3 Then
    Unload Me
End If
End Sub

Private Sub CmdDetKimpia_Click()
limpia
FrmRegistro.buscar_trans
TxtArticulo.SetFocus
End Sub

Private Sub CmdDetSalir_Click()
Label9.Caption = ""
lbcantstk = ""
Unload Me
End Sub

Private Sub Chkserie_Click()
 If chkserie.Value Then
        formIngSerie.Show 1
 End If
End Sub

Private Sub Ctr_AyuAnalitico_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
txEquip = Ctr_AyuAnalitico.xclave
End Sub

Private Sub Ctr_Ayuart_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Dim xsql As New ADODB.Recordset
TxtArticulo.text = Ctr_Ayuart.xclave
Set xsql = VGCNx.Execute(" select stskdis from stkart where stalma='" & VGAlma & "' and stcodigo='" & Ctr_Ayuart.xclave & "'")
lbcantstk = ESNULO(xsql!STSKDIS, 0)
TxtArticulo_KeyPress (13)
End Sub



Private Sub MaskEdBox1_GotFocus()
  MaskEdBox1.SelStart = 0: MaskEdBox1.SelLength = Len(MaskEdBox1)
End Sub

Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If MaskEdBox1 = "__/__/____" Then
        ' MsgBox "Ingrese  Fecha de Vencimiento ", vbExclamation + vbOKOnly, "Advertencia"
         'MaskEdBox1.SetFocus
         'Exit Sub                             'Cambios de la version 6
    End If
    If FechaFabric.Visible Then
         FechaFabric.SetFocus
    Else
        TxtCantidad.SetFocus
'         SendKeys "{tab}"
'         KeyAscii = 0
    End If
End If
End Sub

Private Sub MaskEdBox1_LostFocus()
Dim cValor As String

If MaskEdBox1 = "__/__/____" Then
Else
    cValor = ValidFecha(MaskEdBox1)
    If cValor = "" Then
       MsgBox "Ingrese la Fecha Correctamente", vbExclamation + vbOKOnly, "Advertencia"
       MaskEdBox1 = "__/__/____"
       MaskEdBox1.SetFocus
       Exit Sub
    Else
      MaskEdBox1 = cValor
    End If
    If CDate(MaskEdBox1) < Date Then
        MsgBox "El articulo ya vencio", vbExclamation + vbOKOnly, "Error"
        MaskEdBox1.SetFocus
    End If
End If
End Sub

Private Sub FechaFabric_GotFocus()
  MaskEdBox1.SelStart = 0: MaskEdBox1.SelLength = Len(MaskEdBox1)
End Sub

Private Sub FechaFabric_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
      If FechaFabric = "__/__/____" And VGRegEnt = 1 Then
        MsgBox "Ingrese  Fecha  de Fabricación ", vbExclamation + vbOKOnly, "Advertencia"
        FechaFabric.SetFocus
        'Exit Sub                  'Cambios de la version 6
      End If
      ' TxtCantidad.SetFocus
      SendKeys "{tab}"
      KeyAscii = 0
End If
End Sub

Private Sub FechaFabric_LostFocus()
Dim cValor As String

If FechaFabric = "__/__/____" Then
Else
    cValor = ValidFecha(FechaFabric)
    If cValor = "" Then
        MsgBox "Ingrese la Fecha Correctamente", vbExclamation + vbOKOnly, "Advertencia"
        FechaFabric = "__/__/____"
        FechaFabric.SetFocus
    ElseIf CDate(cValor) > Date Then
'        MsgBox "fecha de fab. es mayor que la fecha actual", vbExclamation + vbOKOnly, "Error"
'        FechaFabric.SetFocus
'    ElseIf FechaFabric < CDate(cValor) Then
'        MsgBox "Ingrese fecha Valida", vbExclamation + vbOKOnly, "Error"
'        FechaFabric.SetFocus
    Else
        FechaFabric = cValor
    End If
End If
End Sub

Private Sub TxtUniRef_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     SendKeys "{tab}"
     KeyAscii = 0
  End If
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
If Cancel Then
   Text6_KeyPress (13)
End If
End Sub

Private Sub txccosto_DblClick()
  Dim Adodc3 As ADODB.Recordset   'Centro de Costos
  Set Adodc3 = New ADODB.Recordset
 Adodc3.Open "SELECT centrocostocodigo,centrocostodescripcion FROM ct_centrocosto where empresacodigo='" & VGParametros.empresacodigo & "' and  centrocostonivel = 3", VGcnxCT, adOpenStatic, adLockOptimistic
 frmReferencia.Conectar Adodc3, "SELECT centrocostocodigo,centrocostodescripcion FROM ct_centrocosto where empresacodigo='" & VGParametros.empresacodigo & "' and  centrocostonivel = 3"
 frmReferencia.Label1.Caption = "Centro de Costos"
 frmReferencia.Show vbModal

        If vGUtil(1) <> "" Then
           txccosto = vGUtil(1)
                 'LblCC = vGUtil(2)
        End If
        If txccosto.text <> "" Then txccosto_KeyPress (13)
End Sub

Private Sub txccosto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then txccosto_DblClick
End Sub

Private Sub txccosto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Existe(3, txccosto.text, "ct_CENTROCOSTO", "centrocostocodigo", False) = False Then
            MsgBox "Centro de Costo no existe", vbInformation, "Mensaje"
            txccosto = ""
            txccosto.SetFocus: Exit Sub
        Else
            Tabula (KeyAscii)
        End If
    End If

End Sub

Private Sub txEquip_KeyPress(KeyAscii As Integer)
Tabula (KeyAscii)
End Sub

Private Sub TxordFab_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then frm_manten_ordfabri.Show 1
End Sub

Private Sub TxordFab_KeyPress(KeyAscii As Integer)
Tabula (KeyAscii)
End Sub

Private Sub TxtArticulo_DblClick()
On Error Resume Next
cant = 0
I = 1
'Load (FormAyuArt)

VGForm1 = 2
FormAyuArt.Show 1
fin = Salida.Rows
If Salida.Rows = 1 Then Exit Sub
LblNroReg.Visible = True
lbEtiNum.Visible = False
LblNroReg = I
DisplayDisp
If flagserie = "S" Or flaglote = "S" Then
          If flaglote = "S" Then
                MaskEdBox1.Enabled = True
                FechaFabric.Enabled = True
                xserie = "N"
                VGcod = TxtArticulo
                Text6.Visible = True
                Text6.Enabled = True
                Text6.SetFocus
                TxtCantidad.Enabled = True
          Else
                xserie = "S"
                chkserie.Enabled = True
                TxtCantidad.Enabled = False
                TxtCantidad = "1"
                CmdDetEnvio.Enabled = True
                If VGRegEnt <> 1 Then
                        Combo1.Visible = True
                        Text6.Visible = False
                        Combo1.SetFocus
                Else
                        If Text6.Enabled = True Then Text6.SetFocus
                End If
          End If
Else
          xserie = "X"
          TxtCantidad.Enabled = True
          TxtCantidad.SetFocus
'          txtcanref.SetFocus
End If
'TxtUniRef.Enabled = True

End Sub

Private Sub TxtArticulo_KeyPress(KeyAscii As Integer)
 Dim rpta As Integer
 Dim criterio As String
 Dim rsa, RSB As New ADODB.Recordset
  Dim cant As Integer
  cant = 0
If KeyAscii = 13 Then
        If Not Validadato(TxtArticulo) Then
          MsgBox "CODIGO NO VALIDO....!!", vbInformation, "AVISO"
          Call VGDllGeneral.Enfoquetexto(TxtArticulo)
          Exit Sub
        
        End If
        criterio = " where a.ACODIGO = " & "'" + TxtArticulo.text + "'"
        Set rsa = VGCNx.Execute("Select A.*,B.PRODUCTOPRECVTA from MAEART A INNER JOIN LISTAPRE1 B ON A.ACODIGO=B.PRODUCTOCODIGO" & criterio)
        If rsa.RecordCount > 0 Then
           Set rsSTKART = VGCNx.Execute("Select * from STKART WHERE STALMA='" & VGAlma & "'")
           Label14.Caption = "" & rsa.Fields("AUNIDAD")
           LblPrecio.Caption = ESNULO(rsa.Fields("productoprecvta"), 0)
           VGabrev = Label14
           lblUniEst = Nombre_Unidad(VGabrev)
           flagserie = IIf(Not IsNull(rsa.Fields("AFSERIE")), rsa.Fields("AFSERIE"), "N")
           flaglote = IIf(Not IsNull(rsa.Fields("AFLOTE")), rsa.Fields("AFLOTE"), "N")

           criterio = " STCODIGO ='" & TxtArticulo.text & "' and  STALMA ='" & VGAlma & "'"
           rsSTKART.Filter = criterio
           If Not rsSTKART.EOF Then
               If stockcomp Then
                 cant = numero(rsSTKART("STSKDIS")) - numero(rsSTKART("STSKcom"))
                 Else
                 cant = numero(rsSTKART("STSKDIS"))
              End If
           Else
               cant = 0
           End If
           lbcantstk = cant
           TxtArticulo.Enabled = False
           ver_serie_lote
           If flagserie = "S" Or flaglote = "S" Then    ' crear funcion
                If flaglote = "S" Then
                        MaskEdBox1.Enabled = True
                        FechaFabric.Enabled = True
                        xserie = "N"
                        VGcod = TxtArticulo
                        SendKeys "{tab}"
                        KeyAscii = 0
                Else
                        xserie = "S"
                        chkserie.Enabled = True
                        If VGRegEnt <> 1 Then
                           Combo1.Visible = True
                           Text6.Visible = False
                           agregar_combo
                           Combo1.SetFocus
                        Else
                            Text6.SetFocus
                        End If
                End If
           Else
                xserie = "X"
                TxtCantidad.SetFocus
                'txtcanref.SetFocus
           End If
        Else
             If Val(TxtArticulo) = 0 Then
                TxtArticulo_DblClick
             ElseIf TxtArticulo <> "" Then
                    MsgBox "El Código de Articulo no existe ", vbExclamation, mensaje1
               End If
             TxtArticulo.SetFocus
             txtcanref.SetFocus
          End If
               
Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
        ' para establecer que no hay nada seleccionado
End Sub

Private Sub TxtUniRef_DblClick()
'Dim db As Database
Dim rs As New ADODB.Recordset
Dim RSQL As String
Dim FACTOR As Double
If VGRegEnt = 1 Then
    Frmayuunidades.Show 1
    If TxtCantidad <> "" Then CmdDetEnvio.Enabled = True
    Carga
    If Not dato_invalido Then Exit Sub
End If
End Sub

Private Sub TxtCanref_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Len(Trim(txtcanref)) = 0 Then txtcanref = 0
      SendKeys "{tab}"
   End If
End Sub

Private Sub TxtCantidad_Change()
FACTOR = 1
If TxtCantidad <> "" Then
    If TxtCantidad.Enabled Then
        If Not IsNumeric(TxtCantidad.text) And TxtArticulo <> "" Then
            MsgBox "Ingrese un Numero", vbOKOnly + vbExclamation, "Error"
            'If TxtCantidad.Visible And TxtCantidad.Enabled Then TxtCantidad.SetFocus
            'MOMENTO DE MODIFICAR
        Else
            If IsNumeric(TxtCantidad) Then
                LblCantidad = Val(TxtCantidad) * FACTOR             'entra siempre al momento de editar
                If Label9 = "" Then Label9 = lblUniEst
                CmdDetEnvio.Enabled = True
            End If
        End If
    End If
End If
End Sub

Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Not IsNumeric(TxtCantidad.text) And TxtArticulo <> "" Then
        MsgBox "Ingrese la cantidad", vbOKOnly + vbExclamation, "Error"
        Tabula (KeyAscii)
    Else
        'debe revisar solo cuando tenga tipo de unidad
        If VGtipocreacion = 1 Then Carga '   devuelve dato_invalido=false  cuando se produjo error
   End If
 Else
        If Chr(KeyAscii) = "." And IsNumeric(TxtCantidad) Then Exit Sub
        If ((Chr$(KeyAscii) < "0" Or Chr(KeyAscii) > "9")) And KeyAscii <> 8 Then KeyAscii = 0
 End If
End Sub

Public Sub DisplayDisp()
  'funcion de llenar los datos de formulario utilizando los datos MSflexGrid
Dim criterio As String
Dim RSQL As New ADODB.Recordset
On Error GoTo Err
If I > Salida.Rows - 1 Then Exit Sub
TxtArticulo = Salida.TextMatrix(I, 0)   'codigo
If TxtArticulo.text = "" Then Exit Sub
VGabrev = Salida.TextMatrix(I, 2)  'UNIDAD
flagserie = Salida.TextMatrix(I, 3) 'serie
flaglote = Salida.TextMatrix(I, 4) 'serie
criterio = "STCODIGO ='" & TxtArticulo.text & "' and STALMA ='" & VGAlma & "'"
   'RMM ****************************************************
criterio = "select * from stkart where " & criterio
   'RMM ****************************************************
   Set RSQL = VGCNx.Execute(criterio)
      
   If RSQL.RecordCount() > 0 Then
      'Data2.Recordset.FindFirst criterio
      If stockcomp Then
         cant = ESNULO(RSQL!STSKDIS, 0) - ESNULO(RSQL!STSKcom, 0)
       Else
         cant = ESNULO(RSQL!STSKDIS, 0)
      End If
   Else
      cant = 0
   End If

Label14 = VGabrev  ' label14 variable auxiliar
lblUniEst = Nombre_Unidad(VGabrev)
lbcantstk = cant
TxtUniRef.text = lblUniEst
 'TxtArticulo.Locked = True
ver_serie_lote
lbEtiNum.Visible = True
Exit Sub
Err:
   MsgBox Err.Description
  Resume
End Sub

Public Sub Carga()
Dim criterio1 As String
''Dim db As Database
Dim rs As New ADODB.Recordset
Dim RSQL As String
Dim FACTOR As Double
FACTOR = 1
If Trim(Label14) <> Trim(VGabrev) Then                          'CONSULTA POR DEFECTO MODIFICAR
        RSQL = "select  p.EQCANTEQUI from TabEqui p where p.EQUNIPRI = '" & VGabrev & "'   and p.EQUNIEQUI = '" & Label14.Caption & "'"
        Set rs = VGCNx.Execute(RSQL)
        If rs.RecordCount = 0 Then
            MsgBox "la unidad de referencia no tiene unidad equivalente"
            lblUniEst = Nombre_Unidad(Label14)
            Exit Sub
        End If
        rs.MoveFirst
        FACTOR = rs.Fields("EQCANTEQUI")
        rs.Close
  Else
        FACTOR = 1
  End If
cant = Val(lbcantstk)
  
  If cant < Val(TxtCantidad.text) * FACTOR And VGRegEnt = 0 Then  ' revisar si validar en creacion
        MsgBox "No hay stock suficente", 48, "Aviso"
        TxtCantidad.SetFocus
        dato_invalido = False
        Exit Sub
  End If
  dato_invalido = True
  LblCantidad = Val(TxtCantidad.text) * FACTOR 'VGcant
  lblUniEst = Nombre_Unidad(Label14)
End Sub

Function coduso(dato As String) As String
Dim RSQL As String
Dim rs As New ADODB.Recordset
RSQL = "select UM_ABREV from TabUniMed where UM_NOMBRE ='" & dato & "'"
'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
Set rs = VGCNx.Execute(RSQL)
If rs.EOF Then
    coduso = ""
Else
    coduso = rs(0)
End If
rs.Close
End Function

Function Nombre_Unidad(dato As String) As String
Dim RSQL As String
Dim rs As New ADODB.Recordset
RSQL = "select UM_NOMBRE from TabUniMed where UM_ABREV ='" & dato & "'" '   AND UM_ESTADO ='A'"
'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
Set rs = VGCNx.Execute(RSQL)
If rs.EOF Then
        Nombre_Unidad = ""
Else
        Nombre_Unidad = rs(0)
End If
rs.Close
End Function

Function preciovta(Cod As String) As Double
Dim RSQL As String
Dim rs As Recordset
RSQL = "select APRECIO from maeart where ACODIGO='" & TxtArticulo & "'"
'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
Set rs = VGCNx.Execute(RSQL)
If rs.EOF Then
    preciovta = 0
Else
    preciovta = rs(0)
End If
rs.Close
End Function

Private Sub limpia()
   LblNroReg = ""
   Label14.Caption = ""
   lblUniEst = ""
   LblCantidad.Caption = ""
   Label9.Caption = ""
   TxtArticulo.text = ""
   Fechavcto.text = ""
   lbcantstk = ""
   TxtUniRef.text = ""
   MaskEdBox1 = "__/__/____"
   Text6.Enabled = True
   Text6.text = ""
   Text6.Enabled = False
   TxtCantidad.Enabled = True
   TxtCantidad.text = ""
   FechaFabric = "__/__/____"
   MaskEdBox1.BackColor = &H80000009
   Text6.BackColor = &H80000009
   FechaFabric.BackColor = &H80000009
   TxtArticulo.Enabled = True
   MaskEdBox1.Enabled = False
   Text6.Enabled = False
   FechaFabric.Enabled = False
   LblNroReg.Visible = False
   lbEtiNum.Visible = False
   CmdDetEnvio.Enabled = False
   Combo1.Clear
   Combo1.Visible = False
   Text6.Visible = True
   chkserie.Enabled = False
End Sub

Private Sub ver_serie_lote()

    If flagserie = "S" Or flaglote = "S" Then
             Text6.Enabled = True
             Text6.Visible = True
             If (VGRegEnt <> 1) And flagserie = "S" And VGtipocreacion = 1 Then 'con guia de salida
                        agregar_combo
                        Combo1.Visible = True
                        Text6.Visible = False
             End If
             If flaglote = "S" Then
                          MaskEdBox1.BackColor = &H80000009
                          FechaFabric.BackColor = &H80000009
                          MaskEdBox1.Enabled = True
                          FechaFabric.Enabled = True
                          VGcod = TxtArticulo
                          Text6.Visible = True
              Else
                         TxtCantidad = "1"
                         TxtCantidad.Enabled = False
                         MaskEdBox1.BackColor = &H8000000F
                         FechaFabric.BackColor = &H80000009
                         MaskEdBox1.Enabled = False
                         FechaFabric.Enabled = True
             End If
   Else
     'Text6.Visible = False
     Text6.Enabled = False
     Text6.BackColor = &H8000000F
     MaskEdBox1.BackColor = &H8000000F
     FechaFabric.BackColor = &H8000000F
     MaskEdBox1.Enabled = False
     FechaFabric.Enabled = False
     TxtUniRef.Enabled = True
     TxtCantidad.Enabled = True
   End If
    

End Sub

Private Sub agregar_combo()
  Dim rs As New ADODB.Recordset
  Dim RSQL As String
  If flagserie = "S" Then
     RSQL = "select stsserie from stkseri where  STSALMA='" & VGAlma & "' and STSCODIGO='" & TxtArticulo & "' and STSSKDIS<>0 "
  End If
  'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
  Set rs = VGCNx.Execute(RSQL)
  If rs.EOF Then Exit Sub                     'revisar porque no entra al bucle
  Combo1.Clear
  While Not rs.EOF
     Combo1.AddItem (rs(0))
     rs.MoveNext
  Wend
  rs.Close
  Combo1.ListIndex = 0
  CmdDetEnvio.Enabled = True
End Sub

Private Sub Text6_Change()
 If Trim(Text6) <> "" Then CmdDetEnvio.Enabled = True
End Sub

Private Sub Text6_DblClick()
  If flaglote = "S" Then
    FormAyuLote.Show 1
  End If
End Sub

Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 112 Then
     Text6_DblClick
 ElseIf KeyCode = vbKeyTab Then
     Text6_KeyPress (13)
 End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
 Dim ncombo As Integer
 If KeyAscii = 13 Then
    Text6 = Trim(Text6)
    If Text6 <> "" Then
                If flaglote = "S" Then
                         existe_lote1 Text6
                         MaskEdBox1.SetFocus
                Else
                         If flagserie = "S" Then
                            If FrmRegistro.DataGrid.RecordCount <> 1 Then
                                For ncombo = 1 To FrmRegistro.DataGrid.RecordCount - 1
                                  If Combo1.Visible Then
                                      If UCase(Combo1.text) = UCase(FrmRegistro.MSFlexGrid1.TextMatrix(ncombo, 2)) Then
                                        MsgBox "Ya se ingreso la serie", vbInformation, "Aviso"
                                        Combo1.SetFocus
                                        Exit Sub
                                      End If
                                  ElseIf Text6 <> "" Then
                                      If UCase(Text6.text) = UCase(FrmRegistro.MSFlexGrid1.TextMatrix(ncombo, 2)) Then
                                        MsgBox "Ya se ingreso la serie", vbInformation, "Aviso"
                                        Text6.SetFocus
                                        Exit Sub
                                      End If
                                  End If
                                Next ncombo
                            End If
                            If existe_serie(Text6) Then Exit Sub
                            TxtCantidad = "1"
                            TxtCantidad.Enabled = False
                            TxtUniRef.Enabled = False
                            CmdDetEnvio.Enabled = True
                           
                         End If
                         SendKeys "{tab}"
                         KeyAscii = 0
                         Text6_Validate (False)
                End If
      Else
             If flaglote = "S" Then
                          MsgBox "Ingrese el número de Lote", vbInformation, "Aviso"
             Else
                          MsgBox "Ingrese el número de Serie", vbInformation, "Aviso"
             End If
             Text6.SetFocus
      End If
 End If
End Sub
              
Private Sub deshabilitartx5_tx3(flag As Boolean)
    TxtCantidad.Enabled = flag
    TxtUniRef.Enabled = flag
End Sub
 
Private Sub existe_lote(text As TextBox)
Dim rs As New ADODB.Recordset
Dim RSQL As String
   RSQL = "select STSLOTE, STSLKDIS,STSFECVEN,STSFECFAB from STKLOTE where  STSALMA ='" & VGAlma & "' and STSCODIGO = '" & TxtArticulo & "' and STSLOTE = '" & text & "'"
   'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
   Set rs = VGCNx.Execute(RSQL)
   If Not rs.EOF Then
     MsgBox "Lote Registrado en Almacen", vbInformation, "Aviso"
     lbcantstk = rs(1)
     MaskEdBox1 = IIf(IsNull(rs(2)), "__/__/____", rs(2))
     FechaFabric = IIf(IsNull(rs(3)), "__/__/____", rs(3))
   End If
End Sub

Function existe_lote1(text As TextBox) As Boolean
Dim rs As New ADODB.Recordset
Dim RSQL As String
   RSQL = "select STSLOTE,STSFECVEN,STSFECFAB from STKLOTE where  STSALMA ='" & VGAlma & "' and STSCODIGO = '" & TxtArticulo & "' and STSLOTE = '" & Text6 & "'"
   'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
   Set rs = VGCNx.Execute(RSQL)
   If Not rs.EOF Then
     If Not graba Then MsgBox "Lote Registrado en Almacen", vbInformation, "Aviso"
     existe_lote1 = True
     If Not IsNull(rs(1)) Then
           MaskEdBox1 = rs(1)
     End If
     If Not IsNull(rs(2)) Then
           FechaFabric = rs(2)
     End If
   Else
     MsgBox "Lote  No Registrado en Almacen", vbInformation, "Aviso"
     existe_lote1 = False
   End If
   rs.Close
End Function

Function existe_serie(text As TextBox) As Boolean
Dim rs As New ADODB.Recordset
Dim RSQL As String
   RSQL = "select STSSERIE,STSFECVEN from STKSERI where  STSALMA ='" & VGAlma & "' and STSCODIGO = '" & TxtArticulo & "' and STSSERIE = '" & Text6 & "'"
   'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
   Set rs = VGCNx.Execute(RSQL)
   If Not rs.EOF Then  'Not graba
     If True Then MsgBox "Serie Registrada en Almacen", vbInformation, "Aviso"
     If Not IsNull(rs(1)) Then
        MaskEdBox1 = rs(1)
     End If
     existe_serie = True
   Else
     If Not graba Then MsgBox "Serie  No Registrada en Almacen", vbInformation, "Aviso"
     existe_serie = False
   End If
   rs.Close
End Function

Private Sub ingreso_salida()
       If VGtipocreacion = 2 Then
           hubo_error = False
           grabadetalle
           If hubo_error Then Exit Sub
'           MsgBox "Se grabo sastifactoriamente", vbInformation, "Aviso"
       End If
       If (VGRegEnt <> 1) And (flagserie = "S") And VGtipocreacion = 1 Then
           serie_lote = Combo1.text
       Else
           serie_lote = Trim(Text6)
       End If
       If VGSeleccion = 2 Then
        LblCantidad = IIf(IsNumeric(LblCantidad), LblCantidad, TxtCantidad)
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = serie_lote  'serie
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = TxtCantidad.text  'çantidad
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = VGabrev
        '**********************************************************************
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 11) = txccosto
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 12) = TxordFab
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 13) = txEquip
      
        '**********************************************************************
       'varform.MsFlexgrid1.TextMatrix(varform.MsFlexgrid1.Row, 4) = Text9.Text   'Precio
        If VGtipocreacion = 1 Then
             MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) = Val(LblCantidad)   'Cantidad informada
        Else
             MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) = Val(FrmModificar.numitem)    'Cantidad informada
        End If
        If xserie = "S" Then
           MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 10) = "S"
           MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 8) = MaskEdBox1
           MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 9) = MaskEdBox1
           xserie = "S"
        ElseIf xserie = "N" Then
           MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 10) = "N"
        Else
           MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 10) = "X"
        End If
      Else
      '                                       0                 1                 2                    3                       4             5               6                   7                      8                9               10
        pro_xserie
        LblCantidad = IIf(IsNumeric(LblCantidad), LblCantidad, TxtCantidad)
        If VGtipocreacion = 1 Then
          MSFlexGrid1.AddItem (TxtArticulo.text & vbTab & Ctr_Ayuart.xnombre & vbTab & serie_lote & vbTab & Format(TxtCantidad.text, "##0.00") & vbTab & VGabrev & vbTab & "" & vbTab & Val(LblCantidad) & vbTab & "" & vbTab & MaskEdBox1 & vbTab & FechaFabric & vbTab & xserie & vbTab & txccosto & vbTab & TxordFab & vbTab & txEquip & vbTab & txtcanref)
        Else
         FrmModificar.MSFlexGrid1.AddItem (TxtArticulo.text & vbTab & Ctr_Ayuart.xnombre & vbTab & serie_lote & vbTab & Format(TxtCantidad.text, "###0.00") & vbTab & VGabrev & vbTab & "" & vbTab & Val(FrmModificar.numitem - 1) & vbTab & "" & vbTab & MaskEdBox1 & vbTab & FechaFabric & vbTab & xserie & txccosto & vbTab & TxordFab & vbTab & txEquip)
        'FrmModificar.MSFlexGrid1.AddItem (TxtArticulo.text & vbTab & LblCodigo & vbTab & serie_lote & vbTab & Format(TxtCantidad.text, "###0.00") & vbTab & VGabrev & vbTab & "" & vbTab & Val(FrmModificar.numitem - 1) & vbTab & "" & vbTab & MaskEdBox1 & vbTab & FechaFabric & vbTab & xserie)
      End If
     End If
End Sub

Private Sub pro_xserie()
  If flagserie = "S" Then
        xserie = "S"
        Exit Sub
  End If
  If flaglote = "S" Then
        xserie = "N"
        Exit Sub
  End If
  xserie = "X"
End Sub

Private Sub modifica_ingreso_salida()

      MaskEdBox1.Enabled = True
      FechaFabric.Enabled = True
      If varform.MSFlexGrid1.Row <> 0 Then
        TxtArticulo.text = varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 0)
        TxtCantidad.text = varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 3)
        cantidadini = CDbl(TxtCantidad)
        Label14 = varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 4)
        xserie = varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 10)
        
        txccosto = varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 11)
        TxordFab = varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 12)
        txEquip = varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 13)
        
        VGabrev = Label14
        If xserie = "S" Then
             Text6 = varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 2)
             deshabilitartx5_tx3 (False)
        ElseIf xserie = "N" Then
             Text6 = varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 2)
             If varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 9) <> "" Then FechaFabric = varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 9)
             If varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 8) <> "" Then MaskEdBox1 = varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 8)
        Else
             flagserie = "N"
             flaglote = "N"
            ' ver_serie_lote
        End If
        lblUniEst = Nombre_Unidad(Label14)
        TxtUniRef = lblUniEst
        If VGRegEnt <> 1 Then TxtUniRef.Enabled = False
        lbcantstk = cant
        TxtArticulo.Enabled = False
        TxtArticulo.TabStop = False
        'formato de pantalla
        If varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 10) = "X" Then
                flaglote = "N"
                flagserie = "N"
                Text6.Enabled = False
                Text6.BackColor = &H8000000F
                MaskEdBox1.BackColor = &H8000000F
                FechaFabric.BackColor = &H8000000F
                MaskEdBox1.Enabled = False
                FechaFabric.Enabled = False
                TxtCantidad.Enabled = True
                TxtCantidad.SetFocus
        Else
                Text6.Enabled = True
                If varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 10) = "S" Then
                     flagserie = "S"
                    flaglote = "N"
                     MaskEdBox1.BackColor = &H8000000F
                     FechaFabric.BackColor = &H80000009
                     MaskEdBox1.Enabled = False
                     FechaFabric.Enabled = True
                     TxtCantidad = "1"
                     TxtCantidad.Enabled = False
                Else
                     flaglote = "S"
                     flagserie = "N"
                     MaskEdBox1.BackColor = &H80000009
                     FechaFabric.BackColor = &H80000009
                     MaskEdBox1.Enabled = True
                     FechaFabric.Enabled = True
                End If
               Text6.SetFocus
        End If
        'TxtCantidad.SetFocus
    End If
End Sub
Private Sub colocastk()
  Dim cadena As String
   cadena = varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 0)
   If varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 10) = "S" Then
        seriestk
   ElseIf varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 10) = "N" Then
        lotestk
   Else
        llenastk
  End If
End Sub

Private Sub llenastk()
Dim RSQL As String
Dim rs As New ADODB.Recordset
  
   RSQL = "select  stskdis, stskmin,stskmax,stpunrep from stkart  WHERE STALMA='" & VGAlma & "' and  stcodigo ='" & codigo & "'"
   'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
   Set rs = VGCNx.Execute(RSQL)
   If Not rs.EOF Then
     lbcantstk = rs(0)
   Else
     lbcantstk = 0
   End If
   rs.Close
End Sub

Private Sub lotestk()
Dim rs As New ADODB.Recordset
Dim RSQL As String
Dim Lote As String
   Lote = varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 2)
   RSQL = "select  STSLKDIS from STKLOTE where STSALMA ='" & VGAlma & "' and STSCODIGO = '" & codigo & "' and STSLOTE = '" & Lote & "'"
   'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
   Set rs = VGCNx.Execute(RSQL)
   If Not rs.EOF Then
        lbcantstk = rs(0)
     Else
        lbcantstk = 0
     End If
   
   rs.Close
End Sub

Private Sub seriestk()
Dim rs As New ADODB.Recordset
Dim RSQL As String
Dim Serie As String
   Serie = varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 2)
   RSQL = "select STSSKDIS from STKSERI where  STSALMA ='" & VGAlma & "' and STSCODIGO = '" & codigo & "' and STSSERIE = '" & Serie & "'"
   Set rs = VGCNx.Execute(RSQL)
   If Not rs.EOF Then
      lbcantstk = rs(0)
   Else
      lbcantstk = 0
   End If
   rs.Close
End Sub

Private Sub grabadetalle()
 Dim Adoreg1 As ADODB.Recordset
 Dim AdoReg2 As ADODB.Recordset
 Dim Rsql1 As String
 Dim criterio As String
 Dim item As Integer
' On Error GoTo GrabErr
  TxtArticulo = Trim(TxtArticulo)
  If VGSeleccion = 2 Then   'indica es el dato se modifica
     Rsql1 = "select * from movalmdet where DETD= '" & FrmModificar.TxDoc & "' AND DENUMDOC= '" & FrmModificar.Lblnumdoc & "'  AND DEALMA= '" & VGAlma & "' and  DEITEM = " & FrmModificar.contador & " and DECODIGO = '" & TxtArticulo & "'"
  Else
     Rsql1 = "select * from movalmdet where DETD='TT'"
  End If
  Set Adoreg1 = New ADODB.Recordset
  Adoreg1.Open Rsql1, VGCNx, adOpenDynamic, adLockOptimistic
  ' Si es nuevo adicciono los datos primary key
  If Adoreg1.RecordCount = 0 Then
        Adoreg1.AddNew
        Adoreg1("dealma") = VGAlma
Retor:
        Rsql1 = "select * from movalmdet where DETD= '" & FrmModificar.TxDoc & "' AND DENUMDOC= '" & FrmModificar.Lblnumdoc & "'  AND DEALMA= '" & VGAlma & "' and  DEITEM = " & FrmModificar.numitem & ""
        Set AdoReg2 = New ADODB.Recordset
        AdoReg2.Open Rsql1, VGCNx, adOpenDynamic, adLockOptimistic
        If Not AdoReg2.EOF Then
           FrmModificar.numitem = AdoReg2("deitem") + 1
           FrmModificar.contador = FrmModificar.numitem
           FormCreacion.LblNroReg = FrmModificar.numitem
           GoTo Retor
        End If
        AdoReg2.Close
        Adoreg1("deitem") = FrmModificar.numitem
        Adoreg1("DECODIGO") = Trim(TxtArticulo)   ' Format(MSFlexGrid1.TextMatrix(contador, 0), "00000000")
        Adoreg1("DEDESCRI") = Ctr_Ayuart.xnombre
        Adoreg1("detd") = FrmModificar.TxDoc
        Adoreg1("denumdoc") = FrmModificar.Lblnumdoc
        
        FrmModificar.numitem = FrmModificar.numitem + 1
  End If
  ' adicciono la nueva cantidad, serie y lote
     Adoreg1("decantid") = Val(TxtCantidad)        '
     CANTIDAD = Val(TxtCantidad)
     If xserie = "S" Then
          actserie
          Adoreg1("DESERIE") = Trim(Text6)
     ElseIf xserie = "N" Then
         grabalote
         Adoreg1("DELOTE") = Trim(Text6)
     End If
     '***********************************
        Adoreg1("DECENCOS") = txccosto
        Adoreg1("DEORDFAB") = TxordFab
        Adoreg1("DEQUIPO") = txEquip
     '***********************************
    
    Adoreg1.Update
    Set Adoreg1 = Nothing
    nuevodet = True   'para que no actualice dos veces
    actualizastk TxtArticulo  'actualizando dos veces serie lote
    ya_grabo_det = True
  
    Exit Sub
GrabErr:
    MsgBox Err.Description
    hubo_error = True
End Sub

Private Sub actualizastk(codigo As String)

Dim criterio As String
Dim canttemp As Double
Dim adoreg As ADODB.Recordset
Dim RSQL As String
  RSQL = "select * from STKART where  STCODIGO= '" & TxtArticulo & "' and STALMA = '" & VGAlma & "'  "
   Set adoreg = New ADODB.Recordset
   adoreg.Open RSQL, VGCNx, adOpenDynamic, adLockOptimistic
   If adoreg.RecordCount = 0 Then
     adoreg.AddNew
     adoreg("stalma") = VGAlma
     adoreg("stcodigo") = TxtArticulo
     adoreg.Update
   Else
     canttemp = adoreg("stskdis")
   End If
   adoreg.Close
   Set adoreg = Nothing
         RSQL = "Update STKART set stskdis= " & IIf(FrmModificar.tipo = "NI", canttemp + Val(TxtCantidad), canttemp - Val(TxtCantidad)) & " where  STCODIGO= '" & TxtArticulo & "' and STALMA = '" & VGAlma & "'"
         VGCNx.Execute RSQL
   ValMes
   nuevodet = False
End Sub

Private Sub reactualizastk(codigo As String)
Dim criterio As String
Dim canttemp As Double
Dim RSQL As String
Dim adoreg As ADODB.Recordset
'On Error GoTo ERR
   RSQL = "select * from STKART where  STCODIGO= '" & TxtArticulo & "' and STALMA = '" & VGAlma & "'  "
   Set adoreg = New ADODB.Recordset
   adoreg.Open RSQL, VGCNx, adOpenDynamic, adLockOptimistic
   If adoreg.RecordCount = 0 Then
     adoreg.AddNew
     adoreg("stalma") = VGAlma
     adoreg("stcodigo") = TxtArticulo
     adoreg.Update
   Else
    canttemp = adoreg("stskdis")
   End If
   
   RSQL = "Update STKART set stskdis=" & IIf(FrmModificar.tipo = "NI", canttemp + cantidadini, canttemp - cantidadini) & " where  STCODIGO= '" & TxtArticulo & "' and STALMA = '" & VGAlma & "'"
   VGCNx.Execute RSQL
   CANTIDAD = cantidadini
   If Not nuevodet Then  ' si no es nuevo tiene que actualizar la serie y lote
        If Text6 <> "" Then
          If xserie = "S" Then actserie  'solo descarga   o carga dependiendo el tipo
          If xserie = "N" Then actlote TxtArticulo  'solo desCarga
        End If
        ValMes             'reactualiza  el moremes
   End If
   adoreg.Close
   nuevodet = False
   Exit Sub
Err:
   MsgBox Err.Description
End Sub

Private Sub grabalote()
Dim uSql As String
Dim Lote As String
Dim nuevo_stk As Double
Dim RSQL As String
Dim rs As New ADODB.Recordset
Dim fecfab As Date
Dim fecven As Date
    On Error GoTo GrabErr
    
    RSQL = "select STSLKDIS FROM STKLOTE where  STSALMA= '" & VGAlma & "' and  STSCODIGO= '" & TxtArticulo & "' and STSLOTE= '" & Text6 & "'" '
    'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
    Set rs = VGCNx.Execute(RSQL)
    If Not rs.EOF Then
       nuevo_stk = IIf(FrmModificar.tipo = "NI", rs(0) + CANTIDAD, rs(0) - CANTIDAD)
       uSql = "Update STKLOTE set STSLKDIS = " & nuevo_stk & " WHERE  STSALMA='" & VGAlma & "' and STSCODIGO='" & codigo & "' AND STSLOTE='" & Text6 & "'"
    Else
        If MaskEdBox1 <> "__/__/____" And (FechaFabric = "__/__/____") Then
            uSql = "insert into STKLOTE (STSALMA,STSCODIGO,STSLOTE,STSLKDIS,STSFECFAB)  VALUES ('" & VGAlma & "','" & TxtArticulo & "','" & Text6 & "' ," & Val(TxtCantidad) & ",'" & DateSQL(FechaFabric) & "') "
        ElseIf MaskEdBox1 = "__/__/____" And FechaFabric <> "__/__/____" Then
            uSql = "insert into STKLOTE  (STSALMA,STSCODIGO,STSLOTE,STSLKDIS,STSFECVEN)VALUES ('" & VGAlma & "','" & TxtArticulo & "','" & Text6 & "' ," & Val(TxtCantidad) & " ,' ','" & DateSQL(MaskEdBox1) & "') "  'SIN FECFAB
        ElseIf MaskEdBox1 <> "__/__/____" And FechaFabric <> "__/__/____" Then
            uSql = "insert into STKLOTE  (STSALMA,STSCODIGO,STSLOTE,STSLKDIS,STSFECFAB,STSFECVEN)  VALUES ('" & VGAlma & "','" & TxtArticulo & "','" & Text6 & "' ," & Val(TxtCantidad) & " ,'" & DateSQL(FechaFabric) & "','" & DateSQL(MaskEdBox1) & "') "
        Else
            uSql = "insert into STKLOTE  (STSALMA,STSCODIGO,STSLOTE,STSLKDIS)  VALUES ('" & VGAlma & "','" & TxtArticulo & "','" & Text6 & "' ," & Val(TxtCantidad) & "') "
        End If
    
    End If
    rs.Close
    Set rs = Nothing
    VGCNx.Execute uSql
    Exit Sub
GrabErr:
    MsgBox Err.Description
    hubo_error = False
    
End Sub

Private Sub actserie()
Dim uSql As String
Dim Serie As String
Dim valor As Integer
Dim rs As New ADODB.Recordset
Dim RSQL As String
Dim fecfab As Date
Dim fecven As Date
   
On Error GoTo Err
    If Combo1.Visible Then
          Text6 = Combo1.text
    End If
    
    RSQL = "select STSSKDIS FROM STKSERI where   STSALMA= '" & VGAlma & "' and  STSCODIGO= '" & codigo & "' and STSSERIE= '" & Text6 & "'" '
    'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
    Set rs = VGCNx.Execute(RSQL)
    If Not rs.EOF Then
       valor = IIf(FrmModificar.tipo = "NI", 1, 0)
       uSql = "update STKSERI set STSSKDIS = " & valor & " WHERE  STSALMA='" & VGAlma & "' and STSCODIGO='" & codigo & "'AND STSSERIE='" & Text6 & "'"
    Else
       uSql = "insert into STKSERI (STSALMA,STSCODIGO,STSSERIE,STSSKDIS) VALUES ('" & VGAlma & "','" & codigo & "','" & Text6 & "' ,1) "
    End If
    rs.Close
    Set rs = Nothing
    VGCNx.Execute uSql
    Exit Sub
Err:
   MsgBox Err.Description, vbExclamation, "Aviso"
End Sub

Private Sub actlote(codigo As String)
Dim uSql As String
Dim nuevo_stk As Double
Dim RSQL As String
Dim rs As New ADODB.Recordset

    RSQL = "select STSLKDIS FROM STKLOTE where   STSALMA= '" & VGAlma & "' and  STSCODIGO= '" & TxtArticulo & "' and STSLOTE= '" & Text6 & "'" '
    'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
    Set rs = VGCNx.Execute(RSQL)
    If Not rs.EOF Then
         nuevo_stk = IIf(FrmModificar.tipo = "NI", rs(0) + CANTIDAD, rs(0) - CANTIDAD)
         uSql = "Update STKLOTE set STSLKDIS = " & nuevo_stk & " WHERE  STSALMA= '" & VGAlma & "' and STSCODIGO='" & codigo & "'AND STSLOTE='" & Text6 & "'"
          VGCNx.Execute uSql
    End If
     rs.Close
End Sub

Private Sub actvalmes()
 
  Dim criterio As String
  Dim Adoreg1 As ADODB.Recordset
  Dim RSQL As String
  Dim Cantent As Double
  Dim Cantsal As Double
  Dim mespro As String
  Dim uSql As String
  On Error GoTo Err
   mespro = Year(FrmModificar.DTPicker1) & Format(Month(FrmModificar.DTPicker1), "00")
   CANTIDAD = Val(TxtCantidad)
   RSQL = "select SMCANENT,SMCANSAL FROM MoResMes where  SMALMA = '" & VGAlma & "' and SMMESPRO = '" & mespro & "' AND  SMCODIGO= '" & Trim(TxtArticulo) & "'" '
   Set Adoreg1 = New ADODB.Recordset
   Adoreg1.Open RSQL, VGCNx, adOpenDynamic, adLockOptimistic
   If Adoreg1.RecordCount <> 0 Then
      If FrmModificar.tipo = "NI" Then
        Cantent = Adoreg1(0) + CANTIDAD
        uSql = "Update MoResMes set SMCANENT = " & Cantent & " where SMALMA='" & VGAlma & "'  and  SMCODIGO ='" & TxtArticulo & "' AND SMMESPRO ='" & mespro & "' "
       Else
        Cantsal = Adoreg1(1) - CANTIDAD
        uSql = "Update MoResMes set SMCANSAL = " & Cantsal & " where SMALMA='" & VGAlma & "' and   SMCODIGO ='" & TxtArticulo & "' AND SMMESPRO ='" & mespro & "' "
       End If
   Else
      If FrmModificar.tipo = "NI" Then
        Cantent = CANTIDAD
        Cantsal = 0
      Else
        Cantsal = CANTIDAD
        Cantent = 0
      End If
       uSql = "insert into MoResMes (SMALMA,SMCODIGO,SMMESPRO,SMCANENT,SMCANSAL,SMSALDOINI) VALUES ('" & VGAlma & "','" & TxtArticulo & "','" & mespro & "' ," & Cantent & "," & Cantsal & ",0) "
   End If
   VGCNx.Execute uSql
   Adoreg1.Close
   Exit Sub
Err:
    MsgBox Err.Description
End Sub


Public Function Validadato(pvalor) As Boolean
    Dim k As Integer
    Dim l As Integer
    Dim txt As String
    Dim compara As String
    
    Validadato = True
    
    txt = UCase(Trim(CStr(pvalor)))
    l = Len(Trim(txt))
    compara = "[?%$',#@" & Chr(34) & "*-+{}!¿¡]"
    For k = 1 To l
      If Mid(txt, k, 1) Like compara Then
          Validadato = False
          Exit For
      End If
    Next k

End Function

