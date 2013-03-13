VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmImpresionNotas 
   Caption         =   "Impresion de Notas "
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7830
   ScaleWidth      =   10545
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   7425
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   13097
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Relacion de Documentos"
      TabPicture(0)   =   "FrmImpresionNotas.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1(3)"
      Tab(0).Control(1)=   "Label5(6)"
      Tab(0).Control(2)=   "Ctr_Ayuda4"
      Tab(0).Control(3)=   "TDBGrid1"
      Tab(0).Control(4)=   "Combo2"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Documento"
      TabPicture(1)   =   "FrmImpresionNotas.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "oCrystalReport"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Panel"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdBotones(2)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdBotones(1)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   -68640
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   660
         Width           =   1665
      End
      Begin VB.CommandButton cmdBotones 
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
         Height          =   990
         Index           =   1
         Left            =   3420
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   6150
         Width           =   1110
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   2
         Left            =   5070
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   6150
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         Height          =   2985
         Left            =   315
         TabIndex        =   32
         Top             =   3060
         Width           =   9735
         Begin VB.TextBox Text2 
            Height          =   1635
            Left            =   150
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   42
            Top             =   510
            Width           =   9435
         End
         Begin VB.Frame Frame3 
            Height          =   675
            Left            =   120
            TabIndex        =   33
            Top             =   2190
            Width           =   9465
            Begin VB.TextBox Text1 
               BackColor       =   &H00C0FFFF&
               Enabled         =   0   'False
               ForeColor       =   &H00404040&
               Height          =   285
               Index           =   0
               Left            =   1410
               MaxLength       =   10
               TabIndex        =   37
               Top             =   240
               Width           =   1005
            End
            Begin VB.TextBox Text1 
               BackColor       =   &H00C0FFFF&
               Enabled         =   0   'False
               ForeColor       =   &H00404040&
               Height          =   285
               Index           =   1
               Left            =   3450
               MaxLength       =   2
               TabIndex        =   36
               Top             =   210
               Width           =   675
            End
            Begin VB.TextBox Text1 
               BackColor       =   &H00C0FFFF&
               Enabled         =   0   'False
               ForeColor       =   &H00404040&
               Height          =   285
               Index           =   2
               Left            =   5910
               MaxLength       =   10
               TabIndex        =   35
               Top             =   240
               Width           =   1005
            End
            Begin VB.TextBox Text1 
               BackColor       =   &H00C0FFFF&
               Enabled         =   0   'False
               ForeColor       =   &H00404040&
               Height          =   285
               Index           =   3
               Left            =   8280
               MaxLength       =   10
               TabIndex        =   34
               Top             =   210
               Width           =   1005
            End
            Begin VB.Label Label2 
               Caption         =   "IGV"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   0
               Left            =   2850
               TabIndex        =   41
               Top             =   240
               Width           =   495
            End
            Begin VB.Label Label2 
               Caption         =   "IMPORTE"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   1
               Left            =   360
               TabIndex        =   40
               Top             =   270
               Width           =   1035
            End
            Begin VB.Label Label2 
               Caption         =   "TOTAL IGV"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   2
               Left            =   4680
               TabIndex        =   39
               Top             =   270
               Width           =   1095
            End
            Begin VB.Label Label2 
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
               Height          =   210
               Index           =   3
               Left            =   7380
               TabIndex        =   38
               Top             =   270
               Width           =   675
            End
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "REFERENCIA"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   180
            TabIndex        =   43
            Top             =   180
            Width           =   9405
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2535
         Left            =   300
         TabIndex        =   1
         Top             =   480
         Width           =   9735
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   8130
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1320
            Width           =   1425
         End
         Begin VB.CommandButton cAyuda 
            Caption         =   "..."
            Height          =   285
            Left            =   3450
            TabIndex        =   3
            Top             =   2100
            Width           =   255
         End
         Begin VB.CheckBox chkInafecto 
            Alignment       =   1  'Right Justify
            Caption         =   "Inaf."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2490
            TabIndex        =   2
            Top             =   1740
            Width           =   735
         End
         Begin MSMask.MaskEdBox MBox1 
            Height          =   285
            Index           =   0
            Left            =   1740
            TabIndex        =   5
            Top             =   -330
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12648447
            Enabled         =   0   'False
            MaxLength       =   6
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "d/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   3
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   6240
            TabIndex        =   6
            Top             =   210
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12648447
            AutoTab         =   -1  'True
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox1 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   8580
            TabIndex        =   7
            Top             =   210
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12648447
            Enabled         =   0   'False
            MaxLength       =   8
            PromptChar      =   "_"
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda1 
            Height          =   345
            Left            =   1110
            TabIndex        =   8
            Top             =   930
            Width           =   8475
            _ExtentX        =   14949
            _ExtentY        =   609
            XcodMaxLongitud =   11
            xcodwith        =   800
            NomTabla        =   "vt_Cliente"
            TituloAyuda     =   "Ayuda de Clientes"
            ListaCampos     =   $"FrmImpresionNotas.frx":0038
            XcodCampo       =   "clientecodigo"
            XListCampo      =   "clienterazonsocial"
            ListaCamposDescrip=   "Codigo,Descripcion,Ruc,Direccion,Distrito,LimiteCred,Saldo,T,P,M,D"
            ListaCamposText =   $"FrmImpresionNotas.frx":011E
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda3 
            Height          =   315
            Left            =   5370
            TabIndex        =   9
            Top             =   2070
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   556
            XcodMaxLongitud =   2
            xcodwith        =   100
            NomTabla        =   "cc_conceptos"
            TituloAyuda     =   "Ayuda de Conceptos"
            ListaCampos     =   "conceptocodigo(1),conceptodescripcion(1)"
            XcodCampo       =   "conceptocodigo"
            XListCampo      =   "conceptodescripcion"
            ListaCamposDescrip=   "Codigo,Descripcion"
            ListaCamposText =   "conceptocodigo,conceptodescripcion"
         End
         Begin MSMask.MaskEdBox MBox 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   3
            EndProperty
            Height          =   300
            Index           =   3
            Left            =   5880
            TabIndex        =   10
            Top             =   1335
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   529
            _Version        =   393216
            ClipMode        =   1
            AllowPrompt     =   -1  'True
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox 
            Height          =   315
            Index           =   1
            Left            =   2940
            TabIndex        =   11
            Top             =   1320
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   3
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox 
            Height          =   315
            Index           =   2
            Left            =   3450
            TabIndex        =   12
            Top             =   1320
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   8
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox 
            Height          =   255
            Index           =   6
            Left            =   1110
            TabIndex        =   13
            Top             =   2100
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   3
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox 
            Height          =   255
            Index           =   7
            Left            =   1560
            TabIndex        =   14
            Top             =   2100
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   3
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox 
            Height          =   255
            Index           =   8
            Left            =   2100
            TabIndex        =   15
            Top             =   2100
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   8
            PromptChar      =   "_"
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda2 
            Height          =   315
            Left            =   6990
            TabIndex        =   16
            Top             =   1680
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   556
            XcodMaxLongitud =   3
            xcodwith        =   200
            NomTabla        =   "vt_vendedor"
            TituloAyuda     =   "Ayuda de Vendedores"
            ListaCampos     =   "vendedorcodigo(1),vendedornombres(1)"
            XcodCampo       =   "vendedorcodigo"
            XListCampo      =   "vendedornombres"
            ListaCamposDescrip=   "Codigo,Descripcion"
            ListaCamposText =   "vendedorcodigo,vendedornombres"
         End
         Begin MSMask.MaskEdBox MBox 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   3
            EndProperty
            Height          =   300
            Index           =   5
            Left            =   4770
            TabIndex        =   17
            Top             =   1695
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   529
            _Version        =   393216
            ClipMode        =   1
            AllowPrompt     =   -1  'True
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox 
            Height          =   315
            Index           =   4
            Left            =   1110
            TabIndex        =   18
            Top             =   1680
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Registro"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   4890
            TabIndex        =   31
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Cambio"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   7470
            TabIndex        =   30
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Planilla"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   480
            TabIndex        =   29
            Top             =   -300
            Width           =   1215
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000009&
            Index           =   0
            X1              =   30
            X2              =   9750
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   1
            X1              =   30
            X2              =   9720
            Y1              =   570
            Y2              =   570
         End
         Begin VB.Label Label3 
            Caption         =   "DETALLE DOCUMENTO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   210
            TabIndex        =   28
            Top             =   630
            Width           =   3795
         End
         Begin VB.Label Label4 
            Caption         =   "Cliente"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   27
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Fe. Emision"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   4860
            TabIndex        =   26
            Top             =   1380
            Width           =   1035
         End
         Begin VB.Label Label5 
            Caption         =   "Moneda"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   7290
            TabIndex        =   25
            Top             =   1380
            Width           =   915
         End
         Begin VB.Label Label5 
            Caption         =   "Documento"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   4
            Left            =   90
            TabIndex        =   24
            Top             =   2100
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "Concepto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   5
            Left            =   4290
            TabIndex        =   23
            Top             =   2130
            Width           =   1305
         End
         Begin VB.Label LblFecDoc 
            Height          =   285
            Left            =   3780
            TabIndex        =   22
            Top             =   2100
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.Label Label5 
            Caption         =   "Vendedor"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   6090
            TabIndex        =   21
            Top             =   1710
            Width           =   1605
         End
         Begin VB.Label Label5 
            Caption         =   "Fe. Vencimiento"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   3
            Left            =   3360
            TabIndex        =   20
            Top             =   1740
            Width           =   1395
         End
         Begin VB.Label Label5 
            Caption         =   "Importe"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   7
            Left            =   120
            TabIndex        =   19
            Top             =   1740
            Width           =   825
         End
      End
      Begin MSComctlLib.StatusBar Panel 
         Height          =   345
         Left            =   -450
         TabIndex        =   46
         Top             =   8625
         Width           =   11085
         _ExtentX        =   19553
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
      Begin Crystal.CrystalReport oCrystalReport 
         Left            =   450
         Top             =   5460
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   5475
         Left            =   -74820
         TabIndex        =   47
         Top             =   1500
         Width           =   9765
         _ExtentX        =   17224
         _ExtentY        =   9657
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
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
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
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda4 
         Height          =   315
         Left            =   -73800
         TabIndex        =   48
         Top             =   690
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   556
         XcodMaxLongitud =   3
         xcodwith        =   200
         NomTabla        =   "co_multiempresas"
         TituloAyuda     =   "Ayuda de Empresas"
         ListaCampos     =   "empresacodigo(1),empresadescripcion(1)"
         XcodCampo       =   "empresacodigo"
         XListCampo      =   "empresadescripcion"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "empresacodigo,empresadescripcion"
      End
      Begin VB.Label Label5 
         Caption         =   "Documento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   6
         Left            =   -69870
         TabIndex        =   51
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Empresa :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   -74700
         TabIndex        =   49
         Top             =   720
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmImpresionNotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adll As New dllgeneral.dll_general

Dim nLongicampo(6) As Integer
Dim rsdeta As New ADODB.Recordset
Dim wCabe(40)

Dim apedido As String
Dim aalmacen As String
Dim alista As String * 2
Dim rs As ADODB.Recordset

Private Sub chkInafecto_Click()
    Text1(0) = numero(MBox(4))
    If Len(Trim(Text1(1))) > 0 And chkInafecto.Value = 0 Then
        Text1(2) = Trim(numero(CDbl(Text1(0)) * CDbl(Text1(1)) / 100))
        Text1(3) = Trim(numero(CDbl(Text1(0)) + CDbl(Text1(2))))
    Else
        Text1(2) = "0"
        Text1(3) = Trim(numero(CDbl(numero(MBox(4)))))
    End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
Call adll.ListarEnTDBGRID(VGCNx, "vt_pedido", TDBGrid1, "pedidotipofac,pedidonrofact,pedidofechafact,clientecodigo,clienterazonsocial,pedidotiporefe,pedidonrorefe", "pedidofechafact", nLongicampo, " pedidotipofac='" & Left(Combo2.Text, 2) & "' and pedidocondicionfactura<>'1' and empresacodigo='" & VGparametros.empresacodigo & "'")

TDBGrid1.Columns(0).Width = 1000
TDBGrid1.Columns(1).Width = 1200
TDBGrid1.Columns(2).Width = 1200
TDBGrid1.Columns(3).Width = 1200
TDBGrid1.Columns(4).Width = 4000
TDBGrid1.Columns(5).Width = 1200


End Sub

Private Sub Ctr_Ayuda4_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
VGparametros.empresacodigo = Ctr_Ayuda4.xclave
Call adll.llenacombo(Combo2, "select * from cc_tipodocumento inner join cc_parametro on cc_tipodocumento.tdocumentocodigo=cc_parametro.tdocumentonotaabono or cc_tipodocumento.tdocumentocodigo=cc_parametro.tdocumentonotacargo or cc_tipodocumento.tdocumentocodigo=cc_parametro.tdocumentonotacarbo or cc_tipodocumento.tdocumentocodigo=cc_parametro.tdocumentonotaabobo where cc_parametro.empresacodigo='" & Ctr_Ayuda4.xclave & "'", VGCNx)
End Sub

Private Sub Form_Load()
MostrarForm Me, "C"
MBox1(1) = Format(Date, "DD/MM/YYYY")
  
Call Ctr_Ayuda1.Conexion(VGCNx)
Call Ctr_Ayuda2.Conexion(VGCNx)
Call Ctr_Ayuda3.Conexion(VGCNx)
Call Ctr_Ayuda4.Conexion(VGCNx)

'Combo2.Visible = False

MBox1(2) = Format(DatoTipoCambio(VGCNx, Date), "##0.00")
Text1(1) = (VGparametros.igv * 100)
cmdBotones(1).Picture = MDIPrincipal.ImageList2.ListImages.Item("Imprimir").Picture
cmdBotones(2).Picture = MDIPrincipal.ImageList2.ListImages.Item("Retornar").Picture
  
Set rs = VGCNx.Execute("select monedacodigo,monedadescripcion from gr_moneda")
Do While Not rs.EOF
    Combo1.AddItem rs!monedacodigo & "-" & rs!monedadescripcion
    rs.MoveNext
Loop


End Sub

Public Function CargarData() As Integer
Dim J As Integer
Dim regi As Long
Dim nsql As String
Dim ltipo As String
Dim lzona As String
Dim Previo As Double
Dim tinafecto As Double
Dim xserie As String * 3
Dim xfactu As String * 5
Dim xtipofac As String * 2
Dim fechasunat As Date
Dim tcargo As String
Dim RsSerie As New ADODB.Recordset
Dim acmd As New ADODB.Command
Dim asql As New ADODB.Recordset
Dim arbusca As New ADODB.Recordset
Dim existedoc As Integer
On Error GoTo vererror

'GrabarData = 0
existedoc = 0
'******** CABECERA DE MOVIMIENTO *****************
For J = 1 To 29
    wCabe(J) = ""
Next J
fechasunat = Date
apedido = MBox(6) & Trim(MBox(7) & MBox(8))

Set asql = VGCNx.Execute("select * from vt_pedido a inner join vt_cargo b on a.pedidonrofact=b.cargonumdoc where a.pedidotipofac='" & Left(Combo2.Text, 2) & "' and pedidonrofact='" & TDBGrid1.Columns(1) & "' and a.empresacodigo='" & VGparametros.empresacodigo & "' ")
'Set asql = VGCNx.Execute("select * from vt_pedido where pedidotipofac='" & MBox(6) & "' and pedidonrofact='" & Trim(MBox(7) & MBox(8)) & "' and empresacodigo='" & VGparametros.empresacodigo & "' ")
If asql.RecordCount > 0 Then
   existedoc = 1
   'apedido = Escadena(asql!pedidonumero)
   'wCabe(1) = Escadena(asql!puntovtacodigo)         'Escadena(asql!p)                       'Pto Venta
   'wCabe(2) = Escadena(asql!pedidonumero)           'Trim(MBox(1))                       'nro pedido
   MBox(1) = Left(Escadena(asql!pedidonrofact), 3)         'Trim(MBox(2))                        'nro factura
   MBox(2) = Right(Escadena(asql!pedidonrofact), 8)        'Trim(MBox(3))                         'nro boleta
   'wCabe(5) = Escadena(asql!pedidonrofact)          'Trim(MBox(4))                         'nro guia
   'wCabe(6) = 0      'MBox(5)                       'dscto gral
   'wCabe(7) = 0      'MBox(6)                       'dscto promocional
   'wCabe(8) = 0      'MBox(7)                       'dscto especial
   Combo1.ListIndex = CInt(asql!pedidomoneda) - 1    'adll.ComboDato(Combo1.Text)           'moneda
   MBox1(2) = asql!pedidotipcambio                       'tipo de cambio
   'wCabe(11) = CDbl(Escadena(asql!pedidolistaprec)) 'dllgeneral.ComboDato(Combo2.Text)       'lista de precios
  ' wCabe(12) = " "                                  'MBox(9)                      'mensajes
  ' wCabe(13) = Escadena(asql!modovtacodigo)         'dllgeneral.ComboDato(Combo3.Text)       'modo de venta
   MBox1(1) = asql!pedidofechafact                             'MBox(10)                     'fecha de atencion
   'wCabe(15) = Escadena(asql!formapagocodigo)       'dllgeneral.ComboDato(Combo4.Text)       'forma de pago
   Ctr_Ayuda1.xclave = asql!clientecodigo: Ctr_Ayuda1.Ejecutar                   'MBox(11)                     'cliente
   Ctr_Ayuda2.xclave = asql!vendedorcodigo: Ctr_Ayuda2.Ejecutar                   'MBox(12)                       'vendedor
   'wCabe(18) = 0    'MBox(13)                       'comision
   'wCabe(19) = Escadena(asql!almacencodigo)         'Ctr_Ayuda3.xclave        'MBox(14)                     'almacen
   'wCabe(20) = 0      'MBox(15)                     'otros gastos
   'wCabe(21) = "0"      'MBox(16)                   'nota pedido
   'wCabe(22) = "0"      'MBox(17)                   'orden de compra
   'wCabe(23) = Escadena(asql!pedidoautorizacion)    'dllgeneral.ComboDato(Combo5.Text)       'autorizacion
   'wCabe(24) = 0       'MBox(18)                    'dias pago
   'wCabe(25) = 0                           'Total Cantidad
   Text1(0) = asql!pedidototbruto    'Round(Text1(0), 2)                         'Total Bruto
   'wCabe(27) = 0    'MBox2(8)              'total fletes --T.D.
   MBox(4) = asql!pedidototneto
   Text1(2) = asql!pedidototimpuesto                        'Round(Text1(2), 2)          'Total Igv
   Text1(3) = asql!pedidototneto                        'Round(Text1(3), 2)         'Neto a Facturar
   wCabe(30) = Escadena(asql!pedidoentrega)    'MBox(19)                    'entrega pedido
   wCabe(31) = Escadena(asql!clienterazonsocial)  'MBox3(1)                    'nombre cliente
   wCabe(32) = Escadena(asql!clientedireccion)    'MBox3(3)                    'direccion
   wCabe(33) = Escadena(asql!ClienteRuc)  'MBox3(2)                    'ruc
   wCabe(34) = MBox(3)                  'Date                           'fechafactura
   wCabe(35) = 0                     'Total Descuentos Globales
   wCabe(36) = 0                     'Total Descuentos Cliente
   wCabe(37) = 0                     'Total Descuentos Oficina
   wCabe(38) = 0                     'Total Descuentos Item
   wCabe(39) = 0                     'Total Descuentos Linea
   wCabe(40) = 0                     'Total Descuentos x Promocion
   'fechasunat = IIf(IsNull(asql!pedidofechasunat), MBox(3), asql!pedidofechasunat)
   MBox(3) = asql!cargoapefecemi
   MBox(5) = asql!cargoapefecvct
   Text2.Text = ESNULO(asql!cargoaperefere, "")
   Ctr_Ayuda3.xclave = ESNULO(asql!conceptocodigo, "")
   
  End If

  Set asql = Nothing
Exit Function
vererror:
   If Err Then
      MsgBox Err.Number & "-" & Err.Description
      MsgBox "Comunicarse con Sistemas ...!!" & Chr(13) & Chr(10) & VGCNx.Errors(0).Number & "-" & VGCNx.Errors(0).Description
      Exit Function
      Resume
   End If
End Function

Private Sub cAyuda_Click()
 nAyuda = "": nDetalle = ""
 If Len(Trim(MBox(6))) > 0 And Len(Trim(MBox(7))) > 0 And Len(Trim(MBox(8))) > 0 Then
    SendKeys "{tab}"
    Exit Sub
 End If
 
 If adll.VerificaDatoExistente(VGCNx, "select * from vt_pedido where clientecodigo='" & Trim(Ctr_Ayuda1.xclave) & "' and empresacodigo='" & VGparametros.empresacodigo & "'") = 1 Then
       Dim gfiltra(2, 2) As String
       gfiltra(1, 1) = g_tipofac: gfiltra(1, 2) = "pedidonrofact"
       gfiltra(2, 1) = g_tipobol: gfiltra(2, 2) = "pedidonroboleta"
       FrmAyudaCli.TipoForma = 1
       FrmAyudaCli.BConexion = VGCNx   'cn
       FrmAyudaCli.Bdata = "0"
       FrmAyudaCli.BTabla = "vt_pedido"
       FrmAyudaCli.BCampos = "pedidotipofac as Tipo,pedidonrofact as Documento,pedidofecha as Fecha,pedidomoneda as Moneda,pedidototneto as Total"
       FrmAyudaCli.BOrden = "pedidofecha"
       FrmAyudaCli.BCondi = "clientecodigo='" & Ctr_Ayuda1.xclave & "' and empresacodigo='" & VGparametros.empresacodigo & "' and pedidotipofac='" & MBox(6).Text & "' "
       FrmAyudaCli.BFiltro = gfiltra
 Else
  If adll.VerificaDatoExistente(VGCNx, "select * from vt_cargo where clientecodigo='" & Trim(Ctr_Ayuda1.xclave) & "'  and cargoapecarabo='C' and isnull(cargoapeflgreg,0)<>1 and empresacodigo='" & VGparametros.empresacodigo & "'") = 1 Then
       Dim ffiltra(1, 1) As String
       ffiltra(1, 1) = g_tipofac: ffiltra(1, 2) = "cargonumdoc"
       FrmAyudaCli.TipoForma = 1
       FrmAyudaCli.BConexion = VGCNx   'cn
       FrmAyudaCli.Bdata = "0"
       FrmAyudaCli.BTabla = "vt_cargo"
       FrmAyudaCli.BCampos = "documentocargo as Tipo,cargonumdoc as Documento,cargoapefecemi as Fecha,monedacodigo as Moneda,cargoapeimpape as Total"
       FrmAyudaCli.BOrden = "cargoapefecemi"
       FrmAyudaCli.BCondi = "clientecodigo='" & Trim(Ctr_Ayuda1.xclave) & "' and empresacodigo='" & VGparametros.empresacodigo & "' and cargoapecarabo='C' and isnull(cargoapeflgreg,0)<>1"
       FrmAyudaCli.BFiltro = ffiltra
   Else
       nAyuda = "": nDetalle = ""
       MsgBox "No existen documentos pendientes...", vbInformation, MsgTitle
       Exit Sub
   End If
 End If
 FrmAyudaCli.Show 1
 If Len(Escadena(nAyuda)) > 0 Then
    MBox(6) = Escadena(nAyuda): MBox(7) = Left(Escadena(nDetalle), 3): MBox(8) = Right(Escadena(nDetalle), 8)
    LblFecDoc.Caption = nfecha
 End If
 nAyuda = "": nDetalle = ""

End Sub

Private Sub cmdBotones_Click(Index As Integer)
   Select Case Index
    Case 1
    If MsgBox("Desea Imprimir la Nota de Credito", vbYesNo + vbInformation, "AVISO") = vbYes Then Call ImprimirNota
    Case 2
    Call adll.ActivaTab(0, 1, SSTab1)
   End Select
End Sub

Private Sub ImprimirNota()
Dim arrform(1) As Variant, arrparm(6) As Variant
Dim NombreRep As String, CadOrden As String

arrparm(0) = VGCNx.DefaultDatabase
arrparm(1) = MBox(1).Text & MBox(2).Text
arrparm(2) = CDbl(Text1(0).Text)
arrparm(3) = CDbl(Text1(2).Text)
arrparm(4) = Ctr_Ayuda4.xclave
arrparm(5) = Left(Combo2.Text, 2)       'MBox(6).Text CODIGO DE DOCUMENTO REFERENCIA
arrform(0) = "letras='" & adll.NUMLET(numero(Round(CDbl(Text1(3)), 2))) & IIf(adll.ComboDato(Combo1.Text) = g_tiposol, "Nuevos Soles", "Dolares Americanos") & "'"

If adll.ComboDato(Combo2.Text) = "07" Then
   NombreRep = "RepNotaCredito_" & Ctr_Ayuda4.xclave & ".rpt"
   'NombreRep = VGparamsistem.RutaReport & "RepNotaCredito_" & Ctr_Ayuda4.xclave & ".rpt"
Else
   NombreRep = "RepNotaDebito_" & Ctr_Ayuda4.xclave & ".rpt"
   'NombreRep = VGparamsistem.RutaReport & "RepNotaDebito_" & Ctr_Ayuda4.xclave & ".rpt"
End If

Call ImpresionRptProc(NombreRep, arrform, arrparm, CadOrden, "Cuenta Corriente por Cliente")

End Sub

Private Sub TDBGrid1_DblClick()
CargarData
SSTab1.Tab = 1
Call adll.ActivaTab(1, 0, SSTab1)

End Sub


