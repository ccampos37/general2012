VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Begin VB.Form frECnpt 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edición de Conceptos"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   Icon            =   "frECnpt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   6105
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmCancela 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   3262
      TabIndex        =   13
      Top             =   5640
      Width           =   1530
   End
   Begin VB.CommandButton cmAcepta 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   1312
      TabIndex        =   12
      Top             =   5640
      Width           =   1530
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5340
      Left            =   105
      TabIndex        =   28
      Top             =   165
      Width           =   5850
      _ExtentX        =   10319
      _ExtentY        =   9419
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "frECnpt.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Afecto A"
      TabPicture(1)   =   "frECnpt.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Contabilidad"
      TabPicture(2)   =   "frECnpt.frx":0342
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame3"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame3 
         Caption         =   "Cuenta Debe"
         Height          =   2235
         Left            =   225
         TabIndex        =   52
         Top             =   480
         Width           =   5400
         Begin MSDataGridLib.DataGrid XdataDebe 
            Height          =   1830
            Left            =   180
            TabIndex        =   53
            Top             =   270
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   3228
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
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
            ColumnCount     =   6
            BeginProperty Column00 
               DataField       =   "SEC"
               Caption         =   "Sec"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "CUENTA"
               Caption         =   "Cuenta"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "TipAsi"
               Caption         =   "Tipo"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "TIPASINOM"
               Caption         =   "Desc de tipo  de Asiento"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "TIPOCTA"
               Caption         =   "TIPOCTA"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "CONCEPT"
               Caption         =   "CONCEPT"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   450.142
               EndProperty
               BeginProperty Column01 
                  Button          =   -1  'True
                  ColumnWidth     =   1454.74
               EndProperty
               BeginProperty Column02 
                  Button          =   -1  'True
                  ColumnWidth     =   420.095
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   2264.882
               EndProperty
               BeginProperty Column04 
               EndProperty
               BeginProperty Column05 
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Cuenta Haber"
         Height          =   2310
         Left            =   210
         TabIndex        =   41
         Top             =   2805
         Width           =   5400
         Begin MSDataGridLib.DataGrid XdataHaber 
            Height          =   1830
            Left            =   195
            TabIndex        =   54
            Top             =   315
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   3228
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            TabAction       =   2
            FormatLocked    =   -1  'True
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
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
            ColumnCount     =   6
            BeginProperty Column00 
               DataField       =   "SEC"
               Caption         =   "Sec"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "CUENTA"
               Caption         =   "Cuenta"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "TipAsi"
               Caption         =   "Tipo"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "TIPASINOM"
               Caption         =   "Desc de tipo  de Asiento"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "TIPOCTA"
               Caption         =   "TIPOCTA"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "CONCEPT"
               Caption         =   "CONCEPT"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   450.142
               EndProperty
               BeginProperty Column01 
                  Button          =   -1  'True
                  ColumnWidth     =   1454.74
               EndProperty
               BeginProperty Column02 
                  Button          =   -1  'True
                  ColumnWidth     =   420.095
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   2264.882
               EndProperty
               BeginProperty Column04 
               EndProperty
               BeginProperty Column05 
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Sumar a (Disponibles en el Cálculo de Planillas)"
         Height          =   3960
         Left            =   -74880
         TabIndex        =   39
         Top             =   435
         Width           =   5520
         Begin VB.CheckBox xIndVac 
            Caption         =   "No Ind."
            Height          =   210
            Left            =   4440
            TabIndex        =   47
            ToolTipText     =   "Concepto no indemnizable / Se considera su total"
            Top             =   3075
            Visible         =   0   'False
            Width           =   930
         End
         Begin VB.CheckBox xIndGra 
            Caption         =   "No Ind."
            Height          =   210
            Left            =   4440
            TabIndex        =   46
            ToolTipText     =   "Concepto no indemnizable / Se considera su total"
            Top             =   2745
            Visible         =   0   'False
            Width           =   930
         End
         Begin VB.CheckBox xIndCTS 
            Caption         =   "No Ind."
            Height          =   210
            Left            =   4440
            TabIndex        =   45
            ToolTipText     =   "Concepto no indemnizable / Se considera su total"
            Top             =   2415
            Visible         =   0   'False
            Width           =   930
         End
         Begin VB.ComboBox Combo3 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frECnpt.frx":035E
            Left            =   2550
            List            =   "frECnpt.frx":0383
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   3000
            Visible         =   0   'False
            Width           =   1845
         End
         Begin VB.ComboBox Combo2 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frECnpt.frx":041A
            Left            =   2550
            List            =   "frECnpt.frx":043F
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   2670
            Visible         =   0   'False
            Width           =   1845
         End
         Begin VB.ComboBox Combo1 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frECnpt.frx":04D9
            Left            =   2550
            List            =   "frECnpt.frx":04FE
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   2340
            Visible         =   0   'False
            Width           =   1845
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Remuneración Asegurable de AFP"
            Height          =   300
            Left            =   225
            TabIndex        =   14
            Top             =   345
            Width           =   3240
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Remuneración de Aportaciones a EsSalud"
            Height          =   300
            Left            =   225
            TabIndex        =   15
            Top             =   615
            Width           =   3585
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Remuneración afecta al I.E.S."
            Height          =   300
            Left            =   225
            TabIndex        =   16
            Top             =   885
            Width           =   3240
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Remuneración afecta a Quinta Categoria - Remuneración Ordinaria"
            Height          =   300
            Left            =   225
            TabIndex        =   18
            Top             =   1425
            Width           =   5160
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Remuneración afecta al S.C.T.R."
            Height          =   300
            Left            =   225
            TabIndex        =   17
            Top             =   1155
            Width           =   3240
         End
         Begin VB.CheckBox Check10 
            Caption         =   "Variable Auxiliar Total02"
            Height          =   300
            Left            =   225
            TabIndex        =   25
            Top             =   3630
            Width           =   2445
         End
         Begin VB.CheckBox Check11 
            Caption         =   "Variable Auxiliar Total03"
            Height          =   300
            Left            =   2715
            TabIndex        =   26
            Top             =   3615
            Width           =   2445
         End
         Begin VB.CheckBox Check12 
            Caption         =   "Remuneración afecta a Quinta Categoria - Promediable"
            Height          =   300
            Left            =   225
            TabIndex        =   20
            Top             =   1950
            Width           =   4335
         End
         Begin VB.CheckBox Check13 
            Caption         =   "Remuneración afecta a Quinta Categoria - Remunaración Variable"
            Height          =   300
            Left            =   225
            TabIndex        =   19
            Top             =   1695
            Width           =   5040
         End
         Begin VB.CheckBox Check9 
            Caption         =   "Variable Auxiliar Total01"
            Height          =   300
            Left            =   225
            TabIndex        =   24
            Top             =   3315
            Width           =   2445
         End
         Begin VB.CheckBox Check8 
            Caption         =   "Provisiones de Vacaciones"
            Height          =   300
            Left            =   225
            TabIndex        =   23
            Top             =   3007
            Width           =   2445
         End
         Begin VB.CheckBox Check7 
            Caption         =   "Provisiones de Gratificación"
            Height          =   300
            Left            =   225
            TabIndex        =   22
            Top             =   2677
            Width           =   2445
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Provisiones de C.T.S."
            Height          =   300
            Left            =   225
            TabIndex        =   21
            Top             =   2340
            Width           =   2445
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000005&
            Index           =   1
            X1              =   165
            X2              =   5385
            Y1              =   2280
            Y2              =   2280
         End
         Begin VB.Line Line1 
            BorderColor     =   &H8000000C&
            Index           =   0
            X1              =   150
            X2              =   5370
            Y1              =   2265
            Y2              =   2265
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Datos Generales"
         Height          =   4680
         Left            =   -74850
         TabIndex        =   27
         Top             =   540
         Width           =   5490
         Begin VB.CheckBox Xpermite 
            Caption         =   "Permitir grabar valores <=0"
            Height          =   210
            Left            =   1545
            TabIndex        =   51
            Top             =   4350
            Width           =   3735
         End
         Begin VB.CheckBox xImpresionFija 
            Caption         =   "Al imprimir se encuentra en una posición fija"
            Height          =   210
            Left            =   1575
            TabIndex        =   11
            Top             =   3645
            Width           =   3735
         End
         Begin VB.CommandButton Command1 
            Caption         =   "..."
            Height          =   240
            Left            =   1290
            TabIndex        =   40
            Top             =   1650
            Width           =   270
         End
         Begin AplisetControlText.Aplitext xEnlace 
            Height          =   300
            Left            =   1575
            TabIndex        =   10
            Top             =   3300
            Visible         =   0   'False
            Width           =   3795
            _ExtentX        =   6694
            _ExtentY        =   529
            Locked          =   -1  'True
            Text            =   ""
         End
         Begin VB.ComboBox xTipoRemu 
            Height          =   315
            ItemData        =   "frECnpt.frx":058E
            Left            =   1575
            List            =   "frECnpt.frx":059B
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   2265
            Width           =   3795
         End
         Begin VB.ComboBox xTipoInfo 
            Height          =   315
            ItemData        =   "frECnpt.frx":05EA
            Left            =   1575
            List            =   "frECnpt.frx":0600
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   1935
            Width           =   3795
         End
         Begin AplisetControlText.Aplitext xFila 
            Height          =   315
            Left            =   4155
            TabIndex        =   2
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            Text            =   "0"
            TipoDato        =   "N"
         End
         Begin AplisetControlText.Aplitext xcolplanilla 
            Height          =   315
            Left            =   1575
            TabIndex        =   8
            Top             =   2610
            Width           =   3795
            _ExtentX        =   6694
            _ExtentY        =   556
            Locked          =   -1  'True
            Text            =   ""
            Seleccionar     =   0   'False
         End
         Begin AplisetControlText.Aplitext xFormula 
            Height          =   525
            Left            =   1575
            TabIndex        =   5
            Top             =   1380
            Width           =   3795
            _ExtentX        =   6694
            _ExtentY        =   926
            MaxLength       =   250
            Text            =   ""
         End
         Begin AplisetControlText.Aplitext xNombre 
            Height          =   315
            Left            =   1575
            TabIndex        =   1
            Top             =   700
            Width           =   3795
            _ExtentX        =   6694
            _ExtentY        =   556
            Text            =   ""
         End
         Begin AplisetControlText.Aplitext xCodigo 
            Height          =   315
            Left            =   1575
            TabIndex        =   0
            Top             =   360
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   556
            MaxLength       =   8
            Text            =   ""
            Seleccionar     =   0   'False
            SinBlancos      =   -1  'True
            TipoCodigo      =   -1  'True
         End
         Begin VB.CheckBox xEsEscrito 
            Caption         =   "Valor será escrito"
            Height          =   225
            Left            =   3855
            TabIndex        =   4
            Top             =   1080
            Width           =   1515
         End
         Begin VB.ComboBox xMoneda 
            Height          =   315
            ItemData        =   "frECnpt.frx":06E7
            Left            =   1575
            List            =   "frECnpt.frx":06F1
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   2955
            Width           =   1860
         End
         Begin VB.ComboBox xTipo 
            Height          =   315
            ItemData        =   "frECnpt.frx":070C
            Left            =   1575
            List            =   "frECnpt.frx":071C
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1040
            Width           =   1860
         End
         Begin AplisetControlText.Aplitext xComentario 
            Height          =   300
            Left            =   1560
            TabIndex        =   49
            Top             =   3900
            Width           =   3810
            _ExtentX        =   6720
            _ExtentY        =   529
            MaxLength       =   250
            Text            =   ""
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Adicional"
            Height          =   195
            Left            =   180
            TabIndex        =   48
            Top             =   3960
            Width           =   645
         End
         Begin VB.Label l1 
            AutoSize        =   -1  'True
            Caption         =   "Enlazar con"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   38
            Top             =   3375
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label l1 
            AutoSize        =   -1  'True
            Caption         =   "Fila"
            Height          =   195
            Index           =   6
            Left            =   3780
            TabIndex        =   37
            Top             =   420
            Width           =   240
         End
         Begin VB.Label l1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Remun."
            Height          =   195
            Index           =   7
            Left            =   180
            TabIndex        =   36
            Top             =   2325
            Width           =   1140
         End
         Begin VB.Label l1 
            Caption         =   "Considerar como"
            Height          =   240
            Index           =   5
            Left            =   180
            TabIndex        =   35
            Top             =   1972
            Width           =   1275
         End
         Begin VB.Label l1 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   34
            Top             =   435
            Width           =   495
         End
         Begin VB.Label l1 
            AutoSize        =   -1  'True
            Caption         =   "Nombre"
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   33
            Top             =   750
            Width           =   555
         End
         Begin VB.Label l1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Concepto"
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   32
            Top             =   1095
            Width           =   1275
         End
         Begin VB.Label l1 
            AutoSize        =   -1  'True
            Caption         =   "Valor del Rubro"
            Height          =   195
            Index           =   4
            Left            =   180
            TabIndex        =   31
            Top             =   1380
            Width           =   1095
         End
         Begin VB.Label l1 
            AutoSize        =   -1  'True
            Caption         =   "Columna Planilla"
            Height          =   195
            Index           =   8
            Left            =   180
            TabIndex        =   30
            Top             =   2670
            Width           =   1155
         End
         Begin VB.Label l1 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Index           =   9
            Left            =   180
            TabIndex        =   29
            Top             =   3015
            Width           =   585
         End
      End
   End
   Begin VB.TextBox xMensaje 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   885
      Left            =   615
      MultiLine       =   -1  'True
      TabIndex        =   50
      Text            =   "frECnpt.frx":074A
      Top             =   4665
      Visible         =   0   'False
      Width           =   4860
   End
End
Attribute VB_Name = "frECnpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSCNPT As New ADODB.Recordset
'NO SQL
Dim WithEvents RSCTADEBE As ADODB.Recordset
Attribute RSCTADEBE.VB_VarHelpID = -1
Dim WithEvents RSCTAHABER As ADODB.Recordset
Attribute RSCTAHABER.VB_VarHelpID = -1
Dim FLAGMOV As Boolean
Dim TECLA As Long

Private Sub Command2_Click()
    'RSCTADEBE.FIELDS("CUENTA").VALUE = TEXT1.TEXT
End Sub

Private Sub CHECK6_CLICK()
    If Check6.Value = 1 Then
        Combo1.ListIndex = 1
        Combo1.Enabled = True
    Else
        Combo1.ListIndex = 0
        Combo1.Enabled = False
    End If
End Sub

Private Sub CHECK7_CLICK()
    If Check7.Value = 1 Then
        Combo2.ListIndex = 1
        Combo2.Enabled = True
    Else
        Combo2.ListIndex = 0
        Combo2.Enabled = False
    End If
End Sub

Private Sub CHECK8_CLICK()
    If Check8.Value = 1 Then
        Combo3.ListIndex = 1
        Combo3.Enabled = True
    Else
        Combo3.ListIndex = 0
        Combo3.Enabled = False
    End If
End Sub

Private Sub cmAcepta_Click()
  If Not VERIFICADATOS Then Exit Sub
  With RSCNPT
    If UCase(VPTAREA) = "NUEVO" Then
        Dim X As Integer
        DBSYSTEM.Execute "UPDATE CONCEPTOS SET FILA=FILA WHERE CODIGO='" & xCodigo.Text & "'", X
        If X <> 0 Then
            MsgBox "EL CóDIGO INGRESADO YA EXISTE, CAMBIELO Y ACEPTE DE NUEVO", vbCritical
            xCodigo.SetFocus
            Exit Sub
        End If
        .AddNew
    End If
    !Codigo = xCodigo.Text
    !NOMBRE = Trim(Left(xNombre.Text & Space(25), 25))
    !TIPO = xTipo.ListIndex
    !ESESCRITO = IIf(xEsEscrito.Value = 1, True, False)
    !FORMULA = IIf(xEsEscrito.Value = 1, "", xFormula.Text)
    !TIPOINFO = xTipoInfo.ListIndex
    !TIPOREMU = xTipoRemu.ListIndex
    !COLPLANILLA = xcolplanilla.Tag
    !Moneda = xMoneda.ListIndex
    !FILA = Val(xFila.Text)
    !SUMAAFP = IIf(Check1.Value = 1, True, False)
    !SUMASALUD = IIf(Check2.Value = 1, True, False)
    !SUMAIES = IIf(Check3.Value = 1, True, False)
    !SUMARENTA = IIf(Check4.Value = 1, True, False)
    !SUMASCTR = IIf(Check5.Value = 1, True, False)
    !SUMACTS = IIf(Check6.Value = 1, True, False)
    !SUMAGRAT = IIf(Check7.Value = 1, True, False)
    !SUMAVAC = IIf(Check8.Value = 1, True, False)
    !SUMAT1 = IIf(Check9.Value = 1, True, False)
    !SUMAT2 = IIf(Check10.Value = 1, True, False)
    !SUMAT3 = IIf(Check11.Value = 1, True, False)
    !SUMAT4 = IIf(Check12.Value = 1, True, False)
    !SUMAT5 = IIf(Check13.Value = 1, True, False)
    !IMPRESIONFIJA = xImpresionFija.Value
    !PERMITE = IIf(Xpermite.Value = 1, True, False)
    !TIPOCTS = Combo1.ListIndex
    !TIPOVAC = Combo2.ListIndex
    !TIPOGRA = Combo3.ListIndex
    !INDCTS = xIndCTS.Value
    !INDGRA = xIndGra.Value
    !INDVAC = xIndVac.Value
    !COMENTARIO = "" & xComentario.Text & " "
    If xTipo.ListIndex <> 1 Then xEnlace.Tag = ""
    !ENLACE = xEnlace.Tag
    .Update
  End With
  cmCancela.Enabled = True
On Error Resume Next
  RSCTADEBE.UpdateBatch
  RSCTAHABER.UpdateBatch
  RSCTADEBE.Update
  RSCTAHABER.Update
  Unload Me
End Sub

Private Sub CMCANCELA_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    frmHelpTmp.Show 1
End Sub

Private Sub Form_Load()
    SSTab1.Tab = 0
    xTipo.ListIndex = 0
    xMoneda.ListIndex = 0
    xTipoInfo.ListIndex = 0
    xTipoRemu.ListIndex = 0
    Combo1.ListIndex = 0
    Combo2.ListIndex = 0
    Combo3.ListIndex = 0
    RSCNPT.Open "CONCEPTOS", DBSYSTEM, adOpenKeyset, adLockOptimistic
    If UCase(VPTAREA) = "EDITAR" Then
        xCodigo.Locked = True
        RSCNPT.FIND "CODIGO='" & VPCODTMP & "'"
        If RSCNPT.EOF Then
            MsgBox "EL REGISTRO YA NO EXISTE. HA SIDO ELIMINADO POR OTRO USUARIO O SE ENCUENTRA EN EDICIóN", vbCritical
            Unload Me
        Else
            CARGADATOS
        End If
    End If
    'NO SQL
    FLAGMOV = False
    
    Set RSCTADEBE = New ADODB.Recordset
    Set RSCTAHABER = New ADODB.Recordset
    RSCTADEBE.Open "SELECT * FROM CTACONCEPTO WHERE TIPOCTA='D' AND CONCEPT='" & VPCODTMP & "'", DBSYSTEM, adOpenDynamic, adLockBatchOptimistic
    RSCTAHABER.Open "SELECT * FROM CTACONCEPTO WHERE TIPOCTA='H' AND CONCEPT='" & VPCODTMP & "'", DBSYSTEM, adOpenDynamic, adLockBatchOptimistic
    Set XdataDebe.DataSource = RSCTADEBE
    Set XdataHaber.DataSource = RSCTAHABER
    FLAGMOV = True
    XdataHaber.Columns("TIPOCTA").Width = 0
    XdataHaber.Columns("CONCEPT").Width = 0
    XdataDebe.Columns("TIPOCTA").Width = 0
    XdataDebe.Columns("CONCEPT").Width = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xMensaje.Visible = False
End Sub

Private Sub FORM_QUERYUNLOAD(CANCEL As Integer, UNLOADMODE As Integer)
    If Not cmCancela.Enabled Then CANCEL = 1
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSCNPT = Nothing
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xMensaje.Visible = False
End Sub

Private Sub RSCTADEBE_MoveComplete(ByVal ADREASON As ADODB.EventReasonEnum, ByVal PERROR As ADODB.Error, ADSTATUS As ADODB.EventStatusEnum, ByVal PRECORDSET As ADODB.Recordset)
    'NO SQL
    On Error GoTo ERRMOV
    If Not FLAGMOV Then Exit Sub
    XdataDebe.Columns("TIPOCTA") = "D"
    XdataDebe.Columns("CONCEPT").Value = Trim(xCodigo.Text)
    Exit Sub
ERRMOV:
    Exit Sub
End Sub

Private Sub RSCTAHABER_MoveComplete(ByVal ADREASON As ADODB.EventReasonEnum, ByVal PERROR As ADODB.Error, ADSTATUS As ADODB.EventStatusEnum, ByVal PRECORDSET As ADODB.Recordset)
On Error GoTo ERRMOV
    If Not FLAGMOV Then Exit Sub
    XdataHaber.Columns("TIPOCTA") = "H"
    XdataHaber.Columns("CONCEPT").Value = Trim(xCodigo.Text)
    Exit Sub
ERRMOV:
    Exit Sub
End Sub
Private Sub SSTAB1_Click(PREVIOUSTAB As Integer)
    If xTipo.ListIndex <> 1 Then
        Frame2.Enabled = False
    Else
        Frame2.Enabled = True
    End If
End Sub

Private Sub xCodigo_LostFocus()
    Dim CADENA As String
    CADENA = "'CODTRAB','NOMBRES','CODAREA','CODCCOSTO','BASICO','ASIGFAM','CODAFP','TASASCTR','APOROBL','SEGURO','TOPESEGURO'," & _
             "'COMISIONRA','SUMAAFP','SUMASALUD','TOTING','TOTEGR','_HORAST','_HOREXTRAS','_QUINTACAT','SUMAIES','SUMARENTA'," & _
             "'SUMASCTR','SUMACTS','SUMAGRAT','SUMAVAC','T1','T2','T3','T4','T5','OTROSING','OTROSEGR','ADELANTO','UBIGEO'," & _
             "'SEXO','TIPOTRAB','FECHAING','SITUACION','CARGO','BANCO','ESSALUDVIDA','RUCEPS','NOPDT','OPCION01','OPCION02'," & _
             "'OPCIONA','OPCIONB','XREDONDEO','AFECTOQUINTA'"
    If InStr(CADENA, "'" & Trim(xCodigo.Text) & "'") > 0 Then
        MsgBox "EL CODIGO : " & xCodigo.Text & " ES PALABRA RESERVADA DEL SISTEMA ", vbExclamation
        xCodigo.SetFocus
        Exit Sub
    End If
'    If Len(Trim(xCodigo.Text)) >= 7 Then
'        If Mid(xCodigo.Text, 1, 2) = "XX" And Mid(xCodigo.Text, 7, 1) = "X" Then
'            MsgBox "El codigo es reservado del sistema", vbExclamation
'            xCodigo.SetFocus
'            Exit Sub
'        End If
'    End If
End Sub

Private Sub XCOLPLANILLA_DBLCLICK()
    Dim RSCOLPLANILLA As New ADODB.Recordset
    RSCOLPLANILLA.Open "SELECT CODIGO, NOMBRE FROM COLUMPL ORDER BY NOMBRE", DBSYSTEM, adOpenKeyset, adLockOptimistic
    frmComun.CONECTAR RSCOLPLANILLA
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xcolplanilla.Text = VGUTIL(1) & " :  " & VGUTIL(2)
        xcolplanilla.Tag = VGUTIL(1)
    End If
    RSCOLPLANILLA.Close
    Set RSCOLPLANILLA = Nothing
End Sub

Private Sub XCOLPLANILLA_KEYPRESS(KeyAscii As Integer)
    If KeyAscii = 32 Then
        xcolplanilla.Text = ""
        xcolplanilla.Tag = ""
    End If
End Sub


Private Sub XCOMENTARIO_GOTFOCUS()
    xMensaje.Visible = False
End Sub

Private Sub XCOMENTARIO_LOSTFOCUS()
    xMensaje.Visible = False
End Sub

Private Sub XCOMENTARIO_MOUSEDOWN(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xMensaje.Visible = False
End Sub
Private Sub XCOMENTARIO_MOUSEMOVE(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xMensaje.Visible = True
End Sub

Private Sub XdataDebe_ButtonClick(ByVal COLINDEX As Integer)
    Dim DESCAUX As String
    XdataHaber.SetFocus
    Screen.MousePointer = 13
    Select Case COLINDEX
        Case 1
            If REGSISTEMA.scTieneStConta Then
                XdataDebe.Columns("CUENTA").Text = SELCUENTA(XdataDebe.Columns("CUENTA").Text)
            End If
        Case 2
            XdataDebe.Columns(2).Text = SELTIPOASIS(XdataDebe.Columns(2).Text, DESCAUX)
            If DESCAUX <> "" Then XdataDebe.Columns(3).Text = DESCAUX
    End Select
    XdataDebe.SetFocus
    Screen.MousePointer = 1
End Sub
Private Function SELCUENTA(TEXTO As String) As String
    Dim RSAUX As New ADODB.Recordset
    Set RSAUX = New ADODB.Recordset
    RSAUX.Open "SELECT PLANCTA_CODIGO,PLANCTA_DESCRIPCION FROM PLAN_CUENTA_NACIONAL WHERE PLANCTA_NIVEL=" & REGSISTEMA.scNivelCta, CONECTARDBSQL(REGSISTEMA.scRutaEmpresaWenco & "BDCONTABILIDAD"), adOpenKeyset, adLockReadOnly
    VGUTIL(1) = ""
    frmComun.CONECTAR RSAUX
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        SELCUENTA = VGUTIL(1)
      Else
        SELCUENTA = TEXTO
    End If
End Function

Private Sub XdataHaber_ButtonClick(ByVal COLINDEX As Integer)
    Dim DESCAUX As String
    Screen.MousePointer = 13
    XdataDebe.SetFocus
    Select Case COLINDEX
        Case 1
            If REGSISTEMA.scTieneStConta Then
                XdataHaber.Columns("CUENTA").Text = SELCUENTA(XdataHaber.Columns("CUENTA").Text)
            End If
        Case 2
            XdataHaber.Columns(2).Text = SELTIPOASIS(XdataHaber.Columns(2).Text, DESCAUX)
            If DESCAUX <> "" Then XdataHaber.Columns(3).Text = DESCAUX
    End Select
    XdataHaber.SetFocus
    Screen.MousePointer = 1
End Sub
Private Function SELTIPOASIS(TEXTO As String, Optional ByRef DESC As String) As String
'NO SQL
    Dim RSTIP As New ADODB.Recordset
    Dim CAMPOS As Variant
    RSTIP.Fields.Append "COD", adInteger
    RSTIP.Fields.Append "DESC", adVarChar, 25
    CAMPOS = Array("COD", "DESC")
    RSTIP.Open
    RSTIP.AddNew CAMPOS, Array("1", "SIMPLE")
    RSTIP.AddNew CAMPOS, Array("2", "POR TRABAJADOR")
    RSTIP.AddNew CAMPOS, Array("3", "POR CENTRO DE COSTOS")
    RSTIP.AddNew CAMPOS, Array("4", "POR A.F.P.")
    RSTIP.AddNew CAMPOS, Array("5", "POR TRABAJADOR Y C.C.")
    VGUTIL(1) = ""
    RSTIP.Filter = "COD='1' OR COD='2' OR COD='3'"
    frmComun.CONECTAR RSTIP
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        SELTIPOASIS = VGUTIL(1)
        DESC = VGUTIL(2)
      Else
        SELTIPOASIS = TEXTO
        DESC = ""
    End If
End Function

Private Sub XENLACE_DBLCLICK()
    Dim RSRUBROS As New ADODB.Recordset
    RSRUBROS.Open "SELECT CODIGO, NOMBRE FROM CONCEPTOS WHERE TIPO=0 ORDER BY NOMBRE", DBSYSTEM, adOpenStatic
    If RSRUBROS.RecordCount = 0 Then
        MsgBox "NO EXISTEN RUBROS INFORMATIVOS. NO SE PODRá UTILIZAR ESTA FUNCIóN APROPIADAMENTE, CREE NUEVOS RUBROS INFORMATIVOS E INTENTE OTRA VEZ", vbCritical
        Set RSRUBROS = Nothing
        Exit Sub
    End If
    frmComun.CONECTAR RSRUBROS
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xEnlace.Text = VGUTIL(1) & ": " & VGUTIL(2)
        xEnlace.Tag = VGUTIL(1)
    End If
    Set RSRUBROS = Nothing
End Sub

Private Sub XENLACE_KEYPRESS(KeyAscii As Integer)
    If KeyAscii = 32 Then
        If MsgBox("REALMENTE DESEA ELIMINAR EL CONCEPTO ENLAZADO", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        xEnlace.Text = ""
        xEnlace.Tag = ""
    End If
End Sub
Private Sub XESESCRITO_CLICK()
    If xEsEscrito.Value = 1 Then
        xFormula.Visible = False
    Else
        xFormula.Visible = True
    End If
End Sub

Public Sub CARGADATOS()
    With RSCNPT
        xCodigo.Text = "" & !Codigo
        If xCodigo.Text = "" Or xCodigo.Text = "NEWCODE" Then xCodigo.Locked = False
        xNombre.Text = "" & !NOMBRE
        xTipo.ListIndex = !TIPO
        xEsEscrito.Value = IIf(!ESESCRITO, 1, 0)
        xFormula.Text = "" & !FORMULA
        xcolplanilla.Tag = "" & !COLPLANILLA
        If Not IsNull(!COLPLANILLA) Or !COLPLANILLA = "''" Then xcolplanilla.Text = "" & !COLPLANILLA & ": " & DevuelveValor("SELECT NOMBRE FROM COLUMPL WHERE CODIGO='" & !COLPLANILLA & "'", DBSYSTEM)
        xMoneda.ListIndex = ESNULO(!Moneda, 0)
        xFila.Text = 0 & !FILA
        xTipoInfo.ListIndex = ESNULO(!TIPOINFO, 0)
        xTipoRemu.ListIndex = ESNULO(!TIPOREMU, 0)
        Combo1.ListIndex = 0
        Combo2.ListIndex = 0
        Combo3.ListIndex = 0
        Check1.Value = IIf(ESNULO(!SUMAAFP, False), 1, 0)
        Check2.Value = IIf(ESNULO(!SUMASALUD, False), 1, 0)
        Check3.Value = IIf(ESNULO(!SUMAIES, False), 1, 0)
        Check4.Value = IIf(ESNULO(!SUMARENTA, False), 1, 0)
        Check5.Value = IIf(ESNULO(!SUMASCTR, False), 1, 0)
        Check6.Value = IIf(ESNULO(!SUMACTS, False), 1, 0)
        Check7.Value = IIf(ESNULO(!SUMAGRAT, False), 1, 0)
        Check8.Value = IIf(ESNULO(!SUMAVAC, False), 1, 0)
        Check9.Value = IIf(ESNULO(!SUMAT1, False), 1, 0)
        Check10.Value = IIf(ESNULO(!SUMAT2, False), 1, 0)
        Check11.Value = IIf(ESNULO(!SUMAT3, False), 1, 0)
        Check12.Value = IIf(ESNULO(!SUMAT4, False), 1, 0)
        Check13.Value = IIf(ESNULO(!SUMAT5, False), 1, 0)
        xIndCTS.Value = 0
        xIndGra.Value = 0
        xIndVac.Value = 0
        xEnlace.Text = "" & !ENLACE
        xImpresionFija = IIf(ESNULO(!IMPRESIONFIJA, False), 1, 0)
        Xpermite.Value = IIf(ESNULO(!PERMITE, False), 1, 0)
        If !FLAG = 1 Then
            xTipo.Enabled = False
            xEsEscrito.Enabled = False
            xTipoInfo.Enabled = False
            xTipoRemu.Enabled = False
            xEnlace.Visible = False
        End If
        xComentario.Text = "" & !COMENTARIO
    End With
End Sub

Public Function VERIFICADATOS() As Boolean
    VERIFICADATOS = False
    If xCodigo.Text = "" Then
        MsgBox "CODIGO NO VALIDO, DEBERá CONTENER UN NOMBRE DE VARIABLE VALIDO.", vbCritical
        xCodigo.SetFocus
        Exit Function
    End If
    If xNombre.Text = "" Then
        MsgBox "DEBERá INGRESAR UN NOMBRE DE CONCEPTO VALIDO", vbCritical
        xNombre.SetFocus
        Exit Function
    End If
    If xTipo.ListIndex = -1 Then
        MsgBox "DEBERá SELECCIONAR UN TIPO DE CONCEPTO VáLIDO", vbCritical
        xTipo.SetFocus
        Exit Function
    End If
    If xEsEscrito.Value = 0 And xFormula.Text = "" Then
        MsgBox "DEBERá INGRESAR UNA FORMULA DE ACCIóN VáLIDA. LA FORMULA DEBERá SER EN VISUALBASIC SCRIPT O JAVASCRIPT-ANTECEDIDO DEL CODIGO JAVA", vbCritical
        xFormula.SetFocus
        Exit Function
    End If
    If xMoneda.ListIndex = -1 Then
        MsgBox "DEBERá SELECCIONAR EL TIPO DE MONEDA PARA EL RUBRO", vbCritical
        xMoneda.SetFocus
        Exit Function
    End If
    If Val(xFila.Text) = 0 Then
        MsgBox "DEBERá FIJAR UN ORDEN DE FILA PARA EL RUBRO", vbCritical
        xFila.SetFocus
        Exit Function
    End If
    If Combo1.ListIndex = 0 And Check6.Value = 1 Then
        MsgBox "NO SE PUEDE PROVISIONAR EL VALOR DE ESTE INGRESO, SI HA SETEADO QUE NO ESTA AFECTO", vbInformation, "CTS"
'        IF COMBO1.VISIBLE = TRUE THEN
            Combo1.SetFocus
'       END IF
        Exit Function
    End If
    If Combo2.ListIndex = 0 And Check7.Value = 1 Then
        MsgBox "NO SE PUEDE PROVISIONAR EL VALOR DE ESTE INGRESO, SI HA SETEADO QUE NO ESTA AFECTO", vbInformation, "VACACIONES"
        If Combo2.Visible = True Then
            Combo2.SetFocus
        End If
        Exit Function
    End If
    If Combo3.ListIndex = 0 And Check8.Value = 1 Then
        MsgBox "NO SE PUEDE PROVISIONAR EL VALOR DE ESTE INGRESO, SI HA SETEADO QUE NO ESTA AFECTO", vbInformation, "GRATIFICACIóN"
        If Combo3.Visible = True Then
            Combo3.SetFocus
        End If
        Exit Function
    End If
    If xcolplanilla.Tag <> "" Then
        Dim X As Integer
        DBSYSTEM.Execute "UPDATE COLUMPL SET TIPO=TIPO WHERE CODIGO='" & UCase(xcolplanilla.Tag) & "'", X
        If X = 0 Then
            MsgBox "LA COLUMNA DE PLANILLA NO EXISTE", vbInformation
            xcolplanilla.SetFocus
            Exit Function
        End If
    End If
    VERIFICADATOS = True
End Function
Private Sub XTIPO_CLICK()
    If xTipo.ListIndex = 1 Then
        l1(0).Visible = True
        xEnlace.Visible = True
    Else
        l1(0).Visible = False
        xEnlace.Visible = False
    End If
End Sub


