VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frTrab 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trabajador"
   ClientHeight    =   6135
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   10440
   Icon            =   "frTrab.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   10440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Importar Trabajador"
      Height          =   345
      Left            =   1275
      TabIndex        =   106
      Top             =   5685
      Width           =   1755
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   90
      Top             =   5580
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frTrab.frx":0E42
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   450
      Left            =   7860
      TabIndex        =   100
      Top             =   5595
      Width           =   2445
   End
   Begin VB.PictureBox Ole1 
      DataField       =   "Foto"
      DataSource      =   "Data1"
      Height          =   1320
      Left            =   7230
      ScaleHeight     =   1260
      ScaleWidth      =   990
      TabIndex        =   97
      Top             =   270
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir Ficha"
      Height          =   345
      Left            =   6000
      TabIndex        =   45
      Top             =   5685
      Width           =   1335
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   75
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   ""
      Top             =   5730
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.CommandButton cmDescargaFoto 
      Caption         =   "Descargar Foto"
      Height          =   315
      Left            =   8505
      TabIndex        =   48
      Top             =   1410
      Width           =   1245
   End
   Begin VB.CommandButton cmCargaFoto 
      Caption         =   "Cargar &Foto"
      Height          =   315
      Left            =   8505
      TabIndex        =   46
      Top             =   1005
      Width           =   1245
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   4560
      TabIndex        =   44
      Top             =   5685
      Width           =   1335
   End
   Begin VB.CommandButton cmAcepta 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   3105
      TabIndex        =   43
      Top             =   5685
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Principal"
      Height          =   1365
      Left            =   135
      TabIndex        =   49
      Top             =   195
      Width           =   6765
      Begin AplisetControlText.Aplitext xCodAlt 
         Height          =   285
         Left            =   5430
         TabIndex        =   4
         Top             =   345
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         MaxLength       =   6
         Text            =   ""
         Entero          =   -1  'True
         SinBlancos      =   -1  'True
      End
      Begin AplisetControlText.Aplitext xCodTrab 
         Height          =   285
         Left            =   930
         TabIndex        =   0
         Top             =   345
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         MaxLength       =   8
         Text            =   ""
         SinBlancos      =   -1  'True
      End
      Begin AplisetControlText.Aplitext xApePat 
         Height          =   285
         Left            =   180
         TabIndex        =   1
         Top             =   930
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         MaxLength       =   20
         Text            =   ""
         SinBlancos      =   -1  'True
      End
      Begin AplisetControlText.Aplitext xApeMat 
         Height          =   285
         Left            =   2160
         TabIndex        =   2
         Top             =   930
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         MaxLength       =   20
         Text            =   ""
         SinBlancos      =   -1  'True
      End
      Begin AplisetControlText.Aplitext xNombre 
         Height          =   285
         Left            =   4140
         TabIndex        =   3
         Top             =   930
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   503
         MaxLength       =   20
         Text            =   ""
         SinBlancos      =   -1  'True
      End
      Begin VB.Label l0 
         AutoSize        =   -1  'True
         Caption         =   "Código Alterno"
         Height          =   195
         Index           =   4
         Left            =   4170
         TabIndex        =   54
         Top             =   390
         Width           =   1035
      End
      Begin VB.Label l0 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Index           =   3
         Left            =   4170
         TabIndex        =   53
         Top             =   690
         Width           =   555
      End
      Begin VB.Label l0 
         AutoSize        =   -1  'True
         Caption         =   "Apellido Materno"
         Height          =   195
         Index           =   2
         Left            =   2190
         TabIndex        =   52
         Top             =   690
         Width           =   1185
      End
      Begin VB.Label l0 
         AutoSize        =   -1  'True
         Caption         =   "Apellido Paterno"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   51
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label l0 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   50
         Top             =   390
         Width           =   495
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3825
      Left            =   120
      TabIndex        =   47
      Top             =   1755
      Width           =   10185
      _ExtentX        =   17965
      _ExtentY        =   6747
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frTrab.frx":1194
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "l1(0)"
      Tab(0).Control(1)=   "l1(1)"
      Tab(0).Control(2)=   "l1(2)"
      Tab(0).Control(3)=   "l1(3)"
      Tab(0).Control(4)=   "l1(4)"
      Tab(0).Control(5)=   "l1(5)"
      Tab(0).Control(6)=   "l1(6)"
      Tab(0).Control(7)=   "Line1(0)"
      Tab(0).Control(8)=   "Line1(1)"
      Tab(0).Control(9)=   "l1(7)"
      Tab(0).Control(10)=   "l1(8)"
      Tab(0).Control(11)=   "l1(9)"
      Tab(0).Control(12)=   "l1(10)"
      Tab(0).Control(13)=   "l1(11)"
      Tab(0).Control(14)=   "l1(12)"
      Tab(0).Control(15)=   "l1(13)"
      Tab(0).Control(16)=   "l1(23)"
      Tab(0).Control(17)=   "xTipDoc"
      Tab(0).Control(18)=   "xDocIden"
      Tab(0).Control(19)=   "xFechaNac"
      Tab(0).Control(20)=   "xDireccion"
      Tab(0).Control(21)=   "xUbigeo"
      Tab(0).Control(22)=   "xTelefono"
      Tab(0).Control(23)=   "xCarnetSeg"
      Tab(0).Control(24)=   "cmdGenCS"
      Tab(0).Control(25)=   "xFondoPens"
      Tab(0).Control(26)=   "xCuspp"
      Tab(0).Control(27)=   "xMesDevengue"
      Tab(0).Control(28)=   "xCtaBanco"
      Tab(0).Control(29)=   "xBanco"
      Tab(0).Control(30)=   "xEstadoCivil"
      Tab(0).Control(31)=   "xSexo"
      Tab(0).Control(32)=   "xFechaIAFP"
      Tab(0).Control(33)=   "xNoCalculo"
      Tab(0).ControlCount=   34
      TabCaption(1)   =   "Laboral"
      TabPicture(1)   =   "frTrab.frx":11B0
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "l1(15)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "l1(16)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "l1(17)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "l1(18)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "l1(19)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "l1(20)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "l1(21)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Line2(0)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "l1(24)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "l1(25)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "l1(26)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "l1(28)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "l1(29)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "l1(30)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "l1(22)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label1"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "l1(14)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Line2(1)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Label5"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Label6"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Label7"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "xOpQuinta"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "xNumFicha"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "xFechaIng"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "xDepartamento"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "xCargo"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "xBasico"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "xFechaCese"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "xCtaCTS"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "xBancoCTS"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "xRucEPS"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "xCodCTR"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "xAsigFam"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "xEsSaludVida"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "xSituacion"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "xContrato"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "xFechaTermino"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "xTipoTrab"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "xCCosto"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "xNoPDT"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "xOpcion01"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "xOpcion02"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "xOpcionA"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "xOpcionB"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "Check1"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).ControlCount=   45
      TabCaption(2)   =   "Cta. Cte."
      TabPicture(2)   =   "frTrab.frx":11CC
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "LbEgreD"
      Tab(2).Control(1)=   "lbingrD"
      Tab(2).Control(2)=   "LbEgre"
      Tab(2).Control(3)=   "Label4"
      Tab(2).Control(4)=   "LbIngr"
      Tab(2).Control(5)=   "Label3"
      Tab(2).Control(6)=   "Frame3"
      Tab(2).Control(7)=   "Frame2"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "Otra Información"
      TabPicture(3)   =   "frTrab.frx":11E8
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Command5"
      Tab(3).Control(1)=   "xLista"
      Tab(3).Control(2)=   "Command4"
      Tab(3).Control(3)=   "QuitarDAto"
      Tab(3).ControlCount=   4
      Begin VB.CheckBox Check1 
         Caption         =   "Afecto a Quinta"
         Height          =   225
         Left            =   8280
         TabIndex        =   107
         Top             =   3330
         Width           =   1830
      End
      Begin VB.CheckBox xNoCalculo 
         Alignment       =   1  'Right Justify
         Caption         =   "No Considerar en el Calculo"
         Height          =   240
         Left            =   -67650
         TabIndex        =   105
         Top             =   3030
         Width           =   2670
      End
      Begin VB.CommandButton QuitarDAto 
         Caption         =   "Quitar dato informativo"
         Height          =   555
         Left            =   -74790
         TabIndex        =   104
         Top             =   1275
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Editar Información"
         Height          =   555
         Left            =   -74790
         TabIndex        =   102
         Top             =   495
         Width           =   1095
      End
      Begin MSComctlLib.ListView xLista 
         Height          =   2820
         Left            =   -73515
         TabIndex        =   101
         Top             =   480
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   4974
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   4941
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Valor"
            Object.Width           =   2540
         EndProperty
      End
      Begin AplisetControlText.Aplitext xOpcionB 
         Height          =   285
         Left            =   6855
         TabIndex        =   41
         Top             =   2925
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         MaxLength       =   10
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xOpcionA 
         Height          =   285
         Left            =   6855
         TabIndex        =   39
         Top             =   2610
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         MaxLength       =   10
         Text            =   ""
      End
      Begin VB.CheckBox xOpcion02 
         Caption         =   "Opcion02"
         Height          =   195
         Left            =   8280
         TabIndex        =   42
         Top             =   3060
         Width           =   1035
      End
      Begin VB.CheckBox xOpcion01 
         Caption         =   "Opcion01"
         Height          =   195
         Left            =   8280
         TabIndex        =   40
         Top             =   2760
         Width           =   1035
      End
      Begin VB.CheckBox xNoPDT 
         Caption         =   "No Declarar al PDT"
         Height          =   210
         Left            =   8280
         TabIndex        =   38
         Top             =   2445
         Width           =   1770
      End
      Begin VB.Frame Frame2 
         Caption         =   "Ingresos"
         Height          =   2385
         Left            =   -74910
         TabIndex        =   89
         Top             =   435
         Width           =   4965
         Begin MSDataGridLib.DataGrid DtIngr 
            Height          =   1725
            Left            =   90
            TabIndex        =   90
            Top             =   390
            Width           =   4680
            _ExtentX        =   8255
            _ExtentY        =   3043
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
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
            ColumnCount     =   4
            BeginProperty Column00 
               DataField       =   "Descripcion"
               Caption         =   "Descripción"
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
               DataField       =   "Mon"
               Caption         =   "Moneda"
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
               DataField       =   "Soles"
               Caption         =   "Capital(S/.)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "Dolares"
               Caption         =   "Capital(US$)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#,##0.00"
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
                  ColumnAllowSizing=   0   'False
               EndProperty
               BeginProperty Column01 
                  Alignment       =   2
                  ColumnAllowSizing=   0   'False
               EndProperty
               BeginProperty Column02 
                  Alignment       =   1
                  ColumnAllowSizing=   -1  'True
                  ColumnWidth     =   1080
               EndProperty
               BeginProperty Column03 
                  Alignment       =   1
                  ColumnAllowSizing=   -1  'True
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Egresos"
         Height          =   2385
         Left            =   -69870
         TabIndex        =   87
         Top             =   450
         Width           =   4965
         Begin MSDataGridLib.DataGrid DtEgre 
            Height          =   1725
            Left            =   120
            TabIndex        =   88
            Top             =   375
            Width           =   4680
            _ExtentX        =   8255
            _ExtentY        =   3043
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
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
            ColumnCount     =   4
            BeginProperty Column00 
               DataField       =   "Descripcion"
               Caption         =   "Descripción"
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
               DataField       =   "Mon"
               Caption         =   "Moneda"
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
               DataField       =   "Soles"
               Caption         =   "Capital(S/.)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "Dolares"
               Caption         =   "Capital(US$)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#,##0.00"
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
                  ColumnAllowSizing=   0   'False
               EndProperty
               BeginProperty Column01 
                  Alignment       =   2
                  ColumnAllowSizing=   0   'False
               EndProperty
               BeginProperty Column02 
                  Alignment       =   1
                  ColumnAllowSizing=   -1  'True
                  ColumnWidth     =   1080
               EndProperty
               BeginProperty Column03 
                  Alignment       =   1
                  ColumnAllowSizing=   -1  'True
               EndProperty
            EndProperty
         End
      End
      Begin AplisetControlText.Aplitext xCCosto 
         Height          =   285
         Left            =   1710
         TabIndex        =   24
         Top             =   1410
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   503
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xTipoTrab 
         Height          =   285
         Left            =   1710
         TabIndex        =   22
         Top             =   795
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   503
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker xFechaIAFP 
         Height          =   285
         Left            =   -68175
         TabIndex        =   17
         Top             =   1710
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   503
         _Version        =   393216
         Format          =   61800449
         CurrentDate     =   36644
      End
      Begin MSComCtl2.DTPicker xFechaTermino 
         Height          =   315
         Left            =   1710
         TabIndex        =   29
         Top             =   2940
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         _Version        =   393216
         Format          =   61800449
         CurrentDate     =   36644
      End
      Begin VB.ComboBox xContrato 
         Height          =   315
         ItemData        =   "frTrab.frx":1204
         Left            =   1710
         List            =   "frTrab.frx":120E
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   2595
         Width           =   3135
      End
      Begin VB.ComboBox xSituacion 
         Height          =   315
         ItemData        =   "frTrab.frx":1233
         Left            =   1710
         List            =   "frTrab.frx":1235
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1095
         Width           =   3135
      End
      Begin VB.ComboBox xSexo 
         Height          =   315
         Left            =   -73680
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2790
         Width           =   3195
      End
      Begin VB.CheckBox xEsSaludVida 
         Caption         =   "EsSalud Vida"
         Height          =   195
         Left            =   8280
         TabIndex        =   36
         Top             =   2130
         Width           =   1335
      End
      Begin AplisetControlText.Aplitext xAsigFam 
         Height          =   285
         Left            =   6855
         TabIndex        =   35
         Top             =   1980
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         Text            =   ""
         Redondear       =   -1  'True
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xCodCTR 
         Height          =   285
         Left            =   6855
         TabIndex        =   34
         Top             =   1680
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   503
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xRucEPS 
         Height          =   285
         Left            =   6855
         TabIndex        =   33
         Top             =   1380
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   503
         MaxLength       =   11
         Text            =   ""
         SinBlancos      =   -1  'True
      End
      Begin AplisetControlText.Aplitext xBancoCTS 
         Height          =   285
         Left            =   6855
         TabIndex        =   32
         Top             =   1080
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   503
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xCtaCTS 
         Height          =   285
         Left            =   6855
         TabIndex        =   31
         Top             =   780
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   503
         MaxLength       =   30
         Text            =   ""
         SinBlancos      =   -1  'True
      End
      Begin MSComCtl2.DTPicker xFechaCese 
         Height          =   285
         Left            =   6855
         TabIndex        =   30
         Top             =   480
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   61800449
         CurrentDate     =   36495
      End
      Begin AplisetControlText.Aplitext xBasico 
         Height          =   285
         Left            =   1710
         TabIndex        =   27
         Top             =   2295
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   503
         MaxLength       =   20
         Text            =   ""
         Redondear       =   -1  'True
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xCargo 
         Height          =   285
         Left            =   1710
         TabIndex        =   26
         Top             =   1995
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   503
         MaxLength       =   25
         Text            =   ""
         SinBlancos      =   -1  'True
      End
      Begin AplisetControlText.Aplitext xDepartamento 
         Height          =   285
         Left            =   1710
         TabIndex        =   25
         Top             =   1695
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   503
         MaxLength       =   25
         Locked          =   -1  'True
         Text            =   ""
         SinBlancos      =   -1  'True
      End
      Begin MSComCtl2.DTPicker xFechaIng 
         Height          =   285
         Left            =   1710
         TabIndex        =   21
         Top             =   495
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   503
         _Version        =   393216
         Format          =   61800449
         CurrentDate     =   36494
      End
      Begin VB.ComboBox xEstadoCivil 
         Height          =   315
         Left            =   -68175
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   2610
         Width           =   3195
      End
      Begin AplisetControlText.Aplitext xBanco 
         Height          =   285
         Left            =   -68175
         TabIndex        =   19
         Top             =   2310
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   503
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xCtaBanco 
         Height          =   285
         Left            =   -68175
         TabIndex        =   18
         Top             =   2010
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   503
         MaxLength       =   30
         Text            =   ""
         SinBlancos      =   -1  'True
      End
      Begin MSComCtl2.DTPicker xMesDevengue 
         Height          =   285
         Left            =   -68175
         TabIndex        =   16
         Top             =   1410
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   61800449
         CurrentDate     =   36495
      End
      Begin AplisetControlText.Aplitext xCuspp 
         Height          =   285
         Left            =   -68175
         TabIndex        =   15
         Top             =   1110
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   503
         MaxLength       =   12
         Text            =   ""
         SinBlancos      =   -1  'True
      End
      Begin AplisetControlText.Aplitext xFondoPens 
         Height          =   285
         Left            =   -68175
         TabIndex        =   14
         Top             =   810
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   503
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin VB.CommandButton cmdGenCS 
         Caption         =   "..."
         Height          =   285
         Left            =   -68505
         TabIndex        =   12
         Top             =   510
         Width           =   315
      End
      Begin AplisetControlText.Aplitext xCarnetSeg 
         Height          =   285
         Left            =   -68175
         TabIndex        =   13
         Top             =   510
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   503
         MaxLength       =   15
         Text            =   ""
         SinBlancos      =   -1  'True
      End
      Begin AplisetControlText.Aplitext xTelefono 
         Height          =   285
         Left            =   -73680
         TabIndex        =   10
         Top             =   2490
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   503
         MaxLength       =   10
         Text            =   ""
         SinBlancos      =   -1  'True
      End
      Begin AplisetControlText.Aplitext xUbigeo 
         Height          =   495
         Left            =   -73680
         TabIndex        =   9
         Top             =   1980
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   873
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xDireccion 
         Height          =   555
         Left            =   -73680
         TabIndex        =   8
         Top             =   1410
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   979
         MaxLength       =   50
         Text            =   ""
         SinBlancos      =   -1  'True
      End
      Begin MSComCtl2.DTPicker xFechaNac 
         Height          =   285
         Left            =   -73680
         TabIndex        =   7
         Top             =   1110
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   503
         _Version        =   393216
         Format          =   61800449
         CurrentDate     =   36494
      End
      Begin AplisetControlText.Aplitext xDocIden 
         Height          =   285
         Left            =   -73680
         TabIndex        =   6
         Top             =   810
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   503
         MaxLength       =   12
         Text            =   ""
         SinBlancos      =   -1  'True
      End
      Begin AplisetControlText.Aplitext xTipDoc 
         Height          =   285
         Left            =   -73680
         TabIndex        =   5
         Top             =   510
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   503
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xNumFicha 
         Height          =   285
         Left            =   6855
         TabIndex        =   37
         Top             =   2295
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         MaxLength       =   10
         Text            =   ""
         SinBlancos      =   -1  'True
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Datos Informativos"
         Height          =   390
         Left            =   -67410
         TabIndex        =   103
         Top             =   480
         Width           =   1980
      End
      Begin AplisetControlText.Aplitext xOpQuinta 
         Height          =   285
         Left            =   6870
         TabIndex        =   108
         Top             =   3285
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         MaxLength       =   10
         Text            =   "0.0"
         TipoDato        =   "N"
      End
      Begin VB.Label Label7 
         Caption         =   "Opcion Quinta"
         Height          =   180
         Left            =   5565
         TabIndex        =   109
         Top             =   3315
         Width           =   1125
      End
      Begin VB.Label Label6 
         Caption         =   "OpcionB"
         Height          =   180
         Left            =   5550
         TabIndex        =   99
         Top             =   3000
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "OpcionA"
         Height          =   195
         Left            =   5550
         TabIndex        =   98
         Top             =   2670
         Width           =   735
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         Index           =   1
         X1              =   5235
         X2              =   5235
         Y1              =   300
         Y2              =   3780
      End
      Begin VB.Label Label3 
         Caption         =   "Total Ingresos"
         Height          =   285
         Left            =   -73815
         TabIndex        =   96
         Top             =   2970
         Width           =   1245
      End
      Begin VB.Label LbIngr 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   -72450
         TabIndex        =   95
         Top             =   2910
         Width           =   1080
      End
      Begin VB.Label Label4 
         Caption         =   "Total Egresos "
         Height          =   285
         Left            =   -68580
         TabIndex        =   94
         Top             =   2985
         Width           =   1095
      End
      Begin VB.Label LbEgre 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   -67350
         TabIndex        =   93
         Top             =   2925
         Width           =   1095
      End
      Begin VB.Label lbingrD 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   -71325
         TabIndex        =   92
         Top             =   2910
         Width           =   1080
      End
      Begin VB.Label LbEgreD 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   -66210
         TabIndex        =   91
         Top             =   2925
         Width           =   1095
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "F.Inscripción Pension"
         Height          =   195
         Index           =   23
         Left            =   -69855
         TabIndex        =   85
         Top             =   1755
         Width           =   1515
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Historial Físico"
         Height          =   195
         Index           =   14
         Left            =   5535
         TabIndex        =   84
         Top             =   2370
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Término de Contrato"
         Height          =   195
         Left            =   180
         TabIndex        =   83
         Top             =   3000
         Width           =   1440
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Contrato"
         Height          =   195
         Index           =   22
         Left            =   180
         TabIndex        =   82
         Top             =   2655
         Width           =   1185
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Asig. Familiar"
         Height          =   195
         Index           =   30
         Left            =   5550
         TabIndex        =   81
         Top             =   2040
         Width           =   1035
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Centro Riesgo"
         Height          =   195
         Index           =   29
         Left            =   5535
         TabIndex        =   80
         Top             =   1725
         Width           =   1125
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "RUC E.P.S."
         Height          =   195
         Index           =   28
         Left            =   5550
         TabIndex        =   79
         Top             =   1425
         Width           =   960
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Banco C.T.S."
         Height          =   195
         Index           =   26
         Left            =   5550
         TabIndex        =   78
         Top             =   1125
         Width           =   1080
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta C.T.S."
         Height          =   195
         Index           =   25
         Left            =   5550
         TabIndex        =   77
         Top             =   840
         Width           =   1125
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Cese"
         Height          =   195
         Index           =   24
         Left            =   5550
         TabIndex        =   76
         Top             =   540
         Width           =   1200
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000003&
         Index           =   0
         X1              =   5220
         X2              =   5220
         Y1              =   300
         Y2              =   3810
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Rem. Básica"
         Height          =   195
         Index           =   21
         Left            =   180
         TabIndex        =   75
         Top             =   2325
         Width           =   900
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Ocupación"
         Height          =   195
         Index           =   20
         Left            =   180
         TabIndex        =   74
         Top             =   2025
         Width           =   780
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Area de Trabajo"
         Height          =   195
         Index           =   19
         Left            =   180
         TabIndex        =   73
         Top             =   1725
         Width           =   1140
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Centro de Costo"
         Height          =   195
         Index           =   18
         Left            =   180
         TabIndex        =   72
         Top             =   1434
         Width           =   1140
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Situación"
         Height          =   195
         Index           =   17
         Left            =   180
         TabIndex        =   71
         Top             =   1140
         Width           =   660
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Trabajador"
         Height          =   195
         Index           =   16
         Left            =   180
         TabIndex        =   70
         Top             =   838
         Width           =   1350
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Ingreso"
         Height          =   195
         Index           =   15
         Left            =   180
         TabIndex        =   69
         Top             =   540
         Width           =   1245
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Estado Civil"
         Height          =   195
         Index           =   13
         Left            =   -69870
         TabIndex        =   68
         Top             =   2670
         Width           =   825
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Banco"
         Height          =   195
         Index           =   12
         Left            =   -69870
         TabIndex        =   67
         Top             =   2370
         Width           =   465
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta Bancaria"
         Height          =   195
         Index           =   11
         Left            =   -69870
         TabIndex        =   66
         Top             =   2070
         Width           =   1185
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Devengue AFP"
         Height          =   195
         Index           =   10
         Left            =   -69870
         TabIndex        =   65
         Top             =   1455
         Width           =   1095
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "C.U.S.P.P."
         Height          =   195
         Index           =   9
         Left            =   -69870
         TabIndex        =   64
         ToolTipText     =   "Código Unico del Sistema Privado de Pensiones"
         Top             =   1155
         Width           =   765
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Pensiones"
         Height          =   195
         Index           =   8
         Left            =   -69870
         TabIndex        =   63
         Top             =   855
         Width           =   735
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Carnet Seguro"
         Height          =   195
         Index           =   7
         Left            =   -69870
         TabIndex        =   62
         Top             =   555
         Width           =   1020
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         Index           =   1
         X1              =   -70110
         X2              =   -70110
         Y1              =   300
         Y2              =   3750
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         Index           =   0
         X1              =   -70125
         X2              =   -70125
         Y1              =   285
         Y2              =   3735
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Sexo"
         Height          =   195
         Index           =   6
         Left            =   -74805
         TabIndex        =   61
         Top             =   2880
         Width           =   360
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Teléfono"
         Height          =   195
         Index           =   5
         Left            =   -74805
         TabIndex        =   60
         Top             =   2520
         Width           =   630
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Ubigeo (INEI)"
         Height          =   195
         Index           =   4
         Left            =   -74805
         TabIndex        =   59
         Top             =   2010
         Width           =   960
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Dirección"
         Height          =   195
         Index           =   3
         Left            =   -74805
         TabIndex        =   58
         Top             =   1440
         Width           =   675
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Nac."
         Height          =   195
         Index           =   2
         Left            =   -74805
         TabIndex        =   57
         ToolTipText     =   "Fecha de Nacimiento"
         Top             =   1170
         Width           =   840
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Documento"
         Height          =   195
         Index           =   1
         Left            =   -74805
         TabIndex        =   56
         Top             =   870
         Width           =   825
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Doc."
         Height          =   195
         Index           =   0
         Left            =   -74805
         TabIndex        =   55
         ToolTipText     =   "Tipo de Documento"
         Top             =   570
         Width           =   930
      End
   End
   Begin Crystal.CrystalReport Reporte 
      Left            =   9945
      Top             =   1245
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Image xFoto 
      BorderStyle     =   1  'Fixed Single
      Height          =   1605
      Left            =   7080
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Trabajador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   8610
      TabIndex        =   86
      Top             =   465
      Width           =   1185
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   9735
      Picture         =   "frTrab.frx":1237
      Top             =   180
      Width           =   240
   End
End
Attribute VB_Name = "frTrab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSAFP As New ADODB.Recordset
Dim RSTIPDOC As New ADODB.Recordset
Dim RSBANCO As New ADODB.Recordset
Dim RSTIPTRAB As New ADODB.Recordset
Dim RSSCTR As New ADODB.Recordset
Dim RSCCOSTO As New ADODB.Recordset
Dim RSTRABS As New ADODB.Recordset
Dim RSUBIGEO As New ADODB.Recordset
Private Sub cmAcepta_Click()
    If Not OKPARAEDITAR Then Exit Sub
    If VPTAREA = "NUEVO" Then
        If Not OKPARANUEVOS Then Exit Sub
            DBSYSTEM.Execute "INSERT INTO TRABAJADORES (CODTRAB) VALUES ('" & xCodTrab.Text & "')"
            GRABAR xCodTrab.Text
    Else
        RSTRABS.MoveFirst
        RSTRABS.FIND "CODTRAB='" & VPTAREA & "'"
        If RSTRABS.EOF Then
            MsgBox "EL REGISTRO EN EDICIÓN YA NO SE ENCUENTRA ACTUALEMENTE EN LA BASE DE DATOS, POSIBLEMENTE SE HA ELIMINADO POR OTRO USUARIO", vbInformation
            Unload Me
            Exit Sub
        End If
        GRABAR xCodTrab.Text
    End If
    frPersonal.cmdEventos.Enabled = True
    frPersonal.Command1.Enabled = True
    Unload Me
End Sub

Private Sub CMCARGAFOTO_Click()
    frOpenGr.Show 1
    If VGUTIL(0) <> "" Then
        xFoto.Picture = LoadPicture(VGUTIL(0))
        xFoto.Tag = VGUTIL(0)
        VGUTIL(0) = ""
    Else
        MsgBox "ACCIÓN CANCELADA", vbInformation
    End If
End Sub

Private Sub CMDESCARGAFOTO_Click()
    Set xFoto.Picture = Nothing
    xFoto.Tag = ""
End Sub

Private Sub CMDGENCS_Click()
    If xSexo.ListIndex = -1 Then
        MsgBox "FALTA ESPECIFICAR EL SEXO DEL TRABAJADOR. SELECCIONE ENTRE MASCULINO Y FEMENINO", vbInformation
        xSexo.SetFocus
        Exit Sub
    End If
    xCarnetSeg.Text = DarCarnetSeg(xFechaNac.Value, xApePat.Text, xApeMat.Text, xNombre.Text, xSexo.ListIndex)
    xCarnetSeg.SetFocus
End Sub

Private Sub CmdImprimir_Click()
    Dim XCUEN As Long
    'ON ERROR RESUME NEXT
    Dim SQLSTR As String
    Set RS_TABLA = New ADODB.Recordset
    Set RS_AUX = New ADODB.Recordset
    If Not ExisteTabla("DATATRAB") Then
        MsgBox "ERROR NO SE ENCONTRO EL ARCHIVO O LA TABLA DATA TRABAJADOR", vbCritical, "INFORMACION"
        Exit Sub
    End If
        If Not ExisteTablaAux(" [##_TMPCNP" & VGL_COMPUTER & "] ") Then
        DBSTARPLAN.Execute "CREATE TABLE  [##_TMPCNP" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8), CODCNP VARCHAR(30), CONCEPTO VARCHAR(100), VALOR VARCHAR(50))"
    End If
        DBSTARPLAN.Execute "DELETE FROM  [##_TMPCNP" & VGL_COMPUTER & "] "
        RS_AUX.Open "SELECT * FROM DATATRAB", DBSYSTEM
        If RS_AUX.RecordCount Then
            While Not RS_AUX.EOF
                SQLSTR = "SELECT CODTRAB, " & RS_AUX.Fields(0) & " FROM TRABAJADORES WHERE CODTRAB='" & xCodTrab.Text & "'"
                RS_TABLA.Open SQLSTR, DBSYSTEM
                If Not IsNull(RS_TABLA.Fields(1)) Then
                    If RS_AUX!TIPODATA = "B" Then
                        SQL = "INSERT INTO  [##_TMPCNP" & VGL_COMPUTER & "]  VALUES('" & RS_TABLA.Fields(0) & "','" & RS_AUX.Fields(0) & "','" & RS_AUX.Fields(1) & "','" & IIf(RS_TABLA.Fields(1), "SI", "NO") & "')"
                    Else
                        SQL = "INSERT INTO  [##_TMPCNP" & VGL_COMPUTER & "]  VALUES('" & RS_TABLA.Fields(0) & "','" & RS_AUX.Fields(0) & "','" & RS_AUX.Fields(1) & "','" & RS_TABLA.Fields(1) & "')"
                    End If
                    DBSTARPLAN.Execute SQL, T
                    T = 0
                End If
                RS_TABLA.Close
                RS_AUX.MoveNext
            Wend
        End If
        RS_AUX.Close

    Screen.MousePointer = 11
    CambiaPanelBD True
    
    If Data1.Recordset.RecordCount > 0 Then
        Data1.Recordset.MoveFirst
        Do While Not Data1.Recordset.EOF
            Data1.Recordset.Delete
            Data1.Recordset.MoveNext
        Loop
    End If
    Data1.Refresh
    Data1.Recordset.AddNew
    If xFoto.Tag <> "" Then
        Set Ole1.Picture = LoadPicture(xFoto.Tag)
    Else
        Set Ole1.Picture = LoadPicture(App.PATH & "\OBJBLANK.BMP")
    End If
    Data1.Recordset.Fields("CODIGO") = xCodTrab.Text
    Data1.Recordset.Update
    With Reporte
        .Reset
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0037.RPT"
        .DataFiles(0) = App.PATH & "\BDAUXCOM.MDB"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowTitle = "PLAN0037 - FICHA DEL TRABAJADOR"
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "XCODIGO='" & xCodTrab.Text & "'"
        .Formulas(2) = "XAPEPAT='" & xApePat.Text & "'"
        .Formulas(3) = "XAPEMAT='" & xApeMat.Text & "'"
        .Formulas(4) = "XNOMB='" & xNombre.Text & "'"
        .Formulas(5) = "XTIPDOC='" & xTipDoc.Text & "'"
        .Formulas(6) = "XNDOC='" & xDocIden.Text & "'"
        .Formulas(7) = "XFECHNAC='" & Format(xFechaNac, "DD/MM/YYYY") & "'"
        .Formulas(8) = "XDIR='" & xDireccion.Text & "'"
        .Formulas(9) = "XUBIGEO='" & Right(xUbigeo.Text, Len(xUbigeo.Text) - Len(Getcad(":", 1, xUbigeo.Text)) - 2) & "'"
        .Formulas(10) = "XTELF='" & xTelefono.Text & "'"
        .Formulas(11) = "XSEX='" & xSexo.Text & "'"
        .Formulas(12) = "XCSEG='" & xCarnetSeg.Text & "'"
        .Formulas(13) = "XPENS='" & xFondoPens.Text & "'"
        .Formulas(14) = "XCVSPP='" & xCuspp.Text & "'"
        .Formulas(15) = "XDEVAFP='" & IIf(xFondoPens.Tag <> "ON", xMesDevengue.Value, "--/--/----") & "'"
        .Formulas(16) = "XFINSC='" & Format(xFechaIAFP, "DD/MM/YYYY") & "'"
        .Formulas(17) = "XCTBANC='" & xCtaBanco.Text & "'"
        If Len(xBanco.Text) > 0 Then
            .Formulas(18) = "XBANCO='" & Right(xBanco.Text, Len(xBanco.Text) - Len(Getcad(":", 1, xBanco.Text)) - 3) & "'"
        End If
        .Formulas(19) = "XESTCIV='" & xEstadoCivil.Text & "'"
        .Formulas(20) = "XFECHING='" & Format(xFechaIng, "DD/MM/YYYY") & "'"
        .Formulas(21) = "XTIPTRAB='" & xTipoTrab.Text & "'"
        .Formulas(22) = "XSITU='" & xSituacion.Text & "'"
        .Formulas(23) = "XCTCOST='" & Right(xCCosto.Text, Len(xCCosto.Text) - Len(Getcad(":", 1, xCCosto.Text)) - 3) & "'"
        .Formulas(24) = "XARTRAB='" & xDepartamento.Text & "'"
        .Formulas(25) = "XCARGO='" & xCargo.Text & "'"
        .Formulas(26) = "XREMBANC='" & xBasico.Text & "'"
        .Formulas(27) = "XTIPCONT='" & xContrato.Text & "'"
        .Formulas(28) = "XTERMCONT='" & IIf(xContrato.ListIndex = 0, "--/--/----", Format(xFechaTermino, "DD/MM/YYYY")) & "'"
        .Formulas(29) = "XFECHCE='" & IIf(xFechaCese.CheckBox, Format(xFechaCese, "MM/DD/YYYY"), "") & "'"
        .Formulas(30) = "XCTCTS='" & xCtaCTS.Text & "'"
        .Formulas(31) = "XBANCCTS='" & Right(xBancoCTS.Text, Len(xBancoCTS.Text) - Len(Getcad(":", 1, xBancoCTS.Text)) - 3) & "'"
        .Formulas(32) = "XRUCEPS='" & xRucEPS.Text & "'"
        .Formulas(33) = "XCENTRIES='" & xCodCTR.Text & "'"
        .Formulas(34) = "XASIGFAM='" & xAsigFam.Text & "'"
        If xEsSaludVida.Value Then
            .Formulas(35) = "XESSALUD='SI'"
        Else: .Formulas(35) = "XESSALUD='NO'"
        End If
        .Formulas(36) = "XHISTFIS='" & xNumFicha.Text & "'"
        .Formulas(37) = "XHORA='" & Format(Time, "HH:MM") & "'"
        .Formulas(38) = "XRUC='" & REGSISTEMA.RUC & "'"
        .SubreportToChange = "PlRH0014.rpt"
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
        If .Status <> 2 Then .Action = 1
    End With
    Screen.MousePointer = 1
    CambiaPanelBD False
End Sub

Private Sub Command1_Click()
    VGLFRM = 2
   frEmpTr.Show 1
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command4_Click()
    frIngDatos.Show 1
    REFRESCAROTRADATA
End Sub

Private Sub Command5_Click()
    frDataTrab.Show 1
End Sub

Private Sub Form_Initialize()
Data1.DatabaseName = App.PATH & "\BDAUXCOM.MDB"
Data1.RecordSource = "FTMPFOTO"
Data1.Refresh
End Sub

Private Sub Form_Load()
SSTab1.Tab = 0
Data1.DatabaseName = App.PATH & "\BDAUXCOM.MDB"
Data1.RecordSource = "FTMPFOTO"
Data1.Refresh
If VAR_SHOW = 1 Then
    Unload frEstCurrTrab
ElseIf VAR_SHOW = 2 Then
    Unload FrmEventos
End If
    Screen.MousePointer = 1
    GetPosition Me
    Set RSAFP = Nothing
    RSAFP.Open "AFPS", DBSYSTEM, adOpenKeyset, adLockOptimistic
    If RSAFP.RecordCount <> 0 Then
        Print "LA VERSIÓN DE LA BASE DE DATOS SE HA ACTUALIZADO"
    End If
    Set RSAFP = Nothing
    RSAFP.Open "EMPRESA", DBSYSTEM, adOpenStatic
    If RSAFP.RecordCount Then
        If Not IsEmpty(RSAFP!CFG0004) Then
            If RSAFP!CFG0004 Then
                xCodTrab.Locked = True
            Else
                xCodTrab.Locked = False
            End If
        End If
    End If
    RSAFP.Close
    RSAFP.Open "AFPS", DBSYSTEM, adOpenKeyset, adLockOptimistic
    RSTIPDOC.Open "DOCUMENTOS", DBSYSTEM, adOpenKeyset, adLockOptimistic
    RSBANCO.Open "BANCOS", DBSYSTEM, adOpenKeyset, adLockOptimistic
    RSTIPTRAB.Open "TIPOSTRAB", DBSYSTEM, adOpenKeyset, adLockOptimistic
    RSSCTR.Open "SELECT CODCAR, NOMBRE, TASA FROM CENTROSAR ORDER BY NOMBRE", DBSYSTEM, adOpenKeyset, adLockOptimistic
    RSCCOSTO.Open "SELECT CODCCOSTO, NOMBRE FROM CCOSTOS ORDER BY CODCCOSTO", DBSYSTEM, adOpenKeyset, adLockReadOnly
    RSTRABS.Open "TRABAJADORES", DBSYSTEM, adOpenKeyset, adLockOptimistic
    RSUBIGEO.Open "VWUBIGEO", DBSYSTEM, adOpenStatic
    For X = 0 To 1
        xSexo.AddItem ARRSEXO(X)
    Next
    For X = 0 To 8
        xSituacion.AddItem ARRSITUACION(X)
    Next
    For X = 0 To 3
        xEstadoCivil.AddItem ARRESTCIVIL(X)
    Next
    SSTab1.TabVisible(2) = False 'CAMBIO
    CMDIMPRIMIR.Enabled = False
    If VPTAREA <> "NUEVO" Then
        CMDIMPRIMIR.Enabled = True
        'CUANDO NO ES NUEVO, ARRASTRA EL CÓDIGO A EDITAR
        SSTab1.TabVisible(2) = True 'CAMBIO
        RSTRABS.FIND "CODTRAB='" & VPTAREA & "'"
        If RSTRABS.EOF Then Unload Me
        VACIARDATOS
    End If
End Sub
Private Sub REFRESCARINGRESOS() 'CAMBIO
    'REFRESCANDO SUS INGRESOS
    Dim RSINGR As New ADODB.Recordset
    Dim I As Integer, SUMINGR As Double, SUMEGRE As Double
    Dim SUMINGRD As Double, SUMEGRED As Double
    SUMINGR = 0: SUMEGRE = 0: SUMINGRD = 0: SUMEGRED = 0
    
    RSINGR.Open "" & _
    "SELECT *, (SELECT MON = CASE MONEDA WHEN 0 THEN 'SOLES' ELSE 'DOLARES' END) AS MON," & _
    "(SELECT SOLES=CASE MONEDA WHEN 0 THEN CAPITAL ELSE 0 END) AS SOLES," & _
    "(SELECT DOLARES=CASE MONEDA WHEN 1 THEN CAPITAL ELSE 0 END) DOLARES" & _
    " FROM MOVICTA WHERE CODTRAB='" & xCodTrab.Text & "' AND TIPOGRUPO=1 AND SALDO<>0", _
    DBSYSTEM, adOpenKeyset
    Set DtIngr.DataSource = RSINGR
    
    'CALCULANDO EL TOTAL DE INGRESOS
    If RSINGR.RecordCount > 0 Then
        RSINGR.MoveFirst
        Do While Not (RSINGR.EOF)
            SUMINGR = SUMINGR + RSINGR!SOLES
            SUMINGRD = SUMINGRD + RSINGR!DOLARES
            RSINGR.MoveNext
        Loop
    End If
    LbIngr.Caption = Format(SUMINGR, "###,###,##0.00")
    lbingrD.Caption = Format(SUMINGRD, "###,###,##0.00")
End Sub
Private Sub REFRESCAEGRESOS() 'CAMBIO
    'REFRESCANDO SUS EGRESOS
    Dim RSEGRE As New ADODB.Recordset
    RSEGRE.Open "" & _
    "SELECT *, (SELECT MON =CASE MONEDA WHEN 0 THEN 'SOLES' ELSE 'DOLARES' END) AS MON," & _
    "(SELECT SOLES =CASE MONEDA WHEN 0 THEN CAPITAL ELSE 0 END) AS SOLES," & _
    "(SELECT DOLARES=CASE MONEDA WHEN 1 THEN CAPITAL ELSE 0 END) AS DOLARES " & _
    "FROM MOVICTA WHERE CODTRAB='" & xCodTrab.Text & "' AND TIPOGRUPO=2 AND SALDO<>0", _
    DBSYSTEM, adOpenKeyset
    Set DtEgre.DataSource = RSEGRE
    
    'CALCULANDO EL TOTAL DE EGRESOS
    If RSEGRE.RecordCount > 0 Then
        RSEGRE.MoveFirst
        For I = 1 To RSEGRE.RecordCount
            SUMEGRE = SUMEGRE + RSEGRE!SOLES
            SUMEGRED = SUMEGRED + RSEGRE!DOLARES
            RSEGRE.MoveNext
        Next
    End If
    LbEgre.Caption = Format(SUMEGRE, "###,###,##0.00")
    LbEgreD.Caption = Format(SUMEGRED, "###,###,##0.00")
End Sub
Private Sub FORM_UNLOAD(CANCEL As Integer)
    SetPosition Me
    Set RSAFP = Nothing
    Set RSTIPDOC = Nothing
    Set RSBANCO = Nothing
    Set RSTIPTRAB = Nothing
    Set RSSCTR = Nothing
    Set RSCCOSTO = Nothing
    Set RSTRABS = Nothing
    Set RSUBIGEO = Nothing
End Sub

Private Sub FRAME4_Click()
    MsgBox "PRODUCTO DE ENTERPRISE SOLUTIONS S.A.", vbInformation
End Sub

Private Sub IMAGE1_CLICK()
    MsgBox "SISTEMA DE PLANILLAS" & Chr(13) & Chr(10) & "DESARROLLADO POR MARFICE S.A. " & Chr(13) & Chr(10) & "PROGRAMADO POR FERNANDO COSSIO ", vbInformation
End Sub

Private Sub QUITARDATO_Click()
On Error GoTo ERRPR
    If xLista.ListItems.Count = 0 Then Exit Sub
    DBSYSTEM.Execute "UPDATE TRABAJADORES SET " & xLista.SelectedItem.KEY & "=NULL WHERE CODTRAB='" & Trim(xCodTrab.Text) & "'"
    REFRESCAROTRADATA
    Exit Sub
ERRPR:
    MsgBox "Este tipo de Campo una vez creado no se puede eliminar", vbExclamation
End Sub

Private Sub SSTAB1_Click(PREVIOUSTAB As Integer)
    Select Case SSTab1.Tab
        Case 2
            REFRESCARINGRESOS
            REFRESCAEGRESOS
        Case 3
            If Not ExisteTabla("DATATRAB") Then
                DBSYSTEM.Execute "CREATE TABLE DATATRAB (CODDATA VARCHAR(15),DESCDATA VARCHAR(30),TIPODATA VARCHAR(1))"
                MsgBox "SE ACTUALIZO EL SISTEMA DE PLANILLAS HA ACTUALIZADO SU SISTEMA", vbInformation
            End If
            REFRESCAROTRADATA
            Command4.SetFocus
    End Select
End Sub

Private Sub XBANCO_DblClick()
    frmComun.CONECTAR RSBANCO
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xBanco.Text = VGUTIL(1) & " :  " & VGUTIL(2)
        xBanco.Tag = VGUTIL(1)
    End If
End Sub

Private Sub XBANCOCTS_DblClick()
    frmComun.CONECTAR RSBANCO
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xBancoCTS.Text = VGUTIL(1) & " :  " & VGUTIL(2)
        xBancoCTS.Tag = VGUTIL(1)
    End If
End Sub

Private Sub XCCOSTO_DBLCLICK()
    frmComun.CONECTAR RSCCOSTO
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xCCosto.Text = VGUTIL(1) & " :  " & VGUTIL(2)
        xCCosto.Tag = VGUTIL(1)
    End If
End Sub

Private Sub XCODCTR_DblClick()
    frmComun.CONECTAR RSSCTR
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xCodCTR.Text = VGUTIL(1) & " :  " & VGUTIL(2)
        xCodCTR.Tag = VGUTIL(1)
    End If
End Sub

Private Sub XCONTRATO_Click()
    If xContrato.ListIndex = 0 Then
        xFechaTermino.Visible = False
        Label1.Visible = False
    Else
        xFechaTermino.Visible = True
        Label1.Visible = True
    End If
End Sub

Private Sub XCONTRATO_KeyDown(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 13 Then SendKeys "{TAB}"
End Sub

Private Sub XDEPARTAMENTO_DblClick()
    Dim RSAREA As New ADODB.Recordset
    RSAREA.Open "SELECT CODCCOSTO, NOMBRE FROM AREASTRAB ORDER BY CODCCOSTO", DBSYSTEM, adOpenStatic
    frmComun.CONECTAR RSAREA
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xDepartamento.Text = VGUTIL(1) & " :  " & VGUTIL(2)
        xDepartamento.Tag = VGUTIL(1)
    End If
    Set RSAREA = Nothing
End Sub

Private Sub XESSALUDVIDA_KeyDown(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 13 Then SendKeys "{TAB}"
End Sub

Private Sub XESTADOCIVIL_KeyDown(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 13 Then SendKeys "{TAB}"
End Sub

Private Sub XFECHACESE_KeyDown(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 13 Then SendKeys "{TAB}"
End Sub

Private Sub XFECHAIAFP_KeyDown(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 13 Then SendKeys "{TAB}"
End Sub

Private Sub XFECHAING_GOTFOCUS()
    If SSTab1.Tab <> 2 Then
        SSTab1.Tab = 1
        xFechaIng.SetFocus
    End If
End Sub

Private Sub XFECHAING_KeyDown(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 13 Then SendKeys "{TAB}"
End Sub

Private Sub XFECHANAC_KeyDown(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 13 Then SendKeys "{TAB}"
End Sub

Private Sub XFECHATERMINO_KeyDown(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 13 Then SendKeys "{TAB}"
End Sub

Private Sub XFONDOPENS_DblClick()
    frmComun.CONECTAR RSAFP
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xFondoPens.Text = VGUTIL(1) & " :  " & VGUTIL(2)
        xFondoPens.Tag = VGUTIL(1)
    End If
End Sub

Private Sub XLISTA_DblClick()
    If xLista.ListItems.Count = 0 Then Exit Sub
    Load frIngDatos
    frIngDatos.xConcepto.Text = xLista.SelectedItem.SubItems(1)
    frIngDatos.xConcepto.Tag = xLista.SelectedItem.Text
    frIngDatos.cmAceptar.Tag = xLista.SelectedItem.Tag
    
    If UCase(frIngDatos.xConcepto.Tag) = "XXCTADES" Then
            frIngDatos.Command1.Visible = True
    Else
            frIngDatos.Command1.Visible = False
    End If
    Select Case xLista.SelectedItem.Tag
        Case "T"
            frIngDatos.tipoT.Visible = True
            frIngDatos.tipoT.Text = xLista.SelectedItem.SubItems(2)
        Case "N"
            frIngDatos.tipoN.Visible = True
            frIngDatos.tipoN.Text = Val(xLista.SelectedItem.SubItems(2))
        Case "F"
            frIngDatos.tipoF.Visible = True
            frIngDatos.tipoF.Value = CDate(xLista.SelectedItem.SubItems(2))
        Case "B"
            frIngDatos.tipoBNo.Visible = True
            frIngDatos.tipoBSi.Visible = True
            If xLista.SelectedItem.SubItems(2) = "SI" Then
                frIngDatos.tipoBSi.Value = True
            Else
                frIngDatos.tipoBNo.Value = True
            End If
    End Select
    frIngDatos.Show 1
    REFRESCAROTRADATA
End Sub

Private Sub XMESDEVENGUE_KeyDown(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 13 Then SendKeys "{TAB}"
End Sub

Private Sub XNOPDT_KeyDown(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 13 Then SendKeys "{TAB}"
End Sub

Private Sub XOPCION01_KeyDown(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 13 Then SendKeys "{TAB}"
End Sub

Private Sub XOPCION02_KeyDown(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 13 Then SendKeys "{TAB}"
End Sub

Private Sub XSEXO_KeyDown(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 13 Then SendKeys "{TAB}"
End Sub

Private Sub XSITUACION_KeyDown(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 13 Then SendKeys "{TAB}"
End Sub

Private Sub XTIPDOC_DblClick()
    frmComun.CONECTAR RSTIPDOC
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xTipDoc.Text = VGUTIL(1) & " :  " & VGUTIL(2)
        xTipDoc.Tag = VGUTIL(1)
    End If
End Sub

Private Sub XTIPDOC_GOTFOCUS()
    If SSTab1.Tab <> 0 Then
        SSTab1.Tab = 0
    End If
End Sub
Private Sub XTIPOTRAB_DblClick()
    frmComun.CONECTAR RSTIPTRAB
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xTipoTrab.Text = VGUTIL(1) & " :  " & VGUTIL(2)
        xTipoTrab.Tag = VGUTIL(1)
    End If
End Sub
Public Function OKPARAEDITAR() As Boolean
Dim strCadenaAux As String
    OKPARAEDITAR = False
    Dim A As Integer
    If xApePat.Text = "" Then
        MsgBox "FALTA APELLIDO PATERNO", vbCritical
        xApePat.SetFocus
        Exit Function
    End If
    If xApeMat.Text = "" Then
        MsgBox "FALTA APELLIDO MATERNO", vbCritical
        xApeMat.SetFocus
        Exit Function
    End If
    If xNombre.Text = "" Then
        MsgBox "FALTA NOMBRE DEL TRABAJADOR", vbCritical
        xNombre.SetFocus
        Exit Function
    End If
    A = DateDiff("YYYY", xFechaNac.Value, Date)
    If A > 99 Or A < 17 Then
        MsgBox "LA EDAD CALCULADA DE ACUERDO A LA FECHA DE INGRESO Y LA FECHA ACTUAL NO ES VÁLIDA", vbCritical
        xFechaNac.SetFocus
        Exit Function
    End If
    A = DateDiff("YYYY", xFechaNac.Value, xFechaIng.Value)
    If A < 17 Then
        MsgBox "FECHA DE INGRESO O FECHA DE NACIMIENTO INCORRECTA. REVISE LOS DATOS INGRESADOS", vbCritical
        xFechaNac.SetFocus
        Exit Function
    End If
    If xUbigeo.Tag = "" Then
        MsgBox "DEBE SELECCIONAR UN CÓDIGO DE UBICACIÓN GEOGRÁFICA", vbCritical
        xUbigeo.SetFocus
        Exit Function
    End If
    If xCCosto.Tag = "" Then
        MsgBox "DEBE SELECCIONAR UN CENTRO DE COSTO VÁLIDO"
        xCCosto.SetFocus
        Exit Function
    End If
    If xTipDoc.Tag = "" Then
        MsgBox "DEBE SELECCIONAR UN TIPO DE DOCUMENTO VÁLIDO", vbCritical
        xTipDoc.SetFocus
        Exit Function
    End If
    If xEstadoCivil.ListIndex = -1 Then
        MsgBox "DEBE SELECCIONAR UN ESTADO CIVIL VÁLIDO", vbCritical
        xEstadoCivil.SetFocus
        Exit Function
    End If
    If xUbigeo.Tag = "" Then
        MsgBox "EL CÓDIGO DE UBICACIÓN GEOGRÁFICA NO ES VÁLIDO", vbCritical
        xUbigeo.SetFocus
        Exit Function
    End If
    If xSexo.ListIndex = -1 Then
        MsgBox "EL TIPO DE SEXO DEL TRABAJADOR NO SE HA DEFINIDO", vbCritical
        xSexo.SetFocus
        Exit Function
    End If
    If xTipoTrab.Tag = "" Then
        MsgBox "NO SE HA DEFINIDO UN TIPO DE TRABAJADOR VÁLIDO", vbCritical
        xTipoTrab.SetFocus
        Exit Function
    End If
    If xSituacion.ListIndex = -1 Then
        MsgBox "LA SITUACIÓN ACTUAL DEL TRABAJADOR NO ES VÁLIDA", vbCritical
        xSituacion.SetFocus
        Exit Function
    End If
    If xDepartamento.Text = "" Then
        MsgBox "EL NOMBRE DEL DEPARTAMENTO NO ES VÁLIDO", vbCritical
        xDepartamento.SetFocus
        Exit Function
    End If
    If xCargo.Text = "" Then
        MsgBox "EL NOMBRE DEL CARGO NO ES VÁLIDO", vbCritical
        xCargo.SetFocus
        Exit Function
    End If
    If xCtaBanco.Text <> "" Then
        'PARA VALIDAR QUE UNA CTA NO SE REPITA
'        strCadenaAux = DevuelveValor("select DATOS=APEPAT +' ' + APEMAT + ' ' + NOMBRE from TRABAJADORES WHERE CTABANCO='" & xCtaBanco.Text & "'", DBSYSTEM)
'        If Trim(strCadenaAux) <> "" Then
'            MsgBox "Ya existe la cta del Banco!! ,le pertenece al Sr(a):: " & strCadenaAux, vbCritical
'            xCtaBanco.SetFocus
'            Exit Function
'        End If
        If xBanco.Tag = "" Then
            MsgBox "HA ASIGNADO UN NÚMERO DE CUENTA BANCARIA DE DEPÓSITO DE REMUNERACIONES PERO NO HA DEFINIDO UN BANCO. ASIGNE UN BANCO Y VUELVA A INTENTAR GRABAR", vbCritical
            xBanco.SetFocus
            Exit Function
        End If
    Else
        xBanco.Tag = "NONE"
    End If
    If xCtaCTS.Text <> "" Then
        If xBancoCTS.Tag = "" Then
            MsgBox "HA ASIGNADO UN NÚMERO DE CUENTA BANCARIA DE DEPÓSITO DE REMUNERACIONES PERO NO HA DEFINIDO UN BANCO. ASIGNE UN BANCO Y VUELVA A INTENTAR GRABAR", vbCritical
            xBancoCTS.SetFocus
            Exit Function
        End If
    Else
        xBancoCTS.Tag = "NONE"
    End If
    If xFondoPens.Tag = "" Then
        MsgBox "LA ADMINISTRADORA DE FONDO DE PENSIONES NO SE HA SELECCIONADO", vbCritical
        xFondoPens.SetFocus
        Exit Function
    End If
    If xMesDevengue.CheckBox Then
        If xMesDevengue.Value > Date Then
            MsgBox "LA FECHA DEL PRIMER DEVENGUE, CORRESPONDIENTE A LAS AFP NO ES VÁLIDO. VALOR NO PUEDE SER MAYOR A LA FECHA ACTUAL", vbCritical
            xMesDevengue.SetFocus
            Exit Function
        End If
    End If
    If xFechaCese.CheckBox Then
        If xFechaCese.Value <= xFechaIng.Value Then
            MsgBox "LA FECHA DE CESE NO PUEDE SER MENOR O IGUAL A LA FECHA DE INGRESO", vbCritical
            xFechaCese.SetFocus
            Exit Function
        End If
    Else
        If InStr(xSituacion.Text, "ACTIVO") > 0 Then
            MsgBox "NO PUEDE TENER UNA FECHA DE CESE VÁLIDO Y ESTAR EN SITUACIÓN ACTIVO", vbCritical
            xSituacion.SetFocus
            Exit Function
        End If
    End If
    If xRucEPS.Text <> "" Then
        If Not Validar_RUC(xRucEPS.Text) Then
            MsgBox "NÚMERO DE RUC DE LA EPS NO ES VÁLIDO", vbCritical
            xRucEPS.SetFocus
            Exit Function
        End If
    End If
    If xCodCTR.Tag = "" Then
        xCodCTR.Tag = "NONE"
    End If
    OKPARAEDITAR = True
End Function
Public Function OKPARANUEVOS() As Boolean
    OKPARANUEVOS = False
    If xCodTrab.Locked Then
        Dim X As Integer, Y As Integer, xCod As String
        X = 1: Y = 1
        Do While X <> 0
            xCod = UCase(Left(xApePat.Text, 1) & Left(xApeMat.Text, 1) & Left(xNombre.Text, 1) & Format(Y, "000"))
            DBSYSTEM.Execute "UPDATE TRABAJADORES SET BASICO=BASICO WHERE CODTRAB='" & xCod & "'", X
            Y = Y + 1
        Loop
        xCodTrab.Text = xCod
    Else
        If xCodTrab.Text = "" Then
            MsgBox "FALTA CÓDIGO DEL TRABAJADOR"
            xCodTrab.SetFocus
            Exit Function
        End If
    End If
    If Not RSTRABS.EOF Then
        RSTRABS.MoveFirst
        RSTRABS.FIND "DOCIDEN='" & xDocIden.Text & "'"
        Dim A As Integer
        If Not RSTRABS.EOF Then
            MsgBox "EL NÚMERO DE DOCUMENTO DE IDENTIDAD YA EXISTE, POR FAVOR INGRESE UN NÚMERO DIFERENTE", vbCritical
            xDocIden.SetFocus
            Exit Function
        End If
        RSTRABS.MoveFirst
        RSTRABS.FIND "CODTRAB='" & xCodTrab.Text & "'"
        If Not RSTRABS.EOF Then
            MsgBox "EL CÓDIGO DEL TRABAJADOR YA EXISTE, POR FAVOR INGRESE UNO DIFERENTE", vbCritical
            xCodTrab.SetFocus
            Exit Function
        End If
    End If
    OKPARANUEVOS = True
End Function
Public Sub GRABAR(ByVal CodigoTrab As String)
Dim FECHACES As String
    If Not IsNull(xFechaCese) Then FECHACES = DateSQL(xFechaCese.Value)
    DBSYSTEM.Execute "UPDATE TRABAJADORES SET APEPAT='" & xApePat.Text & "'," _
        & "APEMAT= '" & xApeMat.Text & "',NOMBRE='" & xNombre.Text & "'," _
        & "TIPDOC='" & xTipDoc.Tag & "',DOCIDEN='" & xDocIden.Text & "'," _
        & "FECHANAC=" & DateSQL(xFechaNac.Value) & "," _
        & "ESTADOCIVIL=" & xEstadoCivil.ListIndex & "," _
        & "NOCALCULO=" & IIf(xNoCalculo, -1, 0) & "," _
        & "UBIGEO='" & xUbigeo.Tag & "',DIRECCIÓN='" & xDireccion.Text & " '," _
        & "TELEFONO='" & xTelefono.Text & " ',SEXO=" & xSexo.ListIndex & "," _
        & "TIPOTRAB='" & xTipoTrab.Tag & "',FECHAING=" & DateSQL(xFechaIng.Value) & "," _
        & "SITUACIÓN='" & xSituacion.ListIndex & "',AREA='" & xDepartamento.Tag & "'," _
        & "CCOSTO='" & xCCosto.Tag & "',DEPARTAMENTO='" & xDepartamento.Text & "'," _
        & "CARGO='" & xCargo.Text & "',CTABANCO='" & xCtaBanco.Text & " '," _
        & "BANCO='" & xBanco.Tag & "',CTACTS='" & xCtaCTS.Text & " '," _
        & "BANCOCTS='" & xBancoCTS.Tag & "',BASICO=" & Val(xBasico.Text) & "," _
        & "NUMFICHA='" & xNumFicha.Text & " ',CARNETSEG='" & xCarnetSeg.Text & "'," _
        & "FONDOPENS='" & xFondoPens.Tag & "',CUSPP ='" & xCuspp.Text & " '," _
        & "MESDEVENGUE=" & DateSQL(IIf(IsNull(xMesDevengue.Value), Date, xMesDevengue.Value)) _
        & " WHERE CODTRAB='" & CodigoTrab & "'"
    DBSYSTEM.Execute "UPDATE TRABAJADORES SET ESSALUDVIDA=" & IIf(xEsSaludVida.Value = 0, 0, -1) & "," _
        & "ASIGFAM=" & Val(xAsigFam.Text) & "," _
        & "ESTADOINTERNO=0,CODIGOALT='" & xCodAlt.Text & " '," _
        & "CODSCTR='" & xCodCTR.Tag & "',RUCEPS='" & xRucEPS.Text & " '," _
        & "TIPOCONTRATO=" & xContrato.ListIndex & "," _
        & IIf(xContrato.ListIndex = 1, "FECHATERMINO=" & DateSQL(xFechaTermino.Value) & ",", "") _
        & "FECHAIAFP=" & DateSQL(xFechaIAFP.Value) & ",TIPOSISTEMA='PL'," _
        & "FECHAEDIT=" & DateSQL(Date) & ",NUMEDIT=NUMEDIT + 1," _
        & "NOPDT=" & xNoPDT.Value & ",OPCION01=" & xOpcion01.Value & "," _
        & "OPCION02=" & xOpcion02.Value & ",AFECTOQUINTA=" & Check1.Value & ",OPCIONA='" & xOpcionA.Text & " '," _
        & "OPCIONB='" & xOpcionB.Text & " '" _
        & IIf(VPTAREA = "NUEVO", ",CODTRAB='" & xCodTrab.Text & "'", "") _
        & IIf(IsNull(xFechaCese.Value), ",FECHACESE=NULL", ",FECHACESE=" & FECHACES) _
        & ", TOTALEXTRA=" & xOpQuinta.Text & " WHERE CODTRAB='" & CodigoTrab & "'"
        On Error GoTo ERRFOTO
        'OBETENER LA EXTENSION DEL ARCHIVO FOTO
        Dim EXT As String
        If xFoto.Tag <> "" Then
            FileCopy xFoto.Tag, (REGSISTEMA.PATHFOTOS & "\" & xCodTrab.Text & ".FTE")  '& EXT)
        Else
            Kill (REGSISTEMA.PATHFOTOS & "\" & xCodTrab.Text & ".FTE") ' & EXT)
        End If
        Exit Sub
ERRFOTO:
        Resume Next
End Sub

Public Sub VACIARDATOS()
On Error GoTo ERRVACIAR
    With RSTRABS
        xCodTrab.Text = !CODTRAB
        xApePat.Text = Trim(!ApePat)
        xApeMat.Text = Trim(!ApeMat)
        xNombre.Text = Trim(!NOMBRE)
        xTipDoc.Tag = !TIPDOC
        xDocIden.Text = "" & !DOCIDEN
        xFechaNac.Value = !FechaNac
        xEstadoCivil.ListIndex = !ESTADOCIVIL
        xNoCalculo.Value = IIf(!NOCALCULO = -1, 1, 0)
        xUbigeo.Tag = "" & !UBIGEO
        xDireccion.Text = Trim("" & !DIRECCIÓN)
        xTelefono.Text = Trim("" & !TELEFONO)
        xSexo.ListIndex = !Sexo
        xTipoTrab.Tag = !TIPOTRAB
        xFechaIng.Value = !FECHAING
        xSituacion.ListIndex = !SITUACIÓN
        xDepartamento.Tag = !AREA
        xCCosto.Tag = !CCosto
        'XDEPARTAMENTO.TEXT = "" & !DEPARTAMENTO
        xCargo.Text = "" & !CARGO
        xCtaBanco.Text = "" & !CTABANCO
        xBanco.Tag = !BANCO
        xCtaCTS.Text = "" & !CTACTS
        xBancoCTS.Tag = !BANCOCTS
        xBasico.Text = Format(!BASICO, "0.00")
        xNumFicha.Text = "" & !NUMFICHA
        xCarnetSeg.Text = "" & !CARNETSEG
        xFondoPens.Tag = !FONDOPENS
        xCuspp.Text = "" & !CUSPP
        xMesDevengue.Value = "" & !MESDEVENGUE
        xFechaIAFP.Value = "" & !FECHAIAFP
        xEsSaludVida.Value = IIf(!ESSALUDVIDA, 1, 0)
        xAsigFam.Text = Format(!ASIGFAM, "0.00")
        xFechaCese.Value = Trim("" & !FECHACESE)
        xCodAlt.Text = Trim("" & !CODIGOALT)
        xCodCTR.Tag = !CODSCTR
        xRucEPS.Text = Trim("" & !RUCEPS)
        xContrato.ListIndex = 0 + !TIPOCONTRATO
        If !TIPOCONTRATO = 1 Then
            If Not IsNull(!FECHATERMINO) Then xFechaTermino.Value = !FECHATERMINO
        End If
        xOpcion01.Value = !OPCION01
        xOpcion02.Value = !OPCION02
        xNoPDT.Value = !NOPDT
        xOpcionA.Text = !OPCIONA
        xOpcionB.Text = !OPCIONB
        Check1.Value = IIf(!AFECTOQUINTA, 1, 0)
        xOpQuinta.Text = !TOTALEXTRA
        
        xFoto.Picture = LoadPicture(REGSISTEMA.PATHFOTOS & "\" & xCodTrab.Text & ".FTE")
        If xFoto.Picture <> 0 Then
             xFoto.Tag = REGSISTEMA.PATHFOTOS & "\" & xCodTrab.Text & ".FTE"
        End If
    End With
    Dim RSAUX2 As New ADODB.Recordset
    RSAUX2.Open "SELECT CODCCOSTO, NOMBRE FROM AREASTRAB WHERE CODCCOSTO='" & xDepartamento.Tag & "'", DBSYSTEM, adOpenStatic
    If RSAUX2.EOF Then
        MsgBox "NO SE ENCUENTRA EL AREA DE TRABAJO DEL TRABAJADOR, SELECCIONAR OTRO", vbCritical
    Else
        xDepartamento.Text = RSAUX2!CODCCOSTO & " : " & RSAUX2!NOMBRE
    End If
    Set RSAUX2 = Nothing
    RSTIPDOC.FIND "TIPDOC='" & xTipDoc.Tag & "'"
    If RSTIPDOC.EOF Then
        MsgBox "EL TIPO DE DOCUMENTO AL QUE SE REFERIA EL REGISTRO YA NO EXISTE, SELECCIONE OTRO", vbCritical
        xTipDoc.Tag = ""
    Else
        xTipDoc.Text = RSTIPDOC!TIPDOC & " :  " & RSTIPDOC!DESCRIP
    End If
    RSUBIGEO.FIND "CODIGO='" & xUbigeo.Tag & "'"
    If RSUBIGEO.EOF Then
        MsgBox "EL CÓDIGO DE UBICACIÓN GEOGRÁFICA YA NO EXISTE, SELECCIONE OTRO", vbCritical
        xUbigeo.Tag = ""
    Else
        xUbigeo.Text = RSUBIGEO!Codigo & " : " & RSUBIGEO!LUGAR
    End If
    RSTIPTRAB.MoveFirst
    RSTIPTRAB.FIND "TIPTRAB='" & xTipoTrab.Tag & "'"
    If RSTIPTRAB.EOF Then
        MsgBox "EL TIPO DE TRABAJADOR AL QUE SE REFIERE EL REGISTRO ACTUAL YA NO EXISTE, SELECCIONE OTRO", vbCritical
        xTipoTrab.Tag = ""
    Else
        xTipoTrab.Text = RSTIPTRAB!TIPTRAB & " :  " & RSTIPTRAB!DESCRIP
    End If
    RSCCOSTO.FIND "CODCCOSTO='" & xCCosto.Tag & "'"
    If RSCCOSTO.EOF Then
        MsgBox "EL CENTRO DE COSTO AL QUE SE REFIERE YA NO EXISTE, SELECCIONE OTRO", vbCritical
        xCCosto.Tag = ""
    Else
        xCCosto.Text = RSCCOSTO!CODCCOSTO & " :  " & RSCCOSTO!NOMBRE
    End If
    RSBANCO.FIND "CODBANCO='" & xBanco.Tag & "'"
    If RSBANCO.EOF Then
        MsgBox "EL CÓDIGO DEL BANCO DE LA CUENTA BANCARIA DE DEPÓSITO DE REMUNERACIONES NO EXISTE, SELECCIONE OTRO", vbCritical
        xBanco.Tag = ""
    Else
        xBanco.Text = RSBANCO!CODBANCO & " :  " & RSBANCO!NOMBRE
    End If
    RSBANCO.MoveFirst
    RSBANCO.FIND "CODBANCO='" & xBancoCTS.Tag & "'"
    If RSBANCO.EOF Then
        MsgBox "EL CÓDIGO DEL BANCO DE LA CUENTA BANCARIA DE DEPÓSITO DE CTS NO EXISTE, SELECCIONE OTRO", vbCritical
        xBancoCTS.Tag = ""
    Else
        xBancoCTS.Text = RSBANCO!CODBANCO & " :  " & RSBANCO!NOMBRE
    End If
    RSAFP.FIND "CODAFP='" & xFondoPens.Tag & "'"
    If RSAFP.EOF Then
        MsgBox "LA ADMINISTRADORA DE FONDO DE PENSIONES NO EXISTE, SELECCIONE OTRA", vbCritical
        xFondoPens.Tag = ""
    Else
        xFondoPens.Text = RSAFP!CODAFP & " :  " & RSAFP!NOMBRE
    End If
    RSSCTR.FIND "CODCAR='" & xCodCTR.Tag & "'"
    If RSSCTR.EOF Then
        MsgBox "EL REGISTRO REFERENCIADO DEL CENTRO DE ALTO RIESGO - SCTR YA NO EXISTE, SELECCIONE OTRO", vbCritical
        xCodCTR.Tag = ""
    Else
        xCodCTR.Text = RSSCTR!CODCAR & " :  " & RSSCTR!NOMBRE
    End If
    Exit Sub
ERRVACIAR:
    Resume Next
End Sub

Private Sub XUBIGEO_DblClick()
    frUbigeo.Show 1
    If VPCODTMP <> "" Then
        xUbigeo.Tag = VPCODTMP
        xUbigeo.Text = VPTRASPRM
    End If
End Sub

Public Sub REFRESCAROTRADATA()
    xLista.ListItems.Clear
    If Trim(xCodTrab.Text) = "" Then
        MsgBox "IMPOSIBLE ACTUALIZAR EL REGISTRO, PORQUE LA FICHA DEL TRABAJADOR NO SE HA GRABADO.", vbInformation
        Exit Sub
    End If
    Dim RSOTROS As New ADODB.Recordset
    Dim RSTRABZ As New ADODB.Recordset
    RSTRABZ.Open "SELECT * FROM TRABAJADORES WHERE CODTRAB='" & xCodTrab.Text & "'", DBSYSTEM, adOpenStatic, adLockReadOnly
    If RSTRABZ.RecordCount = 0 Then
        MsgBox "IMPOSIBLE ACTUALIZAR EL REGISTRO. PUEDE DEBERSE A QUE LA FICHA NUEVA DEL TRABAJADOR NO HA SIDO GRABADA AÚN O QUE EL REGISTRO DEL TRABAJADOR HAYA SIDO ELIMINADO POR OTRO USUARIO", vbInformation
        Set RSTRABZ = Nothing
        Exit Sub
    End If
    Dim XITEM As ListItem
    RSOTROS.Open "SELECT * FROM DATATRAB ORDER BY DESCDATA", DBSYSTEM, adOpenStatic, adLockReadOnly
    If RSOTROS.RecordCount > 0 Then RSOTROS.MoveFirst
    Do While Not RSOTROS.EOF
        If Not IsNull(RSTRABZ.Fields(Trim$(RSOTROS!CODDATA)).Value) Then
            Set XITEM = xLista.ListItems.Add(, RSOTROS!CODDATA, RSOTROS!CODDATA, , 1)
            XITEM.SubItems(1) = RSOTROS!DESCDATA
            Select Case RSOTROS!TIPODATA
                Case "T"
                    XITEM.SubItems(2) = RSTRABZ.Fields(Trim$(RSOTROS!CODDATA)).Value
                Case "N"
                    XITEM.SubItems(2) = Format(RSTRABZ.Fields(Trim$(RSOTROS!CODDATA)).Value, "0.00")
                Case "B"
                    If RSTRABZ.Fields(Trim$(RSOTROS!CODDATA)).Value Then
                        XITEM.SubItems(2) = "SI"
                    Else
                        XITEM.SubItems(2) = "NO"
                    End If
                Case "F"
                    XITEM.SubItems(2) = Format(RSTRABZ.Fields(Trim$(RSOTROS!CODDATA)).Value, "DD/MM/YYYY")
            End Select
            XITEM.Tag = RSOTROS!TIPODATA
        End If
        RSOTROS.MoveNext
    Loop
    Set RSOTROS = Nothing
    Set RSTRABZ = Nothing
End Sub


