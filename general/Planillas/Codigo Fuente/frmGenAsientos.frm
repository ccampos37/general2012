VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmGenAsientos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generar Asientos para Contabilidad"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8595
   Icon            =   "frmGenAsientos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   8595
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar pgbar 
      Height          =   165
      Left            =   75
      TabIndex        =   37
      Top             =   5175
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   8493
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Generar Asientos Mensual"
      TabPicture(0)   =   "frmGenAsientos.frx":044A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "l2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "l1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "xTotPlan"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Xntrab"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Lista"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "xFechaIni"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "xFechaFin"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "xMes"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "CmdSalir"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdGenerar"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Command1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "xDgResult"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "CmdEnviar"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "CmdEliminar"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Command4"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "ChkGenAnx"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Check1"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).ControlCount=   21
      TabCaption(1)   =   "Resultado"
      TabPicture(1)   =   "frmGenAsientos.frx":0466
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Generar Asiento por Adelanto de Quincena"
      TabPicture(2)   =   "frmGenAsientos.frx":0482
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "CmdSalirQuin"
      Tab(2).Control(1)=   "CmdEnviarQuin"
      Tab(2).Control(2)=   "ProgressBar2"
      Tab(2).Control(3)=   "Command6"
      Tab(2).Control(4)=   "CmdExa"
      Tab(2).Control(5)=   "Command2"
      Tab(2).Control(6)=   "chkCrearAnexo"
      Tab(2).Control(7)=   "chkConsiderarAdelantoDetallado"
      Tab(2).Control(8)=   "LstvwCronograma"
      Tab(2).Control(9)=   "dgComprobante"
      Tab(2).Control(10)=   "AplitxtMesTrabajo"
      Tab(2).Control(11)=   "ProgressBar1"
      Tab(2).Control(12)=   "Command3"
      Tab(2).Control(13)=   "lblMessage2"
      Tab(2).Control(14)=   "Label16"
      Tab(2).Control(15)=   "Label14"
      Tab(2).Control(16)=   "Label11"
      Tab(2).Control(17)=   "Label10"
      Tab(2).Control(18)=   "lblTotalPlanilla"
      Tab(2).Control(19)=   "lblNroTrabajadores"
      Tab(2).Control(20)=   "lblMessage"
      Tab(2).ControlCount=   21
      TabCaption(3)   =   "Resultado Quincena"
      TabPicture(3)   =   "frmGenAsientos.frx":049E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame2"
      Tab(3).ControlCount=   1
      Begin VB.CommandButton CmdSalirQuin 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   375
         Left            =   -67935
         TabIndex        =   58
         Top             =   4260
         Width           =   1275
      End
      Begin VB.CommandButton CmdEnviarQuin 
         Caption         =   "&Enviar a Contabilidad"
         Height          =   375
         Left            =   -74880
         TabIndex        =   57
         Top             =   4245
         Width           =   2085
      End
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   180
         Left            =   -74865
         TabIndex        =   79
         Top             =   4575
         Visible         =   0   'False
         Width           =   8145
         _ExtentX        =   14367
         _ExtentY        =   318
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Frame Frame2 
         Height          =   4125
         Left            =   -74775
         TabIndex        =   59
         Top             =   450
         Width           =   8010
         Begin VB.CommandButton CmdAntQuin 
            Caption         =   "<< &Regresar"
            Height          =   330
            Left            =   165
            TabIndex        =   62
            Top             =   3690
            Width           =   1215
         End
         Begin VB.CommandButton CMDIMPRIMIRQuin 
            Caption         =   "&Imprimir"
            Height          =   330
            Left            =   1425
            TabIndex        =   61
            Top             =   3690
            Width           =   1215
         End
         Begin VB.TextBox XNUMQuin 
            Height          =   300
            Left            =   1620
            TabIndex        =   60
            Text            =   "XNUMQuin"
            Top             =   2235
            Visible         =   0   'False
            Width           =   1035
         End
         Begin Crystal.CrystalReport CrystalReport1 
            Left            =   3435
            Top             =   1935
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            PrintFileLinesPerPage=   60
         End
         Begin MSDataGridLib.DataGrid xDGDetaQuin 
            Height          =   2580
            Left            =   165
            TabIndex        =   63
            Top             =   1035
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   4551
            _Version        =   393216
            AllowUpdate     =   0   'False
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
            Caption         =   "Detalle del Asiento"
            ColumnCount     =   6
            BeginProperty Column00 
               DataField       =   "DMOV_SECUE"
               Caption         =   "Secue"
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
               DataField       =   "DMOV_CUENT"
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
               DataField       =   "DMOV_ANEXO"
               Caption         =   "Anexo"
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
               DataField       =   "DMOV_CENCO"
               Caption         =   "C.Costo"
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
               DataField       =   "DMOV_DEBE"
               Caption         =   "DEBE"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "###,###,###.00 "
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "DMOV_HABER"
               Caption         =   "HABER"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "###,###,###.00 "
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
                  ColumnWidth     =   689.953
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1110.047
               EndProperty
               BeginProperty Column02 
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1124.787
               EndProperty
               BeginProperty Column04 
                  Alignment       =   1
                  ColumnWidth     =   1305.071
               EndProperty
               BeginProperty Column05 
                  Alignment       =   1
                  ColumnWidth     =   1319.811
               EndProperty
            EndProperty
         End
         Begin VB.Label Label27 
            Caption         =   "Subdiario"
            Height          =   300
            Left            =   150
            TabIndex        =   76
            Top             =   255
            Width           =   675
         End
         Begin VB.Label Label26 
            Caption         =   "Compr."
            Height          =   315
            Left            =   150
            TabIndex        =   75
            Top             =   645
            Width           =   570
         End
         Begin VB.Label lsubquin 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   900
            TabIndex        =   74
            Top             =   210
            Width           =   705
         End
         Begin VB.Label lcompquin 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   900
            TabIndex        =   73
            Top             =   570
            Width           =   1155
         End
         Begin VB.Label ltipcamquin 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.000 "
            Height          =   300
            Left            =   3630
            TabIndex        =   72
            Top             =   210
            Width           =   1275
         End
         Begin VB.Label Label22 
            Caption         =   "T/C. Venta"
            Height          =   210
            Left            =   2625
            TabIndex        =   71
            Top             =   255
            Width           =   915
         End
         Begin VB.Label Label21 
            Caption         =   "Glosa"
            Height          =   225
            Left            =   2625
            TabIndex        =   70
            Top             =   615
            Width           =   645
         End
         Begin VB.Label lglosaquin 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   3660
            TabIndex        =   69
            Top             =   570
            Width           =   4200
         End
         Begin VB.Label Label19 
            Caption         =   "fecha"
            Height          =   210
            Left            =   5580
            TabIndex        =   68
            Top             =   255
            Width           =   690
         End
         Begin VB.Label lfechaquin 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6585
            TabIndex        =   67
            Top             =   210
            Width           =   1275
         End
         Begin VB.Label Label17 
            Caption         =   "Total"
            Height          =   225
            Left            =   4005
            TabIndex        =   66
            Top             =   3810
            Width           =   510
         End
         Begin VB.Label ldebequin 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00 "
            Height          =   285
            Left            =   4920
            TabIndex        =   65
            Top             =   3765
            Width           =   1350
         End
         Begin VB.Label lhaberquin 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00 "
            Height          =   285
            Left            =   6270
            TabIndex        =   64
            Top             =   3765
            Width           =   1350
         End
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   -72735
         TabIndex        =   56
         Top             =   4260
         Width           =   1275
      End
      Begin VB.CommandButton CmdExa 
         Caption         =   "&E&xaminar Comprobante"
         Height          =   375
         Left            =   -71385
         TabIndex        =   55
         Top             =   4260
         Width           =   2025
      End
      Begin VB.CommandButton Command2 
         Caption         =   "E&xportar Trabajadores a Contabilidad"
         Enabled         =   0   'False
         Height          =   540
         Left            =   -74850
         TabIndex        =   53
         Top             =   1290
         Width           =   2265
      End
      Begin VB.CheckBox chkCrearAnexo 
         Alignment       =   1  'Right Justify
         Caption         =   "Crear anexo al generar"
         Height          =   255
         Left            =   -68820
         TabIndex        =   43
         Top             =   1875
         Width           =   2145
      End
      Begin VB.CheckBox chkConsiderarAdelantoDetallado 
         Alignment       =   1  'Right Justify
         Caption         =   "Considerar Adelanto detallado"
         Height          =   330
         Left            =   -68820
         TabIndex        =   42
         Top             =   2280
         Width           =   2145
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Considerar Adelanto detallado"
         Height          =   330
         Left            =   6135
         TabIndex        =   41
         Top             =   1995
         Width           =   2145
      End
      Begin VB.Frame Frame1 
         Height          =   4125
         Left            =   -74790
         TabIndex        =   21
         Top             =   450
         Width           =   8010
         Begin Crystal.CrystalReport CRREPCOMP 
            Left            =   3435
            Top             =   1935
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            PrintFileLinesPerPage=   60
         End
         Begin VB.TextBox XNUM 
            Height          =   300
            Left            =   1620
            TabIndex        =   40
            Text            =   "XNUM"
            Top             =   2235
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.CommandButton CMDIMPRIMIR 
            Caption         =   "&Imprimir"
            Height          =   330
            Left            =   1425
            TabIndex        =   39
            Top             =   3690
            Width           =   1215
         End
         Begin VB.CommandButton CmdAnt 
            Caption         =   "<< &Regresar"
            Height          =   330
            Left            =   165
            TabIndex        =   36
            Top             =   3690
            Width           =   1215
         End
         Begin MSDataGridLib.DataGrid XDGdetAsi 
            Height          =   2580
            Left            =   165
            TabIndex        =   32
            Top             =   1035
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   4551
            _Version        =   393216
            AllowUpdate     =   0   'False
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
            Caption         =   "Detalle del Asiento"
            ColumnCount     =   6
            BeginProperty Column00 
               DataField       =   "DMOV_SECUE"
               Caption         =   "Secue"
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
               DataField       =   "DMOV_CUENT"
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
               DataField       =   "DMOV_ANEXO"
               Caption         =   "Anexo"
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
               DataField       =   "DMOV_CENCO"
               Caption         =   "C.Costo"
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
               DataField       =   "DMOV_DEBE"
               Caption         =   "DEBE"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "###,###,###.00 "
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "DMOV_HABER"
               Caption         =   "HABER"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "###,###,###.00 "
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
                  ColumnWidth     =   689.953
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1110.047
               EndProperty
               BeginProperty Column02 
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1124.787
               EndProperty
               BeginProperty Column04 
                  Alignment       =   1
                  ColumnWidth     =   1305.071
               EndProperty
               BeginProperty Column05 
                  Alignment       =   1
                  ColumnWidth     =   1319.811
               EndProperty
            EndProperty
         End
         Begin VB.Label LTOTHABER 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00 "
            Height          =   285
            Left            =   6270
            TabIndex        =   35
            Top             =   3765
            Width           =   1350
         End
         Begin VB.Label LTOTDEBE 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00 "
            Height          =   285
            Left            =   4920
            TabIndex        =   34
            Top             =   3765
            Width           =   1350
         End
         Begin VB.Label Label15 
            Caption         =   "Total"
            Height          =   225
            Left            =   4005
            TabIndex        =   33
            Top             =   3810
            Width           =   510
         End
         Begin VB.Label Lfecha 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6585
            TabIndex        =   31
            Top             =   210
            Width           =   1275
         End
         Begin VB.Label Label13 
            Caption         =   "fecha"
            Height          =   210
            Left            =   5580
            TabIndex        =   30
            Top             =   255
            Width           =   690
         End
         Begin VB.Label Lglosa 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   3660
            TabIndex        =   29
            Top             =   570
            Width           =   4200
         End
         Begin VB.Label Label9 
            Caption         =   "Glosa"
            Height          =   225
            Left            =   2625
            TabIndex        =   28
            Top             =   615
            Width           =   645
         End
         Begin VB.Label Label12 
            Caption         =   "T/C. Venta"
            Height          =   210
            Left            =   2625
            TabIndex        =   27
            Top             =   255
            Width           =   915
         End
         Begin VB.Label LTipCam 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.000 "
            Height          =   300
            Left            =   3645
            TabIndex        =   26
            Top             =   210
            Width           =   1275
         End
         Begin VB.Label LComp 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   900
            TabIndex        =   25
            Top             =   570
            Width           =   1155
         End
         Begin VB.Label lsub 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   900
            TabIndex        =   24
            Top             =   210
            Width           =   705
         End
         Begin VB.Label Label6 
            Caption         =   "Compr."
            Height          =   315
            Left            =   150
            TabIndex        =   23
            Top             =   645
            Width           =   570
         End
         Begin VB.Label Label5 
            Caption         =   "Subdiario"
            Height          =   300
            Left            =   150
            TabIndex        =   22
            Top             =   255
            Width           =   675
         End
      End
      Begin VB.CheckBox ChkGenAnx 
         Alignment       =   1  'Right Justify
         Caption         =   "Crear anexo al generar"
         Height          =   315
         Left            =   6135
         TabIndex        =   20
         Top             =   1710
         Width           =   2145
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&E&xaminar Comprobante"
         Height          =   375
         Left            =   3630
         TabIndex        =   19
         Top             =   4335
         Width           =   2025
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   2280
         TabIndex        =   18
         Top             =   4335
         Width           =   1275
      End
      Begin VB.CommandButton CmdEnviar 
         Caption         =   "&Enviar a Contabilidad"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   4335
         Width           =   2085
      End
      Begin MSDataGridLib.DataGrid xDgResult 
         Height          =   1380
         Left            =   120
         TabIndex        =   16
         Top             =   2820
         Width           =   8220
         _ExtentX        =   14499
         _ExtentY        =   2434
         _Version        =   393216
         AllowUpdate     =   0   'False
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
         Caption         =   "Comprobante Generado"
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "SUBDIAR_CODIGO"
            Caption         =   "Sub"
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
            DataField       =   "CMOV_C_COMPR"
            Caption         =   "Comprobante"
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
            DataField       =   "CMOV_FECHA"
            Caption         =   "Fecha"
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
            DataField       =   "CMOV_MONED"
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
         BeginProperty Column04 
            DataField       =   "CMOV_DEBE"
            Caption         =   "Debe"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,###.00 "
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "CMOV_HABER"
            Caption         =   "Haber"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,###.00 "
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
               ColumnWidth     =   645.165
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1260.284
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1995.024
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1709.858
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton Command1 
         Caption         =   "E&xportar Trabajadores a Contabilidad"
         Enabled         =   0   'False
         Height          =   540
         Left            =   120
         TabIndex        =   11
         Top             =   2190
         Width           =   2265
      End
      Begin VB.CommandButton cmdGenerar 
         Caption         =   "Generar>>"
         Height          =   375
         Left            =   7050
         TabIndex        =   2
         ToolTipText     =   "Avanzar al proceso de planillas"
         Top             =   2385
         Width           =   1275
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   7080
         TabIndex        =   1
         Top             =   4335
         Width           =   1275
      End
      Begin AplisetControlText.Aplitext xMes 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Seleccione el mes de trabajo para el proceso"
         Top             =   660
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   503
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker xFechaFin 
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Top             =   1620
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   61997057
         CurrentDate     =   36699
      End
      Begin MSComCtl2.DTPicker xFechaIni 
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   61997057
         CurrentDate     =   36699
      End
      Begin MSComctlLib.ListView Lista 
         Height          =   2115
         Left            =   2445
         TabIndex        =   6
         ToolTipText     =   "Seleccione el periodo de planilla"
         Top             =   645
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   3731
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Periodos en Cronogramas"
            Object.Width           =   5733
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "FechaIni"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "FechaFin"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView LstvwCronograma 
         Height          =   1935
         Left            =   -72510
         TabIndex        =   45
         ToolTipText     =   "Seleccione el periodo de planilla"
         Top             =   735
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   3413
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Periodos en Cronogramas"
            Object.Width           =   5733
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "FechaIni"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "FechaFin"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dgComprobante 
         Height          =   1155
         Left            =   -74865
         TabIndex        =   52
         Top             =   2745
         Width           =   8220
         _ExtentX        =   14499
         _ExtentY        =   2037
         _Version        =   393216
         AllowUpdate     =   0   'False
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
         Caption         =   "Comprobante Generado"
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "SUBDIAR_CODIGO"
            Caption         =   "Sub"
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
            DataField       =   "CMOV_C_COMPR"
            Caption         =   "Comprobante"
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
            DataField       =   "CMOV_FECHA"
            Caption         =   "Fecha"
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
            DataField       =   "CMOV_MONED"
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
         BeginProperty Column04 
            DataField       =   "CMOV_DEBE"
            Caption         =   "Debe"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,###.00 "
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "CMOV_HABER"
            Caption         =   "Haber"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,###.00 "
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
               ColumnWidth     =   645.165
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1260.284
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1995.024
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1709.858
            EndProperty
         EndProperty
      End
      Begin AplisetControlText.Aplitext AplitxtMesTrabajo 
         Height          =   285
         Left            =   -74835
         TabIndex        =   44
         ToolTipText     =   "Seleccione el mes de trabajo para el proceso"
         Top             =   750
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   503
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   195
         Left            =   -74865
         TabIndex        =   77
         Top             =   4170
         Visible         =   0   'False
         Width           =   8130
         _ExtentX        =   14340
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Generar>>"
         Height          =   375
         Left            =   -74520
         TabIndex        =   54
         ToolTipText     =   "Avanzar al proceso de planillas"
         Top             =   2115
         Width           =   1275
      End
      Begin VB.Label lblMessage2 
         Caption         =   "Generando Temporal "
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   -74880
         TabIndex        =   80
         Top             =   4380
         Visible         =   0   'False
         Width           =   4740
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Mes de Trabajo"
         Height          =   195
         Left            =   -74820
         TabIndex        =   51
         Top             =   465
         Width           =   1110
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Periodos en Cronograma"
         Height          =   195
         Left            =   -72495
         TabIndex        =   50
         Top             =   480
         Width           =   1740
      End
      Begin VB.Label Label11 
         Caption         =   "Total de Planilla"
         Height          =   210
         Left            =   -68790
         TabIndex        =   49
         Top             =   510
         Width           =   1290
      End
      Begin VB.Label Label10 
         Caption         =   "Nº de Trabajadores"
         Height          =   255
         Left            =   -68790
         TabIndex        =   48
         Top             =   1185
         Width           =   1500
      End
      Begin VB.Label lblTotalPlanilla 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -68790
         TabIndex        =   47
         Top             =   780
         Width           =   2100
      End
      Begin VB.Label lblNroTrabajadores 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   -68790
         TabIndex        =   46
         Top             =   1455
         Width           =   2100
      End
      Begin VB.Label Xntrab 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   6165
         TabIndex        =   15
         Top             =   1365
         Width           =   2100
      End
      Begin VB.Label xTotPlan 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6165
         TabIndex        =   14
         Top             =   690
         Width           =   2100
      End
      Begin VB.Label Label4 
         Caption         =   "Nº de Trabajadores"
         Height          =   300
         Left            =   6165
         TabIndex        =   13
         Top             =   1095
         Width           =   1500
      End
      Begin VB.Label Label3 
         Caption         =   "Total de Planilla"
         Height          =   210
         Left            =   6165
         TabIndex        =   12
         Top             =   420
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Periodos en Cronograma"
         Height          =   195
         Left            =   2460
         TabIndex        =   10
         Top             =   390
         Width           =   1740
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Mes de Trabajo"
         Height          =   195
         Left            =   135
         TabIndex        =   9
         Top             =   375
         Width           =   1110
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1245
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label l2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label lblMessage 
         Caption         =   "Generando Temporal "
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   -74850
         TabIndex        =   78
         Top             =   3960
         Visible         =   0   'False
         Width           =   4740
      End
   End
   Begin VB.Label lbbar 
      Caption         =   "Generando Temporal "
      ForeColor       =   &H000000C0&
      Height          =   270
      Left            =   75
      TabIndex        =   38
      Top             =   4950
      Width           =   4740
   End
End
Attribute VB_Name = "frmGenAsientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim XITEM As ListItem, CADIN As String
Dim TOTALNETO As Double, TOTALING As Double, TOTALEGR As Double
Dim TOTALAPOR As Double
Dim NTRAB As Long
Dim PERIODO As String


'VARIABLES PARA LOS ASIENTOS
Dim RSCONTCAB As ADODB.Recordset
Dim RSCONTDET As ADODB.Recordset
'XXXXXXXXXXXXXXXXXBASILIO
Dim RSCONTCABQUIN As ADODB.Recordset
Dim RSCONTDETQUIN As ADODB.Recordset
'XXXXXXXXXXXXXXXXXXXXXX
Dim TIPCAM As Double
Dim DESCAM As String
Dim SECUE As String
Dim CTA As String, ANEX As String, CCosto As String
Dim CTADES As String
Dim MONTO As Double
Dim NUM As Integer
Dim VERIFI As Boolean
Dim RSBOLETAS As New ADODB.Recordset
Dim RSAUX As New ADODB.Recordset

Private Sub AplitxtMesTrabajo_DblClick()
    
    Dim RSMESES As New ADODB.Recordset
    RSMESES.Open "SELECT MESACTIVO, NOMBRE FROM MESESACT ORDER BY MESACTIVO", DBSYSTEM, adOpenStatic
    If RSMESES.RecordCount = 0 Then
        MsgBox "No se han encontrado meses en actividad", vbCritical
        Set RSMESES = Nothing
        Exit Sub
    End If
    frmComun.CONECTAR RSMESES
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        AplitxtMesTrabajo.Text = RSMESES!NOMBRE
        AplitxtMesTrabajo.Tag = RSMESES!MESACTIVO
        Set RSCONTCAB = Nothing
        Set xDgResult.DataSource = Nothing
        'xDgResult.ReBind
    Else
        Set RSMESES = Nothing
        Exit Sub
    End If
    
    Set RSMESES = Nothing
    Command1.Enabled = False
    'Reciclaje de RsMeses
    CARGAMESESQUINCENA

End Sub

Private Sub CmdAntQuin_Click()
    SSTab1.TabVisible(3) = False
    SSTab1.TabVisible(2) = True
    SSTab1.Tab = 2
End Sub

Private Sub CmdEnviarQuin_Click()
On Error GoTo handler
    Dim ULTNUMERO As String
    Dim CNXAUX As ADODB.Connection
    Set CNXAUX = New ADODB.Connection
    xFechaFin.Value = AplitxtMesTrabajo.Tag
    If Not VERIFI_CONTA(1, Year(xFechaFin)) Then
        Call MsgBox("El año de Ejercicio para esta planilla no esta aperturada en contabilidad", vbExclamation)
        Exit Sub
    End If
    
    Set CNXAUX = CONECTARDBSQL(REGSISTEMA.scRutaEmpresaWenco & "BDCONT" & Format(Year(xFechaFin), "0000"))
    If RSCONTCABQUIN.RecordCount = 0 Then
        Call MsgBox("No Hay Ningun Registo Para Enviar A Contabilidad", vbExclamation)
        Exit Sub
    End If
    If Not (RSCONTCABQUIN.EOF Or RSCONTCABQUIN.BOF) Then
        If Trim(RSCONTCABQUIN("CMOV_C_COMPR")) <> "" Then
            Call MsgBox("El asiento ha sido enviado a contabilidad " & Chr(13) & _
                         "si desea volverlo a generar elimine el Asiento de planillas " & Chr(13), vbInformation)
            Exit Sub
        End If
    End If
    Screen.MousePointer = 11
    ULTNUMERO = DevuelveValor("SELECT MAX(CMOV_C_COMPR) AS  NUMMAXIMO  " & _
                "FROM CABMOV" & Format(Month(xFechaFin), "00") & " WHERE SUBDIAR_CODIGO='" & Trim(REGSISTEMA.scSubdi) & "'", CNXAUX)
    ULTNUMERO = Format(Valc(ULTNUMERO) + 1, "0000")
    'ENVIANDO LA CABECERA DEL COMPROBANTE STANDAR DE CONTABILIDAD
    Dim RUTAPLAN As String
    RUTAPLAN = "[" & REGSISTEMA.BASESQL & "].dbo."
    CNXAUX.Execute "INSERT INTO CABMOV" & Format(Month(xFechaFin), "00") & _
                   "(SUBDIAR_CODIGO,CMOV_C_COMPR,CMOV_FECHA,CMOV_GLOSA,CMOV_MONED,CMOV_CONVE,CMOV_CAMES,CMOV_FECCA,CMOV_TIPCA,CMOV_DEBE," & _
                   " CMOV_HABER,CMOV_DEBUS,CMOV_HABUS,CMOV_AUTOM,CMOV_COSTO,CMOV_CHEQU,CMOV_L_COMPR,CMOV_VENTA) " & _
                   "SELECT SUBDIAR_CODIGO,'" & ULTNUMERO & "',CMOV_FECHA,CMOV_GLOSA,CMOV_MONED,CMOV_CONVE,CMOV_CAMES,CMOV_FECCA,CMOV_TIPCA,CMOV_DEBE," & _
                   "CMOV_HABER,CMOV_DEBUS,CMOV_HABUS,CMOV_AUTOM,CMOV_COSTO,CMOV_CHEQU,CMOV_L_COMPR,CMOV_VENTA FROM " & RUTAPLAN & "CONTQUICAB WHERE CRONO=" & LstvwCronograma.SelectedItem.Tag
                   
    'INSERTANDO EL DETALLE DEL COMPROBANTE DE CONTABILIDAD
    CNXAUX.Execute "INSERT INTO DETMOV" & Format(Month(xFechaFin), "00") & _
                   "(SUBDIAR_CODIGO, DMOV_C_COMPR, DMOV_SECUE, DMOV_FECHA, DMOV_CUENT, DMOV_ANEXO, DMOV_DOCUM, DMOV_FECDC, DMOV_CENCO, DMOV_DEBE, DMOV_HABER, DMOV_DEBUS, DMOV_HABUS, DMOV_GLOSA, DMOV_CHEQU, DMOV_AUTOM, DMOV_COSTO, DMOV_L_COMPR, DMOV_VENTA, DMOV_TRANS, DMOV_L_DESTI, DMOV_C_DESTI,PROVI) " & _
                   "SELECT SUBDIAR_CODIGO,'" & ULTNUMERO & "', DMOV_SECUE, DMOV_FECHA, DMOV_CUENT, DMOV_ANEXO, DMOV_DOCUM, DMOV_FECDC, CASE WHEN LEN(DMOV_CENCO)=0 THEN ' ' ELSE DMOV_CENCO END, DMOV_DEBE, DMOV_HABER, DMOV_DEBUS, DMOV_HABUS, DMOV_GLOSA, DMOV_CHEQU, DMOV_AUTOM, DMOV_COSTO, DMOV_L_COMPR, DMOV_VENTA, DMOV_TRANS, DMOV_L_DESTI, DMOV_C_DESTI,0 FROM " & RUTAPLAN & "CONTQUIDET WHERE CRONO=" & LstvwCronograma.SelectedItem.Tag & "    AND NUM=" & RSCONTCABQUIN.Fields("NUM")
    'ACTUALIZANDO LOS NUMEROS DE COMPROBANTE
    DBSYSTEM.Execute "UPDATE CONTQUICAB SET CMOV_C_COMPR='" & ULTNUMERO & "' WHERE CRONO=" & LstvwCronograma.SelectedItem.Tag
    DBSYSTEM.Execute "UPDATE CONTQUIDET SET DMOV_C_COMPR='" & ULTNUMERO & "' WHERE CRONO=" & LstvwCronograma.SelectedItem.Tag & "  AND NUM=" & RSCONTCABQUIN.Fields("NUM")
    RSCONTCABQUIN.Requery
     Screen.MousePointer = 1
     
Exit Sub
handler:
MsgBox ERR.Description, vbCritical, "Error de USuario"
End Sub

Private Sub CmdExa_Click()
On Error GoTo handler
 If RSCONTCABQUIN.RecordCount = 0 Then Exit Sub
    SSTab1.TabVisible(3) = True
    SSTab1.TabVisible(2) = False
    SSTab1.Tab = 3
    lsubquin.Caption = RSCONTCABQUIN("SUBDIAR_CODIGO")
    lcompquin.Caption = ESNULO(RSCONTCABQUIN("CMOV_C_COMPR"), "")
    If UCase(RSCONTCABQUIN("CMOV_CONVE")) = "VTA" Then 'NO SQL
        If RSCONTCABQUIN("CMOV_TIPCA") = 0 Then
            ltipcamquin.Caption = "0.00 "
          Else
            ltipcamquin.Caption = Format(Round(1 / RSCONTCABQUIN("CMOV_TIPCA"), 3), "0.000 ")
        End If
    End If
    If UCase(RSCONTCABQUIN("CMOV_CONVE")) = "ESP" Then
        ltipcamquin.Caption = Format(Round(RSCONTCABQUIN("CMOV_CAMES"), 3), "0.000 ")
    End If
    XNUMQuin.Text = RSCONTCABQUIN("NUM")
    lglosaquin.Caption = RSCONTCABQUIN("CMOV_GLOSA")
    lfechaquin = Format(RSCONTCABQUIN("CMOV_FECHA"), "dd/mm/yyyy")
    ldebequin.Caption = Format(RSCONTCABQUIN("CMOV_DEBE"), "###,###,###.00 ")
    lhaberquin.Caption = Format(RSCONTCABQUIN("CMOV_HABER"), "###,###,###.00 ")
    
    Set RSCONTDETQUIN = New ADODB.Recordset
    RSCONTDETQUIN.Open "SELECT * FROM CONTQUIDET WHERE NUM =" & XNUMQuin.Text & " AND CRONO=" & LstvwCronograma.SelectedItem.Tag & "  order by  DMOV_SECUE ", DBSYSTEM, adOpenDynamic, adLockOptimistic
    
    Set xDGDetaQuin.DataSource = RSCONTDETQUIN
    
Exit Sub
handler:
MsgBox ERR.Description, vbCritical, "Error de Usuario"
End Sub

Private Sub CMDIMPRIMIRQuin_Click()
Dim SqlCad As String
Dim RUTA As String
    Screen.MousePointer = 11
If ExisteTablaAux("[##TMPCOMPRO" & VGL_COMPUTER & "]") Then DBSYSTEM.Execute "DROP TABLE [##TMPCOMPRO" & VGL_COMPUTER & "]"
    SqlCad = "SELECT CONTQUIDET.* INTO [##TMPCOMPRO" & VGL_COMPUTER & "] " & _
             "FROM CONTQUICAB INNER JOIN CONTQUIDET ON CONTQUICAB.NUM = CONTQUIDET.NUM " & _
             "WHERE CONTQUICAB.NUM=" & XNUMQuin.Text
    DBSYSTEM.Execute SqlCad
    With CRREPCOMP
        .Reset
        .Connect = "DSN=" & VGL_SERVER & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL & ""
        .WindowTitle = "PLAN0092 - ASIENTO DE PLANILLAS"
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0092.RPT"
        .StoredProcParam(0) = "[##TMPCOMPRO" & VGL_COMPUTER & "]"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "XRUC='" & REGSISTEMA.RUC & "'"
        If .Status <> 2 Then .Action = 1
    End With
    Screen.MousePointer = 1
End Sub

Private Sub CmdSalirQuin_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Dim RSAUX As New ADODB.Recordset
    Dim BRES As Boolean
    Dim CON As Integer
    Dim PERIODOxx As String
    PERIODOxx = Format(Month(AplitxtMesTrabajo.Tag), "00") & Format(Year(AplitxtMesTrabajo.Tag), "00")
    Screen.MousePointer = 11
    RSAUX.Open "SELECT DISTINCT CODTRAB FROM BOL" & PERIODOxx, DBSYSTEM, adOpenStatic, adLockReadOnly
    Do While Not RSAUX.EOF
        BRES = GENERARANEXO(RSAUX("CODTRAB"))
        If BRES Then CON = CON + 1
        RSAUX.MoveNext
    Loop
    MsgBox "Se Migraron :" & CON & " Trabajadores a Contabilidad", vbInformation
    Screen.MousePointer = 1


End Sub
Private Sub DGLISTA_KeyDown(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 13 Then SendKeys "{TAB}"
    
End Sub

Private Sub GENTABLATEMP_ASIENTOS(MES As Integer, ANNO As Integer, NOMBOL As Integer)
    Dim RSTMPASI As New ADODB.Recordset
    Dim SCAD As String, SCADAUX As String, SCW As String
    Dim RST As New ADODB.Recordset
            
    If ExisteTablaAux(" [##TEMPASIENTOS" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TEMPASIENTOS" & VGL_COMPUTER & "] "
    'TEMPORAL PARA ALMACENAR EL POSIBLE
    DBSYSTEM.Execute _
    "CREATE TABLE  [##TEMPASIENTOS" & VGL_COMPUTER & "] (CODTRAB VARCHAR(8),CONCEPTO VARCHAR(25),CUENTA VARCHAR(25),UBIC VARCHAR(1),MONTO FLOAT,TIPASI VARCHAR(1),ANEXO VARCHAR(25),CCOSTO VARCHAR(15),CTADEST VARCHAR(25))"
    RSTMPASI.Open " [##TEMPASIENTOS" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic
    
   If chkConsiderarAdelantoDetallado.Value = 0 Then
    SCAD = "SELECT BOL.CODTRAB,MOV.CONCEPTO,MOV.MONTO FROM BOL" & PERIODO & " BOL " & _
         "INNER JOIN MOV" & PERIODO & " MOV " & _
         "ON BOL.INUMBOL = MOV.INUMBOL WHERE BOL.CODNOMBOL=" & NOMBOL & _
         " Union All " & _
         "SELECT CODTRAB,CONCEPTO='XXADELX',MONTO=SUM(MONTO) FROM ADEL2000  WHERE  ORIGEN = " & NOMBOL & _
         " GROUP BY CODTRAB "
   Else
    SCAD = "SELECT BOL.CODTRAB,MOV.CONCEPTO,MOV.MONTO FROM BOL" & PERIODO & " BOL " & _
         "INNER JOIN MOV" & PERIODO & " MOV " & _
         "ON BOL.INUMBOL = MOV.INUMBOL WHERE BOL.CODNOMBOL=" & NOMBOL & _
         " Union All " & _
         "SELECT CODTRAB,CONCEPTO='XXADELX',MONTO=SUM(MONTO) FROM DETADEL  WHERE IE=1 AND NOMBOL = " & NOMBOL & _
         " GROUP BY CODTRAB "
         
   End If
         
    If DevuelveValor("SELECT CODIGO FROM  CONCEPTOS WHERE CODIGO='XXPAGCXI'", DBSYSTEM) <> 0 Then
         SCW = "1"
       Else: SCW = "9"
    End If
    If DevuelveValor("SELECT CODIGO FROM  CONCEPTOS  WHERE CODIGO='XXPAGCXE'", DBSYSTEM) <> 0 Then
         If SCW <> "" Then
             SCW = SCW & ",2"
             Else: SCW = "2"
         End If
       Else
       'Cuentas Corrientes Programadas de Quincena
          Set RST = New ADODB.Recordset
          RST.Open "SELECT * FROM CONCEPTOS WHERE CODIGO LIKE 'XB%'", DBSYSTEM, adOpenKeyset, adLockReadOnly
          If RST.RecordCount = 0 Then
            SCW = "9"
            Else: SCW = "8"
          End If
    End If
    Select Case SCW
        Case "9"
            SCADAUX = "UNION ALL " & _
         "SELECT PAG.CODTRAB,'XX'+LTRIM(RTRIM(MOV.CODGRUPO)) AS CONCEPTO ," & _
         "SUM(PAG.MONTO) AS MONTO FROM PAGOSCTA PAG,MOVICTA MOV " & _
         "WHERE PAG.CODMOV = MOV.CODMOV AND PAG.CODNOMBOL =" & NOMBOL & _
         " GROUP BY PAG.CODTRAB,MOV.CODGRUPO "
        Case "8"
            'Si tiene Cuentas Corrientes Programadas
            SCADAUX = "UNION ALL " & _
         "SELECT PAG.CODTRAB,CONCEPTO=CASE PAG.TIPOBOLETA " & _
                           " WHEN 'B' THEN 'XB'+LTRIM(RTRIM(MOV.CODGRUPO)) " & _
                           "  WHEN 'A' THEN 'XA'+LTRIM(RTRIM(MOV.CODGRUPO))  END , " & _
         "SUM(PAG.MONTO) AS MONTO FROM PAGOSCTA PAG,MOVICTA MOV " & _
         "WHERE PAG.CODMOV = MOV.CODMOV AND PAG.CODNOMBOL =" & NOMBOL & _
         " GROUP BY PAG.CODTRAB,MOV.CODGRUPO,PAG.TIPOBOLETA"
        Case Else ' TODOS LOS PRESTAMOS EN UNA SOLA CUENTA
         SCADAUX = "UNION ALL " & _
         "SELECT CODTRAB,CASE TIPO WHEN 1 THEN 'XXPAGCXI' WHEN 2 THEN 'XXPAGCXE' END AS CONCEPTO " & _
         ",SUM(PAGOSCTA.MONTO) AS MONTO FROM PAGOSCTA " & _
         "WHERE CODNOMBOL =" & NOMBOL & "AND TIPO IN (" & SCW & ") and tipoboleta='B' GROUP BY CODTRAB,TIPO "
    End Select
    'AGUPAR AQUI POR TRABAJADOR----------
    
    
    Set RSBOLETAS = New ADODB.Recordset
    RSBOLETAS.Open SCAD & SCADAUX, DBSYSTEM, adOpenKeyset, adLockReadOnly
    
    pgbar.Scrolling = ccScrollingStandard
    pgbar.Min = 0: pgbar.Max = RSBOLETAS.RecordCount: pgbar.Value = 0
    Me.Height = 5760
    lbbar.ForeColor = &HC0&
    lbbar.Caption = "Creando temporal para generar los asientos ..."
'    VERIFI = VERIFI_CONTA(2)
    VERIFI = False
    
    Do While Not RSBOLETAS.EOF
        Me.Refresh
        pgbar = pgbar.Value + 1
        'LAS CUENTAS DE LOS CONCEPTOS QUE SE ENCUENTREN EN EL DEBE
        If ESNULO(DevuelveValor("SELECT TIPO FROM CONCEPTOS WHERE CODIGO='" & RSBOLETAS("CONCEPTO") & "'", DBSYSTEM), 0) > 0 Then
            Set RSAUX = New ADODB.Recordset
            Set RSAUX = REGDH(RSBOLETAS, "D")
            If RSAUX.RecordCount > 0 Then
                RSTMPASI.AddNew
                RSTMPASI("CODTRAB") = RSBOLETAS("CODTRAB")
                RSTMPASI("CONCEPTO") = RSBOLETAS("CONCEPTO")
                RSTMPASI("CUENTA") = RSAUX.Fields("CUENTA")
                RSTMPASI("UBIC") = "D"
                RSTMPASI("TIPASI") = RSAUX.Fields("TIPASI")
                RSTMPASI("MONTO") = Round(RSBOLETAS("MONTO"), 2)
                RSTMPASI("ANEXO") = Trim(REGSISTEMA.scTipoAnexo) & RSBOLETAS("CODTRAB")
                'NOTA
                'SIEMPRE Y CUANDO LA CUENTA MANEJE CENTRO DE COSTOS
                RSTMPASI("CCOSTO") = CENROCOSTOS 'FUNCION
                RSTMPASI("CTADEST") = ESNULO(GetValor("SELECT XXCTADES FROM TRABAJADORES WHERE CODTRAB='" & RSBOLETAS("CODTRAB") & "'", DBSYSTEM), "")
                RSTMPASI.Update
            End If
            Set RSAUX = New ADODB.Recordset
            Set RSAUX = REGDH(RSBOLETAS, "H")
            'LAS CUENTAS DE LOS CONCEPTOS QUE SE ENCUENTREN EN EL HABER
            If RSAUX.RecordCount > 0 Then
                RSTMPASI.AddNew
                RSTMPASI("CODTRAB") = RSBOLETAS("CODTRAB")
                RSTMPASI("CONCEPTO") = RSBOLETAS("CONCEPTO")
                RSTMPASI("CUENTA") = RSAUX.Fields("CUENTA")
                RSTMPASI("UBIC") = "H"
                RSTMPASI("TIPASI") = RSAUX.Fields("TIPASI")
                RSTMPASI("MONTO") = Round(RSBOLETAS("MONTO"), 2)
                RSTMPASI("ANEXO") = Trim(REGSISTEMA.scTipoAnexo) & RSBOLETAS("CODTRAB")
                'NOTA
                'SIEMPRE Y CUANDO LA CUENTA MANEJE CENTRO DE COSTOS
                RSTMPASI("CCOSTO") = CENROCOSTOS 'FUNCION
                RSTMPASI("CTADEST") = ESNULO(GetValor("SELECT XXCTADES FROM TRABAJADORES WHERE CODTRAB='" & RSBOLETAS("CODTRAB") & "'", DBSYSTEM), "")
                RSTMPASI.Update
            End If
        End If
        If ChkGenAnx.Value = 1 Then Call GENERARANEXO(RSBOLETAS("CODTRAB"))
        RSBOLETAS.MoveNext
    Loop
End Sub
Private Function GENERARANEXO(Codigo As String) As Boolean
    If Not REGSISTEMA.scTieneStConta Then Exit Function
    Dim Cnx As New ADODB.Connection
    GENERARANEXO = False
    Dim TIPANEX As String
    TIPANEX = REGSISTEMA.scTipoAnexo
    Set Cnx = CONECTARDBSQL(REGSISTEMA.scRutaEmpresaWenco & "BDCONTABILIDAD")
    If DevuelveValor("SELECT ANEX_CODIGO FROM ANEXO WHERE TIPOANEX_CODIGO='" & TIPANEX & "' AND ANEX_CODIGO='" & Codigo & "'", Cnx) = "" Then
        Cnx.Execute "INSERT INTO ANEXO(TIPOANEX_CODIGO,ANEX_CODIGO,ANEX_DESCRIPCION) VALUES " & _
        "('" & TIPANEX & "','" & Codigo & "','" & DevuelveValor("SELECT RTRIM(APEPAT)+ ' ' + RTRIM(APEMAT) + ', ' + RTRIM(NOMBRE) AS NOMBRES FROM TRABAJADORES WHERE CODTRAB='" & Trim(Codigo) & "'", DBSYSTEM) & "')"
        GENERARANEXO = True
    End If
End Function

Private Function REGDH(RS As ADODB.Recordset, TIPMOV As String) As ADODB.Recordset
    Set REGDH = New ADODB.Recordset
    REGDH.Open "SELECT * FROM CTACONCEPTO WHERE SEC=" & ESNULO(GetValor("SELECT TIPCTAX FROM TRABAJADORES WHERE CODTRAB='" & RS!CODTRAB & "'", DBSYSTEM), 1) & _
               " AND CONCEPT='" & Trim(RS!CONCEPTO) & "' AND TIPOCTA='" & Trim(TIPMOV) & "'", DBSYSTEM, adOpenKeyset, adLockReadOnly
End Function

Private Sub CmdAnt_Click()
    SSTab1.TabVisible(0) = True
    SSTab1.TabVisible(1) = False
End Sub

Private Sub CMDELIMINAR_CLICK()
On Error GoTo handler
    If RSCONTCAB.BOF Or RSCONTCAB.EOF Then Exit Sub
    If MsgBox("Esta seguro que desea eliminar el comprobante generado", vbQuestion + vbYesNo) = vbYes Then
        DBSYSTEM.Execute "DELETE FROM CONTCAB WHERE CRONO=" & Lista.SelectedItem.Tag
        DBSYSTEM.Execute "DELETE FROM CONTDET WHERE CRONO=" & Lista.SelectedItem.Tag
        If Trim(RSCONTCAB("CMOV_C_COMPR")) <> "" Then
            Dim CNXAUX As ADODB.Connection
            Set CNXAUX = New ADODB.Connection
            Set CNXAUX = CONECTARDBSQL(REGSISTEMA.scRutaEmpresaWenco & "BDCONT" & Format(Year(xFechaFin), "0000"))
            CNXAUX.Execute "DELETE FROM CABMOV" & Format(Month(xFechaFin), "00") & " WHERE SUBDIAR_CODIGO='" & Trim(REGSISTEMA.scSubdi) & "' AND CMOV_C_COMPR='" & RSCONTCAB("CMOV_C_COMPR") & "'"
            CNXAUX.Execute "DELETE FROM DETMOV" & Format(Month(xFechaFin), "00") & " WHERE SUBDIAR_CODIGO='" & Trim(REGSISTEMA.scSubdi) & "' AND DMOV_C_COMPR='" & RSCONTCAB("CMOV_C_COMPR") & "'"
        End If
         RSCONTCAB.Requery
    End If
Exit Sub
handler:
MsgBox "No existen registros", vbCritical, "Error de USuario"
End Sub

Private Sub CmdEnviar_Click()
On Error GoTo handler
    Dim ULTNUMERO As String
    Dim CNXAUX As ADODB.Connection
    Set CNXAUX = New ADODB.Connection
    
    If Not VERIFI_CONTA(1, Year(xFechaFin)) Then
        Call MsgBox("El año de Ejercicio para esta planilla no esta aperturada en contabilidad", vbExclamation)
        Exit Sub
    End If
    
    Set CNXAUX = CONECTARDBSQL(REGSISTEMA.scRutaEmpresaWenco & "BDCONT" & Format(Year(xFechaFin), "0000"))
    If RSCONTCAB.RecordCount = 0 Then
        Call MsgBox("No Hay Ningun Registo Para Enviar A Contabilidad", vbExclamation)
        Exit Sub
    End If
    If Not (RSCONTCAB.EOF Or RSCONTCAB.BOF) Then
        If Trim(RSCONTCAB("CMOV_C_COMPR")) <> "" Then
            Call MsgBox("El asiento ha sido enviado a contabilidad " & Chr(13) & _
                         "si desea volverlo a generar elimine el Asiento de planillas " & Chr(13), vbInformation)
            Exit Sub
        End If
    End If
    Screen.MousePointer = 11
    ULTNUMERO = DevuelveValor("SELECT MAX(CMOV_C_COMPR) AS  NUMMAXIMO  " & _
                "FROM CABMOV" & Format(Month(xFechaFin), "00") & " WHERE SUBDIAR_CODIGO='" & Trim(REGSISTEMA.scSubdi) & "'", CNXAUX)
    ULTNUMERO = Format(Valc(ULTNUMERO) + 1, "0000")
    'ENVIANDO LA CABECERA DEL COMPROBANTE STANDAR DE CONTABILIDAD
    Dim RUTAPLAN As String
    RUTAPLAN = "[" & REGSISTEMA.BASESQL & "].dbo."
    CNXAUX.Execute "INSERT INTO CABMOV" & Format(Month(xFechaFin), "00") & _
                   "(SUBDIAR_CODIGO,CMOV_C_COMPR,CMOV_FECHA,CMOV_GLOSA,CMOV_MONED,CMOV_CONVE,CMOV_CAMES,CMOV_FECCA,CMOV_TIPCA,CMOV_DEBE," & _
                   " CMOV_HABER,CMOV_DEBUS,CMOV_HABUS,CMOV_AUTOM,CMOV_COSTO,CMOV_CHEQU,CMOV_L_COMPR,CMOV_VENTA) " & _
                   "SELECT SUBDIAR_CODIGO,'" & ULTNUMERO & "',CMOV_FECHA,CMOV_GLOSA,CMOV_MONED,CMOV_CONVE,CMOV_CAMES,CMOV_FECCA,CMOV_TIPCA,CMOV_DEBE," & _
                   "CMOV_HABER,CMOV_DEBUS,CMOV_HABUS,CMOV_AUTOM,CMOV_COSTO,CMOV_CHEQU,CMOV_L_COMPR,CMOV_VENTA FROM " & RUTAPLAN & "CONTCAB WHERE CRONO=" & Lista.SelectedItem.Tag
                   
    'INSERTANDO EL DETALLE DEL COMPROBANTE DE CONTABILIDAD
    CNXAUX.Execute "INSERT INTO DETMOV" & Format(Month(xFechaFin), "00") & _
                   "(SUBDIAR_CODIGO, DMOV_C_COMPR, DMOV_SECUE, DMOV_FECHA, DMOV_CUENT, DMOV_ANEXO, DMOV_DOCUM, DMOV_FECDC, DMOV_CENCO, DMOV_DEBE, DMOV_HABER, DMOV_DEBUS, DMOV_HABUS, DMOV_GLOSA, DMOV_CHEQU, DMOV_AUTOM, DMOV_COSTO, DMOV_L_COMPR, DMOV_VENTA, DMOV_TRANS, DMOV_L_DESTI, DMOV_C_DESTI,PROVI) " & _
                   "SELECT SUBDIAR_CODIGO,'" & ULTNUMERO & "', DMOV_SECUE, DMOV_FECHA, DMOV_CUENT, DMOV_ANEXO, DMOV_DOCUM, DMOV_FECDC, CASE WHEN LEN(DMOV_CENCO)=0 THEN ' ' ELSE DMOV_CENCO END, DMOV_DEBE, DMOV_HABER, DMOV_DEBUS, DMOV_HABUS, DMOV_GLOSA, DMOV_CHEQU, DMOV_AUTOM, DMOV_COSTO, DMOV_L_COMPR, DMOV_VENTA, DMOV_TRANS, DMOV_L_DESTI, DMOV_C_DESTI,0 FROM " & RUTAPLAN & "CONTDET WHERE CRONO=" & Lista.SelectedItem.Tag
    'ACTUALIZANDO LOS NUMEROS DE COMPROBANTE
    DBSYSTEM.Execute "UPDATE CONTCAB SET CMOV_C_COMPR='" & ULTNUMERO & "' WHERE CRONO=" & Lista.SelectedItem.Tag
    DBSYSTEM.Execute "UPDATE CONTDET SET DMOV_C_COMPR='" & ULTNUMERO & "' WHERE CRONO=" & Lista.SelectedItem.Tag
    RSCONTCAB.Requery
    Screen.MousePointer = 1
    
Exit Sub
handler:
DisplayarError ERR, True
MsgBox "No existen registros", vbCritical, "Error de usuario"
End Sub

Private Sub cmdGenerar_Click()
'On Error GoTo cmd_generar
'    If Val(DevuelveValor("SELECT COUNT(*) FROM CTACONCEPTO ", DBSYSTEM)) = 0 Then
'        MsgBox "No se ha configurado cuentas para ninguno de los conceptos ", vbExclamation
'        Exit Sub
'    End If
'    If Not xFechaIni.Visible Then
'        MsgBox "Tiene que seleccionar un cronograma", vbExclamation
'        Lista.SetFocus
'        Exit Sub
'    End If
'    If Not (RSCONTCAB.EOF Or RSCONTCAB.BOF) Then
'        If Trim(RSCONTCAB("CMOV_C_COMPR")) <> "" Then
'            Call MsgBox("El asiento ha sido enviado a contabilidad " & Chr(13) & _
'                         "si desea volverlo a generar elimine el Asiento de planillas " & Chr(13), vbInformation)
'            Exit Sub
'        End If
'    End If
'    'BUSCAR SI YA SE GENERO UN COMPROBANTE EN CONTABILIDAD
'    If DevuelveValor("SELECT CRONO FROM CONTCAB WHERE CRONO=" & Lista.SelectedItem.Tag, DBSYSTEM) <> "" Then
'        If MsgBox("El Comprobante ya ha sido generado " & Chr(13) & "Desea volver a generarlo", vbQuestion + vbOKCancel) = vbOK Then
'            Screen.MousePointer = 11
'            DBSYSTEM.Execute "DELETE FROM CONTCAB WHERE CRONO=" & Lista.SelectedItem.Tag
'            DBSYSTEM.Execute "DELETE FROM CONTDET WHERE CRONO=" & Lista.SelectedItem.Tag
'          Else
'            Exit Sub
'        End If
'    End If
    '&H00C00000&
    Screen.MousePointer = 11
    Call GENTABLATEMP_ASIENTOS(Month(xFechaIni), Year(xFechaIni), Lista.SelectedItem.Tag)
    Call GENASIENTOTEMP
    Screen.MousePointer = 1
    Me.Height = 5325
cmd_generar:
 DisplayarError ERR
End Sub

Private Sub GENASIENTOTEMP()
    Dim RSCONCEPTOS As New ADODB.Recordset
    Dim RSAUX As New ADODB.Recordset
    Dim ORDEN As Integer
'    If VERIFI_CONTA(2) Then
'        TIPCAM = DevuelveValor("SELECT TIPOCAMB_EQCOMPRA FROM TIPO_CAMBIO " & _
'             "WHERE TIPOMON_CODIGO='ME' AND CAST(TIPOCAMB_FECHA AS INT)=" & FechS(xFechaFin.Value, Sqlf), CONECTARDBSQL(REGSISTEMA.scRutaEmpresaWenco & "BDCONTABILIDAD"))
'    End If
    TIPCAM = 0
    If TIPCAM > 0 Then
        DESCAM = "VTA"
     Else
        MsgBox "No se encontro tipo de cambio venta en contabilidad " & Chr(13) & _
               "se Utilizara el tipo de cambio puesto en la barra de planillas", vbInformation
        TIPCAM = 1 / Valc(MDIPrincipal.BarraEstado.Panels(3).Text)
        DESCAM = "ESP"
    End If
        
    'GRABA LA CABECERA DEL ASIENTO DE PLANILLA
    Call GRABACAB(RSCONTCAB)
    NUM = DevuelveValor("SELECT MAX(NUM) FROM CONTCAB", DBSYSTEM)
    
    'GRABA EL REGISTRO DEL NETO DE PLANILLA
    CTA = REGSISTEMA.scCuenta
    MONTO = Valc(xTotPlan)
    SECUE = 1
    CCosto = "": ANEX = "": CTADES = ""
    Call GRABADETASI(RSCONTDET, "H")
        
    RSCONCEPTOS.Open "SELECT DISTINCT CONCEPTO, TIPASI,UBIC " & _
                     "FROM  [##TEMPASIENTOS" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockReadOnly
    pgbar.Min = 0: pgbar.Max = RSCONCEPTOS.RecordCount: pgbar.Value = 0
    Me.Refresh
    lbbar.ForeColor = &HC00000
    lbbar.Caption = "Generando el Asiento Contable ..."
    pgbar.Scrolling = ccScrollingSmooth
    Do While Not RSCONCEPTOS.EOF
        Set RSAUX = New ADODB.Recordset
        pgbar.Value = pgbar.Value + 1
        Select Case RSCONCEPTOS!TIPASI
             Case 1 'SIMPLE (AGRUPAR CUENTA)
                 ORDEN = 1
                 SECUE = SECUE + 1
                 RSAUX.Open "SELECT CUENTA, Sum(MONTO) AS TOTAL " & _
                            "From  [##TEMPASIENTOS" & VGL_COMPUTER & "]  WHERE CONCEPTO='" & RSCONCEPTOS("CONCEPTO") & "' AND UBIC='" & Trim(RSCONCEPTOS!UBIC) & "' GROUP BY CUENTA ", DBSYSTEM
                 CTA = RSAUX("CUENTA"): MONTO = RSAUX("TOTAL")
                 Call GRABADETASI(RSCONTDET, RSCONCEPTOS!UBIC)
             Case 2 'POR TRABAJADOR
                ORDEN = 2
                RSAUX.Open "SELECT CODTRAB, CONCEPTO, Sum(MONTO) AS TOTAL, ANEXO,CUENTA From  [##TEMPASIENTOS" & VGL_COMPUTER & "]  " & _
                           " WHERE CONCEPTO='" & Trim(RSCONCEPTOS("CONCEPTO")) & "' AND UBIC='" & Trim(RSCONCEPTOS!UBIC) & "' GROUP BY CODTRAB, CONCEPTO, ANEXO,CUENTA ", DBSYSTEM, adOpenStatic, adLockReadOnly
                Do While Not RSAUX.EOF
                    SECUE = SECUE + 1
                    CTA = RSAUX("CUENTA"): MONTO = RSAUX("TOTAL"): ANEX = RSAUX("ANEXO")
                    Call GRABADETASI(RSCONTDET, RSCONCEPTOS!UBIC)
                    RSAUX.MoveNext
                Loop
             Case 3 'POR CENTRO DE COSTOS
                ORDEN = 3
                RSAUX.Open "SELECT CCOSTO, CONCEPTO, SUM(MONTO) AS TOTAL, CUENTA,CTADEST FROM  [##TEMPASIENTOS" & VGL_COMPUTER & "]  " & _
                           " WHERE CONCEPTO='" & Trim(RSCONCEPTOS("CONCEPTO")) & "' AND UBIC='" & Trim(RSCONCEPTOS!UBIC) & "' GROUP BY CCOSTO, CONCEPTO, CUENTA,CTADEST ", DBSYSTEM, adOpenStatic, adLockReadOnly
                Do While Not RSAUX.EOF
                    SECUE = SECUE + 1
                    CTA = RSAUX("CUENTA"): MONTO = RSAUX("TOTAL"): CCosto = RSAUX("CCOSTO")
                    ANEX = "": CTADES = RSAUX("CTADEST")
                    Call GRABADETASI(RSCONTDET, RSCONCEPTOS!UBIC)
                    RSAUX.MoveNext
                Loop
             Case 4 'POR A.F.P.
             Case 5 'POR TRABAJADOR Y CENTRO DE COSTOS
        End Select
        RSCONCEPTOS.MoveNext
    Loop
    
    'SE ACTUALIZA EL MONTO TOTAL DEL COMPROBANTE
    Dim RSTOTCAB As ADODB.Recordset, TOTCABDEBE As Double, TOTCABHABER As Double
    Dim TOTCABDEBUS As Double, TOTCABHABUS As Double
    Dim REDON As Double, RETIP As String
    Set RSTOTCAB = New ADODB.Recordset
    RSTOTCAB.Open "SELECT SUM(DMOV_DEBE) AS TOTALDEBE,SUM(DMOV_HABER) AS TOTALHABER, " & _
    "SUM(DMOV_DEBUS) AS TOTALDEBUS,SUM(DMOV_HABUS) AS TOTALHABUS FROM CONTDET WHERE NUM=" & NUM & " AND CRONO=" & Lista.SelectedItem.Tag, DBSYSTEM, adOpenKeyset, adLockReadOnly
    TOTCABDEBE = Round(ESNULO(RSTOTCAB("TOTALDEBE"), 0), 2)
    TOTCABHABER = Round(ESNULO(RSTOTCAB("TOTALHABER"), 0), 2)
    TOTCABDEBUS = Round(ESNULO(RSTOTCAB("TOTALDEBUS"), 0), 2)
    TOTCABHABUS = Round(ESNULO(RSTOTCAB("TOTALHABUS"), 0), 2)
    REDON = TOTCABDEBE - TOTCABHABER
    'PARA LA CUENTA DE REDONDEO NO SQL
    If Abs(REDON) < 1 Then
        CTA = REGSISTEMA.scCtaRedon
        SECUE = SECUE + 1
        CCosto = "": ANEX = ""
        If REDON > 0 Then
            RETIP = "H"
            TOTCABHABER = TOTCABHABER + Abs(REDON)
            TOTCABHABUS = TOTCABHABUS + Round((Abs(REDON) * TIPCAM), 2)
           Else
            RETIP = "D"
            TOTCABDEBE = TOTCABDEBE + Abs(REDON)
            TOTCABDEBUS = TOTCABDEBUS + Round((Abs(REDON) * TIPCAM), 2)
        End If
        MONTO = Abs(REDON)
        Call GRABADETASI(RSCONTDET, RETIP)
    End If
    
    DBSYSTEM.Execute "UPDATE CONTCAB SET CMOV_DEBE=" & TOTCABDEBE & "," & _
                     "CMOV_HABER=" & TOTCABHABER & ",CMOV_DEBUS=" & TOTCABDEBUS & ",CMOV_HABUS=" & TOTCABHABUS & " WHERE CRONO=" & Lista.SelectedItem.Tag
    Call GENERASECUENCIA(ORDEN)
    RSCONTCAB.Requery
    RSCONTDET.Requery
    RSCONTCAB.Filter = "CRONO=" & Lista.SelectedItem.Tag
    MsgBox "El proceso concluyo satisfactoriamente", vbInformation
End Sub

Private Sub GENERASECUENCIA(ORDEN As Integer)
    Dim RSDETAUX As New ADODB.Recordset
    Dim CAD As String
    Dim CONT As Integer
    Set RSDETAUX = New ADODB.Recordset
    Select Case ORDEN
        Case 1: CAD = "SELECT * FROM CONTDET WHERE NUM=" & NUM & "  ORDER BY DMOV_CUENT "
        Case 2: CAD = "SELECT * FROM CONTDET WHERE NUM=" & NUM & " ORDER BY DMOV_ANEXO,DMOV_CUENT "
        Case 3: CAD = "SELECT * FROM CONTDET WHERE NUM=" & NUM & " ORDER BY DMOV_CENCO,DMOV_CUENT "
    End Select
    RSDETAUX.Open CAD, DBSYSTEM, adOpenKeyset, adLockOptimistic
    lbbar.ForeColor = &HC00000
    lbbar.Caption = "GENERANDO EL ORDEN DE LA SECUENCIA SEGUN EL TIPO DE ASIENTO"
    pgbar.Min = 0: pgbar.Max = RSDETAUX.RecordCount: pgbar.Value = 0
    pgbar.Scrolling = ccScrollingSmooth
    Me.Refresh
    CONT = 1
    Do While Not RSDETAUX.EOF
       pgbar.Value = pgbar.Value + 1
       RSDETAUX("DMOV_SECUE") = Format(CONT, "0000")
       CONT = CONT + 1
       RSDETAUX.Update
       RSDETAUX.MoveNext
    Loop
End Sub
Private Sub GRABACAB(RSCAB As ADODB.Recordset)
    'CREANDO LA CABECERA DEL ASIENTO DE PLANILLAS
    With RSCAB
        .AddNew
        .Fields("SUBDIAR_CODIGO") = Trim(REGSISTEMA.scSubdi)
        .Fields("CMOV_C_COMPR") = " "
        .Fields("CMOV_FECHA") = xFechaFin.Value
        .Fields("CMOV_GLOSA") = Trim(Mid("Planillas " & Lista.SelectedItem.Text, 1, 29))
        .Fields("CMOV_MONED") = "MN"
        .Fields("CMOV_CONVE") = DESCAM  'FORMA DE CAMBIO
        .Fields("CMOV_CAMES") = IIf(DESCAM = "ESP", 1 / TIPCAM, 0)
        .Fields("CMOV_FECCA") = xFechaFin.Value
        .Fields("CMOV_TIPCA") = IIf(DESCAM = "VTA", TIPCAM, 0)
        .Fields("CMOV_DEBE") = 0
        .Fields("CMOV_HABER") = 0
        .Fields("CMOV_DEBUS") = 0
        .Fields("CMOV_HABUS") = 0
        .Fields("CMOV_AUTOM") = 0
        .Fields("CMOV_COSTO") = 0
        .Fields("CMOV_CHEQU") = 0
        .Fields("CMOV_L_COMPR") = 0
        .Fields("CMOV_VENTA") = 0
        .Fields("CRONO") = Lista.SelectedItem.Tag
        .Update
    End With
End Sub
Private Sub GRABADETASI(RSDET As ADODB.Recordset, TIPO As String)
    'CREANDO EL DETALLE DEL ASIENTO DE PLANILLAS
    CCosto = IIf((CCosto = ""), "", CCosto)
    If CCosto <> "" Then CCosto = Mid(CCosto, 1, InStr(CCosto, ":") - 1)
    With RSDET
        .AddNew
        .Fields("NUM") = NUM
        .Fields("SUBDIAR_CODIGO") = REGSISTEMA.scSubdi
        .Fields("DMOV_C_COMPR") = " "
        .Fields("DMOV_SECUE") = Format(SECUE, "0000")
        .Fields("DMOV_FECHA") = FechS(xMes.Tag, Adof)
        .Fields("DMOV_CUENT") = ESNULO(CTA, " ")
        .Fields("DMOV_ANEXO") = ESNULO(ANEX, " ")
        .Fields("DMOV_DOCUM") = " "
        .Fields("DMOV_FECDC") = FechS(xMes.Tag, Adof)
        .Fields("DMOV_CENCO") = CCosto
        .Fields("DMOV_DEBE") = Round(IIf(TIPO = "D", MONTO, 0), 2)
        .Fields("DMOV_HABER") = Round(IIf(TIPO = "H", MONTO, 0), 2)
        .Fields("DMOV_DEBUS") = Round(IIf(TIPO = "D", MONTO * TIPCAM, 0), 2)
        .Fields("DMOV_HABUS") = Round(IIf(TIPO = "H", MONTO * TIPCAM, 0), 2)
        .Fields("DMOV_GLOSA") = " "
        .Fields("DMOV_CHEQU") = 0
        .Fields("DMOV_AUTOM") = 0
        .Fields("DMOV_COSTO") = 0
        .Fields("DMOV_L_COMPR") = 0
        .Fields("DMOV_VENTA") = 0
        .Fields("DMOV_TRANS") = 0
        .Fields("DMOV_L_DESTI") = 0
        .Fields("DMOV_C_DESTI") = CTADES
        .Fields("CRONO") = Lista.SelectedItem.Tag
        .Update
    End With
End Sub
Private Sub CmdImprimir_Click()
Dim SqlCad As String
Dim RUTA As String
    Screen.MousePointer = 11
    If ExisteTablaAux("[##TMPCOMPRO" & VGL_COMPUTER & "]") Then DBSYSTEM.Execute "DROP TABLE [##TMPCOMPRO" & VGL_COMPUTER & "]"
    SqlCad = "SELECT CONTDET.* INTO [##TMPCOMPRO" & VGL_COMPUTER & "] " & _
             "FROM CONTCAB INNER JOIN CONTDET ON CONTCAB.NUM = CONTDET.NUM " & _
             "WHERE CONTCAB.NUM=" & XNUM.Text
    DBSYSTEM.Execute SqlCad
    With CRREPCOMP
        .Reset
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
        .WindowTitle = "PLAN0092 - ASIENTO DE PLANILLAS"
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0092.RPT"
        .StoredProcParam(0) = "[##TMPCOMPRO" & VGL_COMPUTER & "]"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "XRUC='" & REGSISTEMA.RUC & "'"
        If .Status <> 2 Then .Action = 1
    End With
    Screen.MousePointer = 1
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
Dim RSAUX As New ADODB.Recordset
    Dim BRES As Boolean
    Dim CON As Integer
    Screen.MousePointer = 11
    RSAUX.Open "SELECT DISTINCT CODTRAB FROM BOL" & PERIODO, DBSYSTEM, adOpenStatic, adLockReadOnly
    Do While Not RSAUX.EOF
        BRES = GENERARANEXO(RSAUX("CODTRAB"))
        If BRES Then CON = CON + 1
        RSAUX.MoveNext
    Loop
    MsgBox "Se Migraron :" & CON & " Trabajadores a Contabilidad", vbInformation
    Screen.MousePointer = 1
End Sub

Private Sub Command3_Click()
'CMD GENERAR ASIENTOS DE QUINCENA
On Error GoTo cmd_generarquin
xFechaFin.Value = AplitxtMesTrabajo.Tag
    If Val(DevuelveValor("SELECT COUNT(*) FROM CTACONCEPTO ", DBSYSTEM)) = 0 Then
        MsgBox "No se ha configurado cuentas para ninguno de los conceptos ", vbExclamation
        Exit Sub
    End If
    If AplitxtMesTrabajo.Text = "" Then
        MsgBox "Tiene que seleccionar un cronograma", vbExclamation
        Lista.SetFocus
        Exit Sub
    End If
    If Not (RSCONTCABQUIN.EOF Or RSCONTCABQUIN.BOF) Then
        If Trim(RSCONTCABQUIN("CMOV_C_COMPR")) <> "" Then
            Call MsgBox("El asiento ha sido enviado a contabilidad " & Chr(13) & _
                         "si desea volverlo a generar elimine el Asiento de planillas " & Chr(13), vbInformation)
            Exit Sub
        End If
    End If
    If DevuelveValor("SELECT CRONO FROM CONTQUICAB WHERE CRONO=" & LstvwCronograma.SelectedItem.Tag, DBSYSTEM) <> "" Then
        If MsgBox("Ya se genero el asiento desea volver a generarlo", vbYesNo) = vbYes Then
            DBSYSTEM.Execute "delete  from CONTQUICAB WHERE NUM=" & DevuelveValor("SELECT NUM from CONTQUICAB WHERE CRONO=" & LstvwCronograma.SelectedItem.Tag, DBSYSTEM) & "   AND  CRONO=" & LstvwCronograma.SelectedItem.Tag
            DBSYSTEM.Execute "delete  from CONTQUICAB WHERE CRONO=" & LstvwCronograma.SelectedItem.Tag
        Else
            Exit Sub
        End If
        
    End If
    'BUSCAR SI YA SE GENERO UN COMPROBANTE EN CONTABILIDAD
    If DevuelveValor("SELECT CRONO FROM CONTCABQUI WHERE CRONO=" & LstvwCronograma.SelectedItem.Tag, DBSYSTEM) <> "" Then
        If MsgBox("El Comprobante ya ha sido generado " & Chr(13) & "Desea volver a generarlo", vbQuestion + vbOKCancel) = vbOK Then
            Screen.MousePointer = 11
            DBSYSTEM.Execute "DELETE  FROM CONTCABQUI WHERE CRONO=" & LstvwCronograma.SelectedItem.Tag
            DBSYSTEM.Execute "DELETE  FROM CONTDETQUI WHERE CRONO=" & LstvwCronograma.SelectedItem.Tag
          Else
            Exit Sub
        End If
    End If
    '&H00C00000&
    CmdEnviarQuin.Visible = False
    CmdExa.Visible = False
    CmdSalirQuin.Visible = False
    Command6.Visible = False
    ProgressBar1.Visible = True
    ProgressBar2.Visible = True
    lblMessage.Visible = True
    lblMessage2.Visible = True
    ProgressBar1.Value = 0
    ProgressBar2.Value = 0
    Screen.MousePointer = 11
        Call GENTABLATEMP_ASIENTOS_QUINCENA(Month(LstvwCronograma.SelectedItem.Tag), Year(LstvwCronograma.SelectedItem.Tag), LstvwCronograma.SelectedItem.Tag)
    CmdExa.Visible = True
    CmdEnviarQuin.Visible = True
    CmdSalirQuin.Visible = True
    Command6.Visible = True
    ProgressBar1.Visible = False
    ProgressBar2.Visible = False
    lblMessage.Visible = False
    lblMessage2.Visible = False
    Screen.MousePointer = 1
    Me.Height = 5325
    
Exit Sub
cmd_generarquin:
    DisplayarError ERR
End Sub

Private Sub Command4_Click()
On Error GoTo handler
    If RSCONTCAB.RecordCount = 0 Then Exit Sub
    
    SSTab1.TabVisible(1) = True
    SSTab1.TabVisible(0) = False
    lsub.Caption = RSCONTCAB("SUBDIAR_CODIGO")
    LComp.Caption = ESNULO(RSCONTCAB("CMOV_C_COMPR"), "")
    If UCase(RSCONTCAB("CMOV_CONVE")) = "VTA" Then 'NO SQL
        If RSCONTCAB("CMOV_TIPCA") = 0 Then
            LTipCam.Caption = "0.00 "
          Else
            LTipCam.Caption = Format(Round(1 / RSCONTCAB("CMOV_TIPCA"), 3), "0.000 ")
        End If
    End If
    If UCase(RSCONTCAB("CMOV_CONVE")) = "ESP" Then
        LTipCam.Caption = Format(Round(RSCONTCAB("CMOV_CAMES"), 3), "0.000 ")
    End If
    XNUM.Text = RSCONTCAB("NUM")
    Lglosa.Caption = RSCONTCAB("CMOV_GLOSA")
    Lfecha = Format(RSCONTCAB("CMOV_FECHA"), "dd/mm/yyyy")
    LTOTDEBE.Caption = Format(RSCONTCAB("CMOV_DEBE"), "###,###,###.00 ")
    LTOTHABER.Caption = Format(RSCONTCAB("CMOV_HABER"), "###,###,###.00 ")
    RSCONTDET.Filter = "CRONO=" & Lista.SelectedItem.Tag & " AND NUM=" & RSCONTCAB("NUM")
Exit Sub
handler:
    MsgBox "No Existen Registros", vbCritical, "Error de Usuario"
    DisplayarError ERR, True
End Sub

Private Sub Command5_Click()

End Sub

Private Sub COMMAND6_Click()
On Error GoTo handler
    If RSCONTCABQUIN.BOF Or RSCONTCABQUIN.EOF Then Exit Sub
    If MsgBox("Esta seguro que desea eliminar el comprobante generado", vbQuestion + vbYesNo) = vbYes Then
        DBSYSTEM.Execute "DELETE FROM CONTQUICAB WHERE CRONO=" & LstvwCronograma.SelectedItem.Tag
        DBSYSTEM.Execute "DELETE FROM CONTQUIDET WHERE CRONO=" & LstvwCronograma.SelectedItem.Tag
        If Trim(RSCONTCABQUIN("CMOV_C_COMPR")) <> "" Then
            Dim CNXAUX As ADODB.Connection
            Set CNXAUX = New ADODB.Connection
            Set CNXAUX = CONECTARDBSQL(REGSISTEMA.scRutaEmpresaWenco & "BDCONT" & Format(Year(AplitxtMesTrabajo.Tag), "0000"))
            CNXAUX.Execute "DELETE FROM CABMOV" & Format(Month(AplitxtMesTrabajo.Tag), "00") & " WHERE SUBDIAR_CODIGO='" & Trim(REGSISTEMA.scSubdi) & "' AND CMOV_C_COMPR='" & RSCONTCABQUIN("CMOV_C_COMPR") & "'"
            CNXAUX.Execute "DELETE FROM DETMOV" & Format(Month(AplitxtMesTrabajo.Tag), "00") & " WHERE SUBDIAR_CODIGO='" & Trim(REGSISTEMA.scSubdi) & "' AND DMOV_C_COMPR='" & RSCONTCABQUIN("CMOV_C_COMPR") & "'"
        End If
         RSCONTCABQUIN.Requery
    End If
    MsgBox "El proceso ha concluido satisfactoriamente", vbInformation
Exit Sub
handler:
MsgBox ERR.Description, vbCritical, "Error de Usuario"
End Sub

Private Sub Form_Load()
    Me.TOP = 0
    Me.Left = 0
    SSTab1.TabVisible(1) = False
    SSTab1.TabVisible(3) = False
    Dim XFEC As String
    XFEC = "01/" & MDIPrincipal.BarraEstado.Panels("Periodo").Text
    If IsDate(XFEC) Then
        xMes.Text = DevuelveValor("SELECT Nombre FROM MesesAct WHERE MesActivo=" & DateSQL(XFEC), DBSYSTEM)
        xMes.Tag = CDate(XFEC)
        CARGAMESES
    End If
'    If Not VERIFI_CONTA(2) Then
        CmdEnviar.Enabled = False
        Command1.Enabled = False
'      Else
'        CmdEnviar.Enabled = True
'        Command1.Enabled = True
'    End If
    '---------------------BASILIO
    If ExisteTablaAux("[##_TMPBOLQUI" & VGL_COMPUTER & "]") Then DBSYSTEM.Execute "DROP TABLE [##_TMPBOLQUI" & VGL_COMPUTER & "]"
    DBSYSTEM.Execute "CREATE TABLE  [##_TMPBOLQUI" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8), NOMBRES VARCHAR(100),ADELANTO  Numeric(20,2) , INGRESOS  Numeric(20,2) , EGRESOS  Numeric(20,2) , NETO  Numeric(20,2) , INUMBOL INT, NOMBOL INT, PERIODO VARCHAR(50),SECUENCIA VARCHAR(1),CTADESTINO VARCHAR(25),CCOSTO VARCHAR(50))"
    
End Sub

Private Sub LABEL7_Click()

End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
    xFechaIni.Visible = True
    xFechaFin.Visible = True
    l1.Visible = True
    l2.Visible = True
    xFechaIni.Value = CDate(Item.SubItems(1))
    xFechaFin.Value = CDate(Item.SubItems(2))
    PERIODO = Trim(Format(Month(xFechaIni), "00")) & Trim(Format(Year(xFechaIni), "0000"))
    Call SUMAPLAN(TOTALING, TOTALEGR, TOTALNETO, Lista.SelectedItem.Tag)
    NTRAB = CUENTA(Lista.SelectedItem.Tag)
    xTotPlan.Caption = Format(TOTALNETO, "###,###,###.00 ")
    Xntrab.Caption = Format(NTRAB, "0 ")
     'APERTURANDO LAS TABLAS DE CONTABILIDAD PARA ARMAR EL ASIENTO DE PLANILLA
    Set RSCONTCAB = New ADODB.Recordset
    Set RSCONTDET = New ADODB.Recordset
    RSCONTCAB.Open "CONTCAB", DBSYSTEM, adOpenKeyset, adLockOptimistic
    RSCONTDET.Open "SELECT * FROM CONTDET ORDER BY DMOV_SECUE", DBSYSTEM, adOpenKeyset, adLockOptimistic
    RSCONTCAB.Filter = "CRONO=" & Lista.SelectedItem.Tag
    Command1.Enabled = True
    Set xDgResult.DataSource = RSCONTCAB
    Set XDGdetAsi.DataSource = RSCONTDET
End Sub
Private Sub SUMAPLAN(ByRef xTotIng As Double, ByRef xTotEgr As Double, ByRef xNeto As Double, Crono As Long)
Dim SqlCad As String
Dim RSAUX As New ADODB.Recordset
    If Not xFechaIni.Visible Then Exit Sub
    Set RSAUX = New ADODB.Recordset
    SqlCad = "SELECT SUM(TOTING) AS INGRESOS, SUM(TOTEGR) AS EGRESOS, SUM(TOTING-TOTEGR) AS NETO " & _
             "FROM  BOL" & PERIODO & " BOLS WHERE CODNOMBOL=" & Crono & ""
    RSAUX.Open SqlCad, DBSYSTEM, adOpenKeyset, adLockReadOnly
    If RSAUX.RecordCount > 0 Then
        xTotIng = ESNULO(RSAUX("INGRESOS"), 0)
        xTotEgr = ESNULO(RSAUX("EGRESOS"), 0)
        xNeto = ESNULO(RSAUX("NETO"), 0)
    End If
End Sub

Private Function CUENTA(Crono As Long) As Double
    Dim RSCUENTA As New ADODB.Recordset
    Set RSCUENTA = New ADODB.Recordset
    RSCUENTA.Open "SELECT COUNT(CODTRAB) AS CUENTA FROM BOL" & PERIODO & " WHERE CODNOMBOL=" & Crono, DBSYSTEM
    CUENTA = ESNULO(RSCUENTA!CUENTA, 0)
End Function

Private Sub LISTA_KeyDown(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 13 Then SendKeys "{TAB}"
End Sub
Private Sub LstvwCronograma_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim SNOMBOL As String
    Dim RSAUX As ADODB.Recordset
    Dim STRSQL As String
    SNOMBOL = Right(Item.KEY, Len(Item.KEY) - 1)
    Dim FMES As Date
    FMES = DevuelveValor("SELECT MES FROM NOMBOL WHERE CODIGO=" & SNOMBOL, DBSYSTEM)
    CambiaPanelBD True
    
'If ExisteTablaSQL(" [##_TMPLSTBOL" & VGL_COMPUTER & "] ", DBAUXCOM) Then DBSYSTEM.Execute "DROP TABLE  [##_TMPLSTBOL" & VGL_COMPUTER & "] "
    
'    If Item.Checked Then
'        Screen.MousePointer = 11
'        DBSYSTEM.Execute "UPDATE BOL" & Format(Month(FMES), "00") & Year(FMES) & " SET XREDONDEO=0 WHERE (XREDONDEO)IS NULL"
'        'If ExisteTablaSQL(" [##_TMPLSTBOL" & VGL_COMPUTER & "] ", DBAUXCOM) Then DBSYSTEM.Execute "DROP TABLE  [##_TMPLSTBOL" & VGL_COMPUTER & "] "
'        DBSYSTEM.Execute "INSERT INTO  [##_TMPBOLQUI" & VGL_COMPUTER & "]  SELECT TR.CODTRAB, NOMBRES, TOTING AS INGRESOS, TOTEGR AS EGRESOS, TOTING-TOTEGR+XREDONDEO AS NETO, INUMBOL, CODNOMBOL AS NOMBOL,'" & Item.Text & "' AS PERIODO  FROM " & REGSISTEMA.BASESQL & ".dbo.VWTRABAJ TR, " & REGSISTEMA.BASESQL & ".dbo.BOL" & Format(Month(FMES), "00") & Year(FMES) & " BOLS  WHERE BOLS.CODTRAB=TR.CODTRAB AND CODNOMBOL=" & SNOMBOL & ""
'        Screen.MousePointer = 1
'    Else
'        DBSYSTEM.Execute "DELETE FROM  [##_TMPBOLQUI" & VGL_COMPUTER & "]  WHERE NOMBOL=" & SNOMBOL
'    End If

    If Item.Checked Then
        Screen.MousePointer = 11
        STRSQL = "INSERT INTO  [##_TMPBOLQUI" & VGL_COMPUTER & "] " _
            & "  (CODTRAB,ADELANTO,NOMBOL,PERIODO,INUMBOL)" _
            & " SELECT  CODTRAB,MONTO,ORIGEN AS NOMBOL,'" & Item.Text & "' AS PERIODO,CODIGO" _
            & " FROM ADEL2000 " _
            & " WHERE ORIGEN=" & SNOMBOL
        DBSYSTEM.Execute STRSQL
        DBSYSTEM.Execute "UPDATE A SET A.NOMBRES = W.NOMBRES FROM  [##_TMPBOLQUI" & VGL_COMPUTER & "]  A, " & REGSISTEMA.BASESQL & ".dbo.VWTRABAJ W WHERE A.CODTRAB = W.CODTRAB "
        
        'Coge Ingresos y Egresos de Cuentas Corrientes
        Set RSAUX = New ADODB.Recordset
        'RSAUX.Open "SELECT CODTRAB, SUM(MONTO) AS TOT1 FROM PAGOSCTA WHERE TIPO=1 AND TIPOBOLETA='A' AND CODNOMBOL=" & SNOMBOL & " GROUP BY CODTRAB", DBSYSTEM, adOpenStatic, adLockReadOnly
        'Do While Not RSAUX.EOF
         '   DBSYSTEM.Execute "UPDATE  [##_TMPBOLQUI" & VGL_COMPUTER & "]  SET INGRESOS=" & RSAUX!Tot1 & " WHERE CODTRAB='" & RSAUX!CODTRAB & "' AND NOMBOL=" & SNOMBOL
         '   RSAUX.MoveNext
        'Loop
        'RSAUX.Close
        'RSAUX.Open "SELECT CODTRAB, SUM(MONTO) AS TOT1 FROM PAGOSCTA WHERE TIPO=2 AND TIPOBOLETA='A' AND CODNOMBOL=" & SNOMBOL & " GROUP BY CODTRAB", DBSYSTEM, adOpenStatic, adLockReadOnly
        'Do While Not RSAUX.EOF
        '    DBSYSTEM.Execute "UPDATE  [##_TMPBOLQUI" & VGL_COMPUTER & "]  SET EGRESOS=" & RSAUX!Tot1 & " WHERE CODTRAB='" & RSAUX!CODTRAB & "' AND NOMBOL=" & SNOMBOL
        '    RSAUX.MoveNext
        'Loop
        'RSAUX.Open "SELECT CODTRAB, SUM(MONTO) AS TOT1 FROM DETADEL WHERE IE=2 AND NOMBOL=" & SNOMBOL & "  GROUP BY CODTRAB", DBSYSTEM, adOpenStatic, adLockReadOnly
        'Do While Not RSAUX.EOF
         '   DBSYSTEM.Execute "UPDATE  [##_TMPBOLQUI" & VGL_COMPUTER & "]  SET EGRESOS=" & RSAUX!Tot1 & " WHERE CODTRAB='" & RSAUX!CODTRAB & "' AND NOMBOL=" & SNOMBOL
         '   RSAUX.MoveNext
        'Loop
        
        
        Set RSAUX = Nothing
            If DevuelveValor("select NOMBOL from [##_TMPBOLQUI" & VGL_COMPUTER & "] WHERE NOMBOL=" & SNOMBOL, DBSYSTEM) <> "" Then
                DBSYSTEM.Execute "UPDATE  [##_TMPBOLQUI" & VGL_COMPUTER & "]  SET INGRESOS=0 WHERE (INGRESOS)IS NULL"
                DBSYSTEM.Execute "UPDATE  [##_TMPBOLQUI" & VGL_COMPUTER & "]  SET INGRESOS=INGRESOS+ADELANTO"
                DBSYSTEM.Execute "UPDATE  [##_TMPBOLQUI" & VGL_COMPUTER & "]  SET EGRESOS=0 WHERE (EGRESOS)IS NULL"
                DBSYSTEM.Execute "UPDATE  [##_TMPBOLQUI" & VGL_COMPUTER & "]  SET NETO=INGRESOS-EGRESOS"
                DBSYSTEM.Execute "UPDATE  [##_TMPBOLQUI" & VGL_COMPUTER & "]  SET SECUENCIA =CASE WHEN LTRIM(RTRIM((SELECT TIPCTAX FROM TRABAJADORES WHERE CODTRAB=[##_TMPBOLQUI" & VGL_COMPUTER & "].CODTRAB)))='' THEN '1' ELSE (SELECT TIPCTAX FROM TRABAJADORES WHERE CODTRAB=[##_TMPBOLQUI" & VGL_COMPUTER & "].CODTRAB) END"
                DBSYSTEM.Execute "UPDATE  [##_TMPBOLQUI" & VGL_COMPUTER & "]  SET CTADESTINO = (SELECT XXCTADES FROM TRABAJADORES WHERE CODTRAB=[##_TMPBOLQUI" & VGL_COMPUTER & "].CODTRAB)"
                DBSYSTEM.Execute "UPDATE [##_TMPBOLQUI" & VGL_COMPUTER & "]  SET CCOSTO = CC.RUC" & _
                                          " FROM [##_TMPBOLQUI" & VGL_COMPUTER & "] LEFT JOIN TRABAJADORES TRA ON TRA.CODTRAB=[##_TMPBOLQUI" & VGL_COMPUTER & "].CODTRAB,CCOSTOS CC WHERE TRA.CCOSTO=CC.CODCCOSTO"
                DBSYSTEM.Execute "UPDATE  [##_TMPBOLQUI" & VGL_COMPUTER & "]  SET CCOSTO = CASE WHEN CCOSTO='' THEN  '' ELSE SUBSTRING(CCOSTO,1,PATINDEX('%:%',CCOSTO)-1) END " & _
                                          "" 'FROM [##_TMPBOLQUI" & VGL_COMPUTER & "])"
                Screen.MousePointer = 1
            End If
            Command2.Enabled = True
    Else
        DBSYSTEM.Execute "DELETE FROM  [##_TMPBOLQUI" & VGL_COMPUTER & "]  WHERE NOMBOL=" & SNOMBOL
    End If


 '----------------------------------------------------------
 Dim codnum As Integer
 Dim CRONOS As Integer
 Dim T As Integer
    
    Call SUMAPLANQUIN(TOTALING, TOTALEGR, TOTALNETO, LstvwCronograma.SelectedItem.Tag)
    NTRAB = CUENTAQUIN(LstvwCronograma.SelectedItem.Tag)
    
    lblTotalPlanilla.Caption = Format(TOTALNETO, "###,###,###.00 ")
    lblNroTrabajadores.Caption = Format(NTRAB, "0")
     
     'APERTURANDO LAS TABLASQUIN DE CONTABILIDAD PARA ARMAR EL ASIENTO DE PLANILLA
     
    PERIODO = Trim(Format(Month(AplitxtMesTrabajo.Tag), "00")) & Trim(Format(Year(AplitxtMesTrabajo.Tag), "0000"))
    codnum = DevuelveValor("select MAX(NUM) from  CONTQUIDET where CRONO=" & LstvwCronograma.SelectedItem.Tag & " ", DBSYSTEM)

    
    Set RSCONTCABQUIN = New ADODB.Recordset
    Set RSCONTDETQUIN = New ADODB.Recordset
    RSCONTCABQUIN.Open "SELECT * FROM CONTQUICAB WHERE NUM=" & codnum & "", DBSYSTEM, adOpenKeyset, adLockOptimistic
    RSCONTDETQUIN.Open "SELECT * FROM CONTQUIDET ORDER BY DMOV_SECUE", DBSYSTEM, adOpenKeyset, adLockOptimistic
'---------------------------------------------------------------------------------------------------
    RSCONTCABQUIN.Requery
    If Not RSCONTCABQUIN.EOF Then
        Set Me.dgComprobante.DataSource = RSCONTCABQUIN
    End If
    RSCONTDETQUIN.Requery
'----------------------------------------------------------------------------------------------------
    
    Command1.Enabled = True
    Set xDgResult.DataSource = RSCONTCAB

CambiaPanelBD False
If codnum = 0 Then Exit Sub
CambiaPanelBD True
    For T = 1 To LstvwCronograma.ListItems.Count
        CRONOS = DevuelveValor("select NUM from  CONTQUIDET where NUM=" & codnum & " AND CRONO=" & LstvwCronograma.ListItems(T).Tag & " ", DBSYSTEM)
        If CRONOS <> 0 Then
            LstvwCronograma.ListItems(T).Checked = True
        Else
            LstvwCronograma.ListItems(T).Checked = False
        End If
    Next T
     
     'Set XDGdetAsi.DataSource = RSCONTDET

CambiaPanelBD False
End Sub

Private Sub XMES_DBLCLICK()
    Dim RSMESES As New ADODB.Recordset
    RSMESES.Open "SELECT MESACTIVO, NOMBRE FROM MESESACT ORDER BY MESACTIVO", DBSYSTEM, adOpenStatic
    If RSMESES.RecordCount = 0 Then
        MsgBox "No se han encontrado meses en actividad", vbCritical
        Set RSMESES = Nothing
        Exit Sub
    End If
    frmComun.CONECTAR RSMESES
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xMes.Text = RSMESES!NOMBRE
        xMes.Tag = RSMESES!MESACTIVO
        Set RSCONTCAB = Nothing
        Set xDgResult.DataSource = Nothing
        xDgResult.ReBind
    Else
        Set RSMESES = Nothing
        Exit Sub
    End If
    Set RSMESES = Nothing
    Command1.Enabled = False
    'Reciclaje de RsMeses
    CARGAMESES
End Sub
Public Sub CARGAMESES()
    Dim RSMESES As New ADODB.Recordset
    Lista.ListItems.Clear
    RSMESES.Open "SELECT CODIGO, NOMBRE, FECHAINI, FECHAFIN FROM NOMBOL WHERE CERRADO=0 AND MES=" & DateSQL(CDate(xMes.Tag)) & " ORDER BY FECHAINI", DBSYSTEM, adOpenStatic
    Do While Not RSMESES.EOF
        Set XITEM = Lista.ListItems.Add(, , RSMESES!NOMBRE, , 1)
        XITEM.SubItems(1) = RSMESES!FECHAINI
        XITEM.SubItems(2) = RSMESES!FECHAFIN
        XITEM.Tag = RSMESES!Codigo
        RSMESES.MoveNext
    Loop
    l1.Visible = False
    l2.Visible = False
    xFechaIni.Visible = False
    xFechaFin.Visible = False
    Set RSMESES = Nothing
End Sub
Private Function CENROCOSTOS() As String
    'FUNCION QUE DEVUELVE EL CENTRO DE COSTOS SIEMPRE Y CUANDO ESTE CONFIGURADO QUE TRABAJA
    'CON UN CENTRO DE COSTOS EN CONTABILIDAD SI NO LO TIENE VA COLOCAR EL CENTRO DE COSTOS
    'COLOCADO EN PLANILLAS
    CENROCOSTOS = ""
    If VERIFI Then
        If ESNULO(GetValor("SELECT PLANCTA_CENTCOST FROM PLAN_CUENTA_NACIONAL " & _
                         "WHERE PLANCTA_CODIGO='" & Trim(RSAUX.Fields("CUENTA")) & "'", CONECTARDBSQL(REGSISTEMA.scRutaEmpresaWenco & "BDCONTABILIDAD")), 0) Then
            CENROCOSTOS = ESNULO(GetValor("SELECT RUC FROM CCOSTOS WHERE CODCCOSTO IN (SELECT CCOSTO FROM TRABAJADORES WHERE CODTRAB='" & Trim(RSBOLETAS("CODTRAB")) & "')", DBSYSTEM), 0)
        End If
      Else
        CENROCOSTOS = ESNULO(GetValor("SELECT RUC FROM CCOSTOS WHERE CODCCOSTO IN (SELECT CCOSTO FROM TRABAJADORES WHERE CODTRAB='" & Trim(RSBOLETAS("CODTRAB")) & "')", DBSYSTEM), 0)
    End If
    If CENROCOSTOS = "0" Then CENROCOSTOS = ""
End Function

Private Sub CARGAMESESQUINCENA()
    Dim RSMESES As New ADODB.Recordset
    LstvwCronograma.ListItems.Clear
    RSMESES.Open "SELECT CODIGO, NOMBRE, FECHAINI, FECHAFIN FROM NOMBOL WHERE CERRADO=0 AND MES=" & DateSQL(CDate(AplitxtMesTrabajo.Tag)) & " ORDER BY FECHAINI", DBSYSTEM, adOpenStatic
    Do While Not RSMESES.EOF
        Set XITEM = LstvwCronograma.ListItems.Add(, "C" & RSMESES!Codigo, RSMESES!NOMBRE, , 1)
        XITEM.SubItems(1) = RSMESES!FECHAINI
        XITEM.SubItems(2) = RSMESES!FECHAFIN
        XITEM.Tag = RSMESES!Codigo
        RSMESES.MoveNext
    Loop
    'l1.Visible = False
    'l2.Visible = False
    'xFechaIni.Visible = False
    'xFechaFin.Visible = False
    Set RSMESES = Nothing
End Sub

Private Sub SUMAPLANQUIN(ByRef xTotIng As Double, ByRef xTotEgr As Double, ByRef xNeto As Double, Crono As Long)
On Error GoTo handler
    Dim SqlCad As String
    Dim RSAUX As New ADODB.Recordset
        
    Set RSAUX = New ADODB.Recordset
        SqlCad = "SELECT SUM(INGRESOS) AS INGRESOS, SUM(EGRESOS) AS EGRESOS, SUM(INGRESOS-EGRESOS) AS NETO " & _
             "FROM  [##_TMPBOLQUI" & VGL_COMPUTER & "] BOLS "
        RSAUX.Open SqlCad, DBSYSTEM, adOpenKeyset, adLockReadOnly
    If RSAUX.RecordCount > 0 Then
        xTotIng = ESNULO(RSAUX("INGRESOS"), 0)
        xTotEgr = ESNULO(RSAUX("EGRESOS"), 0)
        xNeto = ESNULO(RSAUX("NETO"), 0)
    End If
Exit Sub
handler:
End Sub
Private Function CUENTAQUIN(Crono As Long) As Double
    Dim RSCUENTA As New ADODB.Recordset
    Set RSCUENTA = New ADODB.Recordset
    RSCUENTA.Open "SELECT COUNT(CODTRAB) AS CUENTA FROM [##_TMPBOLQUI" & VGL_COMPUTER & "]", DBSYSTEM
    CUENTAQUIN = ESNULO(RSCUENTA!CUENTA, 0)
End Function
Private Sub GENTABLATEMP_ASIENTOS_QUINCENA(MES As Integer, ANNO As Integer, NOMBOL As Integer)
    Dim RSTMPASI As New ADODB.Recordset
    Dim SCAD As String, SCADAUX As String, SCW As String
    Dim RST As New ADODB.Recordset
    Dim CADQUIN As String
    Dim CONT As Integer
    If ExisteTablaAux("[##TEMPASIENTOS" & VGL_COMPUTER & "]") Then DBSYSTEM.Execute "DROP TABLE  [##TEMPASIENTOS" & VGL_COMPUTER & "] "
    'TEMPORAL PARA ALMACENAR EL POSIBLE
    DBSYSTEM.Execute _
    "CREATE TABLE  [##TEMPASIENTOS" & VGL_COMPUTER & "] (CODTRAB VARCHAR(8),CONCEPTO VARCHAR(25),CUENTA VARCHAR(25),UBIC VARCHAR(1),MONTO FLOAT,TIPASI VARCHAR(1),ANEXO VARCHAR(25),CCOSTO VARCHAR(15),CTADEST VARCHAR(25))"
    RSTMPASI.Open "[##TEMPASIENTOS" & VGL_COMPUTER & "]", DBSYSTEM, adOpenKeyset, adLockOptimistic
                                                          'BOL" & PERIODO & "     --MOV" & PERIODO & "
    SCAD = "SELECT CODTRAB,CODCONCEP=CASE WHEN ISNULL(DETADEL.CODCONCEP,'')='' THEN 'XXADELTO' ELSE CODCONCEP END ,MONTO FROM DETADEL WHERE CODCONCEP<>'XXX' AND NOMBOL=" & NOMBOL & _
           "  Union All " & _
           "SELECT DISTINCT(CODTRAB),CODCONCEP='XXADEQUI',TOTAL AS MONTO FROM DETADEL  WHERE CODCONCEP<>'XXX' AND NOMBOL= " & NOMBOL
         
    If DevuelveValor("SELECT CODIGO FROM  CONCEPTOS WHERE CODIGO='XXPAGCXI'", DBSYSTEM) <> 0 Then
         SCW = "1"
      Else: SCW = "9"
    End If
    If DevuelveValor("SELECT CODIGO FROM  CONCEPTOS  WHERE CODIGO='XXPAGCXE'", DBSYSTEM) <> 0 Then
         If SCW <> "" Then
             SCW = SCW & ",2"
             Else: SCW = "2"
         End If
       Else
       'Cuentas Corrientes Programadas de Quincena
          Set RST = New ADODB.Recordset
          RST.Open "SELECT * FROM CONCEPTOS WHERE CODIGO LIKE 'XA%'", DBSYSTEM, adOpenKeyset, adLockReadOnly
          If RST.RecordCount = 0 Then
            SCW = "9"
            Else: SCW = "8"
          End If
    End If
    
    
    Select Case SCW
        Case "9"
            SCADAUX = " UNION ALL " & _
         "SELECT PAG.CODTRAB,'XX'+LTRIM(RTRIM(MOV.CODGRUPO)) AS CONCEPTO ," & _
         "SUM(PAG.MONTO) AS MONTO FROM PAGOSCTA PAG,MOVICTA MOV " & _
         "WHERE PAG.CODMOV = MOV.CODMOV AND PAG.CODNOMBOL =" & NOMBOL & _
         " GROUP BY PAG.CODTRAB,MOV.CODGRUPO "
        Case "8"
            'Si tiene Cuentas Corrientes Programadas
            SCADAUX = " UNION ALL " & _
          "SELECT PAG.CODTRAB,CONCEPTO=CASE PAG.TIPOBOLETA " & _
                           "  WHEN 'A' THEN 'XA'+LTRIM(RTRIM(MOV.CODGRUPO))  END , " & _
          "SUM(PAG.MONTO) AS MONTO FROM PAGOSCTA PAG,MOVICTA MOV " & _
          "WHERE TIPOBOLETA='A' AND PAG.CODMOV = MOV.CODMOV AND PAG.CODNOMBOL =" & NOMBOL & _
          " GROUP BY PAG.CODTRAB,MOV.CODGRUPO,PAG.TIPOBOLETA "
        Case Else ' TODOS LOS PRESTAMOS EN UNA SOLA CUENTA
         SCADAUX = " UNION ALL " & _
         "SELECT CODTRAB,CASE TIPO WHEN 1 THEN 'XXPAGCXI' WHEN 2 THEN 'XXPAGCXE' END AS CONCEPTO " & _
         ",SUM(PAGOSCTA.MONTO) AS MONTO FROM PAGOSCTA " & _
         "WHERE CODNOMBOL =" & NOMBOL & "AND TIPO IN (" & SCW & ") GROUP BY CODTRAB,TIPO "
    End Select
    
    
    
    If ExisteTablaAux("[##TMPQUIN" & VGL_COMPUTER & "]") Then DBSYSTEM.Execute "DROP TABLE [##TMPQUIN" & VGL_COMPUTER & "]"
    DBSYSTEM.Execute "SELECT * INTO [##TMPQUIN" & VGL_COMPUTER & "] FROM (" & SCAD & SCADAUX & ") AS UNIONXXX"
    If ExisteTablaAux("[##TMPQUINCTA" & VGL_COMPUTER & "]") Then DBSYSTEM.Execute "DROP TABLE [##TMPQUINCTA" & VGL_COMPUTER & "]"
    CADQUIN = "SELECT CONCEPTO='XXADELTO',TIPOCTA='D', CUENTA=MONADEL,TIPASI=0  FROM CFGASIENTOS UNION ALL " & _
    "SELECT  CONCEPTO='XXADEQUI',TIPOCTA='H', CUENTA=NETADEL,TIPASI=0  FROM CFGASIENTOS UNION ALL " & _
    "SELECT CONCEPTO=CONCEPT,TIPOCTA,CUENTA,TIPASI FROM CTACONCEPTO   WHERE  CONCEPT LIKE 'XA%' UNION ALL  " & _
    "SELECT CONCEPTO=CONCEPT,TIPOCTA,CUENTA,TIPASI FROM CTACONCEPTOQUIN "
    DBSYSTEM.Execute "SELECT * INTO [##TMPQUINCTA" & VGL_COMPUTER & "] FROM (" & CADQUIN & ") AS UNIONXXX"
        
    If ExisteTablaAux("[##TMPQUIN_AUX" & VGL_COMPUTER & "]") Then DBSYSTEM.Execute "DROP TABLE [##TMPQUIN_AUX" & VGL_COMPUTER & "]"
    DBSYSTEM.Execute "SELECT CODTRAB,CODCONCEP,MONTO INTO [##TMPQUIN_AUX" & VGL_COMPUTER & "] FROM [##TMPQUIN" & VGL_COMPUTER & "]"
    If ExisteTablaAux("[##TMPQUINCTA_AUX" & VGL_COMPUTER & "]") Then DBSYSTEM.Execute "DROP TABLE [##TMPQUINCTA_AUX" & VGL_COMPUTER & "]"
    DBSYSTEM.Execute "SELECT TMPMON.CODTRAB,TMPCTA.CONCEPTO,TMPCTA.CUENTA,TMPCTA.TIPOCTA,TMPMON.MONTO,TMPCTA.TIPASI INTO [##TMPQUINCTA_AUX" & VGL_COMPUTER & "]" & _
    " FROM [##TMPQUIN_AUX" & VGL_COMPUTER & "] TMPMON ,[##TMPQUINCTA" & VGL_COMPUTER & "] TMPCTA Where TMPCTA.CONCEPTO = TMPMON.CODCONCEP"
    'AGREGO UN CAMPO LLAMADO CCOSTO AL AUX
    DBSYSTEM.Execute "ALTER TABLE [##TMPQUINCTA_AUX" & VGL_COMPUTER & "] ADD CCOSTO VARCHAR(50)"
    DBSYSTEM.Execute "UPDATE [##TMPQUINCTA_AUX" & VGL_COMPUTER & "] SET CCOSTO= TMP.CCOSTO FROM [##TMPQUINCTA_AUX" & VGL_COMPUTER & "] TMPAUX LEFT JOIN [##_TMPBOLQUI" & VGL_COMPUTER & "] TMP ON TMPAUX.CODTRAB=TMP.CODTRAB"

Dim xxtipe As String

Set RSBOLETAS = New ADODB.Recordset
    RSBOLETAS.Open SCAD & SCADAUX, DBSYSTEM, adOpenKeyset, adLockReadOnly
    VERIFI = VERIFI_CONTA(2)
        xxtipe = DevuelveValor("select max(TIPASI) from [##TMPQUINCTA_AUX" & VGL_COMPUTER & "]", DBSYSTEM)
    Select Case xxtipe
          Case "1"
                Call asientoSimple
          Case "2"
                Call asientotrabaj
          Case "3"
                Call asientoCCosto(True)
                'PROCESO QUE SE ENCARGA DE AGRUPAR POR CENTRO DE COSTO
                Call AGRUPAR_CCOSTOQUIN(LstvwCronograma.SelectedItem.Tag, DevuelveValor("SELECT MAX(NUM) FROM  CONTQUICAB", DBSYSTEM))
          Case Else
                MsgBox "Este tipo de asiento no se puede generar por que no especifica" & _
                "si es un asiento simple,por trabajador o por C.costo", vbCritical, "Error de Usuario"
                MsgBox "Configure el Tipo de Asiento la Tabla Conceptos Generales y Configuración de " & vbCrLf & _
                " Adelantos asegurese que los Conceptos Afectos al Adelanto contengan el mismo tipo de Asiento", vbInformation
                
    End Select
    
    Set RSCONTCABQUIN = New ADODB.Recordset
    RSCONTCABQUIN.Open "SELECT * FROM CONTQUICAB WHERE CRONO=" & LstvwCronograma.SelectedItem.Tag & "", DBSYSTEM, adOpenDynamic, adLockOptimistic
    RSCONTCABQUIN.Requery
    Set dgComprobante.DataSource = RSCONTCABQUIN
    
    Set RSTMPASI = New ADODB.Recordset
    RSTMPASI.Open "SELECT * FROM  CONTQUIDET WHERE NUM=" & DevuelveValor("SELECT MAX(NUM) FROM  CONTQUICAB", DBSYSTEM) & "  order by DMOV_CUENT", DBSYSTEM, adOpenKeyset, adLockOptimistic
    CONT = 0
    ProgressBar2.Value = 0
    If Not RSTMPASI.EOF Then ProgressBar2.Max = RSTMPASI.RecordCount
    Do While Not RSTMPASI.EOF
        ProgressBar2.Value = ProgressBar2.Value + 1
        CONT = CONT + 1
        lblMessage2.Caption = "Generando Secuencia ::" & CStr(Format(CONT, "0000"))
          RSTMPASI!DMOV_SECUE = CStr(Format(CONT, "0000"))
        RSTMPASI.Update
       RSTMPASI.MoveNext
    Loop
    MsgBox "El proceso ha concluido satisfactoriamente", vbInformation





    
End Sub

Private Sub GENASIENTOTEMPQUIN()
    Dim RSCONCEPTOS As New ADODB.Recordset
    Dim RSAUX As New ADODB.Recordset
    Dim ORDEN As Integer
    If VERIFI_CONTA(2) Then
        TIPCAM = DevuelveValor("SELECT TIPOCAMB_EQCOMPRA FROM TIPO_CAMBIO " & _
             "WHERE TIPOMON_CODIGO='ME' AND CAST(TIPOCAMB_FECHA AS INT)=" & FechS(AplitxtMesTrabajo.Tag, Sqlf), CONECTARDBSQL(REGSISTEMA.scRutaEmpresaWenco & "BDCONTABILIDAD"))
    End If
    If TIPCAM > 0 Then
        DESCAM = "VTA"
    Else
        MsgBox "No se encontro tipo de cambio venta en contabilidad " & Chr(13) & _
               "se Utilizara el tipo de cambio puesto en la barra de planillas", vbInformation
        TIPCAM = 1 / Valc(MDIPrincipal.BarraEstado.Panels(3).Text)
        DESCAM = "ESP"
    End If
        
    'GRABA LA CABECERA DEL ASIENTO DE PLANILLA
    Call GRABACABQUIN(RSCONTCABQUIN)
    NUM = DevuelveValor("SELECT MAX(NUM) FROM CONTQUICAB", DBSYSTEM)
    
    'GRABA EL REGISTRO DEL NETO DE ADELANTOS
    'CTA = REGSISTEMA.scCuenta
    CTA = DevuelveValor("SELECT NETADEL FROM CFGASIENTOS ", DBSYSTEM)
    MONTO = Valc(lblTotalPlanilla)
    SECUE = 1
    CCosto = "": ANEX = "": CTADES = ""
    
    Call GRABADETASI(RSCONTDETQUIN, "H")
        
    Set RSCONCEPTOS = New ADODB.Recordset
    'SELECT CONCEPTO,CUENTA,TIPOCTA,MONTO  FROM [##TMPQUINCTA_AUXPC02]
    RSCONCEPTOS.Open "SELECT TMP.*,TIPASI=TCTA.TIPASI   " & _
                     " FROM [##TMPQUINCTA" & VGL_COMPUTER & "] TMP ,CTACONCEPTOQUIN TCTA WHERE TCTA.CONCEPT=TMP.CONCEPTO", DBSYSTEM, adOpenKeyset, adLockReadOnly
    pgbar.Min = 0: pgbar.Max = RSCONCEPTOS.RecordCount: pgbar.Value = 0
    Me.Refresh
  '  lbbar.ForeColor = &HC00000
  '  lbbar.Caption = "Generando el Asiento Contable ..."
  '  pgbar.Scrolling = ccScrollingSmooth
    Do While Not RSCONCEPTOS.EOF
        Set RSAUX = New ADODB.Recordset
        'pgbar.Value = pgbar.Value + 1
        Select Case RSCONCEPTOS!TIPASI
             Case 1 'SIMPLE (AGRUPAR CUENTA)
                 ORDEN = 1
                 SECUE = SECUE + 1
                 RSAUX.Open "SELECT CUENTA, Sum(MONTO) AS TOTAL " & _
                            "From  [##TEMPASIENTOS" & VGL_COMPUTER & "]  WHERE CONCEPTO='" & RSCONCEPTOS("CONCEPTO") & "' AND UBIC='" & Trim(RSCONCEPTOS!UBIC) & "' GROUP BY CUENTA ", DBSYSTEM
                 CTA = RSAUX("CUENTA"): MONTO = RSAUX("TOTAL")
                 Call GRABADETASIQUIN(RSCONTDETQUIN, RSCONCEPTOS!UBIC)
                 
             Case 2 'POR TRABAJADOR
                ORDEN = 2
                RSAUX.Open "SELECT CONCEPTO,MONTO AS TOTAL,CUENTA,CTADEST=TIPOCTA From  [##TMPQUINCTA_AUX" & VGL_COMPUTER & "]  " & _
                           " WHERE CONCEPTO='" & Trim(RSCONCEPTOS("CONCEPTO")) & "'", DBSYSTEM, adOpenStatic, adLockReadOnly
                Do While Not RSAUX.EOF
                    SECUE = SECUE + 1
                    CTA = RSAUX("CUENTA"): MONTO = RSAUX("TOTAL"): ANEX = RSAUX("ANEXO")
                    Call GRABADETASIQUIN(RSCONTDETQUIN, RSCONCEPTOS!UBIC)
                    RSAUX.MoveNext
                Loop
             Case 3 'POR CENTRO DE COSTOS
                ORDEN = 3
                RSAUX.Open "SELECT CONCEPTO,MONTO AS TOTAL,CUENTA,CTADEST=TIPOCTA FROM  [##TMPQUINCTA_AUX" & VGL_COMPUTER & "]  " & _
                           " WHERE CONCEPTO='" & Trim(RSCONCEPTOS("CONCEPTO")) & "'", DBSYSTEM, adOpenStatic, adLockReadOnly
                Do While Not RSAUX.EOF
                    SECUE = SECUE + 1
                    CTA = RSAUX("CUENTA"): MONTO = RSAUX("TOTAL"): CCosto = RSAUX("CCOSTO")
                    ANEX = "": CTADES = RSAUX("CTADEST")
                    Call GRABADETASIQUIN(RSCONTDETQUIN, RSAUX!CTADEST)
                    RSAUX.MoveNext
                Loop
             Case 4 'POR A.F.P.
             Case 5 'POR TRABAJADOR Y CENTRO DE COSTOS
        End Select
        RSCONCEPTOS.MoveNext
    Loop
    
    'SE ACTUALIZA EL MONTO TOTAL DEL COMPROBANTE
    Dim RSTOTCAB As ADODB.Recordset, TOTCABDEBE As Double, TOTCABHABER As Double
    Dim TOTCABDEBUS As Double, TOTCABHABUS As Double
    Dim REDON As Double, RETIP As String
    Set RSTOTCAB = New ADODB.Recordset
    RSTOTCAB.Open "SELECT SUM(DMOV_DEBE) AS TOTALDEBE,SUM(DMOV_HABER) AS TOTALHABER, " & _
    "SUM(DMOV_DEBUS) AS TOTALDEBUS,SUM(DMOV_HABUS) AS TOTALHABUS FROM CONTQUIDET WHERE NUM=" & NUM & " AND CRONO=" & LstvwCronograma.SelectedItem.Tag, DBSYSTEM, adOpenKeyset, adLockReadOnly
    TOTCABDEBE = Round(ESNULO(RSTOTCAB("TOTALDEBE"), 0), 2)
    TOTCABHABER = Round(ESNULO(RSTOTCAB("TOTALHABER"), 0), 2)
    TOTCABDEBUS = Round(ESNULO(RSTOTCAB("TOTALDEBUS"), 0), 2)
    TOTCABHABUS = Round(ESNULO(RSTOTCAB("TOTALHABUS"), 0), 2)
    REDON = TOTCABDEBE - TOTCABHABER
    'PARA LA CUENTA DE REDONDEO NO SQL
    If Abs(REDON) < 1 Then
        CTA = REGSISTEMA.scCtaRedon
        SECUE = SECUE + 1
        CCosto = "": ANEX = ""
        If REDON > 0 Then
            RETIP = "H"
            TOTCABHABER = TOTCABHABER + Abs(REDON)
            TOTCABHABUS = TOTCABHABUS + Round((Abs(REDON) * TIPCAM), 2)
           Else
            RETIP = "D"
            TOTCABDEBE = TOTCABDEBE + Abs(REDON)
            TOTCABDEBUS = TOTCABDEBUS + Round((Abs(REDON) * TIPCAM), 2)
        End If
        MONTO = Abs(REDON)
        Call GRABADETASIQUIN(RSCONTDETQUIN, RETIP)
    End If
    
    DBSYSTEM.Execute "UPDATE CONTQUICAB SET CMOV_DEBE=" & TOTCABDEBE & "," & _
                     "CMOV_HABER=" & TOTCABHABER & ",CMOV_DEBUS=" & TOTCABDEBUS & ",CMOV_HABUS=" & TOTCABHABUS & " WHERE CRONO=" & LstvwCronograma.SelectedItem.Tag
    
    Call GENERASECUENCIAQUIN(ORDEN)
    RSCONTCABQUIN.Requery
    RSCONTDETQUIN.Requery
    RSCONTCABQUIN.Filter = "CRONO=" & LstvwCronograma.SelectedItem.Tag
    MsgBox "El proceso concluyo satisfactoriamente", vbInformation
End Sub


Private Sub GRABACABQUIN(RSCAB As ADODB.Recordset)
    'CREANDO LA CABECERA DEL ASIENTO DE PLANILLAS
    With RSCAB
        .AddNew
        .Fields("SUBDIAR_CODIGO") = Trim(REGSISTEMA.scSubdi)
        .Fields("CMOV_C_COMPR") = " "
        .Fields("CMOV_FECHA") = AplitxtMesTrabajo.Tag
        .Fields("CMOV_GLOSA") = Trim(Mid("Planillas " & LstvwCronograma.SelectedItem.Text, 1, 29))
        .Fields("CMOV_MONED") = "MN"
        .Fields("CMOV_CONVE") = DESCAM  'FORMA DE CAMBIO
        .Fields("CMOV_CAMES") = IIf(DESCAM = "ESP", 1 / TIPCAM, 0)
        .Fields("CMOV_FECCA") = AplitxtMesTrabajo.Tag
        .Fields("CMOV_TIPCA") = IIf(DESCAM = "VTA", TIPCAM, 0)
        .Fields("CMOV_DEBE") = 0
        .Fields("CMOV_HABER") = 0
        .Fields("CMOV_DEBUS") = 0
        .Fields("CMOV_HABUS") = 0
        .Fields("CMOV_AUTOM") = 0
        .Fields("CMOV_COSTO") = 0
        .Fields("CMOV_CHEQU") = 0
        .Fields("CMOV_L_COMPR") = 0
        .Fields("CMOV_VENTA") = 0
        .Fields("CRONO") = LstvwCronograma.SelectedItem.Tag
        .Update
    End With
End Sub

Private Sub GRABADETASIQUIN(RSDET As ADODB.Recordset, TIPO As String, Optional NUMX As String, Optional SECX As Integer)
    'CREANDO EL DETALLE DEL ASIENTO DE PLANILLAS
 Dim anexo As String, DOCUMENTO As String
 Dim CAD1 As String
 Dim RSDETQUINAUX As ADODB.Recordset
 Dim ANEXOVERDAD As String
 Dim BIT_PLANCTA_CENTCOST As Boolean
 Dim Cnx As New ADODB.Connection
 
 If RSDET!MONTO = 0 Then Exit Sub
 
 If REGSISTEMA.scTieneStConta Then
    'anexo = REGSISTEMA.scTipoAnexo
    Set Cnx = CONECTARDBSQL(REGSISTEMA.scRutaEmpresaWenco & "BDCONTABILIDAD")
 End If
 'DevuelveValor("SELECT ANEX_CODIGO FROM ANEXO WHERE TIPOANEX_CODIGO='" & TIPANEX & "' AND ANEX_CODIGO='" & Codigo & "'", Cnx)
 ANEXOVERDAD = DevuelveValor("SELECT TIPOANEX_CODIGO FROM PLAN_CUENTA_NACIONAL " & _
         " WHERE PLANCTA_CODIGO ='" & RSDET!CUENTA & "'", CONECTARDBSQL(REGSISTEMA.scRutaEmpresaWenco & "BDCONTABILIDAD"))
   
 BIT_PLANCTA_CENTCOST = DevuelveValor("SELECT PLANCTA_CENTCOST FROM PLAN_CUENTA_NACIONAL " & _
         " WHERE PLANCTA_CODIGO ='" & RSDET!CUENTA & "'", CONECTARDBSQL(REGSISTEMA.scRutaEmpresaWenco & "BDCONTABILIDAD"))
 If ANEXOVERDAD <> "" Then
    'anexo = DevuelveValor("SELECT ANEX_CODIGO FROM ANEXO WHERE TIPOANEX_CODIGO='" & ANEXOVERDAD & "' AND ANEX_CODIGO='" & RSDET!CODTRAB & "'", Cnx)
    'anexo = DevuelveValor("SELECT ANEX_CODIGO FROM ANEXO " & _
         " WHERE TIPOANEX_CODIGO ='" & ANEXOVERDAD & "'  AND ANEX_CODIGO='" & RSDET!CODTRAB & "'", CONECTARDBSQL(REGSISTEMA.scRutaEmpresaWenco & "BDCONTABILIDAD"))
    DOCUMENTO = RSDET!CODTRAB
 Else
    DOCUMENTO = ""
 End If
 
 Set RSDETQUINAUX = New ADODB.Recordset
 RSDETQUINAUX.Open "CONTQUIDET", DBSYSTEM, adOpenDynamic, adLockOptimistic
    
    With RSDETQUINAUX
        .AddNew
        .Fields("NUM") = CInt(NUMX)
        .Fields("SUBDIAR_CODIGO") = REGSISTEMA.scSubdi
        .Fields("DMOV_C_COMPR") = " "
        .Fields("DMOV_SECUE") = Format(SECX, "0000")
        .Fields("DMOV_FECHA") = FechS(AplitxtMesTrabajo.Tag, Adof)
        .Fields("DMOV_CUENT") = ESNULO(RSDET!CUENTA, " ")
        .Fields("DMOV_ANEXO") = ESNULO(ANEXOVERDAD, "") & IIf(Trim(ANEXOVERDAD) = "", "", DOCUMENTO)
        .Fields("DMOV_DOCUM") = ""
        .Fields("DMOV_FECDC") = FechS(AplitxtMesTrabajo.Tag, Adof)
        .Fields("DMOV_CENCO") = IIf((BIT_PLANCTA_CENTCOST = True), ESNULO(RSDET!CCosto, ""), "")
        .Fields("DMOV_DEBE") = Round(IIf(TIPO = "D", RSDET!MONTO, 0), 2)
        .Fields("DMOV_HABER") = Round(IIf(TIPO = "H", RSDET!MONTO, 0), 2)
        .Fields("DMOV_DEBUS") = Round(IIf(TIPO = "D", RSDET!MONTO * TIPCAM, 0), 2)
        .Fields("DMOV_HABUS") = Round(IIf(TIPO = "H", RSDET!MONTO * TIPCAM, 0), 2)
        .Fields("DMOV_GLOSA") = " "
        .Fields("DMOV_CHEQU") = 0
        .Fields("DMOV_AUTOM") = 0
        .Fields("DMOV_COSTO") = 0
        .Fields("DMOV_L_COMPR") = 0
        .Fields("DMOV_VENTA") = 0
        .Fields("DMOV_TRANS") = 0
        .Fields("DMOV_C_DESTI") = "" 'DevuelveValor("SELECT XXCTADES FROM TRABAJADORES WHERE XXCTADES='" & RSDET!CUENTA & "' AND CODTRAB='" & DOCUMENTO & "'", DBSYSTEM) & ""
        .Fields("DMOV_L_DESTI") = 0 'IIf(.Fields("DMOV_C_DESTI") = "", 0, 1)
        .Fields("CRONO") = LstvwCronograma.SelectedItem.Tag
        .Update
    End With
End Sub

Private Function REGDHQUIN(RS As ADODB.Recordset, TIPMOV As String) As ADODB.Recordset
    Set REGDHQUIN = New ADODB.Recordset
    REGDHQUIN.Open "SELECT TMP.*,TIPASI=TCTA.TIPASI FROM [##_TMPBOLQUI" & VGL_COMPUTER & "] TMP ,CTACONCEPTOQUIN TCTA WHERE   " & _
                 "TCTA.CONCEPT='" & Trim(RS!CODCONCEP) & "' AND TCTA.TIPOCTA='" & Trim(TIPMOV) & "'", DBSYSTEM, adOpenKeyset, adLockReadOnly
End Function

Private Sub GENERASECUENCIAQUIN(ORDEN As Integer)
    Dim RSDETAUX As New ADODB.Recordset
    Dim CAD As String
    Dim CONT As Integer
    Set RSDETAUX = New ADODB.Recordset
    Select Case ORDEN
        Case 1: CAD = "SELECT * FROM CONTQUIDET WHERE NUM=" & NUM & "  ORDER BY DMOV_CUENT "
        Case 2: CAD = "SELECT * FROM CONTQUIDET WHERE NUM=" & NUM & " ORDER BY DMOV_ANEXO,DMOV_CUENT "
        Case 3: CAD = "SELECT * FROM CONTQUIDET WHERE NUM=" & NUM & " ORDER BY DMOV_CENCO,DMOV_CUENT "
    End Select
    
    RSDETAUX.Open CAD, DBSYSTEM, adOpenKeyset, adLockOptimistic
    'lbbar.ForeColor = &HC00000
    'lbbar.Caption = "GENERANDO EL ORDEN DE LA SECUENCIA SEGUN EL TIPO DE ASIENTO"
    'pgbar.Min = 0: pgbar.Max = RSDETAUX.RecordCount: pgbar.Value = 0
    'pgbar.Scrolling = ccScrollingSmooth
    Me.Refresh
    CONT = 1
    Do While Not RSDETAUX.EOF
       'pgbar.Value = pgbar.Value + 1
       RSDETAUX("DMOV_SECUE") = Format(CONT, "0000")
       CONT = CONT + 1
       RSDETAUX.Update
       RSDETAUX.MoveNext
    Loop
End Sub

Private Sub asientoSimple()

Dim NUMCORRELATIVO As String
If VERIFI_CONTA(2) Then
        TIPCAM = DevuelveValor("SELECT TIPOCAMB_EQCOMPRA FROM TIPO_CAMBIO " & _
             "WHERE TIPOMON_CODIGO='ME' AND CAST(TIPOCAMB_FECHA AS INT)=" & FechS(AplitxtMesTrabajo.Tag, Sqlf), CONECTARDBSQL(REGSISTEMA.scRutaEmpresaWenco & "BDCONTABILIDAD"))
    End If
    If TIPCAM > 0 Then
        DESCAM = "VTA"
    Else
        MsgBox "No se encontro tipo de cambio venta en contabilidad " & Chr(13) & _
               "se Utilizara el tipo de cambio puesto en la barra de planillas", vbInformation
        TIPCAM = 1 / Valc(MDIPrincipal.BarraEstado.Panels(3).Text)
        DESCAM = "ESP"
End If

Set RSAUX = New ADODB.Recordset
RSAUX.Open "SELECT * FROM CONTQUICAB", DBSYSTEM, adOpenDynamic, adLockOptimistic
With RSAUX
    .AddNew
    .Fields("SUBDIAR_CODIGO") = Trim(REGSISTEMA.scSubdi)
    .Fields("CMOV_C_COMPR") = " "
    .Fields("CMOV_FECHA") = AplitxtMesTrabajo.Tag
    .Fields("CMOV_GLOSA") = Trim(Mid("Planillas " & LstvwCronograma.SelectedItem.Text, 1, 29))
    .Fields("CMOV_MONED") = "MN"
    .Fields("CMOV_CONVE") = DESCAM  'FORMA DE CAMBIO
    .Fields("CMOV_CAMES") = IIf(DESCAM = "ESP", 1 / TIPCAM, 0)
    .Fields("CMOV_FECCA") = AplitxtMesTrabajo.Tag
    .Fields("CMOV_TIPCA") = IIf(DESCAM = "VTA", TIPCAM, 0)
    .Fields("CMOV_DEBE") = Round(DevuelveValor("SELECT sum(MONTO) FROM [##TMPQUINCTA_AUX" & VGL_COMPUTER & "]  where TIPOCTA='D'", DBSYSTEM), 2)
    .Fields("CMOV_HABER") = Round(DevuelveValor("SELECT sum(MONTO) FROM [##TMPQUINCTA_AUX" & VGL_COMPUTER & "]  where TIPOCTA='H'", DBSYSTEM), 2)
    .Fields("CMOV_DEBUS") = Round(.Fields("CMOV_DEBE") * TIPCAM, 2)
    .Fields("CMOV_HABUS") = Round(.Fields("CMOV_HABER") * TIPCAM, 2)
    .Fields("CMOV_AUTOM") = 0
    .Fields("CMOV_COSTO") = 0
    .Fields("CMOV_CHEQU") = 0
    .Fields("CMOV_L_COMPR") = 0
    .Fields("CMOV_VENTA") = 0
    .Fields("CRONO") = LstvwCronograma.SelectedItem.Tag
.Update
End With
NUMCORRELATIVO = DevuelveValor("SELECT MAX(NUM) FROM CONTQUICAB", DBSYSTEM)

Dim RSAUX_DET  As ADODB.Recordset
  Set RSAUX_DET = New ADODB.Recordset
            'RSAUX_DET.Open "SELECT DISTINCT(DMOV_CUENT) FROM [ASIENTOCC" & VGL_COMPUTER & "] WHERE DMOV_DEBE=0", DBSYSTEM, adOpenDynamic, adLockOptimistic
            RSAUX_DET.Open "SELECT DISTINCT(CUENTA) FROM [##TMPQUINCTA_AUX" & VGL_COMPUTER & "] WHERE TIPOCTA='D'", DBSYSTEM, adOpenDynamic, adLockOptimistic
            ProgressBar1.Value = 0
            If Not RSAUX_DET.EOF Then ProgressBar1.Max = RSAUX_DET.RecordCount
            Do While Not RSAUX_DET.EOF
                If DevuelveValor("SELECT DMOV_CUENT FROM CONTQUIDET WHERE DMOV_CUENT='" & RSAUX_DET!CUENTA & "' AND NUM=" & NUMCORRELATIVO, DBSYSTEM) = "" Then
                   DBSYSTEM.Execute "INSERT INTO CONTQUIDET  SELECT NUM=" & NUMCORRELATIVO & ",SUBDIAR_CODIGO='" & Trim(REGSISTEMA.scSubdi) & "',DMOV_C_COMPR='',DMOV_SECUE=0,DMOV_FECHA='" & AplitxtMesTrabajo.Tag & "',DMOV_CUENT=CUENTA,DMOV_ANEXO='',DMOV_DOCUM='',DMOV_FECDC='" & AplitxtMesTrabajo.Tag & "'," & _
                                    "DMOV_CENCO='',DMOV_DEBE=0,DMOV_HABER=SUM(MONTO),DMOV_DEBUS=0,DMOV_HABUS=SUM(MONTO)*" & TIPCAM & ",DMOV_GLOSA='',DMOV_CHEQU=0" & _
                                    ",DMOV_AUTOM=0,DMOV_COSTO=0,DMOV_L_COMPR=0,DMOV_VENTA=0,DMOV_TRANS=0,DMOV_L_DESTI=0,DMOV_C_DESTI='',CRONO=" & LstvwCronograma.SelectedItem.Tag & "" & _
                                    " FROM [##TMPQUINCTA_AUX" & VGL_COMPUTER & "] WHERE CUENTA='" & RSAUX_DET!CUENTA & "' AND TIPOCTA='D' GROUP BY CUENTA,TIPOCTA " & _
                                    " "
                End If
                ProgressBar1.Value = ProgressBar1.Value + 1
                lblMessage.Caption = "Generando Montos del Debe :: " & RSAUX_DET!CUENTA
                RSAUX_DET.MoveNext
            Loop
            'HABER
            Set RSAUX_DET = New ADODB.Recordset
            RSAUX_DET.Open "SELECT DISTINCT(CUENTA) FROM [##TMPQUINCTA_AUX" & VGL_COMPUTER & "] WHERE TIPOCTA='H'", DBSYSTEM, adOpenDynamic, adLockOptimistic
            ProgressBar2.Value = 0
            If Not RSAUX_DET.EOF Then ProgressBar2.Max = RSAUX_DET.RecordCount
            Do While Not RSAUX_DET.EOF
                If DevuelveValor("SELECT DMOV_CUENT FROM CONTQUIDET WHERE DMOV_CUENT='" & RSAUX_DET!CUENTA & "' AND NUM=" & NUMCORRELATIVO, DBSYSTEM) = "" Then
                   DBSYSTEM.Execute "INSERT INTO CONTQUIDET  SELECT NUM=" & NUMCORRELATIVO & ",SUBDIAR_CODIGO='" & Trim(REGSISTEMA.scSubdi) & "',DMOV_C_COMPR='',DMOV_SECUE=0,DMOV_FECHA='" & AplitxtMesTrabajo.Tag & "',DMOV_CUENT=CUENTA,DMOV_ANEXO='',DMOV_DOCUM='',DMOV_FECDC='" & AplitxtMesTrabajo.Tag & "'," & _
                                    "DMOV_CENCO='',DMOV_DEBE=SUM(MONTO),DMOV_HABER=0,DMOV_DEBUS=SUM(MONTO)*" & TIPCAM & ",DMOV_HABUS=0,DMOV_GLOSA='',DMOV_CHEQU=0" & _
                                    ",DMOV_AUTOM=0,DMOV_COSTO=0,DMOV_L_COMPR=0,DMOV_VENTA=0,DMOV_TRANS=0,DMOV_L_DESTI=0,DMOV_C_DESTI='',CRONO=" & LstvwCronograma.SelectedItem.Tag & "" & _
                                    " FROM [##TMPQUINCTA_AUX" & VGL_COMPUTER & "] WHERE CUENTA='" & RSAUX_DET!CUENTA & "' AND TIPOCTA='H' GROUP BY CUENTA,TIPOCTA " & _
                                    " "
                End If
                ProgressBar2.Value = ProgressBar2.Value + 1
                lblMessage2.Caption = "Generando Montos del Haber :: " & RSAUX_DET!CUENTA
                RSAUX_DET.MoveNext
            Loop
            
  
End Sub
Private Sub asientotrabaj()
Call asientoCCosto(False)
End Sub
Private Sub asientoCCosto(Optional TIPO_ASIENTO As Boolean)
'VALIDAR SI LOS REGISTROS TIENEN CENTRO DE COSTO
Dim RSAUX As ADODB.Recordset, rsCTA As ADODB.Recordset
Dim TIPCCO As String, NUMCORRELATIVO As String, CADSQL As String
Dim CONTADOR As Integer
'***********************************************
If VERIFI_CONTA(2) Then
        TIPCAM = DevuelveValor("SELECT TIPOCAMB_EQCOMPRA FROM TIPO_CAMBIO " & _
             "WHERE TIPOMON_CODIGO='ME' AND CAST(TIPOCAMB_FECHA AS INT)=" & FechS(AplitxtMesTrabajo.Tag, Sqlf), CONECTARDBSQL(REGSISTEMA.scRutaEmpresaWenco & "BDCONTABILIDAD"))
    End If
    If TIPCAM > 0 Then
        DESCAM = "VTA"
    Else
        MsgBox "No se encontro tipo de cambio venta en contabilidad " & Chr(13) & _
               "se Utilizara el tipo de cambio puesto en la barra de planillas", vbInformation
        TIPCAM = 1 / Valc(MDIPrincipal.BarraEstado.Panels(3).Text)
        DESCAM = "ESP"
End If
'***********************************************
Set RSAUX = New ADODB.Recordset
RSAUX.Open "SELECT COUNT(CCOSTO) FROM [##TMPQUINCTA_AUX" & VGL_COMPUTER & "] WHERE CCOSTO<>''", DBSYSTEM, adOpenStatic, adLockOptimistic
If TIPO_ASIENTO = True Then
    If RSAUX.Fields(0) <> DevuelveValor("SELECT COUNT(*) FROM [##TMPQUINCTA_AUX" & VGL_COMPUTER & "]", DBSYSTEM) Then
        'MsgBox "!Falta declarar el C.Costo en la tabla C.Costo de Planillas"
        'Exit Sub
    End If
End If
'****************************************************
Set RSAUX = New ADODB.Recordset
RSAUX.Open "SELECT * FROM CONTQUICAB", DBSYSTEM, adOpenDynamic, adLockOptimistic
With RSAUX
    .AddNew
    .Fields("SUBDIAR_CODIGO") = Trim(REGSISTEMA.scSubdi)
    .Fields("CMOV_C_COMPR") = " "
    .Fields("CMOV_FECHA") = AplitxtMesTrabajo.Tag
    .Fields("CMOV_GLOSA") = Trim(Mid("Planillas " & LstvwCronograma.SelectedItem.Text, 1, 29))
    .Fields("CMOV_MONED") = "MN"
    .Fields("CMOV_CONVE") = DESCAM  'FORMA DE CAMBIO
    .Fields("CMOV_CAMES") = IIf(DESCAM = "ESP", 1 / TIPCAM, 0)
    .Fields("CMOV_FECCA") = AplitxtMesTrabajo.Tag
    .Fields("CMOV_TIPCA") = IIf(DESCAM = "VTA", TIPCAM, 0)
    .Fields("CMOV_DEBE") = Round(DevuelveValor("SELECT sum(MONTO) FROM [##TMPQUINCTA_AUX" & VGL_COMPUTER & "]  where TIPOCTA='D'", DBSYSTEM), 2)
    .Fields("CMOV_HABER") = Round(DevuelveValor("SELECT sum(MONTO) FROM [##TMPQUINCTA_AUX" & VGL_COMPUTER & "]  where TIPOCTA='H'", DBSYSTEM), 2)
    .Fields("CMOV_DEBUS") = Round(.Fields("CMOV_DEBE") * TIPCAM, 2)
    .Fields("CMOV_HABUS") = Round(.Fields("CMOV_HABER") * TIPCAM, 2)
    .Fields("CMOV_AUTOM") = 0
    .Fields("CMOV_COSTO") = 0
    .Fields("CMOV_CHEQU") = 0
    .Fields("CMOV_L_COMPR") = 0
    .Fields("CMOV_VENTA") = 0
    .Fields("CRONO") = LstvwCronograma.SelectedItem.Tag
.Update
End With

NUMCORRELATIVO = DevuelveValor("SELECT MAX(NUM) FROM CONTQUICAB", DBSYSTEM)
'If ExisteTablaAux("[ASIENTOCC" & VGL_COMPUTER & "]") Then DBSYSTEM.Execute "DROP TABLE [ASIENTOCC" & VGL_COMPUTER & "]"
'DBSYSTEM.Execute "SELECT * INTO [ASIENTOCC" & VGL_COMPUTER & "] FROM CONTQUIDET"
'DBSYSTEM.Execute "DELETE FROM [ASIENTOCC" & VGL_COMPUTER & "]"
'verificar si existe la cta contable y si esta amarrada a un ccosto Si/no
ProgressBar2.Value = 0
Me.Cls
Set RSAUX = New ADODB.Recordset
RSAUX.Open "SELECT DISTINCT(CUENTA) FROM [##TMPQUINCTA_AUX" & VGL_COMPUTER & "]", DBSYSTEM, adOpenDynamic, adLockOptimistic
ProgressBar1.Max = RSAUX.RecordCount
Do While Not RSAUX.EOF
    ProgressBar1.Value = ProgressBar1.Value + 1
    CONTADOR = CONTADOR + 1
    lblMessage.Caption = "Generando Registros para el Detalle :: Cuenta C. " & RSAUX!CUENTA
        ProgressBar2.Value = 0
        CADSQL = "SELECT CODTRAB,CONCEPTO,CUENTA,TIPOCTA,MONTO,CCOSTO FROM  [##TMPQUINCTA_AUX" & VGL_COMPUTER & "] WHERE CUENTA='" & RSAUX!CUENTA & "' AND TIPOCTA='D'"
        Set rsCTA = New ADODB.Recordset
        rsCTA.Open CADSQL, DBSYSTEM, adOpenDynamic, adLockOptimistic
        If Not rsCTA.EOF Then ProgressBar2.Max = rsCTA.RecordCount
        Do While Not rsCTA.EOF
            ProgressBar2.Value = ProgressBar2.Value + 1
            lblMessage2.Caption = "Registrando montos del Trabajador :" & rsCTA!CODTRAB
            Me.Cls
            Call GRABADETASIQUIN(rsCTA, "D", NUMCORRELATIVO, CONTADOR)
            rsCTA.MoveNext
            CONTADOR = CONTADOR + 1
        Loop
        ProgressBar2.Value = 0
        CADSQL = "SELECT CODTRAB,CONCEPTO,CUENTA,TIPOCTA,MONTO,CCOSTO FROM  [##TMPQUINCTA_AUX" & VGL_COMPUTER & "] WHERE CUENTA='" & RSAUX!CUENTA & "' AND TIPOCTA='H'"
        Set rsCTA = New ADODB.Recordset
        rsCTA.Open CADSQL, DBSYSTEM, adOpenDynamic, adLockOptimistic
        If Not rsCTA.EOF Then ProgressBar2.Max = rsCTA.RecordCount
        Do While Not rsCTA.EOF
            ProgressBar2.Value = ProgressBar2.Value + 1
            lblMessage2.Caption = "Registrando montos del Trabajador :" & rsCTA!CODTRAB
            Me.Cls
            Call GRABADETASIQUIN(rsCTA, "H", NUMCORRELATIVO, CONTADOR)
            rsCTA.MoveNext
            CONTADOR = CONTADOR + 1
        Loop
 RSAUX.MoveNext
Loop

'Dim RSAUX_DET As ADODB.Recordset
'
''AGRUPAR POR CTA Y (DEBE O HABER)
''AGRUPAR POR CTA Y (DEBE O HABER) Y CCOSTO
'Set RSAUX = New ADODB.Recordset
'RSAUX.Open "[ASIENTOCC" & VGL_COMPUTER & "]", DBSYSTEM, adOpenDynamic, adLockOptimistic
'Do While Not RSAUX.EOF
'    If ESNULO(RSAUX!DMOV_ANEXO, "") <> "" Then
'      Call GRABADETALLESQUIN(RSAUX, IIf(ESNULO(RSAUX!DMOV_DEBE, 0) = 0, "H", "D"))
'    Else
'        If ESNULO(RSAUX!DMOV_CENCO, "") = "" Then
'            'AGRUPA POR CTA
'            'DEBE
'            Set RSAUX_DET = New ADODB.Recordset
'            RSAUX_DET.Open "SELECT DISTINCT(DMOV_CUENT) FROM [ASIENTOCC" & VGL_COMPUTER & "] WHERE DMOV_DEBE=0", DBSYSTEM, adOpenDynamic, adLockOptimistic
'            Do While Not RSAUX_DET.EOF
'                If DevuelveValor("SELECT DMOV_CUENT FROM CONTQUIDET WHERE DMOV_CUENT='" & RSAUX_DET!DMOV_CUENT & "' AND NUM=" & numCorrelativo, DBSYSTEM) = "" Then
'                   DBSYSTEM.Execute "INSERT INTO CONTQUIDET  SELECT NUM,SUBDIAR_CODIGO,DMOV_C_COMPR,DMOV_SECUE=0,DMOV_FECHA,DMOV_CUENT,DMOV_ANEXO,DMOV_DOCUM='',DMOV_FECDC," & _
'                                    "DMOV_CENCO,DMOV_DEBE=SUM(DMOV_DEBE),DMOV_HABER=SUM(DMOV_HABER),DMOV_DEBUS=SUM(DMOV_DEBUS),DMOV_HABUS=SUM(DMOV_HABUS),DMOV_GLOSA,DMOV_CHEQU=0" & _
'                                    ",DMOV_AUTOM=0,DMOV_COSTO=0,DMOV_L_COMPR=0,DMOV_VENTA=0,DMOV_TRANS=0,DMOV_L_DESTI=0,DMOV_C_DESTI='',CRONO" & _
'                                    " FROM [ASIENTOCC" & VGL_COMPUTER & "] WHERE  DMOV_CUENT='" & RSAUX_DET!DMOV_CUENT & "' AND DMOV_DEBE=0 GROUP BY NUM,SUBDIAR_CODIGO,DMOV_C_COMPR,DMOV_FECHA,DMOV_CUENT,DMOV_ANEXO,DMOV_FECDC," & _
'                                    " DMOV_CENCO , DMOV_GLOSA, DMOV_C_DESTI, CRONO"
'                End If
'                RSAUX_DET.MoveNext
'            Loop
'            'HABER
'            Set RSAUX_DET = New ADODB.Recordset
'            RSAUX_DET.Open "SELECT DISTINCT(DMOV_CUENT) FROM [ASIENTOCC" & VGL_COMPUTER & "] WHERE DMOV_HABER=0", DBSYSTEM, adOpenDynamic, adLockOptimistic
'            Do While Not RSAUX_DET.EOF
'                If DevuelveValor("SELECT DMOV_CUENT FROM CONTQUIDET WHERE DMOV_CUENT='" & RSAUX_DET!DMOV_CUENT & "' AND NUM=" & numCorrelativo, DBSYSTEM) = "" Then
'                   DBSYSTEM.Execute "INSERT INTO CONTQUIDET  SELECT NUM,SUBDIAR_CODIGO,DMOV_C_COMPR,DMOV_SECUE=0,DMOV_FECHA,DMOV_CUENT,DMOV_ANEXO,DMOV_DOCUM='',DMOV_FECDC," & _
'                                    "DMOV_CENCO,DMOV_DEBE=SUM(DMOV_DEBE),DMOV_HABER=SUM(DMOV_HABER),DMOV_DEBUS=SUM(DMOV_DEBUS),DMOV_HABUS=SUM(DMOV_HABUS),DMOV_GLOSA,DMOV_CHEQU=0" & _
'                                    ",DMOV_AUTOM=0,DMOV_COSTO=0,DMOV_L_COMPR=0,DMOV_VENTA=0,DMOV_TRANS=0,DMOV_L_DESTI=0,DMOV_C_DESTI='',CRONO" & _
'                                    " FROM [ASIENTOCC" & VGL_COMPUTER & "] WHERE  DMOV_CUENT='" & RSAUX_DET!DMOV_CUENT & "' AND DMOV_HABER=0 GROUP BY NUM,SUBDIAR_CODIGO,DMOV_C_COMPR,DMOV_FECHA,DMOV_CUENT,DMOV_ANEXO,DMOV_FECDC," & _
'                                    " DMOV_CENCO , DMOV_GLOSA, DMOV_C_DESTI, CRONO"
'                End If
'                RSAUX_DET.MoveNext
'            Loop
'            'SE TERMINA TODO
'            Exit Do
'        Else
'            'AGRUPAR POR CTA Y CCOSTO
'
'        End If
'    End If
'RSAUX.MoveNext
'Loop
''*****************************************SE TERMINA TODO


End Sub
Private Sub GRABADETALLESQUIN(RSAUXX1 As ADODB.Recordset, TIPO As String)
    'CREANDO EL DETALLE DEL ASIENTO DE PLANILLAS
    With RSCONTDETQUIN
        .AddNew
        .Fields("NUM") = RSAUXX1!NUM
        .Fields("SUBDIAR_CODIGO") = REGSISTEMA.scSubdi
        .Fields("DMOV_C_COMPR") = " "
        '.Fields("DMOV_SECUE") = Format(SECUE, "0000")
        .Fields("DMOV_FECHA") = FechS(AplitxtMesTrabajo.Tag, Adof)
        .Fields("DMOV_CUENT") = ESNULO(RSAUXX1!DMOV_CUENT, " ")
        .Fields("DMOV_ANEXO") = ESNULO(RSAUXX1!DMOV_ANEXO, " ")
        .Fields("DMOV_DOCUM") = ESNULO(RSAUXX1!DMOV_DOCUM, "")
        .Fields("DMOV_FECDC") = FechS(AplitxtMesTrabajo.Tag, Adof)
        .Fields("DMOV_CENCO") = ESNULO(RSAUXX1!DMOV_CENCO, "")
        .Fields("DMOV_DEBE") = Round(IIf(TIPO = "D", RSAUXX1!DMOV_DEBE, 0), 2)
        .Fields("DMOV_HABER") = Round(IIf(TIPO = "H", RSAUXX1!DMOV_HABER, 0), 2)
        .Fields("DMOV_DEBUS") = Round(IIf(TIPO = "D", RSAUXX1!DMOV_DEBUS, 0), 2)
        .Fields("DMOV_HABUS") = Round(IIf(TIPO = "H", RSAUXX1!DMOV_HABUS, 0), 2)
        .Fields("DMOV_GLOSA") = " "
        .Fields("DMOV_CHEQU") = 0
        .Fields("DMOV_AUTOM") = 0
        .Fields("DMOV_COSTO") = RSAUXX1!DMOV_COSTO
        .Fields("DMOV_L_COMPR") = 0
        .Fields("DMOV_VENTA") = 0
        .Fields("DMOV_TRANS") = 0
        .Fields("DMOV_L_DESTI") = 0
        '.Fields("DMOV_C_DESTI") = CTADES
        .Fields("CRONO") = LstvwCronograma.SelectedItem.Tag
        .Update
    End With
End Sub

Private Sub AGRUPAR_CCOSTOQUIN(ByVal dCrono As Double, ByVal dNum As Double)
Dim sLadoDebe       As String, sLadoHAber   As String
Dim sSoloAnexos     As String, sCabecera    As String
Dim rsCabAuxQuin    As New ADODB.Recordset, rsDetAnexoAuxQuin   As New ADODB.Recordset
Dim rsLadoDebeQuin  As New ADODB.Recordset, rsLadoHAberAuxQuin  As New ADODB.Recordset
Dim dRsAfectados As Double

sCabecera = "SELECT CONTQUICAB.* FROM CONTQUICAB  Where CONTQUICAB.NUM =" & dNum
rsCabAuxQuin.Open sCabecera, DBSYSTEM, adOpenDynamic, adLockOptimistic

sSoloAnexos = "SELECT CONTQUIDET.* FROM CONTQUICAB INNER JOIN CONTQUIDET ON " & _
            "CONTQUICAB.NUM = CONTQUIDET.NUM WHERE CONTQUICAB.NUM=" & dNum & " AND DMOV_ANEXO<>'' "
rsDetAnexoAuxQuin.Open sSoloAnexos, DBSYSTEM, adOpenDynamic, adLockOptimistic
                
sLadoDebe = "SELECT  " & _
"CONTQUIDET.SUBDIAR_CODIGO,DMOV_C_COMPR='',DMOV_SECUE='',DMOV_FECHA,DMOV_CUENT,DMOV_ANEXO=''," & _
"DMOV_DOCUM,DMOV_FECDC,DMOV_CENCO,SUM(DMOV_DEBE) AS DMOV_DEBE ,SUM(DMOV_HABER) AS DMOV_HABER," & _
"SUM(DMOV_DEBUS) AS DMOV_DEBUS,SUM(DMOV_HABUS) AS DMOV_HABUS,DMOV_GLOSA,DMOV_CHEQU=0,DMOV_AUTOM=0," & _
"DMOV_COSTO=0,DMOV_L_COMPR=0,DMOV_VENTA=0,DMOV_TRANS=0,DMOV_L_DESTI=0,DMOV_C_DESTI=''," & _
"CONTQUIDET.CRONO FROM CONTQUICAB INNER JOIN CONTQUIDET ON CONTQUICAB.NUM = CONTQUIDET.NUM" & _
" WHERE CONTQUICAB.NUM=" & dNum & " AND DMOV_ANEXO='' AND DMOV_DEBE<>0 Group By CONTQUIDET.SUBDIAR_CODIGO," & _
"DMOV_FECHA,DMOV_CUENT,DMOV_DOCUM,DMOV_FECDC,DMOV_CENCO,DMOV_GLOSA,CONTQUIDET.CRONO"
rsLadoDebeQuin.Open sLadoDebe, DBSYSTEM, adOpenKeyset, adLockReadOnly

sLadoHAber = "SELECT " & _
"CONTQUIDET.SUBDIAR_CODIGO,DMOV_C_COMPR='',DMOV_SECUE='',DMOV_FECHA,DMOV_CUENT,DMOV_ANEXO=''," & _
"DMOV_DOCUM,DMOV_FECDC,DMOV_CENCO,SUM(DMOV_DEBE) AS DMOV_DEBE ,SUM(DMOV_HABER) AS DMOV_HABER," & _
"SUM(DMOV_DEBUS) AS DMOV_DEBUS,SUM(DMOV_HABUS) AS DMOV_HABUS,DMOV_GLOSA,DMOV_CHEQU=0,DMOV_AUTOM=0," & _
"DMOV_COSTO=0,DMOV_L_COMPR=0,DMOV_VENTA=0,DMOV_TRANS=0,DMOV_L_DESTI=0,DMOV_C_DESTI=''," & _
"CONTQUIDET.CRONO FROM CONTQUICAB INNER JOIN CONTQUIDET ON CONTQUICAB.NUM = CONTQUIDET.NUM" & _
" WHERE CONTQUICAB.NUM=" & dNum & " AND DMOV_ANEXO='' AND DMOV_HABER<>0 Group By CONTQUIDET.SUBDIAR_CODIGO," & _
"DMOV_FECHA,DMOV_CUENT,DMOV_DOCUM,DMOV_FECDC,DMOV_CENCO,DMOV_GLOSA,CONTQUIDET.CRONO"
rsLadoHAberAuxQuin.Open sLadoHAber, DBSYSTEM, adOpenKeyset, adLockReadOnly

DBSYSTEM.Execute "delete from CONTQUICAB where NUM=" & dNum
DBSYSTEM.Execute "delete from CONTQUIDET where NUM=" & dNum
Dim rsCAVMOVAUX As New ADODB.Recordset, rsDETMOVAUX As New ADODB.Recordset
Dim sSecuencias As String, dUltimoNum As Double
rsCAVMOVAUX.Open "CONTQUICAB", DBSYSTEM, adOpenDynamic, adLockOptimistic
rsDETMOVAUX.Open "CONTQUIDET", DBSYSTEM, adOpenDynamic, adLockOptimistic

rsCAVMOVAUX.AddNew
  rsCAVMOVAUX!SUBDIAR_CODIGO = rsCabAuxQuin!SUBDIAR_CODIGO
  rsCAVMOVAUX!CMOV_C_COMPR = rsCabAuxQuin!CMOV_C_COMPR & " "
  rsCAVMOVAUX!CMOV_FECHA = rsCabAuxQuin!CMOV_FECHA
  rsCAVMOVAUX!CMOV_GLOSA = rsCabAuxQuin!CMOV_GLOSA
  rsCAVMOVAUX!CMOV_MONED = rsCabAuxQuin!CMOV_MONED
  rsCAVMOVAUX!CMOV_CONVE = rsCabAuxQuin!CMOV_CONVE
  rsCAVMOVAUX!CMOV_CAMES = rsCabAuxQuin!CMOV_CAMES
  rsCAVMOVAUX!CMOV_FECCA = rsCabAuxQuin!CMOV_FECCA
  rsCAVMOVAUX!CMOV_TIPCA = rsCabAuxQuin!CMOV_TIPCA
  rsCAVMOVAUX!CMOV_DEBE = rsCabAuxQuin!CMOV_DEBE
  rsCAVMOVAUX!CMOV_HABER = rsCabAuxQuin!CMOV_HABER
  rsCAVMOVAUX!CMOV_DEBUS = rsCabAuxQuin!CMOV_DEBUS
  rsCAVMOVAUX!CMOV_HABUS = rsCabAuxQuin!CMOV_HABUS
  rsCAVMOVAUX!CMOV_AUTOM = rsCabAuxQuin!CMOV_AUTOM
  rsCAVMOVAUX!CMOV_COSTO = rsCabAuxQuin!CMOV_COSTO
  rsCAVMOVAUX!CMOV_CHEQU = rsCabAuxQuin!CMOV_CHEQU
  rsCAVMOVAUX!CMOV_L_COMPR = rsCabAuxQuin!CMOV_L_COMPR
  rsCAVMOVAUX!CMOV_VENTA = rsCabAuxQuin!CMOV_VENTA
  rsCAVMOVAUX!Crono = dCrono
rsCAVMOVAUX.Update
dUltimoNum = DevuelveValor("SELECT MAX(NUM) FROM  CONTQUICAB", DBSYSTEM)
    sSecuencias = 0
    Call RSGRABARQUIN_DET(rsDetAnexoAuxQuin, sSecuencias, dUltimoNum, rsDETMOVAUX)
    Call RSGRABARQUIN_DET(rsLadoDebeQuin, sSecuencias, dUltimoNum, rsDETMOVAUX)
    Call RSGRABARQUIN_DET(rsLadoHAberAuxQuin, sSecuencias, dUltimoNum, rsDETMOVAUX)
End Sub
Private Sub RSGRABARQUIN_DET(ByVal RSdequin As ADODB.Recordset, ByRef sSECUENCIA As String, _
ByVal dNum As Double, ByVal rsDetQuin As ADODB.Recordset)
Do While Not RSdequin.EOF
    sSECUENCIA = Format(CDbl(sSECUENCIA) + 1, "0000")
    rsDetQuin.AddNew
    rsDetQuin!NUM = dNum
    rsDetQuin!SUBDIAR_CODIGO = RSdequin!SUBDIAR_CODIGO
    rsDetQuin!DMOV_C_COMPR = RSdequin!DMOV_C_COMPR & " "
    rsDetQuin!DMOV_SECUE = sSECUENCIA
    rsDetQuin!DMOV_FECHA = RSdequin!DMOV_FECHA
    rsDetQuin!DMOV_CUENT = RSdequin!DMOV_CUENT
    rsDetQuin!DMOV_ANEXO = RSdequin!DMOV_ANEXO
    rsDetQuin!DMOV_DOCUM = RSdequin!DMOV_DOCUM
    rsDetQuin!DMOV_FECDC = RSdequin!DMOV_FECDC
    rsDetQuin!DMOV_CENCO = RSdequin!DMOV_CENCO
    rsDetQuin!DMOV_DEBE = RSdequin!DMOV_DEBE
    rsDetQuin!DMOV_HABER = RSdequin!DMOV_HABER
    rsDetQuin!DMOV_DEBUS = RSdequin!DMOV_DEBUS
    rsDetQuin!DMOV_HABUS = RSdequin!DMOV_HABUS
    rsDetQuin!DMOV_GLOSA = RSdequin!DMOV_GLOSA
    rsDetQuin!DMOV_CHEQU = RSdequin!DMOV_CHEQU
    rsDetQuin!DMOV_AUTOM = RSdequin!DMOV_AUTOM
    rsDetQuin!DMOV_COSTO = RSdequin!DMOV_COSTO
    rsDetQuin!DMOV_L_COMPR = RSdequin!DMOV_L_COMPR
    rsDetQuin!DMOV_VENTA = RSdequin!DMOV_VENTA
    rsDetQuin!DMOV_TRANS = RSdequin!DMOV_TRANS
    rsDetQuin!DMOV_L_DESTI = RSdequin!DMOV_L_DESTI
    rsDetQuin!DMOV_C_DESTI = RSdequin!DMOV_C_DESTI
    rsDetQuin!Crono = RSdequin!Crono
    rsDetQuin.Update
RSdequin.MoveNext
Loop
End Sub

