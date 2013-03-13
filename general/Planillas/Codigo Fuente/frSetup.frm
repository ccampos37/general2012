VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Begin VB.Form frSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración del Sistema"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   Icon            =   "frSetup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   6390
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmRestaura 
      Caption         =   "&Restaurar"
      Height          =   360
      Left            =   5100
      TabIndex        =   16
      Top             =   5235
      Width           =   1140
   End
   Begin VB.CommandButton cmCancela 
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   3870
      TabIndex        =   15
      Top             =   5235
      Width           =   1140
   End
   Begin VB.CommandButton cmAcepta 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   2640
      TabIndex        =   14
      Top             =   5235
      Width           =   1140
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   90
      TabIndex        =   0
      Top             =   75
      Width           =   6150
      _ExtentX        =   10848
      _ExtentY        =   8916
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      Tab             =   5
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frSetup.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame10"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Reportes"
      TabPicture(1)   =   "frSetup.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Planillas"
      TabPicture(2)   =   "frSetup.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame1"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "PDT Sunat"
      TabPicture(3)   =   "frSetup.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame6"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Seguridad"
      TabPicture(4)   =   "frSetup.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame11"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Frame8"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Avanzado"
      TabPicture(5)   =   "frSetup.frx":0396
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "Frame9"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      Begin VB.Frame Frame3 
         Caption         =   "Adelanto de Vacaciones"
         Height          =   1035
         Left            =   -74820
         TabIndex        =   97
         Top             =   3735
         Width           =   5715
         Begin AplisetControlText.Aplitext xAdelVac 
            Height          =   300
            Left            =   2880
            TabIndex        =   98
            Top             =   555
            Width           =   2730
            _ExtentX        =   4815
            _ExtentY        =   529
            Locked          =   -1  'True
            Text            =   ""
         End
         Begin VB.Label Label1 
            Caption         =   "Cargar el Neto de Vacaciones en la siguiente planilla dentro del Concepto como Adelanto de Vac."
            Height          =   630
            Index           =   5
            Left            =   195
            TabIndex        =   99
            Top             =   240
            Width           =   2430
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Programación del Mantenimiento del Sistema"
         Height          =   2130
         Left            =   -74820
         TabIndex        =   35
         Top             =   2730
         Width           =   5760
         Begin VB.CommandButton Command5 
            Caption         =   "Recuperar Backups"
            Height          =   300
            Left            =   3915
            TabIndex        =   3
            Top             =   1125
            Width           =   1710
         End
         Begin VB.ComboBox Combo3 
            Height          =   315
            ItemData        =   "frSetup.frx":03B2
            Left            =   2295
            List            =   "frSetup.frx":03D4
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1590
            Width           =   990
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "frSetup.frx":03F7
            Left            =   1260
            List            =   "frSetup.frx":041C
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1140
            Width           =   945
         End
         Begin VB.OptionButton xPer3 
            Caption         =   "Cada"
            Height          =   240
            Left            =   315
            TabIndex        =   9
            Top             =   1170
            Width           =   825
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "frSetup.frx":044C
            Left            =   1260
            List            =   "frSetup.frx":0465
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   750
            Width           =   2025
         End
         Begin VB.OptionButton xPer1 
            Caption         =   "Diariamente"
            Height          =   240
            Left            =   315
            TabIndex        =   37
            Top             =   390
            Width           =   1350
         End
         Begin VB.OptionButton xPer2 
            Caption         =   "Cada "
            Height          =   240
            Left            =   315
            TabIndex        =   38
            Top             =   780
            Width           =   765
         End
         Begin VB.CommandButton cmCompactar 
            Caption         =   "Mantenimiento Ahora"
            Height          =   345
            Left            =   3930
            TabIndex        =   39
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Número de Backups hist."
            Height          =   195
            Left            =   330
            TabIndex        =   10
            Top             =   1635
            Width           =   1785
         End
         Begin VB.Label Label24 
            Caption         =   "días"
            Height          =   240
            Left            =   2295
            TabIndex        =   13
            Top             =   1170
            Width           =   465
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "de la semana"
            Height          =   195
            Left            =   3360
            TabIndex        =   17
            Top             =   810
            Width           =   945
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Otras configuraciones"
         Height          =   1155
         Left            =   -74760
         TabIndex        =   40
         Top             =   3645
         Width           =   5715
         Begin VB.CheckBox xUsarCronograma 
            Caption         =   "Utilizar el cronograma de pagos en Areas de Trabajo/C.C."
            Height          =   210
            Left            =   345
            TabIndex        =   41
            Top             =   315
            Width           =   4575
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Sistema Personalizado"
         Height          =   4500
         Left            =   210
         TabIndex        =   65
         Top             =   405
         Width           =   5775
         Begin VB.CommandButton Command6 
            Caption         =   "MENU"
            Height          =   375
            Left            =   195
            TabIndex        =   96
            Top             =   3960
            Width           =   1575
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Analizador"
            Height          =   360
            Left            =   4035
            TabIndex        =   75
            Top             =   345
            Width           =   1545
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Utilizar ActiveX Data Object (ADO)"
            Enabled         =   0   'False
            Height          =   195
            Left            =   2865
            TabIndex        =   73
            Top             =   4125
            Value           =   -1  'True
            Width           =   2775
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Utilizar Remote Data Object (RDO)"
            Enabled         =   0   'False
            Height          =   195
            Left            =   2865
            TabIndex        =   72
            Top             =   3885
            Width           =   2775
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Utilizar Data Access Object (DAO)"
            Enabled         =   0   'False
            Height          =   195
            Left            =   2865
            TabIndex        =   71
            Top             =   3630
            Width           =   2775
         End
         Begin VB.CheckBox Check10 
            Caption         =   "Comprobar DLL API de Windows"
            Enabled         =   0   'False
            Height          =   195
            Left            =   2865
            TabIndex        =   70
            Top             =   3300
            Value           =   2  'Grayed
            Width           =   2640
         End
         Begin VB.CheckBox Check9 
            Caption         =   "Comprobar DLL Active X al iniciar"
            Enabled         =   0   'False
            Height          =   195
            Left            =   2865
            TabIndex        =   69
            Top             =   3030
            Value           =   2  'Grayed
            Width           =   2730
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Escribir Parámetros"
            Enabled         =   0   'False
            Height          =   360
            Left            =   210
            TabIndex        =   68
            Top             =   3465
            Width           =   1545
         End
         Begin VB.CommandButton Command2 
            Caption         =   "&Activar DLL"
            Enabled         =   0   'False
            Height          =   360
            Left            =   210
            TabIndex        =   67
            Top             =   3000
            Width           =   1545
         End
         Begin VB.ListBox List1 
            Enabled         =   0   'False
            Height          =   2010
            Left            =   210
            TabIndex        =   66
            Top             =   915
            Width           =   5370
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Componentes instalados"
            Height          =   195
            Left            =   210
            TabIndex        =   74
            Top             =   660
            Width           =   1725
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Seguridad de la Base de Datos"
         Height          =   2010
         Left            =   -74820
         TabIndex        =   57
         Top             =   585
         Width           =   5760
         Begin VB.CheckBox xcfg7 
            Caption         =   "Realizar la copia de seguridad en forma manual"
            Height          =   210
            Left            =   5310
            TabIndex        =   64
            Top             =   1860
            Visible         =   0   'False
            Width           =   3765
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Buscar"
            Height          =   330
            Left            =   4830
            TabIndex        =   63
            Top             =   1365
            Width           =   765
         End
         Begin VB.CheckBox xcfg6 
            Caption         =   "Comprobar la integridad de la copia de seguridad al finalizarla"
            Height          =   225
            Left            =   330
            TabIndex        =   59
            Top             =   690
            Width           =   5055
         End
         Begin VB.CheckBox xcfg5 
            Caption         =   "Realizar copia de seguridad como parte del mantenimiento"
            Height          =   210
            Left            =   330
            TabIndex        =   58
            Top             =   405
            Width           =   5160
         End
         Begin AplisetControlText.Aplitext xRuta 
            Height          =   300
            Left            =   1200
            TabIndex        =   95
            Top             =   1395
            Width           =   3540
            _ExtentX        =   6244
            _ExtentY        =   529
            Text            =   ""
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Ruta"
            Height          =   195
            Left            =   615
            TabIndex        =   61
            Top             =   1425
            Width           =   345
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Ubicación donde almacenar el archivo de copia de seguridad"
            Height          =   195
            Left            =   615
            TabIndex        =   60
            Top             =   1035
            Width           =   4350
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Reportes Personalizados"
         Height          =   4350
         Left            =   -74850
         TabIndex        =   51
         Top             =   525
         Width           =   5865
         Begin VB.ComboBox xClaseBoleta 
            Height          =   315
            ItemData        =   "frSetup.frx":04A5
            Left            =   3690
            List            =   "frSetup.frx":04B2
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   76
            ToolTipText     =   "Seleccione un tipo formato de impresión de boleta"
            Top             =   900
            Width           =   1980
         End
         Begin VB.CommandButton CmdSelRep 
            Caption         =   "&Seleccion de Reportes"
            Height          =   450
            Left            =   165
            TabIndex        =   42
            Top             =   2550
            Width           =   1995
         End
         Begin VB.CommandButton CmdListRep 
            Caption         =   "&Lista de los Reportes de todos los Usuarios"
            Height          =   570
            Left            =   4020
            TabIndex        =   46
            Top             =   3675
            Width           =   1740
         End
         Begin AplisetControlText.Aplitext xFilePlanilla 
            Height          =   285
            Left            =   2550
            TabIndex        =   77
            Top             =   1350
            Width           =   3120
            _ExtentX        =   5503
            _ExtentY        =   503
            Locked          =   -1  'True
            Text            =   ""
         End
         Begin AplisetControlText.Aplitext xFileBoleta 
            Height          =   285
            Left            =   2550
            TabIndex        =   78
            Top             =   525
            Width           =   3120
            _ExtentX        =   5503
            _ExtentY        =   503
            Locked          =   -1  'True
            Text            =   ""
         End
         Begin AplisetControlText.Aplitext XPlanCab 
            Height          =   300
            Left            =   2565
            TabIndex        =   79
            Top             =   1770
            Width           =   3120
            _ExtentX        =   5503
            _ExtentY        =   529
            Locked          =   -1  'True
            Text            =   ""
         End
         Begin VB.Label LblCabPlan 
            Caption         =   "Cabecera de Planilla"
            Height          =   240
            Left            =   210
            TabIndex        =   54
            Top             =   1845
            Width           =   2100
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Formato de Impresión Tipo:"
            Height          =   195
            Left            =   1665
            TabIndex        =   62
            Top             =   975
            Width           =   1920
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Planillas de Remuneraciones"
            Height          =   195
            Left            =   195
            TabIndex        =   53
            Top             =   1410
            Width           =   2040
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Boletas de Remuneraciones"
            Height          =   195
            Left            =   195
            TabIndex        =   52
            Top             =   585
            Width           =   1995
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Enlace de Planillas para el PDT Sunat"
         Height          =   4200
         Left            =   -74805
         TabIndex        =   26
         Top             =   600
         Width           =   5745
         Begin VB.CommandButton cmDirPDT 
            Caption         =   "..."
            Height          =   285
            Left            =   4680
            TabIndex        =   47
            Top             =   3780
            Width           =   390
         End
         Begin VB.CheckBox xPDTRunExe 
            Caption         =   "Ejecutar el PDT Sunat una vez culminada la exportación"
            Height          =   195
            Left            =   360
            TabIndex        =   44
            Top             =   3465
            Width           =   4470
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Ejecutar la tarea de Exportación desde el Cliente"
            Height          =   195
            Left            =   375
            TabIndex        =   27
            Top             =   435
            Value           =   2  'Grayed
            Width           =   3885
         End
         Begin AplisetControlText.Aplitext PDT1 
            Height          =   300
            Index           =   7
            Left            =   2730
            TabIndex        =   86
            Tag             =   "1"
            Top             =   3000
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   529
            Locked          =   -1  'True
            Text            =   "sumaSCTR"
            TipoCodigo      =   -1  'True
         End
         Begin AplisetControlText.Aplitext PDT1 
            Height          =   300
            Index           =   6
            Left            =   2730
            TabIndex        =   87
            Tag             =   "0"
            Top             =   2685
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   529
            Locked          =   -1  'True
            Text            =   ""
            TipoCodigo      =   -1  'True
         End
         Begin AplisetControlText.Aplitext PDT1 
            Height          =   300
            Index           =   5
            Left            =   2730
            TabIndex        =   88
            Tag             =   "1"
            Top             =   2370
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   529
            Locked          =   -1  'True
            Text            =   "SumaRenta"
            TipoCodigo      =   -1  'True
         End
         Begin AplisetControlText.Aplitext PDT1 
            Height          =   300
            Index           =   4
            Left            =   2730
            TabIndex        =   89
            Tag             =   "0"
            Top             =   2055
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   529
            Locked          =   -1  'True
            Text            =   ""
            TipoCodigo      =   -1  'True
         End
         Begin AplisetControlText.Aplitext PDT1 
            Height          =   300
            Index           =   3
            Left            =   2730
            TabIndex        =   90
            Tag             =   "1"
            Top             =   1740
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   529
            Locked          =   -1  'True
            Text            =   "SumaSalud"
            TipoCodigo      =   -1  'True
         End
         Begin AplisetControlText.Aplitext PDT1 
            Height          =   300
            Index           =   2
            Left            =   2730
            TabIndex        =   91
            Tag             =   "1"
            Top             =   1425
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   529
            Locked          =   -1  'True
            Text            =   "SumaAFP"
            TipoCodigo      =   -1  'True
         End
         Begin AplisetControlText.Aplitext PDT1 
            Height          =   300
            Index           =   1
            Left            =   2730
            TabIndex        =   92
            Tag             =   "1"
            Top             =   1110
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   529
            Locked          =   -1  'True
            Text            =   "SumaIES"
            TipoCodigo      =   -1  'True
         End
         Begin AplisetControlText.Aplitext PDT1 
            Height          =   300
            Index           =   0
            Left            =   2745
            TabIndex        =   93
            Tag             =   "1"
            Top             =   780
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   529
            Locked          =   -1  'True
            Text            =   "HorasTrab/8"
            TipoCodigo      =   -1  'True
         End
         Begin AplisetControlText.Aplitext xPDTPathExe 
            Height          =   285
            Left            =   2070
            TabIndex        =   94
            Top             =   3765
            Width           =   2610
            _ExtentX        =   4604
            _ExtentY        =   503
            Text            =   ""
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "(**)"
            Height          =   195
            Index           =   2
            Left            =   4275
            TabIndex        =   56
            Top             =   2760
            Width           =   210
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "(**)"
            Height          =   195
            Index           =   1
            Left            =   4275
            TabIndex        =   55
            Top             =   2123
            Width           =   210
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Directorio PDT00.exe"
            Height          =   195
            Left            =   345
            TabIndex        =   45
            Top             =   3825
            Width           =   1530
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Remuneración S.C.T.R."
            Height          =   195
            Left            =   360
            TabIndex        =   43
            Top             =   3075
            Width           =   1695
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Tributo Quinta Categoria"
            Height          =   195
            Left            =   360
            TabIndex        =   34
            Top             =   2760
            Width           =   1725
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Remuneración Quinta Categ."
            Height          =   195
            Left            =   360
            TabIndex        =   33
            Top             =   2440
            Width           =   2055
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Remuneración Artistas"
            Height          =   195
            Left            =   360
            TabIndex        =   32
            Top             =   2123
            Width           =   1590
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Remuneración Salud"
            Height          =   195
            Left            =   360
            TabIndex        =   31
            Top             =   1806
            Width           =   1485
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Remuneración Pensiones"
            Height          =   195
            Left            =   360
            TabIndex        =   30
            Top             =   1489
            Width           =   1815
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Remuneración IES"
            Height          =   195
            Left            =   360
            TabIndex        =   29
            Top             =   1172
            Width           =   1335
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Dias Trabajados"
            Height          =   195
            Left            =   360
            TabIndex        =   28
            Top             =   855
            Width           =   1155
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Sistema de Planillas"
         Height          =   1230
         Left            =   -74760
         TabIndex        =   21
         Top             =   630
         Width           =   5715
         Begin VB.CheckBox xcfg3 
            Caption         =   "Cada vez que inicie, verificar consistencia del Sistema"
            Height          =   270
            Left            =   330
            TabIndex        =   50
            Top             =   840
            Width           =   4635
         End
         Begin VB.CheckBox xcfg2 
            Caption         =   "Activar protección de Base de Datos"
            Height          =   195
            Left            =   330
            TabIndex        =   49
            Top             =   600
            Width           =   3270
         End
         Begin VB.CheckBox xCFG1 
            Caption         =   "Grabar un Archivo de Registro de Visitas"
            Height          =   195
            Left            =   330
            TabIndex        =   48
            Top             =   315
            Width           =   3210
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Panel de Trabajadores"
         Height          =   1650
         Left            =   -74760
         TabIndex        =   19
         Top             =   1935
         Width           =   5715
         Begin VB.CheckBox xcfg4 
            Caption         =   "Usar Código Autogenerado por el Sistema (ABC###)"
            Height          =   195
            Left            =   330
            TabIndex        =   20
            Top             =   300
            Width           =   4095
         End
         Begin VB.Label Label6 
            Caption         =   "### - Número Correlativo"
            Height          =   180
            Left            =   675
            TabIndex        =   25
            Top             =   1320
            Width           =   1965
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "C  -  Primera letra del Nombre"
            Height          =   195
            Left            =   675
            TabIndex        =   24
            Top             =   1080
            Width           =   2055
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "B  -  Primera letra del Apellido Materno"
            Height          =   195
            Left            =   675
            TabIndex        =   23
            Top             =   840
            Width           =   2685
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "A  -  Primera letra del Apellido Paterno"
            Height          =   195
            Left            =   675
            TabIndex        =   22
            Top             =   600
            Width           =   2655
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Adelantos de Remuneraciones"
         Height          =   945
         Left            =   -74820
         TabIndex        =   11
         Top             =   2670
         Width           =   5715
         Begin AplisetControlText.Aplitext xAdelanto 
            Height          =   315
            Left            =   2880
            TabIndex        =   85
            Top             =   405
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   556
            Locked          =   -1  'True
            Text            =   ""
         End
         Begin VB.Label Label1 
            Caption         =   "Cargar los adelantos de pago en la siguiente columna de planilla:"
            Height          =   420
            Index           =   4
            Left            =   180
            TabIndex        =   12
            Top             =   390
            Width           =   2580
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Planilla de Aportes a la AFP"
         Height          =   2085
         Left            =   -74820
         TabIndex        =   1
         Top             =   480
         Width           =   5715
         Begin AplisetControlText.Aplitext xAFPCruze 
            Height          =   315
            Left            =   2265
            TabIndex        =   80
            Top             =   1575
            Width           =   3270
            _ExtentX        =   5768
            _ExtentY        =   556
            Locked          =   -1  'True
            Text            =   ""
         End
         Begin AplisetControlText.Aplitext xAFPComi 
            Height          =   315
            Left            =   2265
            TabIndex        =   81
            Top             =   1245
            Width           =   3270
            _ExtentX        =   5768
            _ExtentY        =   556
            Locked          =   -1  'True
            Text            =   ""
         End
         Begin AplisetControlText.Aplitext xAFPSeg 
            Height          =   315
            Left            =   2265
            TabIndex        =   82
            Top             =   915
            Width           =   3270
            _ExtentX        =   5768
            _ExtentY        =   556
            Locked          =   -1  'True
            Text            =   ""
         End
         Begin AplisetControlText.Aplitext xAFPApor 
            Height          =   315
            Left            =   2265
            TabIndex        =   83
            Top             =   585
            Width           =   3270
            _ExtentX        =   5768
            _ExtentY        =   556
            Locked          =   -1  'True
            Text            =   ""
         End
         Begin AplisetControlText.Aplitext xAFPRA 
            Height          =   315
            Left            =   2265
            TabIndex        =   84
            Top             =   255
            Width           =   3270
            _ExtentX        =   5768
            _ExtentY        =   556
            Locked          =   -1  'True
            Text            =   "SumaAFP"
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Cruce con Planillas"
            Height          =   195
            Left            =   180
            TabIndex        =   18
            Top             =   1695
            Width           =   1350
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Remuneración Asegurable"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   8
            Top             =   360
            Width           =   1875
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Comisión % sobre R.A."
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   6
            Top             =   1350
            Width           =   1590
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Seguros"
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   4
            Top             =   1020
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Aportación Obligatoria"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   2
            Top             =   690
            Width           =   1560
         End
      End
   End
End
Attribute VB_Name = "frSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsEmp As New ADODB.Recordset
Dim Clase As ClassMenu
Private Sub cmAcepta_Click()
If RsEmp.RecordCount Then
    RsEmp.MoveFirst
Else
    DBSYSTEM.Execute "INSERT INTO EMPRESA (ALIAS) VALUES('')"
End If
    DBSYSTEM.Execute "UPDATE EMPRESA SET AFPAPOR='" & xAFPApor.Text & "',AFPCOMI='" & xAFPComi.Text & "',AFPPLAN='" & xAFPCruze.Text & "',AFPREMU='" & xAFPRA.Text & "',AFPSEG='" & xAFPSeg.Text & "',ADELPLAN='" & xAdelanto.Text & "',PDTARTISTA='" & PDT1(4).Text & "',PDTTRIBUTO='" & PDT1(6).Text & "'"
    DBSYSTEM.Execute "UPDATE EMPRESA SET PDTRUNEXE=" & xPDTRunExe.Value & ",PDTPATHEXE='" & xPDTPathExe.Text & "'"
    DBSYSTEM.Execute "UPDATE EMPRESA SET CFG0001=" & xCFG1.Value & ",CFG0002=" & xcfg2.Value & ",CFG0003=" & xcfg3.Value & ",CFG0004=" & xcfg4.Value & ",CFG0005=" & xcfg5.Value & ",CFG0006=" & xcfg6.Value & ",CFG0007=" & xcfg7.Value & ",USARCRONOGRAMA=" & xUsarCronograma.Value
    DBSYSTEM.Execute "UPDATE EMPRESA SET BKP_RUTA='" & xRuta.Text & "',BKP_PERIODO=" & IIf(xPer1, "1", IIf(xPer2, 2, 3)) & ",BKP_DIASSEMANA=" & Combo1.ListIndex & ",BKP_DIASTRANS=" & Combo2.Text & ",BKP_NUMBACKUPS=" & Combo3.ListIndex + 1
    DBSYSTEM.Execute "UPDATE EMPRESA SET ADELVAC='" & xAdelVac.Text & "'"
    Unload Me
End Sub

Private Sub CMCANCELA_Click()
    Unload Me
End Sub

Private Sub CMCOMPACTAR_Click()
    DBSYSTEM.Close
    If MsgBox("IMPORTANTE: ESTE PROCESO SOLO FUNCIONA CUANDO UD. ES EL UNICO USUARIO CONECTADO A LA BASE DE DATOS DEL SISTEMA. DESEA CONTINUAR", vbInformation + vbYesNo) = vbNo Then Exit Sub
    On Error GoTo ERRUSADO
    If UCase(Dir$(REGSISTEMA.PATHEMPRESA & "\PLANILLA.LDB")) = "PLANILLA.LDB" Then Kill REGSISTEMA.PATHEMPRESA & "\PLANILLA.LDB"
    CambiaPanelBD True
    On Error GoTo ERRNAME
    If Dir$(App.PATH & "\PL2.MDB") = "PL2.MDB" Then Kill App.PATH & "\PL2.MDB"
    DBEngine.CompactDatabase REGSISTEMA.PATHEMPRESA & "\PLANILLA.MDB", App.PATH & "\PL2.MDB", , dbEncrypt & ";USER=ADMIN;PWD=SEGURA"
    'COMPROBAR QUE SE HAYA CREADO BIEN LA BASE DE DATOS
    If Dir$(App.PATH & "\PL2.MDB") = "PL2.MDB" Then
        If Dir$(REGSISTEMA.PATHEMPRESA & "\PLANILLA.BAK") = "PLANILLA.BAK" Then Kill REGSISTEMA.PATHEMPRESA & "\PLANILLA.BAK"
        FileCopy REGSISTEMA.PATHEMPRESA & "\PLANILLA.MDB", REGSISTEMA.PATHEMPRESA & "\PLANILLA.BAK"
        Kill REGSISTEMA.PATHEMPRESA & "\PLANILLA.MDB"
        FileCopy App.PATH & "\PL2.MDB", REGSISTEMA.PATHEMPRESA & "\PLANILLA.MDB"
        Kill App.PATH & "\PL2.MDB"
    End If
    Screen.MousePointer = 1
    DBSYSTEM.Open
    cmAcepta.Enabled = False
    cmRestaura.Enabled = False
    CambiaPanelBD False
    Exit Sub
ERRNAME:
    MsgBox "NO SE HA PODIDO ABRIR LA BASE DE DATOS. CONSULTE AL ADMINISTRADOR DE RED PARA DERECHOS DE ELIMINACIÓN O ADICION DE ARCHIVOS", vbInformation
    MsgBox "HA OCURRIDO UN ERROR AL ABRIR DE NUEVO LA BASE DE DATOS. EL SISTEMA SE CERRARÁ. VUELVA A INGRESAR", vbCritical
    End
ERRUSADO:
    MsgBox "EL SOFTWARE DE PLANILLAS O LA BASE DE DATOS PRINCIPAL DEL SISTEMA, ESTÁ SIENDO USADA POR OTRO USUARIO", vbCritical
    DBSYSTEM.Open
    Exit Sub
End Sub

Private Sub CMDIRPDT_Click()
    frSelDir.Show 1
    If VPTAREA <> "" Then
        If Dir$(VPTAREA & IIf(Right(VPTAREA, 1) = "\", "", "\") & "PDT00.EXE") <> "PDT00.EXE" Then
            MsgBox "RUTA INCORRECTA. NO SE ENCUENTRA EN ESTA RUTA EL ARCHIVO EJECUTABLE DEL PDT SUNAT", vbCritical
            Exit Sub
        End If
        xPDTPathExe.Text = VPTAREA
    End If
End Sub

Private Sub CMDSELREP_Click()
   FrmTablRep.Show 1
End Sub

Private Sub CMRESTAURA_Click()
    CARGAVALORES
End Sub

Private Sub Command1_Click()
    frSelDir.Show 1
    If VPTAREA = "" Then Exit Sub
    xRuta.Text = VPTAREA
End Sub

Private Sub Command4_Click()
    MsgBox "SOLO PERMITIDO PARA EL PROGRAMADOR DEL SISTEMA", vbInformation
    If InputBox("ESCRIBA CLAVE DE ACCESO AL MODULO DEL PROGRAMADOR: ", "CLAVE DEL PROGRAMADOR", "FERNANDOCOSSIO") = (Day(Date) - Month(Date)) & "Nando" Then frAccess.Show 1 Else MsgBox "ERROR DE CLAVE. SE RECUERDA QUE ESTE MODULO SOLO ESTA DISPONIBLE PARA EL PROGRAMADOR", vbCritical
End Sub

Private Sub COMMAND6_Click()
    Screen.MousePointer = 11
    Set Clase = New ClassMenu
    If DBSTARPLAN.ConnectionString = "" Then
       MsgBox "No se ha encontrado activa la BD de configuración inicial del sistema", vbInformation
       Exit Sub
    End If
'    Set Cnx = New ADODB.Connection
'    With Cnx
'     .CursorLocation = adUseClient
'     .Provider = "Microsoft.Jet.OLEDB.3.51"
'     .ConnectionString = "Data Source=" & xCad
'     .Open
'    End With
    '/* Actualizando campo MEN_DESCRI */
    DBSTARPLAN.Execute "ALTER TABLE MENU ALTER COLUMN MEN_DESCRI VARCHAR(100)"
    
    '/*********************************/
    Set Clase.MDIMenu = MDIPrincipal
    Clase.TablaMenu = "MENU"
    Set Clase.Conexion = DBSTARPLAN
    Clase.CrearTablaMenu
 MsgBox "Se termino la creacion de menus"
 Screen.MousePointer = 1
End Sub

Private Sub Command5_Click()
Descomprimir.Show 1
End Sub

Private Sub Form_Load()
    Combo1.ListIndex = 0
    Combo2.ListIndex = 0
    Combo3.ListIndex = 0
    RsEmp.Open "SELECT * FROM EMPRESA", DBSYSTEM, adOpenDynamic, adLockOptimistic
    CARGAVALORES
    List1.AddItem "ADMINISTRADOR DE RESULTADOS DE PLANILLA VERSIÓN 3.0 05.11.2000"
    List1.AddItem "GENERADOR DE PAGOS POR BANCO 33.09.2000"
    List1.AddItem "ADMINISTRADOR DE OTROS DATOS INFORMATIVO VERSIÓN 1.0 10.11.2000"
    List1.AddItem "FILTRO DE AFP POR CENTRO DE COSTO 25.09.2000"
    List1.AddItem "GENERADOR DE PDT SUNAT VERSION 2000 02.04.2000"
    List1.AddItem "GENERADOR DE FOTOCHECKS, FICHA DE PERSONAL VERSIÓN 1.0 01.04.2000"
    List1.AddItem "FORMATOS DE IMPRESION DE BOLETAS V.1 03.10.2000"
    List1.AddItem "ADMINISTRADOR DE BENEFICIOS SOCIALES V. 1.0 04.11.2000"
    List1.AddItem "CONTROLADOR DE VERSIONES ANTERIORES V. 2.0 04.12.1999"
    'MOSTRAR SIEMPRE EL REPORTE ACTIVO
    Call MOSTRARREPACTIVO
    
End Sub
Public Sub MOSTRARREPACTIVO()
    Dim RS2 As New ADODB.Recordset
    If REGSISTEMA.USER = "INVITADO" Then
        frSetup.CmdSelRep.Enabled = False
      Else: frSetup.CmdSelRep.Enabled = True
    End If
    RS2.Open "SELECT * FROM TABLREP WHERE CODEMPRESA='" & REGSISTEMA.RUC & "' AND CODUSU='" & REGSISTEMA.USER & "' AND ACTIVO=-1", DBSTARPLAN
    If RS2.RecordCount = 0 Then
        xFileBoleta.Text = ""
        xFilePlanilla.Text = ""
        XPlanCab.Text = ""
        Exit Sub
    End If
    xFileBoleta.Text = RS2!FILEBOLETA
    xFilePlanilla.Text = RS2!FILEPLANILLA
    XPlanCab.Text = RS2!FILEPLANCAB
    xClaseBoleta.ListIndex = RS2!TipFmt
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RsEmp = Nothing
End Sub

Private Sub PDT1_DblClick(INDEX As Integer)
    If PDT1(INDEX).Tag = 1 Then Exit Sub
    If Columna() <> "" Then
        PDT1(INDEX).Text = VGUTIL(1)
    End If
End Sub

Private Sub PDT1_KeyPress(INDEX As Integer, KeyAscii As Integer)
    If PDT1(INDEX).Tag = 1 Then Exit Sub
    If KeyAscii = 32 Then
        PDT1(INDEX).Text = ""
    End If
End Sub

Private Sub XADELANTO_DblClick()
    Dim RSRUBROS As New ADODB.Recordset
    RSRUBROS.Open "SELECT CODIGO, NOMBRE FROM COLUMPL WHERE TIPO=3 ORDER BY NOMBRE", DBSYSTEM, adOpenStatic
    frmComun.CONECTAR RSRUBROS
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xAdelanto.Text = VGUTIL(1)
    End If
    Set RSRUBROS = Nothing
End Sub

Private Sub xAdelVac_Click()
    Dim RSRUBROS As New ADODB.Recordset
    RSRUBROS.Open "SELECT CODIGO, NOMBRE FROM CONCEPTOS WHERE TIPO=2 ORDER BY NOMBRE", DBSYSTEM, adOpenStatic
    frmComun.CONECTAR RSRUBROS
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xAdelVac.Text = VGUTIL(1)
    End If
    Set RSRUBROS = Nothing
End Sub

Private Sub xAdelVac_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then xAdelVac.Text = ""
End Sub

Private Sub XAFPAPOR_DblClick()
    Dim RSRUBROS As New ADODB.Recordset
    RSRUBROS.Open "SELECT CODIGO, NOMBRE FROM CONCEPTOS WHERE TIPO=2 ORDER BY NOMBRE", DBSYSTEM, adOpenStatic
    frmComun.CONECTAR RSRUBROS
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xAFPApor.Text = VGUTIL(1)
    End If
    Set RSRUBROS = Nothing
End Sub

Private Sub XAFPCOMI_DblClick()
    Dim RSRUBROS As New ADODB.Recordset
    RSRUBROS.Open "SELECT CODIGO, NOMBRE FROM CONCEPTOS WHERE TIPO=2 ORDER BY NOMBRE", DBSYSTEM, adOpenStatic
    frmComun.CONECTAR RSRUBROS
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xAFPComi.Text = VGUTIL(1)
    End If
    Set RSRUBROS = Nothing
End Sub

Private Sub XAFPCRUZE_DblClick()
    Dim RSRUBROS As New ADODB.Recordset
    RSRUBROS.Open "SELECT CODIGO, NOMBRE FROM COLUMPL WHERE TIPO=3 ORDER BY NOMBRE", DBSYSTEM, adOpenStatic
    frmComun.CONECTAR RSRUBROS
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xAFPCruze.Text = VGUTIL(1)
    End If
    Set RSRUBROS = Nothing
End Sub

Private Sub XAFPRA_DblClick()
    MsgBox "ESTE RUBRO ES DE SOLO LECTURA. ESTÁ ESPECIFICADO EN SUMAAFP Y NO PUEDE SER CAMBIADO", vbCritical
End Sub

Private Sub XAFPSEG_DblClick()
    Dim RSRUBROS As New ADODB.Recordset
    RSRUBROS.Open "SELECT CODIGO, NOMBRE FROM CONCEPTOS WHERE TIPO=2 ORDER BY NOMBRE", DBSYSTEM, adOpenStatic
    frmComun.CONECTAR RSRUBROS
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xAFPSeg.Text = VGUTIL(1)
    End If
    Set RSRUBROS = Nothing
End Sub

Public Sub CARGAVALORES()
    On Error GoTo CONTINUAR
    With RsEmp
        xAFPApor.Text = "" & !AFPAPOR
        xAFPComi.Text = "" & !AFPCOMI
        xAFPCruze.Text = "" & !AFPPLAN
        xAFPRA.Text = "" & !AFPREMU
        xAFPSeg.Text = "" & !AFPSEG
        xAdelanto.Text = "" & !ADELPLAN
        PDT1(4).Text = "" & !PDTARTISTA
        PDT1(6).Text = "" & !PDTTRIBUTO
        xPDTRunExe.Value = !PDTRUNEXE
        xPDTPathExe.Text = "" & !PDTPATHEXE
        xUsarCronograma.Value = IIf(!USARCRONOGRAMA, 1, 0)
        xCFG1.Value = IIf(!CFG0001, 1, 0)
        xcfg2.Value = IIf(!CFG0002, 1, 0)
        xcfg3.Value = IIf(!CFG0003, 1, 0)
        xcfg4.Value = IIf(!CFG0004, 1, 0)
        xcfg5.Value = IIf(!CFG0005, 1, 0)
        xcfg6.Value = IIf(!CFG0006, 1, 0)
        xcfg7.Value = IIf(!CFG0007, 1, 0)
        xRuta.Text = "" & !BKP_RUTA
        xAdelVac.Text = "" & !ADELVAC
        Select Case !BKP_PERIODO
            Case 1
                xPer1.Value = True
            Case 2
                xPer2.Value = True
            Case 3
                xPer3.Value = True
            Case Else
                xPer1.Value = True
        End Select
        Combo1.ListIndex = IIf(IsNull(!BKP_DIASSEMANA), 0, !BKP_DIASSEMANA)
        Combo2.Text = IIf(IsNull(!BKP_DIASTRANS), 0, !BKP_DIASTRANS)
        Combo3.Text = IIf(IsNull(!BKP_NUMBACKUPS), 0, !BKP_NUMBACKUPS)
    End With
    Exit Sub
CONTINUAR:
    Resume Next
End Sub

Public Function Columna() As String
    Dim RSAUX As ADODB.Recordset
    Set RSAUX = Nothing
    Set RSAUX = New ADODB.Recordset
    RSAUX.Open "SELECT CODIGO, NOMBRE FROM COLUMPL ORDER BY NOMBRE", DBSYSTEM, adOpenStatic
    If RSAUX.RecordCount = 0 Then
        MsgBox "NO SE HAN ENCONTRADO COLUMNAS DE PLANILLA PARA PODERLAS ENLAZAR CON EL RESULTADO DE LA EXPORTACIÓN DEL PDT SUNAT", vbCritical
        Set RSAUX = Nothing
        Columna = ""
        Exit Function
    End If
    frmComun.CONECTAR RSAUX
    frmComun.Show 1
    Columna = VGUTIL(1)
End Function

