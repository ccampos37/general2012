VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form Frmcliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proveedores"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   9930
   Begin MSComctlLib.StatusBar Panel 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   74
      Top             =   6315
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   15875
            MinWidth        =   15875
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6105
      Left            =   90
      TabIndex        =   37
      Top             =   60
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   10769
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      WordWrap        =   0   'False
      TabCaption(0)   =   "Consulta"
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "TDBGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frmbotones"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Descripcion"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Fr1(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Representante"
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Fr3(0)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame2 
         Caption         =   "B�squeda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   1
         Left            =   -74400
         TabIndex        =   75
         Top             =   480
         Width           =   8595
         Begin VB.ComboBox cmbBusqueda 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   480
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   360
            Width           =   2415
         End
         Begin VB.TextBox txtBusqueda 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3000
            TabIndex        =   1
            Top             =   360
            Width           =   3855
         End
         Begin VB.CommandButton cmdAceptar 
            Caption         =   "&Aceptar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7080
            TabIndex        =   2
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame frmbotones 
         Height          =   585
         Left            =   -73320
         TabIndex        =   70
         Top             =   5340
         Width           =   5970
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Salir"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   4
            Left            =   4725
            Style           =   1  'Graphical
            TabIndex        =   81
            Top             =   180
            Width           =   1170
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Imprimir"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   13
            Left            =   3570
            Style           =   1  'Graphical
            TabIndex        =   77
            Top             =   180
            Width           =   1170
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Eliminar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   2
            Left            =   2415
            Style           =   1  'Graphical
            TabIndex        =   80
            Top             =   180
            Width           =   1170
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "E&ditar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   1260
            Style           =   1  'Graphical
            TabIndex        =   79
            Top             =   180
            Width           =   1170
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Nuevo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   105
            Style           =   1  'Graphical
            TabIndex        =   78
            Top             =   180
            Width           =   1170
         End
      End
      Begin VB.Frame Fr3 
         Height          =   5040
         Index           =   0
         Left            =   360
         TabIndex        =   62
         Top             =   720
         Width           =   9015
         Begin VB.Frame Frame4 
            Height          =   555
            Left            =   2640
            TabIndex        =   72
            Top             =   4335
            Width           =   4110
            Begin VB.CommandButton cmdBotones 
               Caption         =   "&Cancelar"
               Height          =   330
               Index           =   12
               Left            =   2670
               Style           =   1  'Graphical
               TabIndex        =   35
               Top             =   165
               Width           =   1335
            End
            Begin VB.CommandButton cmdBotones 
               Caption         =   "&Acepta"
               Height          =   330
               Index           =   11
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   34
               Top             =   165
               Width           =   1335
            End
            Begin VB.CommandButton cmdBotones 
               Caption         =   "&Anterior"
               Height          =   330
               Index           =   7
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   33
               Top             =   165
               Width           =   1335
            End
         End
         Begin VB.Frame Fr3 
            Height          =   3975
            Index           =   1
            Left            =   240
            TabIndex        =   63
            Top             =   240
            Width           =   8535
            Begin MSMask.MaskEdBox MBox 
               Height          =   255
               Index           =   14
               Left            =   1680
               TabIndex        =   27
               Top             =   720
               Width           =   6495
               _ExtentX        =   11456
               _ExtentY        =   450
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   255
               Index           =   15
               Left            =   1680
               TabIndex        =   28
               Top             =   1320
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   8
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   255
               Index           =   16
               Left            =   4920
               TabIndex        =   29
               Top             =   1320
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   11
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   255
               Index           =   17
               Left            =   1680
               TabIndex        =   30
               Top             =   1875
               Width           =   6495
               _ExtentX        =   11456
               _ExtentY        =   450
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   255
               Index           =   18
               Left            =   1680
               TabIndex        =   31
               Top             =   2400
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   25
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   255
               Index           =   19
               Left            =   5520
               TabIndex        =   32
               Top             =   2400
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   3
               PromptChar      =   "_"
            End
            Begin VB.Label Etiq 
               Caption         =   "Cod. Postal"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   27
               Left            =   4080
               TabIndex        =   69
               Top             =   2400
               Width           =   1215
            End
            Begin VB.Label Etiq 
               Caption         =   "Telefono"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   26
               Left            =   240
               TabIndex        =   68
               Top             =   2400
               Width           =   975
            End
            Begin VB.Label Etiq 
               Caption         =   "Reg.Unico"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   25
               Left            =   3720
               TabIndex        =   67
               Top             =   1305
               Width           =   1575
            End
            Begin VB.Label Etiq 
               Caption         =   "Direccion"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   24
               Left            =   240
               TabIndex        =   66
               Top             =   1875
               Width           =   975
            End
            Begin VB.Label Etiq 
               Caption         =   "Doc. Identidad"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   16
               Left            =   240
               TabIndex        =   65
               Top             =   1305
               Width           =   1575
            End
            Begin VB.Label Etiq 
               Caption         =   "Propietario"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   22
               Left            =   240
               TabIndex        =   64
               Top             =   735
               Width           =   975
            End
         End
      End
      Begin VB.Frame Fr1 
         Height          =   5670
         Index           =   0
         Left            =   -74760
         TabIndex        =   38
         Top             =   330
         Width           =   9255
         Begin Crystal.CrystalReport oCrystalReport 
            Left            =   135
            Top             =   4860
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            PrintFileLinesPerPage=   60
         End
         Begin VB.Frame Frame1 
            Height          =   570
            Left            =   3480
            TabIndex        =   71
            Top             =   5025
            Width           =   2970
            Begin VB.CommandButton cmdBotones 
               Caption         =   "&Cancelar"
               Height          =   315
               Index           =   6
               Left            =   1545
               Style           =   1  'Graphical
               TabIndex        =   26
               Top             =   180
               Width           =   1320
            End
            Begin VB.CommandButton cmdBotones 
               Caption         =   "&Siguiente"
               Height          =   300
               Index           =   5
               Left            =   180
               Style           =   1  'Graphical
               TabIndex        =   25
               Top             =   180
               Width           =   1260
            End
         End
         Begin VB.Frame Frame3 
            Height          =   1575
            Left            =   240
            TabIndex        =   52
            Top             =   3390
            Width           =   8775
            Begin VB.CommandButton cAyuda 
               Caption         =   "..."
               Height          =   375
               Index           =   3
               Left            =   8040
               TabIndex        =   22
               Top             =   600
               Width           =   285
            End
            Begin MSComCtl2.DTPicker DTP_FecAnu 
               Height          =   300
               Left            =   4695
               TabIndex        =   21
               Top             =   720
               Width           =   1620
               _ExtentX        =   2858
               _ExtentY        =   529
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   24641537
               CurrentDate     =   37621
            End
            Begin MSComCtl2.DTPicker DTP_FecAct 
               Height          =   300
               Left            =   1680
               TabIndex        =   20
               Top             =   705
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   529
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   24641537
               CurrentDate     =   37480
            End
            Begin VB.ComboBox Combo7 
               Height          =   315
               Left            =   2280
               Style           =   2  'Dropdown List
               TabIndex        =   23
               Top             =   1120
               Width           =   1095
            End
            Begin VB.ComboBox Combo5 
               Height          =   315
               Left            =   7815
               Style           =   2  'Dropdown List
               TabIndex        =   19
               Top             =   225
               Width           =   855
            End
            Begin VB.ComboBox Combo4 
               Height          =   315
               Left            =   6465
               Style           =   2  'Dropdown List
               TabIndex        =   18
               Top             =   225
               Width           =   735
            End
            Begin VB.ComboBox Combo3 
               Height          =   315
               Left            =   3480
               Style           =   2  'Dropdown List
               TabIndex        =   17
               Top             =   240
               Width           =   1770
            End
            Begin VB.ComboBox Combo2 
               Height          =   315
               Left            =   1305
               Style           =   2  'Dropdown List
               TabIndex        =   16
               Top             =   240
               Width           =   1575
            End
            Begin MSMask.MaskEdBox MBox 
               BeginProperty DataFormat 
                  Type            =   0
                  Format          =   "0%"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
               Height          =   255
               Index           =   21
               Left            =   4800
               TabIndex        =   24
               Top             =   1200
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   450
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin VB.Label Etiq 
               Caption         =   "% Descuento"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   3
               Left            =   3480
               TabIndex        =   76
               Top             =   1200
               Width           =   1215
            End
            Begin VB.Label Etiq 
               Caption         =   "M�ximo Dias de Pago"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   21
               Left            =   120
               TabIndex        =   61
               Top             =   1150
               Width           =   2055
            End
            Begin VB.Label Etiq 
               Caption         =   "Mult.Direccion"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   18
               Left            =   6600
               TabIndex        =   60
               Top             =   720
               Width           =   1335
            End
            Begin VB.Label Etiq 
               Caption         =   "Aval"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   20
               Left            =   7320
               TabIndex        =   59
               Top             =   240
               Width           =   495
            End
            Begin VB.Label Etiq 
               Caption         =   "Fec.Baja"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   19
               Left            =   3600
               TabIndex        =   58
               Top             =   720
               Width           =   855
            End
            Begin VB.Label Etiq 
               Caption         =   "Fec.Activaci�n"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   17
               Left            =   120
               TabIndex        =   57
               Top             =   720
               Width           =   1335
            End
            Begin VB.Label Etiq 
               Caption         =   "Suspendido"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   15
               Left            =   5310
               TabIndex        =   56
               Top             =   240
               Width           =   1155
            End
            Begin VB.Label Etiq 
               Caption         =   "Pais"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   14
               Left            =   3000
               TabIndex        =   55
               Top             =   255
               Width           =   615
            End
            Begin VB.Label Etiq 
               Caption         =   "Personer�a"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   5
               Left            =   120
               TabIndex        =   54
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.Frame Frame2 
            Height          =   2055
            Index           =   0
            Left            =   240
            TabIndex        =   45
            Top             =   1320
            Width           =   8775
            Begin MSMask.MaskEdBox MBox 
               Height          =   255
               Index           =   5
               Left            =   1440
               TabIndex        =   8
               Top             =   240
               Width           =   7095
               _ExtentX        =   12515
               _ExtentY        =   450
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   255
               Index           =   6
               Left            =   7680
               TabIndex        =   10
               Top             =   580
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   3
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   255
               Index           =   7
               Left            =   1440
               TabIndex        =   11
               Top             =   960
               Width           =   3495
               _ExtentX        =   6165
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   30
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   255
               Index           =   8
               Left            =   5520
               TabIndex        =   12
               Top             =   960
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   30
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   255
               Index           =   9
               Left            =   1440
               TabIndex        =   13
               Top             =   1320
               Width           =   3495
               _ExtentX        =   6165
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   25
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   255
               Index           =   10
               Left            =   5520
               TabIndex        =   14
               Top             =   1320
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   10
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   255
               Index           =   11
               Left            =   1920
               TabIndex        =   15
               Top             =   1680
               Width           =   6615
               _ExtentX        =   11668
               _ExtentY        =   450
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   255
               Index           =   20
               Left            =   1440
               TabIndex        =   9
               Top             =   600
               Width           =   4935
               _ExtentX        =   8705
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   30
               PromptChar      =   "_"
            End
            Begin VB.Label Etiq 
               Caption         =   "Distrito"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   23
               Left            =   120
               TabIndex        =   73
               Top             =   580
               Width           =   1095
            End
            Begin VB.Label Etiq 
               Caption         =   "Cod. Postal"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   13
               Left            =   6480
               TabIndex        =   53
               Top             =   580
               Width           =   1095
            End
            Begin VB.Label Etiq 
               Caption         =   "Correo Electronico"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   11
               Left            =   120
               TabIndex        =   51
               Top             =   1680
               Width           =   2055
            End
            Begin VB.Label Etiq 
               Caption         =   "Direccion"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   7
               Left            =   120
               TabIndex        =   50
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Etiq 
               Caption         =   "Provincia"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   6
               Left            =   120
               TabIndex        =   49
               Top             =   960
               Width           =   1095
            End
            Begin VB.Label Etiq 
               Caption         =   "Dpto"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   8
               Left            =   5040
               TabIndex        =   48
               Top             =   960
               Width           =   495
            End
            Begin VB.Label Etiq 
               Caption         =   "Telefono"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   9
               Left            =   120
               TabIndex        =   47
               Top             =   1320
               Width           =   1095
            End
            Begin VB.Label Etiq 
               Caption         =   "Fax"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   10
               Left            =   5040
               TabIndex        =   46
               Top             =   1320
               Width           =   1095
            End
         End
         Begin VB.Frame Fr1 
            Height          =   1095
            Index           =   1
            Left            =   240
            TabIndex        =   39
            Top             =   120
            Width           =   8775
            Begin MSMask.MaskEdBox MBox 
               Height          =   255
               Index           =   0
               Left            =   1680
               TabIndex        =   3
               Top             =   240
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   11
               PromptChar      =   "_"
            End
            Begin VB.ComboBox Combo1 
               Height          =   315
               Left            =   6600
               Style           =   2  'Dropdown List
               TabIndex        =   5
               Top             =   240
               Width           =   1935
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   255
               Index           =   1
               Left            =   3960
               TabIndex        =   4
               Top             =   240
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   11
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   255
               Index           =   2
               Left            =   1680
               TabIndex        =   6
               Top             =   600
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   60
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   255
               Index           =   3
               Left            =   5880
               TabIndex        =   7
               Top             =   600
               Width           =   2655
               _ExtentX        =   4683
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   20
               PromptChar      =   "_"
            End
            Begin VB.Label Etiq 
               Caption         =   "Negocio"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   12
               Left            =   5640
               TabIndex        =   44
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Etiq 
               Caption         =   "Siglas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   4
               Left            =   5160
               TabIndex        =   43
               Top             =   600
               Width           =   1095
            End
            Begin VB.Label Etiq 
               Caption         =   "Razon Social"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   42
               Top             =   600
               Width           =   1215
            End
            Begin VB.Label Etiq 
               Caption         =   "R.u.c"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   3240
               TabIndex        =   41
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Etiq 
               Caption         =   "Codigo"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   40
               Top             =   240
               Width           =   1095
            End
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   3735
         Left            =   -74520
         TabIndex        =   36
         Top             =   1560
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   6588
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=29,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
Attribute VB_Name = "Frmcliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nListacampo As String
Dim nLongicampo(3) As Integer
Dim adll As New dllgeneral.dll_general
'''' Busqueda
Dim ArregloBusqueda()
Dim i_indexComboBusqueda As Integer
Dim modoinsert, modoedit As Boolean
Dim i_codigocliente As String
'Picture

Private Sub cAyuda_Click(Index As Integer)
    If Not (Trim(MBox(0)) = "") Then
        FrmMultidireccion.codigo = MBox(0)
        FrmMultidireccion.razon = MBox(2)
        ''''
        If modoinsert = True Then
            FrmMultidireccion.MODONUEVO = True
        ElseIf modoedit = True Then
            FrmMultidireccion.MODONUEVO = False
        End If
        ''''
        FrmMultidireccion.Show
        FrmMultidireccion.SetFocus
    Else
        MsgBox "Ingrese C�digo del Proveedor", vbCritical, "AVISO"
    End If
End Sub
Private Sub cmbBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmdAceptar_Click()
     Call fncBusqueda(VGcnxCP, TDBGrid1)
End Sub

Private Sub cmdBotones_Click(Index As Integer)
   Dim lcondi As String
   Dim OBJ As Object
   Dim SQL As String

   Select Case Index
    Case 0
        adll.ActivaTab 1, 2, SSTab1
        'Limpiartexto MBox, 0, 19, 12, 13
        
        'MBox(0).Enabled = True: MBox(1).Enabled = True
        
        For Each OBJ In Me.Controls
                If TypeOf OBJ Is ComboBox Then
                    OBJ.ListIndex = -1
                End If
                If TypeOf OBJ Is MaskEdBox Then
                    OBJ.Text = ""
                End If
        Next
        'VALORES DEFAULT:
        Combo4.ListIndex = 1 'SUSPENDIDO
        DTP_FecAct.Value = Date 'FECHA ACTIVACION
        DTP_FecAnu.Value = Date
        DTP_FecAnu.Value = ""  'FECHA BAJA
        ''
        modoinsert = True
        ''
        SQL = "DELETE FROM TEMPO_proveedordireccion"
        VGcnxCP.Execute SQL
        ''
        MBox(0).SetFocus
    Case 1
       If TDBGrid1.Row >= 0 Then
          adll.ActivaTab 1, 2, SSTab1
          CargaData
          'MBox(0).Enabled = False: MBox(1).Enabled = False
          modoedit = True
          i_codigocliente = Trim(TDBGrid1.Columns.Item(0).Text)
          MBox(0).SetFocus
       Else
          MsgBox MsgEdit, vbInformation, MsgTitle
       End If
    Case 2   'Boton Eliminar
       If TDBGrid1.Row >= 0 Then
          If MsgBox("Desea eliminar el registro?", vbYesNo + vbDefaultButton2, "AVISO") = vbYes Then
                If Eliminar_Cliente(VGcnxCP) = True Then
                    VGcnxCP.CommitTrans
                    Call fncBusqueda(VGcnxCP, TDBGrid1)
                Else
                    VGcnxCP.RollbackTrans
                End If
          End If
        End If
    
    Case 3   'Boton Busqueda
       'FrmBuscar.Show 1
       'If Len(Trim(Cadenabusca)) > 0 Then
       '  adll.ListarEnTDBGRID VGCNXCP, "cp_proveedor", TDBGrid1, nListacampo, "clientecodigo", nLongicampo, "clientecodigo='" & Cadenabusca & "'"
       'Else
       '   Listado
       'End If
       'Cadenabusca = ""
    Case 4
       Unload Me
    Case 5    ' Boton siguiente
        adll.ActivaTab 2, 2, SSTab1
        MBox(14).SetFocus
    Case 6, 12
        adll.ActivaTab 0, 2, SSTab1
        modoinsert = False
        modoedit = False
    Case 7
        adll.ActivaTab 1, 2, SSTab1
    'Case 8
        'Listado
    Case 11
        If modoinsert = True Then
            If Verificar_Codigo(0) = False Then
                    'MsgBox MsgAdd, vbInformation, MsgTitle
                    MsgBox "El C�digo ya existe...VERIFIQUE", vbInformation, MsgTitle
                    Exit Sub
            End If
        ElseIf modoedit = True Then
            If Verificar_Codigo(1) = False Then
                    'MsgBox MsgAdd, vbInformation, MsgTitle
                    MsgBox "El C�digo ya existe...VERIFIQUE", vbInformation, MsgTitle
                    Exit Sub
            End If
        End If
    
        GrabaCliente
        MsgBox MsgGraba, vbInformation, MsgTitle
        adll.ActivaTab 0, 2, SSTab1
        ''
        modoinsert = False
        modoedit = False
        
    Case 13
        'Call ("RepcpMantProveedor.rpt")
        'Despues se hara otro reporte
   End Select
End Sub

Private Sub Combo1_Click()
 'Combo1.ListIndex = (BuscaCombo(Combo1, rsc!negociocodigo))
 cmdBotones(11).Enabled = Validar_DatosNulos()
 'MsgBox "Indice: " & Combo1.ListIndex
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Combo2_Click()
    cmdBotones(11).Enabled = Validar_DatosNulos()
End Sub

Private Sub Combo2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Combo3_Click()
    cmdBotones(11).Enabled = Validar_DatosNulos()
End Sub

Private Sub Combo3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Combo4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Combo5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub


Private Sub Combo7_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub DTP_FecAct_GotFocus()
'     MsgBox "Fecha : " & DTP_FecAct.Value
End Sub
Private Sub DTP_FecAct_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub DTP_FecAnu_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Activate()
    'Listado
End Sub

Private Sub Form_Load()
   'MostrarForm Me, "C2"
    Me.Width = 10020
    Me.Height = 7065
   adll.ActivaTab 0, 2, SSTab1
   'Listado
    cmdBotones(11).Enabled = False
    DTP_FecAnu.CheckBox = True
    DTP_FecAct.CheckBox = True
    Cargacombo
    Call fncCargaArregloComboBusqueda(ArregloBusqueda, cmbBusqueda)
    Call fncBusqueda(VGcnxCP, TDBGrid1)
End Sub

Public Function Listado()
   nListacampo = "clientecodigo as Codigo,clienterazonsocial as Razon,clienteruc as Ruc"
   nLongicampo(1) = 1500
   nLongicampo(2) = 5000
   nLongicampo(3) = 1500
   Cargacombo
   adll.ListarEnTDBGRID VGcnxCP, "cp_proveedor", TDBGrid1, nListacampo, "clientecodigo", nLongicampo
End Function

Public Function Cargacombo()
   Dim rscom As New ADODB.Recordset
   Dim j As Integer
   
   Combo1.Clear
   Set rscom = VGcnxCP.Execute("select * from cp_negocio")
   If rscom.RecordCount > 0 Then
     Do Until rscom.EOF
        Combo1.AddItem rscom!negociocodigo & "-" & rscom!negociodescripcion
        rscom.MoveNext
     Loop
   End If
   Set rscom = Nothing
   
   Combo2.Clear
   Combo2.AddItem "1-NATURAL"
   Combo2.AddItem "2-JURIDICA"
   
   Combo3.Clear
   Combo3.AddItem "1-PERUANA"
   Combo3.AddItem "2-EXTRANJERA"
   
   Combo4.Clear
   Combo4.AddItem "S-SI"
   Combo4.AddItem "N-NO"
   
   Combo5.Clear
   Combo5.AddItem "S-SI"
   Combo5.AddItem "N-NO"

'   Combo6.Clear
'   Combo6.AddItem "S-SI"
'   Combo6.AddItem "N-NO"
   
   Combo7.Clear
   For j = 1 To 180
     Combo7.AddItem j
   Next j

End Function

Public Function CargaData()
  Dim rsc As New ADODB.Recordset
  
  Set rsc = VGcnxCP.Execute("select * from cp_proveedor where clientecodigo='" & TDBGrid1.Columns(0).Text & "'")
  If rsc.RecordCount > 0 Then
        MBox(0) = rsc!clientecodigo
        MBox(1) = Escadena(rsc!clienteruc)
        Combo1.ListIndex = (BuscaCombo(Combo1, rsc!negociocodigo))
        MBox(2) = Escadena(rsc!clienterazonsocial)
        MBox(3) = Escadena(rsc!clientesiglas)
        MBox(5) = Escadena(rsc!clientedireccion)
        MBox(6) = Escadena(rsc!clientecodpostal)
        MBox(7) = Escadena(rsc!clienteprovincia)
        MBox(8) = Escadena(rsc!clientedepartamento)
        MBox(9) = Escadena(rsc!clientetelefono)
        MBox(10) = Escadena(rsc!clientefax)
        MBox(11) = Escadena(rsc!clientemail)
        Combo2.ListIndex = (BuscaCombo(Combo2, rsc!clientetipopersona))
        Combo3.ListIndex = (BuscaCombo(Combo3, rsc!clientetipopais))
        'Combo4.ListIndex = (BuscaCombo(Combo4, rsc!clientesuspendido))
        Combo4.ListIndex = IIf(rsc!clientesuspendido = 0, 1, 0)
        Combo5.ListIndex = (BuscaCombo(Combo5, rsc!clienteaval))
        
        If rsc!clientefechaactivacion <> "" Then
            DTP_FecAct.Value = rsc!clientefechaactivacion
        Else
            DTP_FecAct.CheckBox = True
            DTP_FecAct.Value = ""
        End If
        If rsc!clientefechabajaoanula <> "" Then
            DTP_FecAnu.Value = rsc!clientefechabajaoanula
        Else
            DTP_FecAnu.CheckBox = True
            DTP_FecAnu.Value = ""
        End If
        
        'Combo6.ListIndex = (BuscaCombo(Combo6, rsc!clientemultidireccion))
        Combo7.ListIndex = BuscaCombo2(Combo7, Trim(rsc!clientediasmaxpagocont))
        MBox(14) = Escadena(rsc!clientepropietario)
        MBox(15) = Escadena(rsc!clienteprople)
        MBox(16) = Escadena(rsc!clientepropruc)
        MBox(17) = Escadena(rsc!clientepropdirecc)
        MBox(18) = Escadena(rsc!clienteproptelefono)
        MBox(19) = Escadena(rsc!clientepropcodpostal)
        MBox(20) = Escadena(rsc!clientedistrito)
        
        MBox(21) = IIf(IsNull(rsc!clientedescuento) = True, 0, rsc!clientedescuento * 100)
        
  End If
  Set rsc = Nothing
End Function

Public Function BuscaCombo(xcombo As ComboBox, ByVal xcampo As String) As Integer
   Dim j As Integer
   Dim np As Integer
   Dim xbusca As String
   
    For j = 0 To xcombo.ListCount - 1
       xcombo.ListIndex = j
       np = InStr(xcombo.Text, "-")
       If np > 1 Then
         xbusca = Left(xcombo.Text, np - 1)
       Else
         xbusca = Combo1.Text
       End If
       
       If xcampo = xbusca Then
          BuscaCombo = j
          Exit For
       End If
    Next j
    
End Function

Private Sub MBox_Change(Index As Integer)
    cmdBotones(11).Enabled = Validar_DatosNulos()
End Sub

Private Sub MBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub MBox_KeyPress(Index As Integer, KeyAscii As Integer)
   
    cmdBotones(11).Enabled = Validar_DatosNulos()
    
'   If KeyAscii = 13 Then
'     If Index = 0 Then
'        If MODOINSERT = True Then
'            lcondi = "select * from cp_proveedor where clientecodigo='" & MBox(0).Text & "'"
'            If adll.VerificaDatoExistente(VGCNXCP, lcondi) = 1 Then
'                MsgBox MsgAdd, vbInformation, MsgTitle
'                Exit Sub
'            End If
'        ElseIf MODOEDIT = True Then
'            lcondi = "select * from cp_proveedor " & _
'                    "where clientecodigo ='" & Trim(MBox(0).Text) & "'" & _
'                    " and clientecodigo <> '" & Trim(i_codigocliente) & "'"
'            If adll.VerificaDatoExistente(VGCNXCP, lcondi) = 1 Then
'                MsgBox MsgAdd, vbInformation, MsgTitle
'                Exit Sub
'            End If
'        End If
'     ElseIf Index = 1 Then
'        If Not (Len(MBox(Index)) = 11) Then
'          MsgBox "Ingrese un numero de RUC v�lido..", vbInformation, MsgTitle
'          Exit Sub
'        End If
'     End If
''     Seguir MBox(index), KeyAscii
'   End If
   
   'Ingresar Mayusculas:
    If KeyAscii > 96 And KeyAscii < 123 Then
        KeyAscii = KeyAscii - 32
    End If

End Sub

Private Sub MBox_LostFocus(Index As Integer)
'Dim lcondi As String

  If MBox(Index) <> "" Then
        
     If Index = 0 Then
        
        If modoinsert = True Then
            Call Formatear_Codigo(Index)
             
             If Verificar_Codigo(0) = False Then
                'MsgBox MsgAdd, vbInformation, MsgTitle
                MsgBox "El C�digo ya existe...VERIFIQUE", vbInformation, MsgTitle
                Exit Sub
             End If
        ElseIf modoedit = True Then
            If Verificar_Codigo(1) = False Then
                'MsgBox MsgAdd, vbInformation, MsgTitle
                MsgBox "El C�digo ya existe...VERIFIQUE", vbInformation, MsgTitle
                Exit Sub
            End If
        End If
     ElseIf Index = 1 Then
        If Not (Len(MBox(Index)) = 11) Then
          MsgBox "Ingrese un numero de RUC v�lido..", vbInformation, MsgTitle
          Exit Sub
        End If
        
     ElseIf Index = 21 Then
            MBox(Index).Text = Format(CDbl(MBox(Index).Text), "##,##0.00")
     
     End If
      
 End If

End Sub

 Public Function BuscaPrefijo(xTexto As String) As String
    Dim j As Integer
    j = InStr(xTexto, "-")
    If j > 1 Then
        BuscaPrefijo = Trim(Left(xTexto, j - 1))
    Else
        BuscaPrefijo = Trim(xTexto) & " "
    End If
End Function

Public Function GrabaCliente()
  Dim opc As Integer
  Dim nsql As String
  Dim s_fechaactivacion As String
  Dim s_fechaanulacion As String
  Dim d_descuento As Double
  
  On Error GoTo CONTROLERRORES
  
  opc = adll.VerificaDatoExistente(VGcnxCP, "select * from cp_proveedor where clientecodigo='" & MBox(0) & "'")
  
  s_fechaactivacion = IIf(IsNull(DTP_FecAct), "Null", DTP_FecAct)
  If Not IsNull(DTP_FecAct) Then
    s_fechaactivacion = "'" & s_fechaactivacion & "'"
  End If
  
  s_fechaanulacion = IIf(IsNull(DTP_FecAnu), "Null", DTP_FecAnu)
  If Not IsNull(DTP_FecAnu) Then
    s_fechaanulacion = "'" & s_fechaanulacion & "'"
  End If

 ' d_descuento = IIf(Len(Trim(MBox(21))) = 0, 0, CDbl(MBox(21)) / 100)
  If MBox(21) = "" Then
      d_descuento = 0
  Else
      d_descuento = MBox(21) / 100
  End If
      

  If opc = 0 Then
            
              nsql = "INSERT INTO cp_proveedor " & _
             "(clientecodigo,clienteruc,negociocodigo," & _
             "clienterazonsocial,clientesiglas," & _
             "clientedireccion," & _
             "clientecodpostal,clienteprovincia," & _
             "clientedepartamento,clientetelefono," & _
             "clientefax,clientemail," & _
             "clientetipopersona,clientetipopais," & _
             "clientesuspendido,clienteaval," & _
             "clientefechaactivacion,clientefechabajaoanula," & _
             "clientediasmaxpagocont," & _
             "clientepropietario,clienteprople," & _
             "clientepropruc,clientepropdirecc," & _
             "clienteproptelefono," & _
             "clientepropcodpostal,fechaact,usuariocodigo,clientedistrito,clientedescuento)" & _
             " VALUES(" & _
             "'" & MBox(0) & "','" & MBox(1) & "','" & BuscaPrefijo(Combo1.Text) & "'," & _
             "'" & MBox(2) & "','" & MBox(3) & "'," & _
             "'" & MBox(5) & "'," & _
             "'" & MBox(6) & "','" & MBox(7) & "'," & _
             "'" & MBox(8) & "','" & MBox(9) & "','" & MBox(10) & "','" & MBox(11) & "','" & BuscaPrefijo(Combo2.Text) & "','" & BuscaPrefijo(Combo3.Text) & "'," & _
             IIf(BuscaPrefijo(Combo4.Text) = "S", 1, 0) & ",'" & BuscaPrefijo(Combo5.Text) & "'," & s_fechaactivacion & "," & s_fechaanulacion & "," & _
             IIf(Combo7.Text = "", 0, Combo7.Text) & ",'" & MBox(14) & "','" & MBox(15) & "','" & MBox(16) & "','" & _
             MBox(17) & "','" & MBox(18) & "','" & MBox(19) & "','" & _
             Date & "','" & VGusuario & "','" & MBox(20) & "'," & d_descuento & ")"
       
             
  ElseIf opc = 1 Then
             
       ' SACAR EL CLIENTE SUSPENDIDO
             
       nsql = "UPDATE cp_proveedor " & _
             " Set clientecodigo='" & MBox(0) & "',clienteruc='" & MBox(1) & "'," & _
             "     negociocodigo='" & BuscaPrefijo(Combo1.Text) & "',clienterazonsocial='" & MBox(2) & "'," & _
             "     clientesiglas='" & MBox(3) & "'," & _
             "     clientedireccion='" & MBox(5) & "',clientecodpostal='" & MBox(6) & "'," & _
             "     clienteprovincia='" & MBox(7) & "',clientedepartamento='" & MBox(8) & "'," & _
             "     clientetelefono='" & MBox(9) & "',clientefax='" & MBox(10) & "',clientemail='" & MBox(11) & "'," & _
             "     clientetipopersona='" & BuscaPrefijo(Combo2.Text) & "',clientetipopais='" & BuscaPrefijo(Combo3.Text) & "'," & _
             "     clientesuspendido='" & IIf(BuscaPrefijo(Combo4.Text) = "S", "1", "0") & "',clienteaval='" & BuscaPrefijo(Combo5.Text) & "'," & _
             "     clientefechaactivacion=" & s_fechaactivacion & ",clientefechabajaoanula=" & s_fechaanulacion & "," & _
             "     clientediasmaxpagocont=" & IIf(Combo7.ListIndex = -1, 0, Trim(Combo7.Text)) & "," & _
             "     clientepropietario='" & MBox(14) & "',clienteprople='" & MBox(15) & "'," & _
             "     clientepropruc='" & MBox(16) & "',clientepropdirecc='" & MBox(17) & "'," & _
             "     clienteproptelefono='" & MBox(18) & "'," & _
             "     clientepropcodpostal='" & MBox(19) & "',fechaact='" & Date & "',usuariocodigo='" & VGusuario & "'," & "     clientedistrito='" & _
                   MBox(20) & "', " & _
             "      clientedescuento=" & d_descuento & _
             "   WHERE clientecodigo='" & _
             MBox(0) & "'"
      
  End If
  
  If Transaccion(nsql, VGcnxCP) = True Then
    VGcnxCP.CommitTrans
  Else
    VGcnxCP.RollbackTrans
  End If
  
  'Listado
  cmbBusqueda.ListIndex = -1
  txtBusqueda = ""
  Call fncBusqueda(VGcnxCP, TDBGrid1)
  
Exit Function
CONTROLERRORES:
  If Err Then
     MsgBox VGcnxCP.Errors(0).Number & "-" & VGcnxCP.Errors(0).Description
     Err = 0
     Resume Next
     'Exit Function
  End If
End Function
Private Function Transaccion(SQL As String, con As Connection) As Boolean
Dim SENsql As String
Dim rs As ADODB.Recordset
On Error GoTo CONTROLERRORES

    Transaccion = False
    con.BeginTrans
    
    con.Execute SQL
    
    If modoinsert = True Then
        SENsql = "INSERT INTO cp_proveedordireccion " & _
            "SELECT * FROM TEMPO_proveedordireccion " & _
            "WHERE CLIENTECODIGO ='" & Trim(MBox(0)) & "'"
        con.Execute (SENsql)
    End If
    
'    SENsql = "DELETE FROM cp_proveedordireccion " & _
'            "WHERE CLIENTECODIGO ='" & MBox(0) & "'"
'    con.Execute (SENsql)
    
    SENsql = "SELECT Count(*) FROM cp_proveedordireccion " & _
            "WHERE CLIENTECODIGO ='" & Trim(MBox(0)) & "'"
            
    Set rs = con.Execute(SENsql)
    If rs(0) > 0 Then
        SENsql = "UPDATE cp_proveedor " & _
                "SET CLIENTEMULTIDIRECCION = 1 " & _
                               "WHERE CLIENTECODIGO ='" & Trim(MBox(0)) & "'"
    Else
            SENsql = "UPDATE cp_proveedor " & _
                "SET CLIENTEMULTIDIRECCION = 0 " & _
                               "WHERE CLIENTECODIGO ='" & Trim(MBox(0)) & "'"
    End If
    con.Execute SENsql
    Transaccion = True
Exit Function

CONTROLERRORES:
   If Err Then
      Exit Function
   End If
End Function


Private Function Validar_DatosNulos() As Boolean

Validar_DatosNulos = False
                If Trim(MBox(0)) <> "" And Trim(MBox(1)) <> "" _
                And Combo1.ListIndex <> -1 And Combo2.ListIndex <> -1 _
                And Combo3.ListIndex <> -1 Then
                    Validar_DatosNulos = True
                    Exit Function
                End If

End Function

Private Function fncCargaArregloComboBusqueda(ArrayBusqueda As Variant, cmb As ComboBox)
Dim I As Integer
    ReDim ArrayBusqueda(0 To 2, 0 To 7)

   'Nombre Campos:
   ArrayBusqueda(0, 0) = "clientecodigo"
   ArrayBusqueda(0, 1) = "clienterazonsocial"
   ArrayBusqueda(0, 2) = "clienteruc"
   ArrayBusqueda(0, 3) = "clientedistrito"
   ArrayBusqueda(0, 4) = "clienteprovincia"
   ArrayBusqueda(0, 5) = "clientedepartamento"
   ArrayBusqueda(0, 6) = "clientetelefono"
   ArrayBusqueda(0, 7) = "clientesuspendido"
   'Nombres de Campo(Combo Busqueda):
   ArrayBusqueda(1, 0) = "C�digo"
   ArrayBusqueda(1, 1) = "Razon Social"
   ArrayBusqueda(1, 2) = "RUC"
   ArrayBusqueda(1, 3) = "Distrito"
   ArrayBusqueda(1, 4) = "Provincia"
   ArrayBusqueda(1, 5) = "Departamento"
   ArrayBusqueda(1, 6) = "Telefono"
   ArrayBusqueda(1, 7) = "Suspendido"
   'Tipo de Dato:
   ArrayBusqueda(2, 0) = "C"
   ArrayBusqueda(2, 1) = "C"
   ArrayBusqueda(2, 2) = "C"
   ArrayBusqueda(2, 3) = "C"
   ArrayBusqueda(2, 4) = "C"
   ArrayBusqueda(2, 5) = "C"
   ArrayBusqueda(2, 6) = "C"
   ArrayBusqueda(2, 7) = "B"
   
   cmb.Clear
   For I = 0 To UBound(ArrayBusqueda, 2)
    cmb.AddItem (Trim(ArrayBusqueda(1, I)))
   Next I
    
End Function

Private Function fncBusqueda(conexion As Connection, grid As TDBGrid)
    Dim SQL As String
    Dim where As String
    Dim condicion As String
    Dim rs As Recordset
    Dim I As Integer
    
    where = ""
    condicion = ""
    
    SQL = "SELECT " & _
         "clientecodigo as 'C�d.Cliente'," & _
         "clienteruc as 'RUC'," & _
         "clienterazonsocial as 'Raz�n Social'," & _
         "clientesiglas as 'Siglas'," & _
         "clientedireccion as 'Direcci�n'," & _
         "clientetelefono as Tel�fono, " & _
         "clientesuspendido as 'Suspendido' " & _
         "FROM cp_proveedor "
    
    If cmbBusqueda.ListIndex <> -1 Then
       where = " WHERE " & _
              Trim(ArregloBusqueda(0, cmbBusqueda.ListIndex))
       Select Case ArregloBusqueda(2, cmbBusqueda.ListIndex)
         Case "C"
            condicion = " LIKE '%" & Trim(txtBusqueda) & "%'"
         Case "N"
            condicion = " = " & Trim(txtBusqueda)
         Case "B"
            If Left(txtBusqueda, 1) = "S" Then
                condicion = " = 1"
            ElseIf Left(txtBusqueda, 1) = "N" Then
                condicion = " = 0"
            End If
       End Select
    End If
       
    SQL = SQL & where & condicion
     
    Set rs = conexion.Execute(SQL)
    Set TDBGrid1.DataSource = rs
    
 ''''''''''''''''''''''''''''''''''' Tipo Columna
      'For i = 0 To grid.Columns.Count - 1
         'grid.Columns(i).Width = i_width * (Len(a_Arreglo(1, i)) / i_total)
         'If ArregloBusqueda(2, i) = "B" Then
         '   grid.Columns(i).ValueItems.Presentation = dbgCheckBox
         'Else
         '   grid.Columns(i).ValueItems.Presentation = dbgNormal
         'End If
      'Next i
      
      TDBGrid1.Columns(6).ValueItems.Presentation = dbgCheckBox
      TDBGrid1.Refresh
    
End Function
Public Function Formatear_Codigo(indice As Integer) As String
Dim CADENA As String
Dim I As Integer

'cadena = ""
'For I = 0 To MBox(indice).MaxLength
'    cadena = cadena & "0"
'Next I

'MBox(indice) = Right(cadena & Trim(MBox(indice)), MBox(indice).MaxLength)

End Function

Private Function Verificar_Codigo(operacion As Integer) As Boolean
Dim lcondi As String

Verificar_Codigo = True

If operacion = 0 Then           'ingreso

            lcondi = "select * from cp_proveedor where clientecodigo='" & MBox(0).Text & "'"
            If adll.VerificaDatoExistente(VGcnxCP, lcondi) = 1 Then
                Verificar_Codigo = False
                Exit Function
            End If

ElseIf operacion = 1 Then       'edicion

            lcondi = "select * from cp_proveedor " & _
                    "where clientecodigo ='" & Trim(MBox(0).Text) & "'" & _
                    " and clientecodigo <> '" & Trim(i_codigocliente) & "'"
            If adll.VerificaDatoExistente(VGcnxCP, lcondi) = 1 Then
                    Verificar_Codigo = False
                Exit Function
            End If
End If
End Function
Private Function Eliminar_Cliente(con As Connection) As Boolean
Dim SENsql As String
On Error GoTo CONTROLERRORES

    Eliminar_Cliente = False
    con.BeginTrans
    
    SENsql = "DELETE FROM cp_proveedordireccion " & _
                "WHERE CLIENTECODIGO = '" & TDBGrid1.Columns(0).Text & "'"
    con.Execute SENsql
    
    SENsql = "DELETE FROM cp_proveedor " & _
                "WHERE CLIENTECODIGO = '" & TDBGrid1.Columns(0).Text & "'"
    
    con.Execute SENsql
    Eliminar_Cliente = True

Exit Function

CONTROLERRORES:
   If Err Then
      Exit Function
   End If
End Function

Public Function BuscaCombo2(xcombo As ComboBox, ByVal xcampo As String) As Integer
   Dim j As Integer
   Dim CADENA As String
   
    For j = 0 To xcombo.ListCount - 1
       xcombo.ListIndex = j
       CADENA = Trim(xcombo.Text)
       If CADENA = Trim(xcampo) Then
         BuscaCombo2 = j
         Exit For
       End If
    Next j
    
End Function

Private Sub txtBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Public Function Escadena(pDato) As String
   If IsNull(pDato) Or Len(Trim(pDato)) = 0 Then
     Escadena = ""
   Else
     Escadena = Trim(pDato)
   End If
End Function
