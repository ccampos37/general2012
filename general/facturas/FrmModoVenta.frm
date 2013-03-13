VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmModoVenta 
   Caption         =   "Modo Venta"
   ClientHeight    =   8790
   ClientLeft      =   -1275
   ClientTop       =   450
   ClientWidth     =   12150
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   12150
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   7215
      Left            =   120
      TabIndex        =   27
      Top             =   120
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   12726
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Consulta"
      TabPicture(0)   =   "FrmModoVenta.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "TDBGridModoVta"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Mantenimiento"
      TabPicture(1)   =   "FrmModoVenta.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   6495
         Left            =   120
         TabIndex        =   35
         Top             =   600
         Width           =   11625
         Begin VB.Frame Frame4 
            Height          =   2295
            Left            =   2640
            TabIndex        =   75
            Top             =   2640
            Width           =   2895
            Begin VB.TextBox txtguia 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2040
               MaxLength       =   8
               TabIndex        =   77
               Top             =   960
               Width           =   585
            End
            Begin VB.TextBox txthoja 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2040
               MaxLength       =   8
               TabIndex        =   76
               Top             =   240
               Width           =   585
            End
            Begin VB.Label lbl 
               Caption         =   "Guía Remisión"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   3
               Left            =   240
               TabIndex        =   79
               Top             =   960
               Width           =   1515
            End
            Begin VB.Label lbl 
               Caption         =   "Hoja Trabajo"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Index           =   7
               Left            =   120
               TabIndex        =   78
               Top             =   240
               Width           =   1425
            End
         End
         Begin VB.Frame FrmCopias 
            Caption         =   "Copias"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2295
            Left            =   240
            TabIndex        =   65
            Top             =   2640
            Width           =   2895
            Begin VB.OptionButton optCopias 
               Height          =   495
               Index           =   8
               Left            =   960
               TabIndex        =   81
               Top             =   1710
               Width           =   315
            End
            Begin VB.TextBox txt 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   8
               Left            =   1440
               MaxLength       =   8
               TabIndex        =   80
               Top             =   1830
               Width           =   585
            End
            Begin VB.TextBox txt 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   7
               Left            =   1440
               MaxLength       =   8
               TabIndex        =   71
               Top             =   1380
               Width           =   585
            End
            Begin VB.OptionButton optCopias 
               Height          =   495
               Index           =   7
               Left            =   960
               TabIndex        =   70
               Top             =   1260
               Width           =   315
            End
            Begin VB.TextBox txt 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   5
               Left            =   1440
               MaxLength       =   8
               TabIndex        =   69
               Top             =   480
               Width           =   585
            End
            Begin VB.OptionButton optCopias 
               Height          =   495
               Index           =   5
               Left            =   990
               TabIndex        =   68
               Top             =   360
               Width           =   375
            End
            Begin VB.OptionButton optCopias 
               Height          =   495
               Index           =   6
               Left            =   960
               TabIndex        =   67
               Top             =   810
               Width           =   375
            End
            Begin VB.TextBox txt 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   6
               Left            =   1440
               MaxLength       =   8
               TabIndex        =   66
               Top             =   930
               Width           =   585
            End
            Begin VB.Label lbl 
               Caption         =   "Ticket"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   26
               Left            =   120
               TabIndex        =   82
               Top             =   1830
               Width           =   735
            End
            Begin VB.Label lbl 
               Caption         =   "Varios"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   25
               Left            =   120
               TabIndex        =   74
               Top             =   1380
               Width           =   885
            End
            Begin VB.Label lbl 
               Caption         =   "Factura"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   5
               Left            =   120
               TabIndex        =   73
               Top             =   480
               Width           =   915
            End
            Begin VB.Label lbl 
               Caption         =   "Boleta"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   6
               Left            =   120
               TabIndex        =   72
               Top             =   930
               Width           =   765
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Modo"
            Height          =   705
            Left            =   9840
            TabIndex        =   62
            Top             =   5070
            Width           =   1605
            Begin VB.CheckBox cmodo 
               Height          =   375
               Left            =   960
               TabIndex        =   64
               Top             =   270
               Width           =   375
            End
            Begin VB.Label Label1 
               Caption         =   "Canje"
               Height          =   285
               Left            =   270
               TabIndex        =   63
               Top             =   330
               Width           =   615
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Almacenes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   6960
            TabIndex        =   60
            Top             =   5040
            Width           =   2835
            Begin VB.TextBox Text1 
               Height          =   285
               Left            =   120
               TabIndex        =   61
               Top             =   360
               Width           =   2595
            End
         End
         Begin VB.Frame FrmFlags 
            Height          =   4095
            Left            =   5640
            TabIndex        =   43
            Top             =   120
            Width           =   5895
            Begin VB.CheckBox chk 
               Height          =   255
               Index           =   15
               Left            =   5280
               TabIndex        =   20
               Top             =   3720
               Width           =   375
            End
            Begin VB.CheckBox chk 
               Height          =   255
               Index           =   14
               Left            =   5280
               TabIndex        =   15
               Top             =   1320
               Width           =   375
            End
            Begin VB.CheckBox chk 
               Height          =   255
               Index           =   13
               Left            =   5280
               TabIndex        =   18
               Top             =   2760
               Width           =   375
            End
            Begin VB.CheckBox chk 
               Height          =   255
               Index           =   12
               Left            =   5280
               TabIndex        =   19
               Top             =   3240
               Width           =   375
            End
            Begin VB.CheckBox chk 
               Height          =   255
               Index           =   11
               Left            =   2280
               TabIndex        =   11
               Top             =   3240
               Width           =   375
            End
            Begin VB.CheckBox chk 
               Height          =   375
               Index           =   10
               Left            =   2280
               TabIndex        =   10
               Top             =   2760
               Width           =   375
            End
            Begin VB.CheckBox chk 
               Height          =   375
               Index           =   7
               Left            =   5280
               TabIndex        =   16
               Top             =   1800
               Width           =   375
            End
            Begin VB.CheckBox chk 
               Height          =   375
               Index           =   6
               Left            =   2280
               TabIndex        =   9
               Top             =   2280
               Width           =   375
            End
            Begin VB.CheckBox chk 
               Height          =   375
               Index           =   5
               Left            =   5280
               TabIndex        =   13
               Top             =   360
               Width           =   375
            End
            Begin VB.CheckBox chk 
               Height          =   375
               Index           =   2
               Left            =   2280
               TabIndex        =   7
               Top             =   1320
               Width           =   375
            End
            Begin VB.CheckBox chk 
               Height          =   255
               Index           =   4
               Left            =   2280
               TabIndex        =   12
               Top             =   3720
               Width           =   375
            End
            Begin VB.CheckBox chk 
               Height          =   375
               Index           =   3
               Left            =   2280
               TabIndex        =   8
               Top             =   1800
               Width           =   375
            End
            Begin VB.CheckBox chk 
               Height          =   375
               Index           =   0
               Left            =   2280
               TabIndex        =   5
               Top             =   360
               Width           =   375
            End
            Begin VB.CheckBox chk 
               Height          =   375
               Index           =   1
               Left            =   2280
               TabIndex        =   6
               Top             =   840
               Width           =   375
            End
            Begin VB.CheckBox chk 
               Height          =   375
               Index           =   8
               Left            =   5280
               TabIndex        =   14
               Top             =   840
               Width           =   375
            End
            Begin VB.CheckBox chk 
               Height          =   375
               Index           =   9
               Left            =   5280
               TabIndex        =   17
               Top             =   2280
               Width           =   375
            End
            Begin VB.Label lbl 
               Caption         =   "Factor Conversión"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   8
               Left            =   2880
               TabIndex        =   59
               Top             =   3720
               Width           =   1995
            End
            Begin VB.Label lbl 
               Caption         =   "Emite Hoja Trabajo"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   24
               Left            =   2880
               TabIndex        =   58
               Top             =   840
               Width           =   2085
            End
            Begin VB.Label lbl 
               Caption         =   "Emite Guia Remisión"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   23
               Left            =   2880
               TabIndex        =   57
               Top             =   1320
               Width           =   2085
            End
            Begin VB.Label lbl 
               Caption         =   "Solo Emite Factura"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   9
               Left            =   2880
               TabIndex        =   56
               Top             =   2760
               Width           =   2085
            End
            Begin VB.Label lbl 
               Caption         =   "Ing. Hasta Hoj.Trab."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   22
               Left            =   2880
               TabIndex        =   55
               Top             =   3240
               Width           =   1935
            End
            Begin VB.Label lbl 
               Caption         =   "Ingresa Hasta Fact."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   21
               Left            =   240
               TabIndex        =   54
               Top             =   3240
               Width           =   1995
            End
            Begin VB.Label lbl 
               Caption         =   "Ingresa Forma Pago"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   4
               Left            =   240
               TabIndex        =   53
               Top             =   2760
               Width           =   1995
            End
            Begin VB.Label lbl 
               Caption         =   "Ingresa Pedido"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   12
               Left            =   240
               TabIndex        =   52
               Top             =   360
               Width           =   1515
            End
            Begin VB.Label lbl 
               Caption         =   "Ingresa Hoja Trab."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   13
               Left            =   240
               TabIndex        =   51
               Top             =   840
               Width           =   1875
            End
            Begin VB.Label lbl 
               Caption         =   "Ingresa Guia Rem."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   14
               Left            =   240
               TabIndex        =   50
               Top             =   1335
               Width           =   1995
            End
            Begin VB.Label lbl 
               Caption         =   "Control Inventario"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   15
               Left            =   240
               TabIndex        =   49
               Top             =   1815
               Width           =   1875
            End
            Begin VB.Label lbl 
               Caption         =   "Impuestos"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   16
               Left            =   960
               TabIndex        =   48
               Top             =   3720
               Width           =   1275
            End
            Begin VB.Label lbl 
               Caption         =   "Controla Correlativo"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   17
               Left            =   2880
               TabIndex        =   47
               Top             =   360
               Width           =   1995
            End
            Begin VB.Label lbl 
               Caption         =   "Ingresa Cod.Cliente"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   18
               Left            =   240
               TabIndex        =   46
               Top             =   2280
               Width           =   1995
            End
            Begin VB.Label lbl 
               Caption         =   "Actualiza Cta.Cte."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Index           =   19
               Left            =   2880
               TabIndex        =   45
               Top             =   1800
               Width           =   1995
            End
            Begin VB.Label lbl 
               Caption         =   "Numeración Automat."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   20
               Left            =   2880
               TabIndex        =   44
               Top             =   2280
               Width           =   2205
            End
         End
         Begin VB.ComboBox cmbDocumento 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1680
            Width           =   3255
         End
         Begin VB.Frame Frmdescuento 
            Caption         =   "Descuento"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   360
            TabIndex        =   42
            Top             =   5040
            Width           =   2895
            Begin VB.OptionButton optDescuento 
               Caption         =   "Reparto"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   1500
               TabIndex        =   22
               Top             =   270
               Width           =   1215
            End
            Begin VB.OptionButton optDescuento 
               Caption         =   "Oficina"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   0
               Left            =   240
               TabIndex        =   21
               Top             =   270
               Width           =   1095
            End
         End
         Begin VB.Frame FrmUnidadMedida 
            Caption         =   "Unidad Medida"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   3300
            TabIndex        =   41
            Top             =   5040
            Width           =   3615
            Begin VB.OptionButton optUnidad 
               Caption         =   "Ventas"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   2280
               TabIndex        =   24
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton optUnidad 
               Caption         =   "Referencial"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   0
               Left            =   480
               TabIndex        =   23
               Top             =   240
               Width           =   1575
            End
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   2040
            MaxLength       =   2
            TabIndex        =   0
            Top             =   240
            Width           =   1185
         End
         Begin VB.CommandButton cAcepta 
            Caption         =   "&Aceptar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4200
            TabIndex        =   25
            Top             =   5880
            Width           =   1335
         End
         Begin VB.CommandButton cCancela 
            Caption         =   "&Cancelar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   5880
            TabIndex        =   26
            Top             =   5880
            Width           =   1335
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   11
            Left            =   2040
            MaxLength       =   8
            TabIndex        =   4
            Top             =   2160
            Width           =   1185
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   2040
            MaxLength       =   30
            TabIndex        =   1
            Top             =   720
            Width           =   3225
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   2040
            MaxLength       =   15
            TabIndex        =   2
            Top             =   1200
            Width           =   3225
         End
         Begin VB.Label lbl 
            Caption         =   "Código"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   360
            TabIndex        =   40
            Top             =   240
            Width           =   960
         End
         Begin VB.Label lbl 
            Caption         =   "Descripción"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   360
            TabIndex        =   39
            Top             =   720
            Width           =   1395
         End
         Begin VB.Label lbl 
            Caption         =   "Descrip. Corta"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   2
            Left            =   360
            TabIndex        =   38
            Top             =   1200
            Width           =   1605
         End
         Begin VB.Label lbl 
            Caption         =   "Documento"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   10
            Left            =   360
            TabIndex        =   37
            Top             =   1680
            Width           =   1365
         End
         Begin VB.Label lbl 
            Caption         =   "Item por Docum."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   11
            Left            =   360
            TabIndex        =   36
            Top             =   2160
            Width           =   1635
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGridModoVta 
         Height          =   5775
         Left            =   -74880
         TabIndex        =   28
         Top             =   480
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   10186
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
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2752"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2752"
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=975,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Arial"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=975,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=Arial"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=975,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=Arial"
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
   Begin VB.Frame frmbotones 
      Height          =   1050
      Left            =   3180
      TabIndex        =   29
      Top             =   7470
      Width           =   5655
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Nuevo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   0
         Left            =   225
         Picture         =   "FrmModoVenta.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   180
         Width           =   915
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "E&ditar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   1
         Left            =   1320
         Picture         =   "FrmModoVenta.frx":047A
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   180
         Width           =   915
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Eliminar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   2
         Left            =   2385
         Picture         =   "FrmModoVenta.frx":08BC
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   180
         Width           =   870
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   4
         Left            =   4590
         Picture         =   "FrmModoVenta.frx":0CFE
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   180
         Width           =   915
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Imprimir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   3
         Left            =   3510
         Picture         =   "FrmModoVenta.frx":1140
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   180
         Width           =   870
      End
   End
End
Attribute VB_Name = "FrmModoVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim modoinsert As Boolean
Dim modoedit As Boolean
Dim i_filaorigen As Integer
Dim i_valorcodigo As String
Dim Arreglo_Doc()

Private Sub cAcepta_Click()
   Dim rs As New ADODB.Recordset
   Dim SQL As String
   Dim J As Integer
   
   Dim s_modovtadescto, s_modovtaunidad, s_codigodocumento As String
   Dim i_copiasfact, i_copiasbol, i_copiasguia, i_copiashoja, i_copiasticket As Integer
   
   On Error GoTo CONTROLERRORES

   ''''''''
   
                If txt(5) = "" Then i_copiasfact = 0 Else i_copiasfact = txt(5)
                If txt(6) = "" Then i_copiasbol = 0 Else i_copiasbol = txt(6)
                If txt(8) = "" Then i_copiasticket = 0 Else i_copiasticket = txt(8)
                If txthoja = "" Then i_copiashoja = 0 Else i_copiashoja = txthoja
                If txtguia.Text = "" Then
                   i_copiasguia = 0
                 Else
                  i_copiasguia = txtguia.Text
                End If
                If optUnidad(0).Value = True Then
                    s_modovtaunidad = "R"
                ElseIf optUnidad(1).Value = True Then
                    s_modovtaunidad = "V"
                End If
                    
                If optDescuento(0).Value = True Then
                    s_modovtadescto = "O"
                ElseIf optDescuento(1).Value = True Then
                    s_modovtadescto = "R"
                End If
                
                If cmbDocumento.ListIndex <> -1 Then
                    s_codigodocumento = Arreglo_Doc(0, cmbDocumento.ListIndex)
                Else
                    s_codigodocumento = ""
                End If
                
   If modoinsert = True Then
   
         If Validar_CodigosDuplicados("INSERT") = True Then
            MsgBox "Código ya existe", vbCritical, "Error"
            cAcepta.Enabled = False
            Exit Sub
          End If
          
         If Validar_DescripcionesDuplicadas("INSERT") = True Then
            MsgBox "Descripción ya existe", vbCritical, "Error"
            cAcepta.Enabled = False
            Exit Sub
          End If
             
          SQL = "INSERT INTO vt_modoventa " & _
               "(modovtacodigo,modovtadescripcion,modovtadescrcorta,modovtaingpedido," & _
               "modovtainghojatrab,modovtaingguiarem,modovtactrlinventario,modovtaunidadmedida," & _
               "modovtadscto,modovtacopiasfact,modovtacopiasboleta,modovtaimpuestos,modovtacontrolcorr," & _
               "modovtacopiashojatrab,modovtaemitehoja,modovtasolemitfact,documentocodigo,modovtaingcodclie," & _
               "modovtaactctacte,modovtaitemxdoc,fechaact,usuariocodigo,modovtacopiasguiarem,modovtanumautom," & _
               "modovtaingformapag,modovtainghastafact,modovtainghastahoja,modovtaemiteguia,modovtausafactconv,modovtaalmacen,modovtacanje,modovtacopiasticket)" & _
               " VALUES " & _
               "('" & txt(0) & "','" & txt(1) & "','" & txt(2) & "'," & chk(0).Value & "," & _
               chk(1).Value & "," & chk(2).Value & "," & chk(3).Value & ",'" & s_modovtaunidad & "','" & _
               s_modovtadescto & "'," & i_copiasfact & "," & i_copiasbol & "," & chk(4).Value & "," & chk(5).Value & "," & _
               i_copiashoja & "," & chk(8).Value & "," & chk(13).Value & ",'" & _
               s_codigodocumento & "'," & chk(6) & "," & _
               chk(7).Value & "," & txt(11) & "," & _
               "'" & Date & "','" & g_usuario & "'," & i_copiasguia & "," & chk(9).Value & _
               "," & chk(10).Value & "," & chk(11).Value & "," & chk(12).Value & "," & chk(14).Value & "," & chk(15).Value & ",'" & IIf(IsNull(Text1), "", Text1) & "','" & cmodo.Value & "'," & i_copiasticket & ")"

          VGCNx.Execute SQL
                   
   ElseIf modoedit = True Then
   
             If Validar_CodigosDuplicados("UPDATE", i_filaorigen) = True Then
               MsgBox "Código ya existe", vbCritical, "Error"
               cAcepta.Enabled = False
               Exit Sub
             End If
                          
             If Validar_DescripcionesDuplicadas("UPDATE", i_filaorigen) = True Then
               MsgBox "Descripción ya existe", vbCritical, "Error"
               cAcepta.Enabled = False
               Exit Sub
             End If
                          
            SQL = "UPDATE vt_modoventa SET " & _
               "modovtacodigo='" & txt(0) & "',modovtadescripcion='" & txt(1) & "'," & _
               "modovtadescrcorta='" & txt(2) & "'," & _
               "modovtaingpedido=" & chk(0).Value & ",modovtainghojatrab=" & chk(1).Value & _
               ",modovtaingguiarem=" & chk(2).Value & ",modovtactrlinventario=" & chk(3).Value & _
               ",modovtaunidadmedida='" & s_modovtaunidad & "',modovtadscto='" & s_modovtadescto & "',modovtacopiasfact=" & i_copiasfact & _
               ",modovtacopiasboleta ='" & i_copiasbol & "',modovtaimpuestos=" & chk(4).Value & "," & _
               "modovtacontrolcorr=" & chk(5).Value & ",modovtacopiashojatrab=" & i_copiashoja & _
               ",modovtaemitehoja=" & chk(8) & ",modovtasolemitfact=" & chk(13).Value & ",documentocodigo='" & s_codigodocumento & _
               "',modovtaingcodclie=" & chk(6).Value & ",modovtaactctacte=" & chk(7).Value & "," & _
               "modovtaitemxdoc=" & txt(11) & ",fechaact='" & Date & "',usuariocodigo='" & g_usuario & "'," & _
               "modovtacopiasguiarem=" & i_copiasguia & ",modovtanumautom=" & chk(9).Value & "," & _
               "modovtaingformapag=" & chk(10).Value & ",modovtainghastafact=" & chk(11).Value & "," & _
               "modovtainghastahoja=" & chk(12).Value & ",modovtaemiteguia=" & chk(14).Value & " " & _
               ",modovtausafactconv=" & chk(15).Value & " " & _
               ",modovtaalmacen='" & IIf(IsNull(Text1), "", Text1) & "' " & _
               ",modovtacanje='" & cmodo.Value & "', modovtacopiasticket=" & i_copiasticket & " " & _
               "WHERE modovtacodigo ='" & i_valorcodigo & "'"
               
            VGCNx.Execute SQL
              
 '******************************************************************************************
        
 End If
      
 Mostrar_Data
 MostrarOcultar_Botones (True)
 '''''''''
 modoinsert = False
 modoedit = False
 '''''''''
 'rs.Close
 'Set rs = Nothing
 TDBGridModoVta.Refresh
 SSTab1.TabEnabled(0) = True
 
Exit Sub
CONTROLERRORES:
   If Err Then
       MsgBox Err.Number & "-" & Err.Description, vbInformation, "ERROR"
       Err = 0
       'VGgeneral.RollbackTrans
       Resume Next
    End If
       
End Sub

Private Sub cCancela_Click()
    SSTab1.TabEnabled(0) = True
    SSTab1.Tab = 0
    SSTab1.SetFocus
    MostrarOcultar_Botones (True)
    modoinsert = False
    modoedit = False
End Sub


Private Sub chk_Click(Index As Integer)
     cAcepta.Enabled = Validar_DatosNulos()
End Sub

Private Sub chk_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmbDocumento_Click()
    cAcepta.Enabled = Validar_DatosNulos()
End Sub

Private Sub cmbDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmdBotones_Click(Index As Integer)
  Dim J As Integer
  Dim SQL As String
  Dim OBJ As Object
  
  On Error GoTo CONTROLERRORES
  
  SSTab1.TabEnabled(1) = True
  '''''
  Select Case Index
     Case 0   'nuevo
            For Each OBJ In Me.Controls
                  If TypeOf OBJ Is TextBox Then
                    OBJ.Text = ""
                  ElseIf TypeOf OBJ Is CheckBox Then
                    OBJ.Value = 0
                  ElseIf TypeOf OBJ Is OptionButton Then
                    OBJ.Value = False
                  ElseIf TypeOf OBJ Is ComboBox Then
                    OBJ.ListIndex = -1
                  End If
            Next
            SSTab1.Tab = 1
            modoinsert = True
            MostrarOcultar_Botones (False)
            txt(5).Enabled = False
            txt(6).Enabled = False
            txt(7).Enabled = False
        '    txtguia.Enabled = False
        '    txthoja.Enabled = False
            txt(0).SetFocus
        
     Case 1   'modificar
     
         If TDBGridModoVta.Row < 0 Then
            Exit Sub
         End If
         
             Call fncSeleccionaCombo(Trim(TDBGridModoVta.Columns(4).Text), cmbDocumento, Arreglo_Doc)
             i_valorcodigo = Trim(TDBGridModoVta.Columns(0).Text)
            
             txt(0) = Trim(TDBGridModoVta.Columns(0).Text) 'Codigo
             txt(1) = Trim(TDBGridModoVta.Columns(1).Text)  'Desc.
             txt(2) = Trim(TDBGridModoVta.Columns(2).Text)  'Desc. Corta
             txt(11) = Trim(TDBGridModoVta.Columns(20).Text) 'item x doc.
             Text1 = Trim(TDBGridModoVta.Columns(27).Text)   'Almacenes.
             cmodo.Value = IIf(TDBGridModoVta.Columns(28) = "0" Or Len(Trim(TDBGridModoVta.Columns(28))) = 0, 0, 1)  'Canjes
             
            'Solo Emite Fact.
            If TDBGridModoVta.Columns(18).Value = False Then
                 chk(13).Value = 0
            ElseIf TDBGridModoVta.Columns(18).Value = True Then
                 chk(13).Value = 1
            End If
            'Ing. pedido
            If TDBGridModoVta.Columns(3).Value = False Then
                 chk(0).Value = 0
            ElseIf TDBGridModoVta.Columns(3).Value = True Then
                 chk(0).Value = 1
            End If
            'Ing. Hoja
            If TDBGridModoVta.Columns(8).Value = False Then
                 chk(1).Value = 0
            ElseIf TDBGridModoVta.Columns(8).Value = True Then
                 chk(1).Value = 1
            End If
            'Ing. Guia
            If TDBGridModoVta.Columns(9).Value = False Then
                 chk(2).Value = 0
            ElseIf TDBGridModoVta.Columns(9).Value = True Then
                 chk(2).Value = 1
            End If
            'Ctrl. Inv.
            If TDBGridModoVta.Columns(12).Value = False Then
                 chk(3).Value = 0
            ElseIf TDBGridModoVta.Columns(12).Value = True Then
                 chk(3).Value = 1
            End If
            'Ing.Cod.Clie.
            If TDBGridModoVta.Columns(19).Value = False Then
                 chk(6).Value = 0
            ElseIf TDBGridModoVta.Columns(19).Value = True Then
                 chk(6).Value = 1
            End If
            'Control Corr.
            If TDBGridModoVta.Columns(16).Value = False Then
                 chk(5).Value = 0
            ElseIf TDBGridModoVta.Columns(16).Value = True Then
                 chk(5).Value = 1
            End If
            'Emite Hoja Trab,
            If TDBGridModoVta.Columns(17).Value = False Then
                 chk(8).Value = 0
            ElseIf TDBGridModoVta.Columns(17).Value = True Then
                 chk(8).Value = 1
            End If
            'Imptos.
            If TDBGridModoVta.Columns(15).Value = False Then
                 chk(4).Value = 0
            ElseIf TDBGridModoVta.Columns(15).Value = True Then
                 chk(4).Value = 1
            End If
            'Act.Cta.Cte.
            If TDBGridModoVta.Columns(11).Value = False Then
                 chk(7).Value = 0
            ElseIf TDBGridModoVta.Columns(11).Value = True Then
                 chk(7).Value = 1
            End If
            'Num.Autom.
            If TDBGridModoVta.Columns(21).Value = False Then
                 chk(9).Value = 0
            ElseIf TDBGridModoVta.Columns(21).Value = True Then
                 chk(9).Value = 1
            End If
            'Ing.Forma Pago
            If TDBGridModoVta.Columns(22).Value = False Then
                 chk(10).Value = 0
            ElseIf TDBGridModoVta.Columns(22).Value = True Then
                 chk(10).Value = 1
            End If
             'Ing.Hasta Fact.
            If TDBGridModoVta.Columns(23).Value = False Then
                 chk(11).Value = 0
            ElseIf TDBGridModoVta.Columns(23).Value = True Then
                 chk(11).Value = 1
            End If
            'Ing.Hasta Hoja
            If TDBGridModoVta.Columns(24).Value = False Then
                 chk(12).Value = 0
            ElseIf TDBGridModoVta.Columns(24).Value = True Then
                 chk(12).Value = 1
            End If
             'Emite Guia
            If TDBGridModoVta.Columns(25).Value = False Then
                 chk(14).Value = 0
            ElseIf TDBGridModoVta.Columns(25).Value = True Then
                 chk(14).Value = 1
            End If
             'Factor Conversion
            If TDBGridModoVta.Columns(26).Value = False Then
                 chk(15).Value = 0
            ElseIf TDBGridModoVta.Columns(26).Value = True Then
                 chk(15).Value = 1
            End If
                 
            'Unidad Medida
            If Trim(TDBGridModoVta.Columns(13).Value) = "R" Then
                 optUnidad(0).Value = True
            ElseIf Trim(TDBGridModoVta.Columns(21).Value) = "V" Then
                 optUnidad(0).Value = True
            End If
            'Descuento
            If Trim(TDBGridModoVta.Columns(14).Value) = "O" Then
                 optDescuento(0).Value = True
            ElseIf Trim(TDBGridModoVta.Columns(14).Value) = "R" Then
                 optDescuento(1).Value = True
            End If
            
            'Copias
            If Trim(TDBGridModoVta.Columns(5).Text) <> 0 Then
                 optCopias(5).Value = True
                 txt(5).Text = Trim(TDBGridModoVta.Columns(5).Text)
            ElseIf Trim(TDBGridModoVta.Columns(6).Text) <> 0 Then
                  optCopias(6).Value = True
                  txt(6).Text = Trim(TDBGridModoVta.Columns(6).Text)
            ElseIf Trim(TDBGridModoVta.Columns(7).Text) <> 0 Then
                  optCopias(7).Value = True
                  txt(7).Text = Trim(TDBGridModoVta.Columns(7).Text)
            ElseIf Trim(TDBGridModoVta.Columns(10).Text) <> 0 Then
                  optCopias(12).Value = True
                  txt(12).Text = Trim(TDBGridModoVta.Columns(10).Text)
            ElseIf Trim(TDBGridModoVta.Columns(29).Text) <> 0 Then
                  optCopias(8).Value = True
                  txt(8).Text = Trim(TDBGridModoVta.Columns(29).Text)
            End If
                 
        modoedit = True
        SSTab1.Tab = 1
        MostrarOcultar_Botones (False)
        i_filaorigen = TDBGridModoVta.Row
        txt(0).SetFocus
      
        '''''''''
      
     Case 2   'eliminar
       If MsgBox("Desea eliminar el registro?", vbYesNo + vbDefaultButton2, "AVISO") = vbYes Then
          SQL = "DELETE FROM vt_modoventa WHERE modovtacodigo = '" & TDBGridModoVta.Columns(0).Text & "'"
          VGCNx.Execute SQL
          Mostrar_Data
       End If
        
     Case 3   'imprimir
            Call imprimir("RepvtMantModoVta.rpt")
     Case 4  ' salir
       Unload Me
  End Select
Exit Sub
CONTROLERRORES:
   If Err Then
       MsgBox Err.Number & "-" & Err.Description, vbInformation, "ERROR"
       Err = 0
       'VGgeneral.RollbackTrans
       Resume Next
    End If
End Sub

Private Sub Form_Load()
 MostrarFormVentas Me, "C2"
 Mostrar_Data
 Setear_Controles
 cAcepta.Enabled = False
 SSTab1.TabEnabled(1) = False
End Sub

Public Function Mostrar_Data()
  Dim SQL As String
  Dim rs As New ADODB.Recordset
  Dim i As Integer
      
       SQL = "SELECT modovtacodigo as 'Cód.', modovtadescripcion as Descripción," & _
      "modovtadescrcorta as 'Desc.Corta', modovtaingpedido as 'Ing.Pedido'," & _
      "documentocodigo as 'Documento', modovtacopiasfact as 'C.Fact.'," & _
      "modovtacopiasboleta as 'C.Bol.', modovtacopiasguiarem as 'C.Guia Rem.'," & _
      "modovtainghojatrab as 'Ing.HojaTrab.', modovtaingguiarem as 'Ing.Guia Rem.'," & _
      "modovtacopiashojatrab as 'C.Hoj.Trab.', modovtaactctacte as 'Act.Cta.Cte'," & _
      "modovtactrlinventario as 'Ctrl.Invent.', modovtaunidadmedida as 'Unid.Medida'," & _
      "modovtadscto as 'Dscto.', " & _
      "modovtaimpuestos as 'Imptos.', modovtacontrolcorr as 'Control.Corr.'," & _
      "modovtaemitehoja as 'Emit.Hoj.Trab.'," & _
      "modovtasolemitfact as 'Sol.Emit.Fact.',modovtaingcodclie " & _
      "as 'Ing.Cod.Clie.', modovtaitemxdoc as 'Item x Docum.'," & _
      "modovtanumautom as 'Num.Auto.'," & _
      "modovtaingformapag as 'Ing.Form.Pag.',modovtainghastafact as 'Ing.Hast.Fact'," & _
      "modovtainghastahoja as 'Ing.Hast.Hoj.Trab.',modovtaemiteguia as 'Emit.Guia', " & _
      "modovtausafactconv as 'Factor Conversión', " & _
      "modovtaalmacen as 'Almacenes', " & _
      "modovtacanje as 'Canje',modovtacopiasticket as CTicket " & _
      "FROM vt_modoventa ORDER BY modovtacodigo"
      
      Set rs = VGCNx.Execute(SQL)
      Set TDBGridModoVta.DataSource = rs
      
      '' COMBO DOCUMENTO:
      SQL = "SELECT documentocodigo,documentodescripcion " & _
      "FROM vt_documento " & _
      "ORDER BY documentocodigo "
      Set rs = VGCNx.Execute(SQL)
      If rs.RecordCount > 0 Then
        ReDim Arreglo_Doc(0 To 1, 0 To rs.RecordCount - 1)
        Call fncLlenarArreglo_Combo(rs, cmbDocumento, Arreglo_Doc, 1)
      End If
      Setear_Controles
    
 TDBGridModoVta.Refresh
 rs.Close
 Set rs = Nothing
 SSTab1.Tab = 0
  
End Function

Private Function Setear_Controles()
Dim i As Integer
Dim i_total As Integer
Dim i_width As Integer

    For i = 0 To TDBGridModoVta.Columns.Count - 1
        Select Case i
            'Case 0, 1, 2, 7, 8, 9, 10, 13, 14, 15, 16, 19, 20
            Case 1
                TDBGridModoVta.Columns(i).Width = 2000
                
            Case 0, 1, 2, 4, 5, 6, 7, 10, 13, 14, 20
                TDBGridModoVta.Columns(i).ValueItems.Presentation = dbgNormal
                TDBGridModoVta.Columns(i).Width = 600
                
            'Case 3, 4, 5, 6, 11, 12, 17, 18
            Case 3, 8, 9, 11, 12, 15, 16, 17, 18, 19, 21, 22, 23, 24, 25, 26
                TDBGridModoVta.Columns(i).ValueItems.Presentation = dbgCheckBox
                TDBGridModoVta.Columns(i).Width = 500
    End Select
    Next i
    
End Function

Private Function Validar_DatosNulos() As Boolean

Validar_Ingreso = False

                If Trim(txt(0)) <> "" And Trim(txt(1)) <> "" And Trim(txt(2)) <> "" _
                And (optUnidad(0).Value = True Or optUnidad(1).Value = True) _
                And (optDescuento(0).Value = True Or optDescuento(1).Value = True) _
                And cmbDocumento.ListIndex <> -1 Then
                    Validar_DatosNulos = True
                    Exit Function
                End If

End Function

Private Sub optCopias_Click(Index As Integer)
    cAcepta.Enabled = Validar_DatosNulos()
    txt(5).Enabled = False
    txt(6).Enabled = False
    txt(7).Enabled = False
'    txt(12).Enabled = False
    txt(5).Text = ""
    txt(6).Text = ""
    txt(7).Text = ""
'    txt(12).Text = ""
    txt(Index).Enabled = True
    txt(Index).SetFocus
End Sub

Private Sub optCopias_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub optDescuento_Click(Index As Integer)
    cAcepta.Enabled = Validar_DatosNulos()
End Sub

Private Sub optDescuento_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub optUnidad_Click(Index As Integer)
    cAcepta.Enabled = Validar_DatosNulos()
End Sub

Private Sub optUnidad_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    SSTab1.TabEnabled(PreviousTab) = False
    cAcepta.Enabled = False
End Sub

Private Sub txt_Change(Index As Integer)
    cAcepta.Enabled = Validar_DatosNulos()
End Sub
Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)  ' Salta con Enter
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    cAcepta.Enabled = Validar_DatosNulos()
    'Ingresar Mayusculas:
    If KeyAscii > 96 And KeyAscii < 123 Then
        KeyAscii = KeyAscii - 32
    End If

End Sub

Private Function Validar_CodigosDuplicados(operacion As String, Optional ByVal filaorigen As Integer) As Boolean
               
Validar_CodigosDuplicados = False
                        
   TDBGridModoVta.MoveFirst
     Do Until TDBGridModoVta.EOF
         If operacion = "INSERT" Then
            If Trim(txt(0)) = _
               Trim(TDBGridModoVta.Columns.Item(0).Value) Then
                  Validar_CodigosDuplicados = True
                  Exit Function
            End If
         ElseIf operacion = "UPDATE" Then
             If Trim(txt(0)) = _
                Trim(TDBGridModoVta.Columns.Item(0).Value) And _
                TDBGridModoVta.Row <> filaorigen Then
                   Validar_CodigosDuplicados = True
                   Exit Function
             End If
         End If
         TDBGridModoVta.MoveNext
   Loop
End Function

Private Function MostrarOcultar_Botones(valor As Boolean)
    frmbotones.Visible = valor
End Function

Private Function fncLlenarArreglo_Combo(rs As Recordset, Cbo As ComboBox, Arreglo As Variant, dimensiones As Integer)
Dim i As Integer
Dim J As Integer

    i = 0
    Cbo.Clear
    Do Until rs.EOF
        Cbo.AddItem (Trim(rs(1)))
        For J = 0 To dimensiones
            Arreglo(J, i) = Trim(rs(J))
        Next J
        rs.MoveNext
        i = i + 1
    Loop
End Function

Private Function fncSeleccionaCombo(ValorCodigo As String, Cbo As ComboBox, Arreglo As Variant)
Dim i As Integer
    For i = 0 To UBound(Arreglo, 2)
       If ValorCodigo = Arreglo(0, i) Then
         Cbo.ListIndex = i
         Exit Function
       End If
    Next i
End Function

Private Sub txt_LostFocus(Index As Integer)
Select Case Index
    Case 5, 6, 7, 12
        If txt(Index) <> "" Then
            If Not IsNumeric(txt(Index)) Then
                MsgBox "Ingrese valores numéricos...", vbInformation, "AVISO"
                txt(Index) = ""
            End If
        End If
End Select

If txt(Index) <> "" Then
    If Index = 0 Then
        Call Formatear_Codigo(Index)
    End If
End If

End Sub
Public Function Formatear_Codigo(indice As Integer) As String
Dim cadena As String
Dim i As Integer

cadena = ""
For i = 0 To txt(indice).MaxLength
    cadena = cadena & "0"
Next i

txt(indice) = Right(cadena & Trim(txt(indice)), txt(indice).MaxLength)

End Function

Private Function Validar_DescripcionesDuplicadas(operacion As String, Optional ByVal filaorigen As Integer) As Boolean
               
Validar_DescripcionesDuplicadas = False
                        
    TDBGridModoVta.MoveFirst
    Do Until TDBGridModoVta.EOF
        If operacion = "INSERT" Then
           If Trim(txt(1)) = _
              Trim(TDBGridModoVta.Columns.Item(1).Value) Then
                Validar_DescripcionesDuplicadas = True
                Exit Function
           End If
        ElseIf operacion = "UPDATE" Then
           If Trim(txt(1)) = _
              Trim(TDBGridModoVta.Columns.Item(1).Value) And _
              TDBGridModoVta.Row <> filaorigen Then
                 Validar_DescripcionesDuplicadas = True
                 Exit Function
           End If
        End If
        TDBGridModoVta.MoveNext
    Loop
               
End Function


