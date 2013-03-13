VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmPedido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pedido"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   12960
   Begin TabDlg.SSTab SSTab1 
      Height          =   7815
      Left            =   510
      TabIndex        =   0
      Top             =   360
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   13785
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   12632256
      TabCaption(0)   =   "Consulta"
      TabPicture(0)   =   "FrmPedido.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "TDBGrid2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Datos Generales"
      TabPicture(1)   =   "FrmPedido.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "Fr2(0)"
      Tab(1).Control(2)=   "Fr1"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Detalle Pedido"
      TabPicture(2)   =   "FrmPedido.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label4"
      Tab(2).Control(1)=   "TDBGrid1"
      Tab(2).Control(2)=   "Fr2(1)"
      Tab(2).Control(3)=   "Fr2(2)"
      Tab(2).ControlCount=   4
      Begin VB.Frame Fr1 
         Height          =   2415
         Left            =   -74610
         TabIndex        =   65
         Top             =   540
         Width           =   11535
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   5760
            Style           =   2  'Dropdown List
            TabIndex        =   67
            Top             =   1320
            Width           =   1455
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   9840
            Style           =   2  'Dropdown List
            TabIndex        =   66
            Top             =   1320
            Width           =   1455
         End
         Begin MSMask.MaskEdBox MBox 
            Height          =   255
            Index           =   0
            Left            =   1800
            TabIndex        =   68
            Top             =   240
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   2
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox 
            Height          =   255
            Index           =   1
            Left            =   9840
            TabIndex        =   69
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   8
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox 
            Height          =   255
            Index           =   2
            Left            =   1800
            TabIndex        =   70
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   8
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox 
            Height          =   255
            Index           =   3
            Left            =   5760
            TabIndex        =   71
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   8
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox 
            Height          =   255
            Index           =   4
            Left            =   9840
            TabIndex        =   72
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   8
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox 
            Height          =   255
            Index           =   5
            Left            =   1800
            TabIndex        =   73
            Top             =   960
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   10
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox 
            Height          =   255
            Index           =   6
            Left            =   5760
            TabIndex        =   74
            Top             =   960
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   10
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox 
            Height          =   255
            Index           =   7
            Left            =   9840
            TabIndex        =   75
            Top             =   960
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   10
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox 
            Height          =   255
            Index           =   8
            Left            =   1800
            TabIndex        =   76
            Top             =   1320
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   10
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox 
            Height          =   255
            Index           =   9
            Left            =   1800
            TabIndex        =   77
            Top             =   1920
            Width           =   9375
            _ExtentX        =   16536
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   45
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Caption         =   "Punto Venta"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   89
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "No .Factura"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   88
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Dcto. Genral."
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   87
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo de Cambio"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   86
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "No .Boleta"
            Height          =   255
            Index           =   4
            Left            =   4560
            TabIndex        =   85
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Dcto. Promoc."
            Height          =   255
            Index           =   5
            Left            =   4560
            TabIndex        =   84
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Moneda"
            Height          =   255
            Index           =   6
            Left            =   4560
            TabIndex        =   83
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "No. Pedido"
            Height          =   255
            Index           =   7
            Left            =   8280
            TabIndex        =   82
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "No. Guia Remision"
            Height          =   255
            Index           =   8
            Left            =   8280
            TabIndex        =   81
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Dcto. Especial"
            Height          =   255
            Index           =   9
            Left            =   8280
            TabIndex        =   80
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Lista Precios"
            Height          =   255
            Index           =   10
            Left            =   8280
            TabIndex        =   79
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Mensajes"
            Height          =   255
            Index           =   11
            Left            =   240
            TabIndex        =   78
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00008080&
            Index           =   0
            X1              =   0
            X2              =   11520
            Y1              =   1750
            Y2              =   1750
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   1
            X1              =   0
            X2              =   11520
            Y1              =   1765
            Y2              =   1765
         End
      End
      Begin VB.Frame Fr2 
         Height          =   2535
         Index           =   0
         Left            =   -74610
         TabIndex        =   36
         Top             =   3060
         Width           =   11535
         Begin VB.ComboBox Combo3 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   180
            Width           =   2655
         End
         Begin VB.ComboBox Combo4 
            Height          =   315
            Left            =   9120
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   180
            Width           =   2175
         End
         Begin VB.ComboBox Combo5 
            Height          =   315
            Left            =   8070
            TabIndex        =   40
            Text            =   "Combo5"
            Top             =   2040
            Width           =   975
         End
         Begin VB.CommandButton cAyuda 
            Caption         =   "..."
            Height          =   255
            Index           =   0
            Left            =   3480
            TabIndex        =   39
            Top             =   600
            Width           =   375
         End
         Begin VB.CommandButton cAyuda 
            Caption         =   "..."
            Height          =   255
            Index           =   1
            Left            =   2640
            TabIndex        =   38
            Top             =   960
            Width           =   375
         End
         Begin VB.CommandButton cAyuda 
            Caption         =   "..."
            Height          =   255
            Index           =   2
            Left            =   2400
            TabIndex        =   37
            Top             =   1320
            Width           =   375
         End
         Begin MSMask.MaskEdBox MBox 
            Height          =   255
            Index           =   10
            Left            =   6240
            TabIndex        =   43
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox 
            Height          =   255
            Index           =   11
            Left            =   1920
            TabIndex        =   44
            Top             =   600
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   11
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox 
            Height          =   255
            Index           =   12
            Left            =   1920
            TabIndex        =   45
            Top             =   960
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   3
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox 
            Height          =   255
            Index           =   13
            Left            =   9840
            TabIndex        =   46
            Top             =   960
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   10
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox 
            Height          =   255
            Index           =   14
            Left            =   1920
            TabIndex        =   47
            Top             =   1320
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   2
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox 
            Height          =   255
            Index           =   15
            Left            =   1920
            TabIndex        =   48
            Top             =   1680
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   10
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox 
            Height          =   255
            Index           =   16
            Left            =   1920
            TabIndex        =   49
            Top             =   2040
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   10
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox 
            Height          =   255
            Index           =   17
            Left            =   5190
            TabIndex        =   50
            Top             =   2040
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   10
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox 
            Height          =   255
            Index           =   18
            Left            =   10410
            TabIndex        =   91
            Top             =   2040
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   3
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Caption         =   "Dias Pago"
            Height          =   255
            Index           =   18
            Left            =   9570
            TabIndex        =   92
            Top             =   2070
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Modo de la Venta"
            Height          =   255
            Index           =   12
            Left            =   240
            TabIndex        =   64
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha de Atencion"
            Height          =   255
            Index           =   13
            Left            =   4800
            TabIndex        =   63
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Forma de Pago"
            Height          =   255
            Index           =   14
            Left            =   7920
            TabIndex        =   62
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Codigo del Cliente"
            Height          =   255
            Index           =   15
            Left            =   240
            TabIndex        =   61
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Codigo del Vendedor"
            Height          =   255
            Index           =   16
            Left            =   240
            TabIndex        =   60
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Codigo del Almacen"
            Height          =   255
            Index           =   17
            Left            =   240
            TabIndex        =   59
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Otros Gastos"
            Height          =   255
            Index           =   19
            Left            =   240
            TabIndex        =   58
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Nota de Pedido"
            Height          =   255
            Index           =   20
            Left            =   240
            TabIndex        =   57
            Top             =   2040
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Orden de Compra"
            Height          =   255
            Index           =   21
            Left            =   3810
            TabIndex        =   56
            Top             =   2040
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Autorizacion"
            Height          =   255
            Index           =   22
            Left            =   7080
            TabIndex        =   55
            Top             =   2040
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "% Comision"
            Height          =   255
            Index           =   23
            Left            =   8880
            TabIndex        =   54
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Etiq 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label2(0)"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   3840
            TabIndex        =   53
            Top             =   600
            Width           =   7455
         End
         Begin VB.Label Etiq 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label2(1)"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   3000
            TabIndex        =   52
            Top             =   960
            Width           =   5775
         End
         Begin VB.Label Etiq 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label2(2)"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   2760
            TabIndex        =   51
            Top             =   1320
            Width           =   8535
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         BorderStyle     =   0  'None
         Caption         =   "Detalles del Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   -74610
         TabIndex        =   29
         Top             =   5820
         Width           =   11535
         Begin VB.Label Dclie 
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente"
            ForeColor       =   &H8000000C&
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   35
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label Dclie 
            BackStyle       =   0  'Transparent
            Caption         =   "Direccion"
            ForeColor       =   &H8000000C&
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   34
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Dclie 
            BackStyle       =   0  'Transparent
            Caption         =   "Distrito"
            ForeColor       =   &H8000000C&
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   33
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Dclie 
            BackStyle       =   0  'Transparent
            Caption         =   "Provincia"
            ForeColor       =   &H8000000C&
            Height          =   255
            Index           =   3
            Left            =   7920
            TabIndex        =   32
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Dclie 
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo US$"
            ForeColor       =   &H8000000C&
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   31
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label Dclie 
            BackStyle       =   0  'Transparent
            Caption         =   "Limite Cred US$"
            ForeColor       =   &H8000000C&
            Height          =   255
            Index           =   5
            Left            =   7920
            TabIndex        =   30
            Top             =   1200
            Width           =   1815
         End
      End
      Begin VB.Frame Fr2 
         BackColor       =   &H00808000&
         BorderStyle     =   0  'None
         Height          =   975
         Index           =   2
         Left            =   -74670
         TabIndex        =   17
         Top             =   6510
         Width           =   11535
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   6
            Left            =   360
            TabIndex        =   18
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            BackColor       =   49344
            ForeColor       =   16777215
            Enabled         =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   7
            Left            =   2520
            TabIndex        =   19
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            BackColor       =   49344
            ForeColor       =   16777215
            Enabled         =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   8
            Left            =   4920
            TabIndex        =   20
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            BackColor       =   49344
            ForeColor       =   16777215
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   9
            Left            =   7440
            TabIndex        =   21
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            BackColor       =   49344
            ForeColor       =   16777215
            Enabled         =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   10
            Left            =   9720
            TabIndex        =   22
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            BackColor       =   49344
            ForeColor       =   16777215
            Enabled         =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FF80&
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   27
            Top             =   675
            Width           =   1335
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Total Bruto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FF80&
            Height          =   255
            Index           =   1
            Left            =   2760
            TabIndex        =   26
            Top             =   680
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Total Dctos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FF80&
            Height          =   255
            Index           =   2
            Left            =   5160
            TabIndex        =   25
            Top             =   680
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Total I.G.V."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FF80&
            Height          =   255
            Index           =   3
            Left            =   7680
            TabIndex        =   24
            Top             =   680
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Neto Factura"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FF80&
            Height          =   255
            Index           =   4
            Left            =   9840
            TabIndex        =   23
            Top             =   680
            Width           =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            Index           =   0
            X1              =   2175
            X2              =   2175
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            Index           =   1
            X1              =   4440
            X2              =   4440
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            Index           =   2
            X1              =   6960
            X2              =   6960
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            Index           =   3
            X1              =   9360
            X2              =   9360
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            Index           =   4
            X1              =   2160
            X2              =   2160
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            Index           =   5
            X1              =   4420
            X2              =   4420
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            Index           =   6
            X1              =   6940
            X2              =   6940
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            Index           =   7
            X1              =   9340
            X2              =   9340
            Y1              =   120
            Y2              =   1215
         End
      End
      Begin VB.Frame Fr2 
         Height          =   1095
         Index           =   1
         Left            =   -74670
         TabIndex        =   1
         Top             =   510
         Width           =   11535
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   2
            Top             =   480
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   3
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   1
            Left            =   960
            TabIndex        =   3
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   8
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   2
            Left            =   6840
            TabIndex        =   4
            Top             =   480
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   8
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   3
            Left            =   7920
            TabIndex        =   5
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   8
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   4
            Left            =   9360
            TabIndex        =   6
            Top             =   480
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   8
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   5
            Left            =   10320
            TabIndex        =   7
            Top             =   480
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   8
            PromptChar      =   "_"
         End
         Begin VB.Label deta 
            Alignment       =   2  'Center
            Caption         =   "Codigo"
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
            Index           =   0
            Left            =   960
            TabIndex        =   15
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label deta 
            Alignment       =   2  'Center
            Caption         =   "Descripción"
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
            Index           =   1
            Left            =   2400
            TabIndex        =   14
            Top             =   240
            Width           =   4335
         End
         Begin VB.Label deta 
            Alignment       =   2  'Center
            Caption         =   "U.M."
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
            Index           =   2
            Left            =   6960
            TabIndex        =   13
            Top             =   240
            Width           =   855
         End
         Begin VB.Label deta 
            Alignment       =   2  'Center
            Caption         =   "Precio Vta"
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
            Index           =   3
            Left            =   7920
            TabIndex        =   12
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label deta 
            Alignment       =   2  'Center
            Caption         =   "Dscto"
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
            Index           =   4
            Left            =   9480
            TabIndex        =   11
            Top             =   240
            Width           =   735
         End
         Begin VB.Label deta 
            Alignment       =   2  'Center
            Caption         =   "%Com"
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
            Index           =   5
            Left            =   10320
            TabIndex        =   10
            Top             =   240
            Width           =   975
         End
         Begin VB.Label deta 
            Alignment       =   2  'Center
            Caption         =   "Cant."
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
            Index           =   6
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label2"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   2400
            TabIndex        =   8
            Top             =   480
            Width           =   4335
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   4455
         Left            =   -74670
         TabIndex        =   16
         Top             =   1710
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   7858
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Item"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Codigo"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Descripcion"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Cant"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Precio Vta"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "% Dscto"
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Total Item"
         Columns(6).DataField=   ""
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "%"
         Columns(7).DataField=   ""
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   8
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(21)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(25)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(29)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
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
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
         _StyleDefs(68)  =   "Named:id=33:Normal"
         _StyleDefs(69)  =   ":id=33,.parent=0"
         _StyleDefs(70)  =   "Named:id=34:Heading"
         _StyleDefs(71)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(72)  =   ":id=34,.wraptext=-1"
         _StyleDefs(73)  =   "Named:id=35:Footing"
         _StyleDefs(74)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(75)  =   "Named:id=36:Selected"
         _StyleDefs(76)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(77)  =   "Named:id=37:Caption"
         _StyleDefs(78)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(79)  =   "Named:id=38:HighlightRow"
         _StyleDefs(80)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(81)  =   "Named:id=39:EvenRow"
         _StyleDefs(82)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(83)  =   "Named:id=40:OddRow"
         _StyleDefs(84)  =   ":id=40,.parent=33"
         _StyleDefs(85)  =   "Named:id=41:RecordSelector"
         _StyleDefs(86)  =   ":id=41,.parent=34"
         _StyleDefs(87)  =   "Named:id=42:FilterBar"
         _StyleDefs(88)  =   ":id=42,.parent=33"
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
         Height          =   6525
         Left            =   330
         TabIndex        =   90
         Top             =   690
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   11509
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Proceso Venta"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "No. Pedido"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Fecha Atn."
         Columns(2).DataField=   ""
         Columns(2).NumberFormat=   "Short Date"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Cotizacion"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Descripcion Referencial"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   5
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=5"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=260"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=260"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=260"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=260"
         Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(25)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=260"
         Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
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
         DeadAreaBackColor=   12632256
         RowDividerColor =   12632256
         RowSubDividerColor=   12632256
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   1785.26
         ViewColumnWidth =   9689.953
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=18,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HC0FFFF&,.bold=0,.fontsize=825"
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
         _StyleDefs(24)  =   "Splits(0).Style:id=67,.parent=1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=76,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=68,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=69,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=70,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=72,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=71,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=73,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=74,.parent=9"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=75,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=77,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=78,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=82,.parent=67"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=79,.parent=68,.alignment=0"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=80,.parent=69"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=81,.parent=71"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=86,.parent=67"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=83,.parent=68,.alignment=0"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=84,.parent=69"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=85,.parent=71"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=102,.parent=67"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=99,.parent=68,.alignment=0"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=100,.parent=69"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=101,.parent=71"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=106,.parent=67"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=103,.parent=68,.alignment=0"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=104,.parent=69"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=105,.parent=71"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=110,.parent=67"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=107,.parent=68,.alignment=0"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=108,.parent=69"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=109,.parent=71"
         _StyleDefs(56)  =   "Named:id=33:Normal"
         _StyleDefs(57)  =   ":id=33,.parent=0"
         _StyleDefs(58)  =   "Named:id=34:Heading"
         _StyleDefs(59)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(60)  =   ":id=34,.wraptext=-1"
         _StyleDefs(61)  =   "Named:id=35:Footing"
         _StyleDefs(62)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(63)  =   "Named:id=36:Selected"
         _StyleDefs(64)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(65)  =   "Named:id=37:Caption"
         _StyleDefs(66)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(67)  =   "Named:id=38:HighlightRow"
         _StyleDefs(68)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(69)  =   "Named:id=39:EvenRow"
         _StyleDefs(70)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(71)  =   "Named:id=40:OddRow"
         _StyleDefs(72)  =   ":id=40,.parent=33"
         _StyleDefs(73)  =   "Named:id=41:RecordSelector"
         _StyleDefs(74)  =   ":id=41,.parent=34"
         _StyleDefs(75)  =   "Named:id=42:FilterBar"
         _StyleDefs(76)  =   ":id=42,.parent=33"
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "<< T O T A L E S  D E  L A  V E N T A  >>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -74670
         TabIndex        =   28
         Top             =   6210
         Width           =   11505
      End
   End
End
Attribute VB_Name = "FrmPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nLongicampo(7) As Integer

Private Sub Form_Load()
   
   nLongicampo(1) = 0  '1000: nLongicampo(2) = 600:  nLongicampo(3) = 1200:   nLongicampo(4) = 3500:   nLongicampo(5) = 600:  nLongicampo(6) = 1200
   Listar Cn, "pedido", TDBGrid2, "pedidofecha as Fecha, documentocodigo as Tipo,seriedocnumero+'-'+pedidonumero as Documento,clienterazonsocial as Cliente,pedidomoneda as Mnd,pedidototbruto as Total", "pedidofecha,documentocodigo", nLongicampo
   
End Sub

