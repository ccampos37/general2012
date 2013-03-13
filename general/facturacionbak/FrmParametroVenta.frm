VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmParametroVenta 
   Caption         =   "Parametros de Venta"
   ClientHeight    =   7725
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmbotones 
      Height          =   1050
      Left            =   3480
      TabIndex        =   28
      Top             =   6600
      Width           =   5655
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
         Picture         =   "FrmParametroVenta.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   33
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
         Picture         =   "FrmParametroVenta.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   32
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
         Picture         =   "FrmParametroVenta.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   180
         Width           =   870
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
         Picture         =   "FrmParametroVenta.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   180
         Width           =   915
      End
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
         Picture         =   "FrmParametroVenta.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   180
         Width           =   915
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   11245
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
      TabPicture(0)   =   "FrmParametroVenta.frx":154A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "TDBGridModoVta"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Mantenimiento"
      TabPicture(1)   =   "FrmParametroVenta.frx":1566
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   5655
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   11625
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
            Height          =   360
            Index           =   9
            Left            =   8040
            MaxLength       =   8
            TabIndex        =   50
            Top             =   2640
            Width           =   585
         End
         Begin VB.Frame FraChk 
            Height          =   615
            Left            =   240
            TabIndex        =   43
            Top             =   3600
            Width           =   11055
            Begin VB.CheckBox Chk 
               Height          =   195
               Index           =   6
               Left            =   9240
               TabIndex        =   47
               Top             =   240
               Width           =   795
            End
            Begin VB.CheckBox Chk 
               Height          =   255
               Index           =   4
               Left            =   3840
               TabIndex        =   17
               Top             =   240
               Width           =   375
            End
            Begin VB.CheckBox Chk 
               Height          =   255
               Index           =   3
               Left            =   7560
               TabIndex        =   16
               Top             =   240
               Width           =   255
            End
            Begin VB.CheckBox Chk 
               Height          =   255
               Index           =   2
               Left            =   1800
               TabIndex        =   15
               Top             =   240
               Width           =   375
            End
            Begin VB.CheckBox Chk 
               Height          =   255
               Index           =   5
               Left            =   5280
               TabIndex        =   14
               Top             =   240
               Width           =   375
            End
            Begin VB.Label lbl 
               Caption         =   "Comis.Vendedor"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   24
               Left            =   5880
               TabIndex        =   49
               Top             =   240
               Width           =   1725
            End
            Begin VB.Label Label1 
               Caption         =   "Ingreso Masivo"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   8040
               TabIndex        =   48
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label lbl 
               Caption         =   "Forma Emisión"
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
               Left            =   2280
               TabIndex        =   46
               Top             =   240
               Width           =   1725
            End
            Begin VB.Label lbl 
               Caption         =   "Lista Precios"
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
               Left            =   360
               TabIndex        =   45
               Top             =   240
               Width           =   1395
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
               Height          =   285
               Index           =   19
               Left            =   4560
               TabIndex        =   44
               Top             =   240
               Width           =   675
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
            Height          =   360
            Index           =   8
            Left            =   8040
            MaxLength       =   8
            TabIndex        =   13
            Top             =   1920
            Width           =   2025
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
            Index           =   7
            Left            =   2040
            MaxLength       =   70
            MultiLine       =   -1  'True
            TabIndex        =   8
            Top             =   4320
            Width           =   9225
         End
         Begin VB.ComboBox cmbAlmacen 
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
            TabIndex        =   7
            Top             =   3120
            Width           =   3255
         End
         Begin VB.CheckBox Chk 
            Height          =   255
            Index           =   1
            Left            =   2040
            TabIndex        =   5
            Top             =   2280
            Width           =   375
         End
         Begin VB.ComboBox cmbMoneda 
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
            Left            =   8040
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   1440
            Width           =   3255
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
            Height          =   360
            Index           =   6
            Left            =   2040
            MaxLength       =   8
            TabIndex        =   6
            Top             =   2640
            Width           =   2025
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
            Height          =   360
            Index           =   5
            Left            =   8040
            MaxLength       =   10
            TabIndex        =   11
            Top             =   1005
            Width           =   2025
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
            Height          =   360
            Index           =   4
            Left            =   8040
            MaxLength       =   25
            TabIndex        =   10
            Top             =   600
            Width           =   2025
         End
         Begin VB.CheckBox Chk 
            Height          =   375
            Index           =   0
            Left            =   2040
            TabIndex        =   3
            Top             =   1560
            Width           =   375
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
            Height          =   360
            Index           =   2
            Left            =   8040
            MaxLength       =   30
            TabIndex        =   9
            Top             =   240
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
            Height          =   360
            Index           =   1
            Left            =   2040
            MaxLength       =   15
            TabIndex        =   2
            Top             =   1200
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
            Height          =   360
            Index           =   3
            Left            =   2040
            MaxLength       =   8
            TabIndex        =   4
            Top             =   1920
            Width           =   2025
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
            Left            =   6240
            TabIndex        =   19
            Top             =   4920
            Width           =   1335
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
            Left            =   4560
            TabIndex        =   18
            Top             =   4920
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
            Height          =   360
            Index           =   0
            Left            =   2040
            MaxLength       =   35
            TabIndex        =   1
            Top             =   720
            Width           =   3225
         End
         Begin VB.ComboBox cmbEmpresa 
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
            TabIndex        =   0
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label lbl 
            Caption         =   "Codigo Transaccion ventas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Index           =   9
            Left            =   5880
            TabIndex        =   51
            Top             =   2520
            Width           =   1995
         End
         Begin VB.Label lbl 
            Caption         =   "Tip.Camb.Refer."
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
            Left            =   6360
            TabIndex        =   42
            Top             =   2040
            Width           =   1635
         End
         Begin VB.Label lbl 
            Caption         =   "Mensaje"
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
            Left            =   360
            TabIndex        =   41
            Top             =   4440
            Width           =   1395
         End
         Begin VB.Label lbl 
            Caption         =   "Almacen"
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
            Left            =   360
            TabIndex        =   40
            Top             =   3120
            Width           =   960
         End
         Begin VB.Label lbl 
            Caption         =   "IGV"
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
            Left            =   360
            TabIndex        =   39
            Top             =   2280
            Width           =   555
         End
         Begin VB.Label lbl 
            Caption         =   "Moneda"
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
            Index           =   7
            Left            =   6360
            TabIndex        =   38
            Top             =   1560
            Width           =   960
         End
         Begin VB.Label lbl 
            Caption         =   "Tasa IGV"
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
            Index           =   6
            Left            =   360
            TabIndex        =   37
            Top             =   2640
            Width           =   1275
         End
         Begin VB.Label lbl 
            Caption         =   "Fax"
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
            Index           =   5
            Left            =   6360
            TabIndex        =   36
            Top             =   1080
            Width           =   1035
         End
         Begin VB.Label lbl 
            Caption         =   "Telefonos"
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
            Index           =   3
            Left            =   6360
            TabIndex        =   35
            Top             =   720
            Width           =   1635
         End
         Begin VB.Label lbl 
            Caption         =   "Descuento"
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
            Left            =   360
            TabIndex        =   34
            Top             =   1560
            Width           =   1395
         End
         Begin VB.Label lbl 
            Caption         =   "Tasa Dscto."
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
            TabIndex        =   26
            Top             =   1920
            Width           =   1635
         End
         Begin VB.Label lbl 
            Caption         =   "Dirección"
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
            Left            =   6360
            TabIndex        =   25
            Top             =   360
            Width           =   1365
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
            TabIndex        =   24
            Top             =   1200
            Width           =   1485
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
            TabIndex        =   23
            Top             =   840
            Width           =   1275
         End
         Begin VB.Label lbl 
            Caption         =   "Empresa"
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
            TabIndex        =   22
            Top             =   360
            Width           =   960
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGridModoVta 
         Height          =   5655
         Left            =   -74880
         TabIndex        =   27
         Top             =   480
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   9975
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
End
Attribute VB_Name = "FrmParametroVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim modoinsert As Boolean
Dim modoedit As Boolean
Dim i_filaorigen As Integer
Dim i_valorcodigo As String
Dim ArregloEmpresa()
Dim ArregloMoneda()
Dim ArregloAlmacen()

Private Sub cAcepta_Click()
   Dim rs As New ADODB.Recordset
   Dim SQL As String
   Dim J As Integer
   
   Dim s_codigoempresa, s_codigomoneda, s_codigoalmacen As String
   Dim d_descuento, d_igv, d_tipocambio As Double
   
   On Error GoTo CONTROLERRORES

   ''''''''
                If Validar_IGV_DSCTO_TIPO = False Then
                    Exit Sub
                End If
                
                If txt(3) = "" Then
                    d_descuento = 0
                Else
                    d_descuento = txt(3) / 100
                End If
                If txt(6) = "" Then
                    d_igv = 0
                Else
                    d_igv = txt(6)
                End If
                If txt(8) = "" Then d_tipocambio = 0 Else d_tipocambio = txt(8)
                
                If cmbEmpresa.ListIndex <> -1 Then
                    s_codigoempresa = ArregloEmpresa(0, cmbEmpresa.ListIndex)
                Else
                    s_codigoempresa = ""
                End If
                If cmbMoneda.ListIndex <> -1 Then
                    s_codigomoneda = ArregloMoneda(0, cmbMoneda.ListIndex)
                Else
                    s_codigomoneda = ""
                End If
                If cmbAlmacen.ListIndex <> -1 Then
                    s_codigoalmacen = ArregloAlmacen(0, cmbAlmacen.ListIndex)
                Else
                    s_codigoalmacen = ""
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
             
          SQL = "INSERT INTO vt_parametroventa " & _
               "(empresacodigo,paramvtadesc,paramvtadescor,paramvtadirec," & _
               "paramvtaestdesc,paramvtadescto,paramvtatelefonos,paramvtafax," & _
               "monedacodigo,paramvtaestigv,paramvtaporcigv,almacencodigo,paramvtamensaje," & _
               "paramvtalistaprec,paramvtatipcambref,paramvtacomisionvendedor," & _
               "paramvtaformaemision,paramvtaboleta,paramvtamasivo," & _
               "codigotransaccionventas,usuariocodigo,fechaact) " & _
               " VALUES " & _
               "('" & s_codigoempresa & "','" & txt(0) & "','" & txt(1) & "','" & txt(2) & "'," & _
               chk(0).Value & "," & d_descuento & ",'" & txt(4) & "','" & txt(5) & "','" & _
               s_codigomoneda & "'," & chk(1).Value & "," & d_igv & ",'" & s_codigoalmacen & "','" & txt(7) & "'," & _
               chk(2).Value & "," & d_tipocambio & "," & chk(3).Value & "," & _
               chk(4).Value & "," & chk(5).Value & "," & chk(6).Value & ",'" & _
               txt(6) & "','" & g_usuario & "','" & Date & "')"

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
                          
            SQL = "UPDATE vt_parametroventa SET " & _
               "empresacodigo='" & s_codigoempresa & "',paramvtadesc='" & txt(0) & "'," & _
               "paramvtadescor='" & txt(1) & "'," & _
               "paramvtadirec='" & txt(2) & "',paramvtaestdesc=" & chk(0).Value & _
               ",paramvtadescto=" & d_descuento & ",paramvtatelefonos='" & txt(4) & _
               "',paramvtafax='" & txt(5) & "',monedacodigo='" & s_codigomoneda & "',paramvtaestigv=" & chk(1).Value & _
               ",paramvtaporcigv=" & d_igv & ",almacencodigo='" & s_codigoalmacen & "'," & _
               "paramvtamensaje='" & txt(7) & "',paramvtalistaprec=" & chk(2).Value & _
               ",paramvtatipcambref=" & d_tipocambio & ",paramvtacomisionvendedor=" & chk(3).Value & _
               ",paramvtaformaemision=" & chk(4).Value & ",paramvtaboleta=" & chk(5).Value & _
               ",paramvtamasivo=" & chk(6).Value & _
               ",fechaact='" & Date & "',usuariocodigo='" & g_usuario & "' " & _
               "WHERE empresacodigo ='" & i_valorcodigo & "'"
               
            VGCNx.Execute SQL
        
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
    Select Case Index
    Case 0
        If chk(Index).Value = 0 Then
            txt(3) = ""
        End If
    Case 1
        If chk(Index).Value = 0 Then
            txt(6) = ""
        End If
    End Select
    cAcepta.Enabled = Validar_DatosNulos()
End Sub

Private Sub cmbDocumento_Click()
    cAcepta.Enabled = Validar_DatosNulos()
End Sub

Private Sub chk_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmbAlmacen_Click()
cAcepta.Enabled = Validar_DatosNulos()
End Sub

Private Sub cmbAlmacen_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmbEmpresa_Click()
cAcepta.Enabled = Validar_DatosNulos()
End Sub

Private Sub cmbEmpresa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmbMoneda_Click()
cAcepta.Enabled = Validar_DatosNulos()
End Sub

Private Sub cmbMoneda_KeyDown(KeyCode As Integer, Shift As Integer)
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
                  ElseIf TypeOf OBJ Is ComboBox Then
                    OBJ.ListIndex = -1
                  End If
            Next
            SSTab1.Tab = 1
            modoinsert = True
            MostrarOcultar_Botones (False)
            cmbEmpresa.SetFocus
        
     Case 1   'modificar
     
         If TDBGridModoVta.Row < 0 Then
            Exit Sub
         End If
          
         i_valorcodigo = Trim(TDBGridModoVta.Columns(0).Text)
         
         Call fncSeleccionaCombo(Trim(TDBGridModoVta.Columns(0).Text), cmbEmpresa, ArregloEmpresa)
         Call fncSeleccionaCombo(Trim(TDBGridModoVta.Columns(7).Text), cmbMoneda, ArregloMoneda)
         Call fncSeleccionaCombo(Trim(TDBGridModoVta.Columns(13).Text), cmbAlmacen, ArregloAlmacen)
            
         txt(0) = Trim(TDBGridModoVta.Columns(2).Text)  'Desc.
         txt(1) = Trim(TDBGridModoVta.Columns(3).Text)  'Desc. Corta
         txt(2) = Trim(TDBGridModoVta.Columns(4).Text)  'Direccion
         txt(3) = Trim(TDBGridModoVta.Columns(10).Text) 'Dscto.
         txt(4) = Trim(TDBGridModoVta.Columns(5).Text)  'Telef.
         txt(5) = Trim(TDBGridModoVta.Columns(6).Text)  'Fax
         txt(6) = Trim(TDBGridModoVta.Columns(12).Text) '% IGV
         txt(7) = Trim(TDBGridModoVta.Columns(15).Text) 'Mensaje
         txt(8) = Trim(TDBGridModoVta.Columns(20).Text) 'Tip.Cambio
         txt(9) = Trim(TDBGridModoVta.Columns(22).Text) ' codigo transaccion vetas
         
         'Dscto.
         If TDBGridModoVta.Columns(9).Value = False Then
            chk(0).Value = 0
         ElseIf TDBGridModoVta.Columns(9).Value = True Then
            chk(0).Value = 1
         End If
         'IGV
         If TDBGridModoVta.Columns(11).Value = False Then
            chk(1).Value = 0
         ElseIf TDBGridModoVta.Columns(11).Value = True Then
            chk(1).Value = 1
         End If
         'Lista Precios
         If TDBGridModoVta.Columns(16).Value = False Then
            chk(2).Value = 0
         ElseIf TDBGridModoVta.Columns(16).Value = True Then
            chk(2).Value = 1
            End If
         'Com. Vendedor
         If TDBGridModoVta.Columns(17).Value = False Then
            chk(3).Value = 0
         ElseIf TDBGridModoVta.Columns(17).Value = True Then
            chk(3).Value = 1
         End If
         'Forma Emision
         If TDBGridModoVta.Columns(18).Value = False Then
            chk(4).Value = 0
         ElseIf TDBGridModoVta.Columns(18).Value = True Then
            chk(4).Value = 1
         End If
         'Boleta
         If TDBGridModoVta.Columns(19).Value = False Then
            chk(5).Value = 0
         ElseIf TDBGridModoVta.Columns(19).Value = True Then
            chk(5).Value = 1
         End If
         'ingreso masivo
         If TDBGridModoVta.Columns(20).Value = False Then
            chk(6).Value = 0
         ElseIf TDBGridModoVta.Columns(20).Value = True Then
            chk(6).Value = 1
         End If
                 
        modoedit = True
        SSTab1.Tab = 1
        MostrarOcultar_Botones (False)
        i_filaorigen = TDBGridModoVta.Row
        cmbEmpresa.SetFocus
      
        '''''''''
      
     Case 2   'eliminar
       If MsgBox("Desea eliminar el registro?", vbYesNo + vbDefaultButton2, "AVISO") = vbYes Then
          SQL = "DELETE FROM vt_parametroventa WHERE empresacodigo = " & TDBGridModoVta.Columns(0).Text
          VGCNx.Execute SQL
          Mostrar_Data
       End If
        
     Case 3   'imprimir
            Call Imprimir("RepvtParamVta.rpt")
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
 MostrarForm Me, "C2"
 Mostrar_Data
 Setear_Controles
 cAcepta.Enabled = False
 SSTab1.TabEnabled(1) = False
End Sub

Public Function Mostrar_Data()
  Dim SQL As String
  Dim rs As New ADODB.Recordset
  Dim i As Integer
      
       SQL = "SELECT a.empresacodigo as 'Cód.Emp.', b.empresadescripcion as 'Desc.Emp.'," & _
      "a.paramvtadesc as 'Desc.Param', a.paramvtadescor as 'Desc.Corta'," & _
      "a.paramvtadirec as 'Direcc.', a.paramvtatelefonos as 'Telef.'," & _
      "a.paramvtafax as 'Fax', a.monedacodigo as 'Cod.Mon.'," & _
      "c.monedadescripcion as 'Desc.Mon.', a.paramvtaestdesc as 'Dscto.'," & _
      "a.paramvtadescto*100 as 'Tasa Dscto.', a.paramvtaestigv as IGV," & _
      "a.paramvtaporcigv as 'Tasa IGV', a.almacencodigo as 'Cod.Alm.'," & _
      "d.almacendescripcion as 'Desc.Alm.'," & _
      "a.paramvtamensaje as 'Mensaje', a.paramvtalistaprec as 'List.Prec.'," & _
      "a.paramvtacomisionvendedor as 'Comis.Vend.'," & _
      "a.paramvtaformaemision as 'Form.Emis.',a.paramvtaboleta " & _
      "as 'Boleta', a.paramvtatipcambref as 'Tip.Camb.Ref.',a.paramvtamasivo as 'Ing. Masivo', " & _
      "a.codigotransaccionventa as 'Cod.Trans.Vtas' FROM vt_parametroventa a " & _
      "LEFT JOIN co_multiempresas b ON a.empresacodigo=b.empresacodigo " & _
      "LEFT JOIN gr_moneda c  ON a.monedacodigo=c.monedacodigo " & _
      "LEFT JOIN vt_almacen d ON a.almacencodigo=d.almacencodigo " & _
      "ORDER BY a.empresacodigo"
      
      Set rs = VGCNx.Execute(SQL)
      Set TDBGridModoVta.DataSource = rs
      
      ' COMBO EMPRESA: antes gr_empresa ahora cambiado por co_multiempresas
      SQL = "SELECT empresacodigo,empresadescripcion " & _
      "FROM co_multiempresas " & _
      "ORDER BY empresacodigo "
      Set rs = VGCNx.Execute(SQL)
      If rs.RecordCount > 0 Then
        ReDim ArregloEmpresa(0 To 1, 0 To rs.RecordCount - 1)
        Call fncLlenarArreglo_Combo(rs, cmbEmpresa, ArregloEmpresa, 1)
      End If
      ' COMBO MONEDA:
      SQL = "SELECT monedacodigo,monedadescripcion " & _
      "FROM gr_moneda " & _
      "ORDER BY monedacodigo "
      Set rs = VGCNx.Execute(SQL)
      If rs.RecordCount > 0 Then
        ReDim ArregloMoneda(0 To 1, 0 To rs.RecordCount - 1)
        Call fncLlenarArreglo_Combo(rs, cmbMoneda, ArregloMoneda, 1)
      End If
      ' COMBO ALMACEN:
      SQL = "SELECT almacencodigo,almacendescripcion " & _
      "FROM vt_almacen " & _
      "ORDER BY almacencodigo "
      Set rs = VGCNx.Execute(SQL)
      If rs.RecordCount > 0 Then
        ReDim ArregloAlmacen(0 To 1, 0 To rs.RecordCount - 1)
        Call fncLlenarArreglo_Combo(rs, cmbAlmacen, ArregloAlmacen, 1)
      End If

      Setear_Controles
      
      'oCrystalReport.ReportFileName = RutaRep & "MantParametroVenta.rpt"
    
 TDBGridModoVta.Refresh
 rs.Close
 Set rs = Nothing
 SSTab1.Tab = 0
  
End Function

Private Function Setear_Controles()
Dim i As Integer

    For i = 0 To TDBGridModoVta.Columns.Count - 1
        Select Case i
            Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 10, 12, 13, 14, 15, 20
                TDBGridModoVta.Columns(i).ValueItems.Presentation = dbgNormal
                TDBGridModoVta.Columns(i).Width = 600
            Case 9, 11, 16, 17, 18, 19, 20, 22
                TDBGridModoVta.Columns(i).ValueItems.Presentation = dbgCheckBox
                TDBGridModoVta.Columns(i).Width = 500
    End Select
    Next i
    
End Function

Private Function Validar_DatosNulos() As Boolean


                If Trim(txt(0)) <> "" And _
                cmbEmpresa.ListIndex <> -1 Then
                    Validar_DatosNulos = True
                    Exit Function
                End If

End Function

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
    Select Case Index
    Case 3
        If chk(0).Value <> 1 Then
            KeyAscii = 0
            Exit Sub
        End If
    Case 6
        If chk(1).Value <> 1 Then
            KeyAscii = 0
            Exit Sub
        End If
    End Select
    
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
            If cmbEmpresa.ListIndex <> -1 Then
                If ArregloEmpresa(0, cmbEmpresa.ListIndex) = _
                    Trim(TDBGridModoVta.Columns.Item(0).Value) Then
                    Validar_CodigosDuplicados = True
                    Exit Function
                End If
            End If
         ElseIf operacion = "UPDATE" Then
            If cmbEmpresa.ListIndex <> -1 Then
                If ArregloEmpresa(0, cmbEmpresa.ListIndex) = _
                Trim(TDBGridModoVta.Columns.Item(0).Value) And _
                TDBGridModoVta.Row <> filaorigen Then
                   Validar_CodigosDuplicados = True
                   Exit Function
                End If
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
    Case 3, 6, 8
        If txt(Index) <> "" Then
            If Not IsNumeric(txt(Index)) Then
                MsgBox "Ingrese valores numéricos...", vbInformation, "AVISO"
                txt(Index) = ""
                Exit Sub
            End If
                'txt(Index).Text = Format(CDbl(txt(Index).Text), "##,##0.00")
                txt(Index).Text = Format(CDbl(txt(Index).Text), "#0.00")
        End If
End Select
End Sub

Private Function Validar_DescripcionesDuplicadas(operacion As String, Optional ByVal filaorigen As Integer) As Boolean
               
Validar_DescripcionesDuplicadas = False
                        
    TDBGridModoVta.MoveFirst
    Do Until TDBGridModoVta.EOF
        If operacion = "INSERT" Then
           If Trim(txt(0)) = _
              Trim(TDBGridModoVta.Columns.Item(2).Value) Then
                Validar_DescripcionesDuplicadas = True
                Exit Function
           End If
        ElseIf operacion = "UPDATE" Then
           If Trim(txt(0)) = _
              Trim(TDBGridModoVta.Columns.Item(2).Value) And _
              TDBGridModoVta.Row <> filaorigen Then
                 Validar_DescripcionesDuplicadas = True
                 Exit Function
           End If
        End If
        TDBGridModoVta.MoveNext
    Loop
               
End Function

Public Function Validar_IGV_DSCTO_TIPO() As Boolean
Validar_IGV_DSCTO_TIPO = True

    If chk(0).Value = 1 Then
        If Val(txt(3)) <= 0 Then
            MsgBox "Ingrese Tasa de Descuento", vbCritical, "AVISO"
            Validar_IGV_DSCTO_TIPO = False
            Exit Function
        End If
    End If
    If chk(1).Value = 1 Then
        If Val(txt(6)) <= 0 Then
            MsgBox "Ingrese Tasa de IGV", vbCritical, "AVISO"
            Validar_IGV_DSCTO_TIPO = False
            Exit Function
        End If
    End If
    'If cmbMoneda.ListIndex <> -1 Then
    '    If txt(8) = "" Then
    '        MsgBox "Ingrese Tipo de Cambio", vbCritical, "AVISO"
    '        Validar_IGV_DSCTO_TIPO = False
    '        Exit Function
    '    End If
    'End If
End Function
