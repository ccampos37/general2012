VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmEmpresa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Empresa"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   10095
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmbotones 
      Height          =   1050
      Left            =   2430
      TabIndex        =   33
      Top             =   7620
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
         Picture         =   "FrmEmpresa.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   38
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
         Picture         =   "FrmEmpresa.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   37
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
         Picture         =   "FrmEmpresa.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   36
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
         Picture         =   "FrmEmpresa.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   35
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
         Picture         =   "FrmEmpresa.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   180
         Width           =   915
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7395
      Left            =   210
      TabIndex        =   0
      Top             =   60
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   13044
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
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
      TabPicture(0)   =   "FrmEmpresa.frx":154A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "TDBGrid1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Mantenimiento"
      TabPicture(1)   =   "FrmEmpresa.frx":1566
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   6900
         Left            =   195
         TabIndex        =   2
         Top             =   375
         Width           =   9285
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
            Index           =   13
            Left            =   2985
            MaxLength       =   2
            TabIndex        =   55
            Top             =   4200
            Width           =   615
         End
         Begin VB.TextBox txt 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00%"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   5
            EndProperty
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
            Index           =   12
            Left            =   3000
            MaxLength       =   6
            TabIndex        =   53
            Top             =   6240
            Width           =   1035
         End
         Begin VB.CheckBox chk 
            Height          =   195
            Index           =   9
            Left            =   3000
            TabIndex        =   51
            Top             =   5760
            Width           =   345
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
            Left            =   7680
            TabIndex        =   50
            Top             =   6240
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
            Left            =   6120
            TabIndex        =   49
            Top             =   6240
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
            Index           =   11
            Left            =   6720
            MaxLength       =   6
            TabIndex        =   48
            Top             =   5760
            Width           =   1035
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
            Index           =   10
            Left            =   3000
            MaxLength       =   6
            TabIndex        =   46
            Top             =   3720
            Width           =   1815
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
            Index           =   9
            Left            =   6345
            MaxLength       =   10
            TabIndex        =   12
            Top             =   5115
            Width           =   720
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
            Left            =   2985
            MaxLength       =   2
            TabIndex        =   11
            Top             =   5115
            Width           =   615
         End
         Begin VB.CheckBox chk 
            Height          =   195
            Index           =   8
            Left            =   6810
            TabIndex        =   43
            Top             =   4800
            Width           =   345
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
            Index           =   7
            Left            =   7260
            MaxLength       =   6
            TabIndex        =   10
            Top             =   3270
            Width           =   1755
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
            Left            =   3000
            MaxLength       =   6
            TabIndex        =   9
            Top             =   3270
            Width           =   1815
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
            Left            =   3000
            MaxLength       =   15
            TabIndex        =   8
            Top             =   2400
            Width           =   3525
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
            Left            =   3000
            MaxLength       =   11
            TabIndex        =   7
            Top             =   1980
            Width           =   2025
         End
         Begin VB.CheckBox chk 
            Height          =   195
            Index           =   4
            Left            =   8580
            TabIndex        =   32
            Top             =   3870
            Width           =   345
         End
         Begin VB.CheckBox chk 
            Enabled         =   0   'False
            Height          =   195
            Index           =   7
            Left            =   8760
            TabIndex        =   30
            Top             =   2490
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.CheckBox chk 
            Height          =   195
            Index           =   6
            Left            =   8340
            TabIndex        =   18
            Top             =   2490
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.CheckBox chk 
            Height          =   195
            Index           =   5
            Left            =   3000
            TabIndex        =   17
            Top             =   4800
            Width           =   375
         End
         Begin VB.CheckBox chk 
            Height          =   195
            Index           =   3
            Left            =   8760
            TabIndex        =   16
            Top             =   2910
            Width           =   345
         End
         Begin VB.CheckBox chk 
            Height          =   315
            Index           =   2
            Left            =   8070
            TabIndex        =   15
            Top             =   2040
            Width           =   255
         End
         Begin VB.CheckBox chk 
            Height          =   195
            Index           =   1
            Left            =   5700
            TabIndex        =   14
            Top             =   2940
            Width           =   225
         End
         Begin VB.CheckBox chk 
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
            Left            =   3000
            TabIndex        =   13
            Top             =   2880
            Width           =   255
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
            Left            =   3000
            MaxLength       =   60
            TabIndex        =   6
            Top             =   1560
            Width           =   6045
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
            Left            =   3000
            MaxLength       =   30
            TabIndex        =   5
            Top             =   1110
            Width           =   6045
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
            Left            =   3000
            MaxLength       =   50
            TabIndex        =   4
            Top             =   660
            Width           =   6045
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
            Left            =   3000
            MaxLength       =   2
            TabIndex        =   3
            Top             =   210
            Width           =   615
         End
         Begin VB.Label lbl 
            Caption         =   "Cod. Oper.Gener.Transf."
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
            TabIndex        =   56
            Top             =   4245
            Width           =   2610
         End
         Begin VB.Label lbl 
            Caption         =   "Porcentaje de Retencion "
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
            Left            =   120
            TabIndex        =   54
            Top             =   6240
            Width           =   2640
         End
         Begin VB.Label lbl 
            Caption         =   "Agente de Retencion Codigo de Caja"
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
            TabIndex        =   52
            Top             =   5760
            Width           =   2490
         End
         Begin VB.Label lbl 
            Caption         =   "Codigo Doc. de Retencion "
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
            Index           =   20
            Left            =   3840
            TabIndex        =   47
            Top             =   5760
            Width           =   2640
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Numeracion Transferencias"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   19
            Left            =   240
            TabIndex        =   45
            Top             =   3780
            Width           =   2715
         End
         Begin VB.Label lbl 
            Caption         =   "Transferencia Ingreso"
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
            Left            =   4005
            TabIndex        =   44
            Top             =   5190
            Width           =   2175
         End
         Begin VB.Label lbl 
            Caption         =   "Control Cuenta Contable"
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
            Left            =   3990
            TabIndex        =   42
            Top             =   4770
            Width           =   2580
         End
         Begin VB.Label lbl 
            Caption         =   "Numeracion Egresos"
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
            Left            =   5070
            TabIndex        =   41
            Top             =   3330
            Width           =   2100
         End
         Begin VB.Label lbl 
            Caption         =   "Ciudad"
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
            Left            =   240
            TabIndex        =   40
            Top             =   2520
            Width           =   2400
         End
         Begin VB.Label lbl 
            Caption         =   "Ruc"
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
            TabIndex        =   39
            Top             =   2100
            Width           =   2370
         End
         Begin VB.Label lbl 
            Caption         =   "Controla Saldo Contable/Disponible"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   17
            Left            =   4950
            TabIndex        =   31
            Top             =   3870
            Width           =   3450
         End
         Begin VB.Label lbl 
            Caption         =   "Transferencia Egreso"
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
            TabIndex        =   29
            Top             =   5160
            Width           =   2490
         End
         Begin VB.Label lbl 
            Caption         =   "NO Control Cobranza Chq."
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
            Left            =   240
            TabIndex        =   28
            Top             =   4740
            Width           =   2640
         End
         Begin VB.Label lbl 
            Caption         =   "Controla Codigo de Caja"
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
            Left            =   6000
            TabIndex        =   27
            Top             =   2910
            Width           =   2610
         End
         Begin VB.Label lbl 
            Caption         =   "Numeracion Ingresos"
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
            Left            =   240
            TabIndex        =   26
            Top             =   3330
            Width           =   2490
         End
         Begin VB.Label lbl 
            Caption         =   "Num. Automatica"
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
            Left            =   6150
            TabIndex        =   25
            Top             =   2040
            Width           =   1890
         End
         Begin VB.Label lbl 
            Caption         =   "Controla Referencia"
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
            Left            =   3450
            TabIndex        =   24
            Top             =   2880
            Width           =   2160
         End
         Begin VB.Label lbl 
            Caption         =   "Para Reporte"
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
            Left            =   240
            TabIndex        =   23
            Top             =   2910
            Width           =   2400
         End
         Begin VB.Label lbl 
            Caption         =   "Direccion"
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
            Left            =   240
            TabIndex        =   22
            Top             =   1650
            Width           =   2400
         End
         Begin VB.Label lbl 
            Caption         =   "Descripcion Corta"
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
            Index           =   2
            Left            =   240
            TabIndex        =   21
            Top             =   1200
            Width           =   2400
         End
         Begin VB.Label lbl 
            Caption         =   "Descripcion"
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
            Left            =   240
            TabIndex        =   20
            Top             =   720
            Width           =   2400
         End
         Begin VB.Label lbl 
            Caption         =   "Codigo"
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
            Left            =   240
            TabIndex        =   19
            Top             =   270
            Width           =   2400
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   6735
         Left            =   -74760
         TabIndex        =   1
         Top             =   450
         Width           =   8985
         _ExtentX        =   15849
         _ExtentY        =   11880
         _LayoutType     =   0
         _RowHeight      =   15
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
         CellTipsWidth   =   104.882
         DeadAreaBackColor=   13160660
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
Attribute VB_Name = "FrmEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adll As New dllgeneral.dll_general
Dim modoinsert As Boolean
Dim modoedit As Boolean
Dim i_filaorigen As Integer
Dim nLongicampo(2) As Integer

Private Sub cAcepta_Click()
 If adll.VerificaDatoExistente(VGCNx, "select * from te_parametroempresa Where empresacodigo='" & txt(0) & "'") = 1 And modoinsert = True Then
    MsgBox "Ya existe el Codigo...!!!", vbInformation, MsgTitle
    Exit Sub
 End If

 If modoinsert = True Then
       VGCNx.Execute "Insert Into te_parametroempresa " & _
                  "(empresacodigo,empresarazonsocial,empresasiglas, " & _
                  "empresadireccion,empresaruc,empresaciudad,empresareporte," & _
                  "empresacontrolarefe,empresanumeauto,empresanumeingreso,empresanumegreso," & _
                  "empresacontrolacodcaja,empresacontrolasaldocontabledispo,empresanocontrolcobranzacheque," & _
                  "empresalistaestadoclientes,empresalistaestadoproveedor,empresacontrolactacontable,empresatransaccionegreso," & _
                  "empresatransaccioningreso,empresanumtransferencia,empresacodigoretencion," & _
                  "porcentajeretencion,empresaretencion,codigooperaciontransferencia,usuariocodigo,fechaact)" & _
                  "VALUES(" & _
                  "'" & txt(0) & "','" & txt(1) & "','" & txt(2) & "','" & txt(3) & "'," & _
                  "'" & txt(4) & "','" & txt(5) & "'," & _
                  "'" & IIf(chk(0).Value = 1, "1", "0") & "'," & _
                  "'" & IIf(chk(1).Value = 1, "1", "0") & "'," & _
                  "'" & IIf(chk(2).Value = 1, "1", "0") & "'," & _
                  "'" & txt(6) & "','" & txt(7) & "'," & _
                  "'" & IIf(chk(3).Value = 1, "1", "0") & "'," & _
                  "'" & IIf(chk(4).Value = 1, "1", "0") & "'," & _
                  "'" & IIf(chk(5).Value = 1, "1", "0") & "'," & _
                  "'" & IIf(chk(6).Value = 1, "1", "0") & "'," & _
                  "'" & IIf(chk(7).Value = 1, "1", "0") & "'," & _
                  "'" & IIf(chk(8).Value = 1, "1", "0") & "'," & _
                  "'" & txt(8).Text & "','" & txt(9).Text & "'," & _
                  "'" & txt(10).Text & "','" & txt(11).Text & "','" & txt(12).Text & "'," & _
                  "'" & IIf(chk(9).Value = 1, "1", "0") & "'," & _
                  "'" & txt(13).Text & "','" & VGUsuario & "','" & Date & "')"
 'CDbl(IIf(IsNull(txt(9)) Or Len(Trim(txt(9))) = 0, 0, txt(9)))
 ElseIf modoedit = True Then
       VGCNx.Execute "Update te_parametroempresa " & _
                  " Set  empresarazonsocial='" & txt(1) & "'," & _
                  "empresasiglas='" & txt(2) & "'," & _
                  "empresadireccion='" & txt(3) & "'," & _
                  "empresaruc='" & txt(4) & "'," & _
                  "empresaciudad='" & txt(5) & "'," & _
                  "empresareporte='" & IIf(chk(0).Value = 1, "1", "0") & "'," & _
                  "empresacontrolarefe='" & IIf(chk(1).Value = 1, "1", "0") & "'," & _
                  "empresanumeauto='" & IIf(chk(2).Value = 1, "1", "0") & "'," & _
                  "empresanumeingreso='" & txt(6) & "'," & _
                  "empresanumegreso='" & txt(7) & "'," & _
                  "empresacontrolacodcaja='" & IIf(chk(3).Value = 1, "1", "0") & "'," & _
                  "empresacontrolasaldocontabledispo='" & IIf(chk(4).Value = 1, "1", "0") & "'," & _
                  "empresanocontrolcobranzacheque='" & IIf(chk(5).Value = 1, "1", "0") & "'," & _
                  "empresalistaestadoclientes='" & IIf(chk(6).Value = 1, "1", "0") & "'," & _
                  "empresalistaestadoproveedor='" & IIf(chk(7).Value = 1, "1", "0") & "'," & _
                  "empresacontrolactacontable='" & IIf(chk(8).Value = 1, "1", "0") & "'," & _
                  "empresaretencion='" & IIf(chk(9).Value = 1, "1", "0") & "'," & _
                  "empresatransaccionegreso='" & txt(8) & "'," & _
                  "empresatransaccioningreso='" & txt(9).Text & "'," & _
                  "empresanumtransferencia='" & txt(10).Text & "', " & _
                  "empresacodigoretencion='" & txt(11).Text & "', " & _
                  "porcentajeretencion='" & txt(12).Text & "', " & _
                  "codigooperaciontransferencia='" & txt(13).Text & "' " & _
                  "Where empresacodigo='" & txt(0) & "'"
 
 End If
 'CDbl(IIf(IsNull(txt(9)) Or Len(Trim(txt(9))) = 0, 0, txt(9))) & ","
 modoedit = False
 modoinsert = False
 Call Listado
End Sub


Public Function Listado()
    TDBGrid1.ClearFields
    Set TDBGrid1.DataSource = Nothing
    Call adll.ListarEnTDBGRID(VGCNx, "te_parametroempresa", TDBGrid1, "empresacodigo,empresarazonsocial,empresasiglas,empresaruc", "empresacodigo", nLongicampo)
    Call ConfiguraGrid
    Call adll.ActivaTab(0, 1, SSTab1)
    frmbotones.Visible = True

End Function



Private Sub cCancela_Click()
  Call adll.ActivaTab(0, 1, SSTab1)
  frmbotones.Visible = True
End Sub

Private Sub cmdBotones_Click(Index As Integer)
  Dim J As Integer
  Dim spos As Integer
  Dim SQL As String
  Dim d_estado As Double
  ''''''''''
  Dim rs As New ADODB.Recordset
  Dim error As ADODB.Errors
  '''''''''''
  On Error GoTo CONTROLERRORES
  
  SSTab1.TabEnabled(1) = True

 modoedit = False
 modoinsert = False

  Select Case Index
  
     Case 0   'nuevo
        SSTab1.Tab = 1
        If txt(0).Visible = True Then
            txt(0).SetFocus
        ElseIf chk(0).Visible = True Then
            chk(0).SetFocus
        End If
        Call Limpia_textos
        Call adll.ActivaTab(1, 1, SSTab1)
        
        frmbotones.Visible = False
        modoinsert = True
        txt(0).SetFocus
        
     Case 1   'modificar
        If TDBGrid1.Row < 0 Then
          Exit Sub
        End If
        
        Call Limpia_textos
        
        Set rs = VGCNx.Execute("select * from te_parametroempresa Where empresacodigo='" & TDBGrid1.Columns(0).Text & "'")
        If rs.RecordCount > 0 Then
           txt(0) = Escadena(rs!empresacodigo)
           txt(1) = Escadena(rs!empresarazonsocial)
           txt(2) = Escadena(rs!empresasiglas)
           txt(3) = Escadena(rs!empresadireccion)
           txt(4) = Escadena(rs!empresaruc)
           txt(5) = Escadena(rs!empresaciudad)
           chk(0).Value = IIf(Escadena(rs!empresareporte) = "1", 1, 0)
           chk(1).Value = IIf(Escadena(rs!empresacontrolarefe) = "1", 1, 0)
           chk(2).Value = IIf(Escadena(rs!empresanumeauto) = "1", 1, 0)
           txt(6) = Escadena(rs!empresanumeingreso)
           txt(7) = Escadena(rs!empresanumegreso)
           chk(3).Value = IIf(Escadena(rs!empresacontrolacodcaja) = "1", 1, 0)
           chk(4).Value = IIf(Escadena(rs!empresacontrolasaldocontabledispo) = "1", 1, 0)
           chk(5).Value = IIf(Escadena(rs!empresanocontrolcobranzacheque) = "1", 1, 0)
           chk(6).Value = IIf(Escadena(rs!empresalistaestadoclientes) = "1", 1, 0)
           chk(7).Value = IIf(Escadena(rs!empresalistaestadoproveedor) = "1", 1, 0)
           chk(8).Value = IIf(Escadena(rs!empresacontrolactacontable) = "1", 1, 0)
           chk(9).Value = IIf(Escadena(rs!empresaretencion) = "1", 1, 0)
           txt(8) = Escadena(rs!empresatransaccionegreso)
           txt(9) = Escadena(rs!empresatransaccioningreso)
           txt(10) = Escadena(rs!empresanumtransferencia)
           txt(11) = Escadena(rs!empresacodigoretencion)
           txt(12) = Escadena(rs!porcentajeretencion)
           txt(13) = Escadena(rs!codigooperaciontransferencia)
           'txt(9) = Numero(IIf(IsNull(rs!empresatipocambio) Or Len(Trim(rs!empresatipocambio)) = 0, 0, rs!empresatipocambio))
        End If
        rs.Close
        Set rs = Nothing
        Call adll.ActivaTab(1, 1, SSTab1)
        frmbotones.Visible = False
        SSTab1.Tab = 1
        
        i_filaorigen = TDBGrid1.Row
        modoedit = True
        If txt(0).Visible = True Then
            txt(0).SetFocus
        ElseIf chk(0).Visible = True Then
            chk(0).SetFocus
        End If

     Case 2   'eliminar
          If TDBGrid1.Row < 0 Then
              Exit Sub
          End If
         
          If MsgBox("Desea Eliminar el Registro?", vbYesNo, MsgTitle) = vbYes Then
              VGCNx.Execute "Delete From  te_parametroempresa where empresacodigo='" & TDBGrid1.Columns(0).Text & "'"
          End If
          Call Listado
     Case 3  'Imprimir
       Call Imprimir("RepMantempresa.rpt")
     Case 4  ' salir
       Unload Me
  End Select
  
  
'RaiseEvent Click(Index)

Exit Sub

CONTROLERRORES:
   If Err Then
       MsgBox Err.Number & "-" & Err.Description, vbInformation, "ERROR"
       Err = 0
       'VGGeneral.RollbackTrans
       Resume Next
    End If

End Sub

Public Function Limpia_textos()
 Dim J As Integer
   For J = 0 To txt.Count - 1
      txt(J).Text = ""
   Next J
   For J = 0 To chk.Count - 1
      chk(J).Value = 0
   Next J
End Function

Private Sub Form_Load()
   MostrarForm Me, "C2"
   Call adll.ActivaTab(0, 1, SSTab1)
   nLongicampo(1) = 0
   
   Call adll.ListarEnTDBGRID(VGCNx, "te_parametroempresa", TDBGrid1, "empresacodigo,empresarazonsocial,empresasiglas,empresaruc", "empresacodigo", nLongicampo)
   Call ConfiguraGrid
   
End Sub

Public Function ConfiguraGrid()
   With TDBGrid1
    .Columns(0).Width = 1200
    .Columns(0).Caption = "Codigo"
    .Columns(1).Width = 3500
    .Columns(1).Caption = "Descripcion"
    .Columns(2).Width = 2000
    .Columns(2).Caption = "Desc. Corta"
    .Columns(3).Width = 1000
    .Columns(3).Caption = "R.U.C."
    .Refresh
   End With
End Function

Private Sub txt_Change(Index As Integer)
  Select Case Index
   Case 0, 9, 6, 7
      If Not adll.ValidaCadena(txt(Index), "N") Then
        If Len(Trim(txt(Index))) > 0 Then
          txt(Index) = Left(txt(Index), Len(txt(Index)) - 1)
        End If
        txt(Index).SetFocus
      End If
      Exit Sub
   Case 1, 2, 3, 4, 5, 8
      If Not adll.ValidaCadena(txt(Index), "C") Then
        If Len(Trim(txt(Index))) > 0 Then
          txt(Index) = Left(txt(Index), Len(txt(Index)) - 1)
        End If
        txt(Index).SetFocus
      End If
      Exit Sub

  End Select

End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
 
 If KeyAscii = 13 Then
   txt(Index) = UCase(txt(Index))

   Call Seguir(txt(Index), KeyAscii)
 End If
 
End Sub

Private Sub txt_LostFocus(Index As Integer)
   If Index = 6 Or Index = 7 Then
       txt(Index) = Right("000000000000000" & txt(Index), txt(Index).MaxLength)
   ElseIf Index = 0 Then
       txt(Index) = Right("000000000000000" & txt(Index), txt(Index).MaxLength)
   ElseIf Index Like "[12345]" Then
       txt(Index) = UCase(txt(Index))
   End If
End Sub
