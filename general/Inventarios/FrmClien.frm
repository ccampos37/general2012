VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmArClien 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualiza Datos Generales de Clientes"
   ClientHeight    =   5085
   ClientLeft      =   1425
   ClientTop       =   1980
   ClientWidth     =   8640
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   8640
   Begin VB.CommandButton Transf 
      Caption         =   "Transf. a  Contab."
      Height          =   825
      Left            =   7110
      Picture         =   "FrmClien.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   4170
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.CommandButton CmdEli 
      Caption         =   "&Eliminar"
      Height          =   675
      Left            =   3480
      Picture         =   "FrmClien.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4230
      Width           =   775
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   675
      Left            =   5940
      Picture         =   "FrmClien.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4230
      Width           =   775
   End
   Begin VB.CommandButton CmdModi 
      Caption         =   "&Modificar"
      Height          =   675
      Left            =   2160
      Picture         =   "FrmClien.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4230
      Width           =   775
   End
   Begin VB.CommandButton CmdIng 
      Caption         =   "&Ingreso"
      Height          =   675
      Left            =   900
      Picture         =   "FrmClien.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4230
      Width           =   775
   End
   Begin VB.CommandButton CmdFicha 
      Caption         =   "&Ficha"
      Enabled         =   0   'False
      Height          =   675
      Left            =   4665
      Picture         =   "FrmClien.frx":154A
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4230
      Width           =   775
   End
   Begin VB.CommandButton CmdSalir2 
      Caption         =   "&Salir"
      Height          =   675
      Left            =   4995
      Picture         =   "FrmClien.frx":198C
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4230
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Height          =   675
      Left            =   2760
      Picture         =   "FrmClien.frx":1DCE
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   4230
      Visible         =   0   'False
      Width           =   775
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4095
      Left            =   120
      TabIndex        =   38
      Top             =   0
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   7223
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Datos Requeridos"
      TabPicture(0)   =   "FrmClien.frx":2210
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LbBanco"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label19"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label18"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label23"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label22"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label21"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label20"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "LbDistrito"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label9"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label8"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label7"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label6"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label5"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label4"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label3"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label2"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "TxCta"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "TxBanco"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "TxHost"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "TxEmail"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "TxFax"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "TxTelefono"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "TxDistrito"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "TxProvincia"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "TxDepartamento"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "TxPais"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "TxEntrega"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "TxFactura"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "TxRazon"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "TxRuc"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "TxCodigo"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "chkAsignarRUC"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).ControlCount=   33
      TabCaption(1)   =   "Datos Adicionales"
      TabPicture(1)   =   "FrmClien.frx":222C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Combo1"
      Tab(1).Control(1)=   "TxGiro"
      Tab(1).Control(2)=   "TxRepresentante"
      Tab(1).Control(3)=   "TxDscto"
      Tab(1).Control(4)=   "TxPago"
      Tab(1).Control(5)=   "TxPrecio"
      Tab(1).Control(6)=   "TxLimite"
      Tab(1).Control(7)=   "TxVendedor"
      Tab(1).Control(8)=   "Label15"
      Tab(1).Control(9)=   "Label17"
      Tab(1).Control(10)=   "LbGiro"
      Tab(1).Control(11)=   "Label24"
      Tab(1).Control(12)=   "Label25"
      Tab(1).Control(13)=   "Label26"
      Tab(1).Control(14)=   "Label27"
      Tab(1).Control(15)=   "Label28"
      Tab(1).Control(16)=   "Label30"
      Tab(1).Control(17)=   "Label31"
      Tab(1).Control(18)=   "LbPago"
      Tab(1).Control(19)=   "LbPrecio"
      Tab(1).Control(20)=   "LbVendedor"
      Tab(1).Control(21)=   "LbFecha"
      Tab(1).ControlCount=   22
      Begin VB.CheckBox chkAsignarRUC 
         Caption         =   "Asignar a RUC?"
         Enabled         =   0   'False
         Height          =   315
         Left            =   3120
         TabIndex        =   1
         Top             =   600
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FrmClien.frx":2248
         Left            =   -72570
         List            =   "FrmClien.frx":2252
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox TxCodigo 
         Height          =   285
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   0
         Text            =   "12345678901"
         Top             =   600
         Width           =   1380
      End
      Begin VB.TextBox TxRuc 
         Height          =   285
         Left            =   5700
         MaxLength       =   11
         TabIndex        =   2
         Text            =   "12345678901"
         Top             =   600
         Width           =   1545
      End
      Begin VB.TextBox TxRazon 
         Height          =   285
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   3
         Text            =   "TxRazon"
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox TxFactura 
         Height          =   285
         Left            =   5700
         MaxLength       =   70
         TabIndex        =   4
         Text            =   "TxFactura"
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox TxEntrega 
         Height          =   285
         Left            =   1665
         MaxLength       =   70
         TabIndex        =   5
         Text            =   "TxEntrega"
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox TxPais 
         Height          =   285
         Left            =   1680
         MaxLength       =   21
         TabIndex        =   6
         Text            =   "TxPais"
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox TxDepartamento 
         Height          =   285
         Left            =   5700
         MaxLength       =   21
         TabIndex        =   7
         Text            =   "TxDepartamento"
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox TxProvincia 
         Height          =   285
         Left            =   1680
         MaxLength       =   21
         TabIndex        =   8
         Text            =   "TxProvincia"
         Top             =   2040
         Width           =   2535
      End
      Begin VB.TextBox TxDistrito 
         Height          =   285
         Left            =   5700
         MaxLength       =   2
         TabIndex        =   9
         Text            =   "TxDistrito"
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox TxTelefono 
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   10
         Text            =   "TxTelefono"
         Top             =   2400
         Width           =   2235
      End
      Begin VB.TextBox TxFax 
         Height          =   285
         Left            =   5700
         MaxLength       =   15
         TabIndex        =   11
         Text            =   "TxFax"
         Top             =   2400
         Width           =   2235
      End
      Begin VB.TextBox TxEmail 
         Height          =   285
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   12
         Text            =   "TxEmail"
         Top             =   2760
         Width           =   2235
      End
      Begin VB.TextBox TxHost 
         Height          =   285
         Left            =   5700
         MaxLength       =   50
         TabIndex        =   13
         Text            =   "TxHost"
         Top             =   2760
         Width           =   2235
      End
      Begin VB.TextBox TxGiro 
         Height          =   285
         Left            =   -72570
         MaxLength       =   2
         TabIndex        =   16
         Text            =   "TxGiro"
         Top             =   2250
         Width           =   495
      End
      Begin VB.TextBox TxRepresentante 
         Height          =   285
         Left            =   -72585
         MaxLength       =   20
         TabIndex        =   17
         Text            =   "TxRepresentante"
         Top             =   2895
         Width           =   2655
      End
      Begin VB.TextBox TxBanco 
         Height          =   285
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   14
         Text            =   "TxBanco"
         Top             =   3120
         Width           =   495
      End
      Begin VB.TextBox TxCta 
         Height          =   285
         Left            =   5700
         MaxLength       =   12
         TabIndex        =   15
         Text            =   "TxCta"
         Top             =   3120
         Width           =   2235
      End
      Begin VB.TextBox TxDscto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -72570
         MaxLength       =   5
         TabIndex        =   18
         Text            =   "TxDsc"
         Top             =   570
         Width           =   975
      End
      Begin VB.TextBox TxPago 
         Height          =   285
         Left            =   -72570
         MaxLength       =   2
         TabIndex        =   19
         Text            =   "TxPago"
         Top             =   900
         Width           =   495
      End
      Begin VB.TextBox TxPrecio 
         Height          =   285
         Left            =   -72570
         MaxLength       =   2
         TabIndex        =   20
         Text            =   "TxPrecio"
         Top             =   1230
         Width           =   495
      End
      Begin VB.TextBox TxLimite 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -72570
         MaxLength       =   12
         TabIndex        =   21
         Text            =   "TxLimite"
         Top             =   1920
         Width           =   1305
      End
      Begin VB.TextBox TxVendedor 
         Height          =   285
         Left            =   -72570
         MaxLength       =   2
         TabIndex        =   22
         Text            =   "TxVendedor"
         Top             =   2580
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Código :"
         Height          =   255
         Left            =   240
         TabIndex        =   69
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "R.U.C. :"
         Height          =   195
         Left            =   5160
         TabIndex        =   68
         Top             =   600
         Width           =   570
      End
      Begin VB.Label Label3 
         Caption         =   "Razón Social :"
         Height          =   255
         Left            =   240
         TabIndex        =   67
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Dirección Factura :"
         Height          =   255
         Left            =   4380
         TabIndex        =   66
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Dirección Entrega :"
         Height          =   255
         Left            =   240
         TabIndex        =   65
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "País :"
         Height          =   255
         Left            =   240
         TabIndex        =   64
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Departamento :"
         Height          =   255
         Left            =   4380
         TabIndex        =   63
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Provincia :"
         Height          =   255
         Left            =   240
         TabIndex        =   62
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Distrito :"
         Height          =   255
         Left            =   4380
         TabIndex        =   61
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label LbDistrito 
         Caption         =   "LbDistrito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6360
         TabIndex        =   60
         Top             =   2040
         Width           =   1950
      End
      Begin VB.Label Label20 
         Caption         =   "Tlf. Cliente :"
         Height          =   255
         Left            =   240
         TabIndex        =   59
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label21 
         Caption         =   "Fax :"
         Height          =   255
         Left            =   4380
         TabIndex        =   58
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label22 
         Caption         =   "Email Cliente :"
         Height          =   255
         Left            =   240
         TabIndex        =   57
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label23 
         Caption         =   "Host Cliente :"
         Height          =   255
         Left            =   4380
         TabIndex        =   56
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label15 
         Caption         =   "Giro Cliente :"
         Height          =   255
         Left            =   -74685
         TabIndex        =   55
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "Representante :"
         Height          =   255
         Left            =   -74700
         TabIndex        =   54
         Top             =   2925
         Width           =   1695
      End
      Begin VB.Label LbGiro 
         Caption         =   "LbGiro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71970
         TabIndex        =   53
         Top             =   2250
         Width           =   2175
      End
      Begin VB.Label Label18 
         Caption         =   "Banco :"
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label19 
         Caption         =   "Cta. Cte. :"
         Height          =   255
         Left            =   4380
         TabIndex        =   51
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label LbBanco 
         Caption         =   "LbBanco"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   50
         Top             =   3120
         Width           =   2055
      End
      Begin VB.Label Label24 
         Caption         =   "Dscto. :"
         Height          =   255
         Left            =   -74685
         TabIndex        =   49
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label25 
         Caption         =   "Tipo Pago :"
         Height          =   255
         Left            =   -74685
         TabIndex        =   48
         Top             =   930
         Width           =   1095
      End
      Begin VB.Label Label26 
         Caption         =   "Precio Factura :"
         Height          =   255
         Left            =   -74685
         TabIndex        =   47
         Top             =   1245
         Width           =   1455
      End
      Begin VB.Label Label27 
         Caption         =   "Moneda Crédito :"
         Height          =   255
         Left            =   -74685
         TabIndex        =   46
         Top             =   1575
         Width           =   1575
      End
      Begin VB.Label Label28 
         Caption         =   "Limite Crédito :"
         Height          =   255
         Left            =   -74685
         TabIndex        =   45
         Top             =   1905
         Width           =   1335
      End
      Begin VB.Label Label30 
         Caption         =   "Vendedor :"
         Height          =   255
         Left            =   -74670
         TabIndex        =   44
         Top             =   2610
         Width           =   1095
      End
      Begin VB.Label Label31 
         Caption         =   "Fecha Registro :"
         Height          =   255
         Left            =   -74685
         TabIndex        =   43
         Top             =   3240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label LbPago 
         Caption         =   "LbPago"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71970
         TabIndex        =   42
         Top             =   900
         Width           =   2625
      End
      Begin VB.Label LbPrecio 
         Caption         =   "LbPrecio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71970
         TabIndex        =   41
         Top             =   1230
         Width           =   2820
      End
      Begin VB.Label LbVendedor 
         Caption         =   "LbVendedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71970
         TabIndex        =   40
         Top             =   2580
         Width           =   2055
      End
      Begin VB.Label LbFecha 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LbFecha"
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
         Left            =   -72585
         TabIndex        =   39
         Top             =   3240
         Visible         =   0   'False
         Width           =   1110
      End
   End
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   120
      TabIndex        =   32
      Top             =   0
      Width           =   8295
      Begin VB.ComboBox CmbOrden 
         Height          =   315
         ItemData        =   "FrmClien.frx":2266
         Left            =   6120
         List            =   "FrmClien.frx":2273
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox TxFiltro 
         Height          =   300
         Left            =   2040
         TabIndex        =   33
         Text            =   "TxFiltro"
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label33 
         Caption         =   "Orden :"
         Height          =   255
         Left            =   5160
         TabIndex        =   36
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label32 
         Caption         =   "Buscar   :"
         Height          =   375
         Left            =   720
         TabIndex        =   35
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ficha Técnica"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      TabIndex        =   30
      Top             =   120
      Visible         =   0   'False
      Width           =   8295
      Begin VB.TextBox TxTecnica 
         Height          =   3135
         Left            =   195
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         Text            =   "FrmClien.frx":229B
         Top             =   360
         Visible         =   0   'False
         Width           =   7815
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FrmClien.frx":22A1
      Height          =   3255
      Left            =   120
      TabIndex        =   37
      Top             =   720
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   5741
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   17
      TabAction       =   2
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "clientecodigo"
         Caption         =   "cliente"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "clienterazonsocial"
         Caption         =   "razon"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         ScrollBars      =   2
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   5504.882
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmArClien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adodc1 As ADODB.Recordset
Dim cSel1 As ADODB.Recordset
Dim cSel2 As ADODB.Recordset
Dim cSql1 As String, CSQL2 As String
Dim cSql3 As String, nT As Integer
Dim cCod As String, cDes As String
Dim nCom As Integer, nExiste As Integer
Dim nTra2 As Integer, nCursor As Integer
Dim nTra As Integer
Dim cBase As String
Private Sub OculObj01(ntipo As Boolean) ' Ficha Tecnica
Frame2.Visible = ntipo
TxTecnica.Visible = ntipo
End Sub
Private Sub OculObj02(ntipo As Boolean)  'Grabar,Eliminar y salir
Cmdgrabar.Visible = ntipo
CmdSalir2.Visible = ntipo
End Sub
Private Sub OculObj03(ntipo As Boolean) ' Todos los datos
SSTab1.Visible = ntipo
If ntipo Then SSTab1.Tab = 0
End Sub
Private Sub OculObj04(ntipo As Boolean) ' Botones principales
Transf.Visible = ntipo
CmdIng.Visible = ntipo
CmdModi.Visible = ntipo
CmdEli.Visible = ntipo
CmdFicha.Visible = ntipo
CmdSalir.Visible = ntipo
End Sub
Private Sub OculObj05(ntipo As Boolean)  'Orden y Filtro
Frame5.Visible = ntipo
End Sub
Private Sub OculObj06(ntipo As Boolean) 'Datagrid
DataGrid1.Visible = ntipo
End Sub

Private Sub chkAsignarRUC_Click()
    If chkAsignarRUC.Value Then
        txRuc = TxCodigo
    Else
        txRuc = ""
    End If
End Sub

Private Sub chkAsignarRUC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        chkAsignarRUC_Click
        SendKeys "{tab}"
    End If
End Sub

Private Sub CmbOrden_Click()            ' Ordenar por
nCom = CmbOrden.ListIndex
Set adodc1 = New ADODB.Recordset
Select Case nCom

Case 0
    adodc1.Open "Select clientecodigo,clienterazonsocial,clienteruc FROM VT_CLIENTE ORDER BY clientecodigo", VGCNx, adOpenStatic
Case 1
    adodc1.Open "Select clientecodigo,clienterazonsocial,clienteruc FROM VT_CLIENTE ORDER BY clienterazonsocial", VGCNx, adOpenStatic
Case 2
    adodc1.Open "Select clientecodigo,clienterazonsocial,clienteruc FROM VT_CLIENTE ORDER BY clienteruc", VGCNx, adOpenStatic
End Select
TxFiltro = ""
Set DataGrid1.DataSource = adodc1

If DataGrid1.Visible Then DataGrid1.SetFocus
End Sub
Private Sub CmdEli_Click()              ' Elimina
Dim nPosi As Integer
On Error GoTo EliErr

If adodc1.RecordCount > 0 Then
    If MsgBox("Desea Eliminar Datos ?", vbQuestion + vbOKCancel, "Mensaje") = vbOK Then
       cBase = cRuta4
       If Existe(1, adodc1(0), "FacCab", "CFCODCLI", False) Then
            MsgBox "No se puede eliminar el Cliente, porque tiene documentos Anexados", vbInformation, "Información"
            Exit Sub
       Else
           cSql1 = "Delete from VT_CLIENTE where clientecodigo = '" & adodc1(0) & "'"
           CSQL2 = "Delete from Dire_Cliente where clientecodigo = '" & adodc1(0) & "'"
           nPosi = Pos_Dato(adodc1)
           nTra = 1
           VGCNx.BeginTrans
           VGCNx.Execute cSql1
    '       Vgcnx.Execute CSQL2
           VGCNx.CommitTrans
           nTra = 0
           adodc1.Requery
           adodc1.AbsolutePosition = nPosi
       End If
    End If
    If DataGrid1.Visible Then DataGrid1.SetFocus
Else
    MsgBox "No existe registros para Eliminar", vbInformation, "Mensaje"
    Exit Sub
End If
Exit Sub
EliErr:
    MsgBox Err.Description
    If nTra = 1 Then VGCNx.RollbackTrans
End Sub

Private Sub CmdFicha_Click()            ' Ficha Tecnica
If adodc1.RecordCount > 0 Then
    nT = 3
    TxTecnica = ""
    OculObj04 (False)
    OculObj05 (False)
    OculObj06 (False)
    OculObj01 (True)
    OculObj02 (True)
    cCod = adodc1("clientecodigo")
    cDes = adodc1("clienterazonsocial")
    cSql1 = "Select clientecodigo,comenta from VT_CLIENTE where clientecodigo = '" & cCod & "'"
    Set cSel1 = New ADODB.Recordset
    cSel1.Open cSql1, VGCNx, adOpenStatic
    If cSel1.RecordCount > 0 Then
        If Not IsNull(cSel1("comenta")) Then TxTecnica = cSel1("comenta")
        Frame2.Caption = "FICHA TECNICA :   " & cDes
    Else
        MsgBox "El registro ha sido eliminado", vbInformation, "Mensaje"
        cSel1.Close: CmdSalir2_Click
        Exit Sub
    End If
    cSel1.Close
    TxTecnica.SetFocus
Else
    MsgBox "No existe Registros", vbInformation, "Mensaje"
    CmdSalir2_Click
End If
End Sub

Private Sub CmdGrabar_Click()           ' Grabar
Dim cMon As String
On Error GoTo GrabErr

If nT <> 3 Then
    If nT = 1 Then
        If CodigoC(TxCodigo) = False Then
            If Trim(TxCodigo) = "" Then
                MsgBox "Ingrese Código", vbInformation, "Mensaje"
            Else
                MsgBox "Código ya existe", vbInformation, "Mensaje"
            End If
            TxCodigo.SetFocus: Exit Sub
        End If
    End If
    
    If Trim(TxRazon) = "" Then
        MsgBox "Ingrese Razón Social", vbInformation, "Mensaje"
        
        TxRazon.SetFocus: Exit Sub
    End If
    
    If Trim(TxLimite) = "" Then TxLimite = 0
    If Trim(TxDscto) = "" Then TxDscto = 0
    
    If txRuc <> "" Then
          If Validar_RUC(txRuc) = False Then
             txRuc.SetFocus: Exit Sub
          End If
    End If
    'Validacion de Textos
    If TxDistrito <> "" Then   'Distritos
        If Val_Ayu(TxDistrito, "13") = "" Then
  '         MsgBox "Código de Distrito no existe", vbInformation, mensaje1
  '         LbDistrito = ""
  '         TxDistrito.SetFocus: Exit Sub
        End If
    End If
    If TxGiro <> "" Then   'Giros
        If Val_Ayu(TxGiro, "62") = "" Then
'           MsgBox "Código de Giro del Cliente no existe", vbInformation, mensaje1
'           LbGiro = ""
'           TxGiro.SetFocus: Exit Sub
        End If
    End If
    
    If Trim(TxBanco) <> "" Then
       If Existe(1, TxBanco, "BANCO", "BAN_CODIGO", False) = False Then
          MsgBox "El código de Banco, no existe", vbInformation, "Mensaje"
          TxBanco.SetFocus: Exit Sub
       End If
    End If
    If Trim(TxPrecio) <> "" Then
        LbPrecio = Mid(fPre(TxPrecio), 1, 15)
        If Trim(LbPrecio) = "" Then
           MsgBox "El código de Precio, no existe", vbInformation, "Mensaje"
           TxPrecio = "": TxPrecio.SetFocus: Exit Sub
        End If
    End If
    
    If Trim(TxPago) <> "" Then
       If Existe(1, TxPago, "FORMA_PAGO", "COD_FP", False) = False Then
          MsgBox "El código de Forma de Pago, no existe", vbInformation, "Mensaje"
          TxPago.SetFocus: Exit Sub
       End If
    End If
    If Trim(TxVendedor) <> "" Then
       If Existe(1, TxVendedor, "VENDEDOR", "Cod_Ven", False) = False Then
          MsgBox "El código de Vendedor no existe", vbInformation, "Mensaje"
          TxVendedor = "": TxVendedor.SetFocus: Exit Sub
       End If
    End If
End If


If MsgBox("Es correcta la Información ?", vbQuestion + vbOKCancel, "Mensaje") = vbOK Then
    'CESTADO,OBSERV,Dfecins  '''DFECMOD,
    If Combo1.ListIndex = 0 Then
       cMon = "02"
    Else
       cMon = "01"
    End If
       
    
    
    If nT = 1 Then      'Ingreso
        CSQL2 = "Insert Into VT_CLIENTE (clientecodigo,clienterazonsocial,clientedireccion,clientetelefono,clienteruc,"
        CSQL2 = CSQL2 & "clientepropietario,Clientedistrito,Usuariocodigo,"
        CSQL2 = CSQL2 & "clientetipopais,clientedepartamento,clienteprovincia,negociocodigo,"
        CSQL2 = CSQL2 & "clientemail,fechaact,clientetipopersona,"
        CSQL2 = CSQL2 & "clientefax,clientelimitecreddolar,clientelimitecredsoles,clientefechaactivacion,clientesuspendido) VALUES "
        CSQL2 = CSQL2 & "('" & TxCodigo & "','" & SupCadSQL(TxRazon) & "','" & SupCadSQL(TxFactura) & "',"
        CSQL2 = CSQL2 & "'" & TxTelefono & "','" & txRuc & "',"
        CSQL2 = CSQL2 & "'" & SupCadSQL(TxRepresentante) & "','" & TxDistrito & "','" & SupCadSQL(VGUsuario) & "',"
        CSQL2 = CSQL2 & "'" & TxPais & "','" & SupCadSQL(TxDepartamento) & "',"
        CSQL2 = CSQL2 & "'" & TxProvincia & "','" & TxGiro & "',"
        CSQL2 = CSQL2 & "'" & TxEmail & "','" & CStr(Date) & "','N',"
        CSQL2 = CSQL2 & "'" & TxFax & "',"
        If Trim(cMon) <> "01" Then
            CSQL2 = CSQL2 & "" & Val(TxLimite) & ",0,"
        Else
            CSQL2 = CSQL2 & "0," & Val(TxLimite) & ","
        End If
        
        
        If nT = 1 Then
            CSQL2 = CSQL2 & "'" & CStr(Date) & "','0')"
        End If
        cCod = TxCodigo
    
    

' clientecodigo clienteruc  clienterazonsocial
'clientedireccion      Clientedistrito                clienteprovincia
'clientedepartamento            negociocodigo estadoreg clientesiglas        clientetelefono
'clientefax clientemail  clientecodpostal clientetipopersona clientetipopais
'clientesuspendido clientelimitecredsoles  clientelimitecreddolar  clientesaldosoles
'clientesaldodolares  clienteaval clientediasmaxpagocont clientenumcopias clientecuentacontable
' clientefechabajaoanula clientefechaactivacion clientefechaultimavta  clientemultidireccion
' clientepropietario  clientepropdirecc clientepropcodpostal clienteproptelefono
' clientepropruc clienteprople clientedescuento Usuariocodigo fechaact
    
    
    
    
    
    
    ElseIf nT = 2 Then     'Modificar
        CSQL2 = "Update VT_CLIENTE Set clientecodigo='" & TxCodigo & "',clienterazonsocial='" & SupCadSQL(TxRazon) & "',"
        CSQL2 = CSQL2 & "clientedireccion='" & SupCadSQL(TxFactura) & "',clientetelefono='" & TxTelefono & "',clienteruc='" & txRuc & "',"
        CSQL2 = CSQL2 & "fechaact=" & Date & ",clientepropietario='" & TxRepresentante & "',"
        CSQL2 = CSQL2 & "Clientedistrito='" & SupCadSQL(TxDistrito) & "',Usuariocodigo='" & SupCadSQL(VGUsuario) & "',"
        CSQL2 = CSQL2 & "clientetipopais='" & TxPais & "',clientedepartamento='" & SupCadSQL(TxDepartamento) & "',"
        CSQL2 = CSQL2 & "clienteprovincia='" & TxProvincia & "',negociocodigo='" & TxGiro & "',"
         CSQL2 = CSQL2 & "clientemail='" & TxEmail & "',"
        CSQL2 = CSQL2 & "clientefax='" & TxFax & "',"
        
        If nT = 2 Then
            CSQL2 = CSQL2 & "clientefechaactivacion=" & Date & ",clientesuspendido='0' "
        End If
        CSQL2 = CSQL2 & "Where clientecodigo = '" & Trim(TxCodigo) & "'"
        
        cCod = TxCodigo
    ElseIf nT = 3 Then      'Ficha Tecnica
        CSQL2 = "Update VT_CLIENTE set Comenta = '" & SupCadSQL(TxTecnica) & "' "
        CSQL2 = CSQL2 & "Where clientecodigo = '" & Trim(cCod) & "'"
    End If
    
    nTra = 1
    VGCNx.BeginTrans
    VGCNx.Execute CSQL2
    VGCNx.CommitTrans
    nTra = 0
    adodc1.Requery
    
    Dim Nombre As String
    If nT = 1 Then
        cBase = cRuta4
        If UCase(Dir$(cBase)) = VGNameCont & ".MDB" Then
            'Se hace un enlace con los archivos de contabilidad, se busca y se graba
            
             
                Nombre = "ANEXOCLIE"
                cSql1 = "Select ConcGral_Contec from Conceptos_Generales Where ConcGral_Codigo= '" & UCase(Nombre) & "'"
                Set cSel1 = New ADODB.Recordset
                cSel1.Open cSql1, VGcnxCT, adOpenStatic
                If Not cSel1.EOF Then
                   cAnexo = cSel1("ConcGral_Contec")
                End If
                cSel1.Close
                
                cSql1 = "Select * from ANEXO Where TIPOANEX_CODIGO= '" & cAnexo & "' and ANEX_CODIGO = '" & Trim(TxCodigo) & "'"
                Set cSel1 = New ADODB.Recordset
                cSel1.Open cSql1, VGcnxCT, adOpenStatic
                If cSel1.RecordCount = 0 Then
                    cSql3 = "Insert Into ANEXO (TIPOANEX_CODIGO,ANEX_CODIGO,ANEX_DESCRIPCION,ANEX_RUC,ANEX_DIRECCION,"
                    cSql3 = cSql3 & "ANEX_TELEFONO,ANEX_REPRESENTANTE) values ('" & cAnexo & "','" & SupCadSQL(TxCodigo) & "','" & IIf(Trim(TxRazon) <> "", SupCadSQL(TxRazon), "0") & "','" & IIf(Trim(txRuc) <> "", txRuc, "0") & "',"
                    cSql3 = cSql3 & "'" & IIf(Trim(TxFactura) <> "", SupCadSQL(TxFactura), "0") & "','" & IIf(Trim(TxTelefono) <> "", TxTelefono, "0") & "','" & IIf(Trim(TxRepresentante) <> "", SupCadSQL(TxRepresentante), "0") & "')"
                    nTra2 = 1
                    VGcnxCT.BeginTrans
                    VGcnxCT.Execute cSql3
                    VGcnxCT.CommitTrans
                    nTra2 = 0
                End If
                cSel1.Close
            
        End If
    ElseIf nT = 2 Then
            cBase = cRuta4
             If UCase(Dir$(cBase)) = VGNameCont & ".MDB" Then
                 Nombre = "ANEXOCLIE"
                 cSql1 = "Select ConcGral_Contec from Conceptos_Generales Where ConcGral_Codigo= '" & UCase(Nombre) & "'"
                 Set cSel1 = New ADODB.Recordset
                 cSel1.Open cSql1, VGcnxCT, adOpenStatic
                 If Not cSel1.EOF Then
                    cAnexo = cSel1("ConcGral_Contec")
                 End If
                 cSel1.Close
                
                    cSql3 = "Update ANEXO Set TIPOANEX_CODIGO ='" & cAnexo & "' ,ANEX_CODIGO='" & TxCodigo & "' ,"
                    cSql3 = cSql3 & "ANEX_DESCRIPCION='" & IIf(Trim(FrmArClien.TxRazon) <> "", SupCadSQL(TxRazon), "0") & "',ANEX_RUC='" & IIf(Trim(txRuc) <> "", txRuc, "0") & "',ANEX_DIRECCION='" & IIf(Trim(TxRazon) <> "", SupCadSQL(TxRazon), "0") & "',"
                    cSql3 = cSql3 & "ANEX_TELEFONO='" & IIf(Trim(TxTelefono) <> "", TxTelefono, "0") & "',ANEX_REPRESENTANTE='" & IIf(Trim(TxRepresentante) <> "", SupCadSQL(TxRepresentante), "0") & "' Where TIPOANEX_CODIGO = '" & cAnexo & "' and ANEX_CODIGO = '" & Trim(TxCodigo) & "'"
                    nTra2 = 1
                    VGcnxCT.BeginTrans
                    VGcnxCT.Execute cSql3
                    VGcnxCT.CommitTrans
                    nTra2 = 0
           End If
    End If
    adodc1.Find "clientecodigo = '" & cCod & "'"
End If

If nT = 1 Then
    Limpiar
    LbFecha = Date
    SSTab1.Tab = 0
    TxCodigo.SetFocus
ElseIf nT = 2 Or nT = 3 Then
    CmdSalir2_Click
End If
Exit Sub
GrabErr:
    MsgBox Err.Description
    
    If nTra = 1 Then VGCNx.RollbackTrans
    If nTra2 = 1 Then VGcnxCT.RollbackTrans
End Sub

Private Sub CmdIng_Click()      'Ingresar
chkAsignarRUC.Enabled = True
chkAsignarRUC.Value = 1
nT = 1
Me.Caption = "Ingreso de Datos del Cliente"
OculObj04 (False)
OculObj05 (False)
OculObj06 (False)
OculObj02 (True)
OculObj03 (True)
OculObj01 (False)
Limpiar
LbFecha = Date
TxCodigo.Enabled = True
TxCodigo.SetFocus
End Sub

Private Sub CmdModi_Click()     'Modificar
If adodc1.RecordCount > 0 Then
    nT = 2
    Me.Caption = "Modificación de Datos"
    OculObj04 (False)
    OculObj05 (False)
    OculObj06 (False)
    OculObj02 (True)
    OculObj03 (True)
    Limpiar
    cCod = adodc1("clientecodigo")
    TxCodigo.Enabled = False
    Mostrar (cCod)
    If txRuc.Visible Then txRuc.SetFocus
Else
    MsgBox "No existen registros", vbInformation, "Mensaje"
End If
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub CmdSalir2_Click()   'Salida de la segunda pantalla
Me.Caption = "Actualiza Datos Generales del Cliente"
OculObj01 (False)
OculObj02 (False)
OculObj03 (False)
OculObj04 (True)
OculObj05 (True)
OculObj06 (True)
InhabObj (True)
Cmdgrabar.Enabled = True
Cmdgrabar.Visible = False
DataGrid1.SetFocus
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then TxLimite.SetFocus
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyBack Then
 If Len(TxFiltro) - 1 > 0 Then
  TxFiltro = Left(TxFiltro, Len(TxFiltro) - 1)
 Else
  TxFiltro = ""
 End If
 KeyAscii = 0
ElseIf KeyAscii <> 13 Then
 TxFiltro = TxFiltro & Chr(KeyAscii)
End If
End Sub

Private Sub Form_Activate()
TxFiltro = ""
CmbOrden.ListIndex = 0
Select Case FrmAyuda.cCod
Case "13"
      TxDistrito = FrmAyuda.cC
      LbDistrito = Mid(FrmAyuda.cD, 1, 15)
Case "62"
      TxGiro = FrmAyuda.cC
      LbGiro = Mid(FrmAyuda.cD, 1, 15)
End Select
 FrmAyuda.cCod = ""
 If DataGrid1.Visible Then DataGrid1.SetFocus
End Sub
Private Sub Form_Load()
central Me             ' Centrar Formulario
Set adodc1 = New ADODB.Recordset

Init_ControlDataGrid DataGrid1
LbFecha = Format(Now, "DD/MM/YYYY")

Limpiar
OculObj01 (False)
OculObj02 (False)
OculObj03 (False)
OculObj04 (True)
OculObj05 (True)
OculObj06 (True)

adodc1.Open "Select clientecodigo,clienterazonsocial,clienteruc  FROM VT_CLIENTE ORDER BY clientecodigo", VGCNx, adOpenStatic
Set DataGrid1.DataSource = adodc1
DataGrid1.Refresh
End Sub
Private Sub Limpiar()   'Limpia variables
TxCodigo = "": TxRazon = "": TxFactura = ""
TxEntrega = "": TxPais = "": TxProvincia = ""
 TxBanco = "": LbBanco = "": TxTelefono = "": TxEmail = ""
TxDscto = "0": TxPrecio = "": LbPrecio = ""
TxLimite = "0": TxVendedor = "": LbVendedor = ""
txRuc = "": TxDepartamento = "": TxDistrito = ""
LbDistrito = "": TxGiro = "": LbGiro = "": TxRepresentante = "": TxCta = ""
TxFax = "": TxHost = "": TxPago = "": LbPago = "": Combo1.ListIndex = 0
LbFecha = ""
End Sub

Private Sub Transf_Click()
Dim Nombre As String
cBase = cRuta4
 If UCase(Dir$(cBase)) = VGNameCont & ".MDB" Then
    'Se hace un enlace con los archivos de contabilidad, se busca y se graba
   If Not adodc1.EOF Then
      adodc1.MoveFirst
      Nombre = "ANEXOCLIE"
      cSql1 = "Select ConcGral_Contec from Conceptos_Generales Where ConcGral_Codigo= '" & UCase(Nombre) & "'"
      Set cSel1 = New ADODB.Recordset
      cSel1.Open cSql1, VGcnxCT, adOpenStatic
      If Not cSel1.EOF Then
         cAnexo = cSel1("ConcGral_Contec")
      End If
      cSel1.Close
      
         Do While Not adodc1.EOF
            cSql1 = "Select * from ANEXO Where TIPOANEX_CODIGO= '" & cAnexo & "' and ANEX_CODIGO = '" & Trim(adodc1("clientecodigo")) & "'"
            Set cSel1 = New ADODB.Recordset
            cSel1.Open cSql1, VGcnxCT, adOpenStatic
            
            cSql1 = "Select clientecodigo,clienterazonsocial,clienteruc,CDIRCLI,CTELEFO,CNOMREP FROM VT_CLIENTE Where clientecodigo= '" & Trim(adodc1("clientecodigo")) & "'"
            Set cSel2 = New ADODB.Recordset
            cSel2.Open cSql1, VGCNx, adOpenStatic
            If Not cSel2.EOF Then
               If cSel1.RecordCount = 0 Then
                  cSql3 = "Insert Into ANEXO (TIPOANEX_CODIGO,ANEX_CODIGO,ANEX_DESCRIPCION,ANEX_RUC,ANEX_DIRECCION,"
                  cSql3 = cSql3 & "ANEX_TELEFONO,ANEX_REPRESENTANTE) values ('" & cAnexo & "','" & Trim(cSel2("clientecodigo")) & "','" & IIf(Trim(cSel2("clienterazonsocial")) <> "", Trim(Mid(cSel2("clienterazonsocial"), 1, 50)), "0") & "','" & IIf(Trim(cSel2("clienteruc")) <> "", cSel2("clienteruc"), "0") & "',"
                  cSql3 = cSql3 & "'" & IIf(Trim(cSel2("CDIRCLI")) <> "", Mid(cSel2("CDIRCLI"), 1, 50), "0") & "','" & IIf(Trim(cSel2("CTELEFO")) <> "", Mid(cSel2("CTELEFO"), 1, 15), "0") & "','" & IIf(Trim(cSel2("CNOMREP")) <> "", cSel2("CNOMREP"), "0") & "')"
                  VGcnxCT.Execute cSql3
               Else
                   cSql3 = "Update ANEXO Set TIPOANEX_CODIGO ='" & cAnexo & "' ,ANEX_CODIGO='" & Trim(cSel2("clientecodigo")) & "' ,"
                   cSql3 = cSql3 & "ANEX_DESCRIPCION='" & IIf(Trim(cSel2("clienterazonsocial")) <> "", Trim(Mid(cSel2("clienterazonsocial"), 1, 50)), "0") & "',ANEX_RUC='" & IIf(cSel2("clienteruc") <> "", cSel2("clienteruc"), "0") & "',ANEX_DIRECCION='" & IIf(Trim(cSel2("CDIRCLI")) <> "", Mid(cSel2("CDIRCLI"), 1, 50), "0") & "',"
                   cSql3 = cSql3 & "ANEX_TELEFONO='" & IIf(Trim(cSel2("CTELEFO")) <> "", Mid(cSel2("CTELEFO"), 1, 15), "0") & "',ANEX_REPRESENTANTE='" & IIf(Trim(cSel2("CNOMREP")) <> "", Trim(cSel2("CNOMREP")), "0") & "' Where TIPOANEX_CODIGO = '" & cAnexo & "' and ANEX_CODIGO = '" & Trim(cSel2("clientecodigo")) & "'"
                   VGcnxCT.Execute cSql3
                End If
             End If
            cSel1.Close
            cSel2.Close
            adodc1.MoveNext
         Loop
       End If
      adodc1.MoveFirst
 End If
End Sub

Private Sub TxBanco_DblClick()
Static Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset

Adodc2.Open "SELECT BAN_CODIGO,BAN_DESCRIPCION FROM BANCO", VGCNx, adOpenStatic, adLockOptimistic
frmReferencia.Conectar Adodc2, "SELECT BAN_CODIGO,BAN_DESCRIPCION FROM BANCO"
frmReferencia.Label1.Caption = "Tabla de Bancos"
frmReferencia.Show vbModal
Adodc2.Close
If vGUtil(1) <> "" Then
  TxBanco.text = (vGUtil(1))
  LbBanco = vGUtil(2)
End If
End Sub
Private Sub TxBanco_GotFocus()
Enfoque TxBanco
End Sub

Private Sub TxBanco_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
   TxBanco_DblClick
Else
    If KeyCode = 46 Then LbBanco = ""
End If
End Sub

Private Sub TxBanco_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(TxBanco) <> "" Then
        If Existe(1, TxBanco, "BANCO", "BAN_CODIGO", False) = False Then
            MsgBox "El código de Banco, no existe", vbInformation, "Mensaje"
            TxBanco.SetFocus: Exit Sub
        Else
            LbBanco = Devolver_Dato(1, TxBanco, "BANCO", "BAN_CODIGO", False, "BAN_DESCRIPCION")
        End If
    End If
    TxCta.SetFocus
Else
    If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
End If
End Sub
Private Sub TxCodigo_GotFocus()
Enfoque TxCodigo
End Sub
Private Sub TxCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If CodigoC(TxCodigo) Then
        SendKeys "{tab}"
        'TxRuc.SetFocus
    Else
        MsgBox "Codigo ya existe", vbInformation, "Mensaje"
        TxCodigo.SetFocus
    End If
Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub
Private Sub TxCta_GotFocus()
Enfoque TxCta
End Sub
Private Sub TxCta_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    SSTab1.Tab = 1
    TxDscto.SetFocus
Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If (Chr$(KeyAscii) > "9" Or Chr$(KeyAscii) < "0") And Chr$(KeyAscii) <> "-" And KeyAscii <> 8 Then KeyAscii = 0
End If
End Sub
Private Sub TxDepartamento_GotFocus()
Enfoque TxDepartamento
End Sub
Private Sub TxDepartamento_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxProvincia.SetFocus
Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub

Private Sub TxDistrito_DblClick()
FrmAyuda.cCod = "13"
FrmAyuda.Show 1
End Sub
Private Sub TxDistrito_GotFocus()
Enfoque TxDistrito
End Sub

Private Sub TxDistrito_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
   TxDistrito_DblClick
Else
    If KeyCode = 46 Then LbDistrito = ""
End If
End Sub

Private Sub TxDistrito_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(TxDistrito) <> "" Then LbDistrito = Mid(fDis(TxDistrito), 1, 15)
    If Trim(LbDistrito) = "" And Trim(TxDistrito) <> "" Then
 '       MsgBox "El código de Distrito no existe", vbInformation, "Mensaje"
 '       TxDistrito = "": TxDistrito.SetFocus
    Else
        TxTelefono.SetFocus
    End If
End If
End Sub
Private Sub TxDscto_GotFocus()
Enfoque TxDscto
End Sub
Private Sub TxDscto_KeyPress(KeyAscii As Integer)
Dim I As Integer

If KeyAscii = 13 Then
   TxPago.SetFocus
Else
  If (Chr$(KeyAscii) > "9" Or Chr$(KeyAscii) < "0") And Chr$(KeyAscii) <> "." And KeyAscii <> 8 Then
     KeyAscii = 0
  Else
     If Chr$(KeyAscii) = "." Then
        For I = 1 To Len(TxDscto)
            If Mid(TxDscto, I, 1) = "." Then KeyAscii = 0: Exit Sub
        Next
        
     End If
  End If
End If

End Sub
Private Sub TxEmail_GotFocus()
Enfoque TxEmail
End Sub
Private Sub TxEmail_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then TxHost.SetFocus
End Sub

Private Sub TxEntrega_DblClick()
FrmDire.cCliente = Trim(TxCodigo)
FrmDire.Show 1
End Sub

Private Sub TxEntrega_GotFocus()
Enfoque TxEntrega
End Sub
Private Sub TxEntrega_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then TxEntrega_DblClick
End Sub

Private Sub TxEntrega_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxPais.SetFocus
Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub
Private Sub TxFactura_GotFocus()
Enfoque TxFactura
End Sub
Private Sub TxFactura_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxEntrega.SetFocus
Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub
Private Sub Txfax_GotFocus()
Enfoque TxFax
End Sub
Private Sub Txfax_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   TxEmail.SetFocus
Else
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub

Private Sub TxFiltro_Change()
'If adodc1.RecordCount > 0 Then
'    If Trim(TxFiltro) <> "" And TxFiltro.Visible Then
'        nCursor = adodc1.Bookmark
'        adodc1.AbsolutePosition = 1
'        adodc1.MoveFirst
        
 '       Select Case CmbOrden.ListIndex
 '       Case 0
 '           adodc1.Find "clientecodigo LIKE '%" & Trim(UCase(TxFiltro)) & "%'"
 '      Case 1
 '           adodc1.Find "clienterazonsocial LIKE '%" & Trim(UCase(TxFiltro)) & "%' "
 '       Case 2
 '           adodc1.Find "clienteruc LIKE '%" & Trim(UCase(TxFiltro)) & "%'"
 '       End Select
 '       If adodc1.EOF Then
 '          adodc1.AbsolutePosition = nCursor
 '        Else
 '          DataGrid1.Refresh
 '       End If
 '     Else
 '       adodc1.Find "clientecodigo like '%%'"
 '   End If
'End If
End Sub

Private Sub TxGiro_DblClick()
FrmAyuda.cCod = "62"
FrmAyuda.Show 1
End Sub
Private Sub TxGiro_GotFocus()
Enfoque TxGiro
End Sub

Private Sub TxGiro_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then TxGiro_DblClick
End Sub

Private Sub TxGiro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(TxGiro) <> "" Then LbGiro = fGir(TxGiro)
    If Trim(LbGiro) = "" And Trim(TxGiro) <> "" Then
        MsgBox "El código no existe", vbInformation, "Mensaje"
        TxGiro = "": TxGiro.SetFocus
    Else
        TxVendedor.SetFocus
    End If
End If
End Sub
Private Sub TxHost_GotFocus()
Enfoque TxHost
End Sub
Private Sub TxHost_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then TxBanco.SetFocus
End Sub
Private Sub TxLimite_GotFocus()
Enfoque TxLimite
End Sub
Private Sub TxLimite_KeyPress(KeyAscii As Integer)
Dim I As Integer

If KeyAscii = 13 Then
   TxGiro.SetFocus
Else
  If (Chr$(KeyAscii) > "9" Or Chr$(KeyAscii) < "0") And Chr$(KeyAscii) <> "." And KeyAscii <> 8 Then
     KeyAscii = 0
  Else
     If Chr$(KeyAscii) = "." Then
        For I = 1 To Len(TxGiro)
            If Mid(TxGiro, I, 1) = "." Then KeyAscii = 0: Exit Sub
        Next
        
     End If
  End If
End If
End Sub
Private Sub TxPago_DblClick()
Static Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset

Adodc2.Open "SELECT COD_FP,DES_FP FROM FORMA_PAGO", VGCNx, adOpenStatic, adLockOptimistic
frmReferencia.Conectar Adodc2, "SELECT COD_FP,DES_FP FROM FORMA_PAGO"
frmReferencia.Label1.Caption = "Formas de Pago"
frmReferencia.Show vbModal
Adodc2.Close
If vGUtil(1) <> "" Then
  TxPago.text = (vGUtil(1))
  LbPago = vGUtil(2)
End If
End Sub
Private Sub TxPago_GotFocus()
Enfoque TxPago
End Sub

Private Sub TxPago_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
   TxPago_DblClick
Else
    If KeyCode = 46 Then LbPago = ""
End If
End Sub

Private Sub TxPago_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(TxPago) <> "" Then
        If Existe(1, TxPago, "FORMA_PAGO", "COD_FP", False) = False Then
            MsgBox "El código de Forma de Pago, no existe", vbInformation, "Mensaje"
            TxPago.SetFocus: Exit Sub
        Else
            LbPago = Devolver_Dato(1, TxPago, "FORMA_PAGO", "COD_FP", False, "DES_FP")
        End If
    End If
    TxPrecio.SetFocus
End If
End Sub
Private Sub TxPais_GotFocus()
Enfoque TxPais
End Sub
Private Sub TxPais_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxDepartamento.SetFocus
Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub
Private Sub TxPrecio_DblClick()
Static Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset

Adodc2.Open "SELECT Cod_LisPre,Des_LisPre FROM TIPO_PRECIO", VGCNx, adOpenStatic, adLockOptimistic
frmReferencia.Conectar Adodc2, "SELECT Cod_LisPre,Des_LisPre FROM TIPO_PRECIO"
frmReferencia.Label1.Caption = "Tipos de Precios"
frmReferencia.Show vbModal
Adodc2.Close
If vGUtil(1) <> "" Then
  TxPrecio.text = (vGUtil(1))
  LbPrecio = vGUtil(2)
End If

End Sub
Private Sub TxPrecio_GotFocus()
Enfoque TxPrecio
End Sub

Private Sub TxPrecio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
   TxPrecio_DblClick
Else
    If KeyCode = 46 Then LbPrecio = ""
End If
End Sub

Private Sub TxPrecio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(TxPrecio) <> "" Then LbPrecio = Mid(fPre(TxPrecio), 1, 15)
    If Trim(LbPrecio) = "" And Trim(TxPrecio) <> "" Then
        MsgBox "El código de Precio, no existe", vbInformation, "Mensaje"
        TxPrecio = "": TxPrecio.SetFocus
    Else
        Combo1.SetFocus
    End If
End If
End Sub
Private Sub TxProvincia_GotFocus()
Enfoque TxProvincia
End Sub
Private Sub TxProvincia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxDistrito.SetFocus
Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub
Private Sub TxRazon_GotFocus()
Enfoque TxRazon
End Sub
Private Sub TxRazon_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(TxRazon) <> "" Then
        InhTex (True)
        TxFactura.SetFocus
    Else
        InhTex (False)
        MsgBox "Ingrese Razón Social", vbInformation, "Mensaje"
        TxRazon.SetFocus
    End If
Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub
Private Sub TxRepresentante_GotFocus()
Enfoque TxRepresentante
End Sub
Private Sub TxRepresentante_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Cmdgrabar.SetFocus
Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub
Private Sub TxRuc_GotFocus()
Enfoque txRuc
End Sub
Private Sub TxRuc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If txRuc <> "" Then
      If Validar_RUC(txRuc) = False Then
         txRuc.SetFocus: Exit Sub
      End If
   End If
   TxRazon.SetFocus
End If
End Sub
Private Sub TxTelefono_GotFocus()
Enfoque TxTelefono
End Sub
Private Sub TxTelefono_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   TxFax.SetFocus
   
Else
  If (Chr$(KeyAscii) > "9" Or Chr$(KeyAscii) < "0") And Chr$(KeyAscii) <> "-" And KeyAscii <> 8 Then KeyAscii = 0
End If
End Sub
Private Sub TxVendedor_DblClick()
Static Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset

Adodc2.Open "SELECT Cod_Ven,Des_Ven FROM VENDEDOR", VGCNx, adOpenStatic, adLockOptimistic
frmReferencia.Conectar Adodc2, "SELECT Cod_Ven,Des_Ven FROM VENDEDOR"
frmReferencia.Label1.Caption = "Vendedores"
frmReferencia.Show vbModal
Adodc2.Close
If vGUtil(1) <> "" Then
  TxVendedor.text = (vGUtil(1))
  LbVendedor = vGUtil(2)
End If
End Sub
Private Sub TxVendedor_GotFocus()
Enfoque TxVendedor
End Sub

Private Sub TxVendedor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
   TxVendedor_DblClick
Else
    If KeyCode = 46 Then LbVendedor = ""
End If
End Sub

Private Sub TxVendedor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(TxVendedor) <> "" Then
        If Existe(1, TxVendedor, "VENDEDOR", "Cod_Ven", False) = False Then
            MsgBox "El código de Vendedor no existe", vbInformation, "Mensaje"
            TxVendedor = "": TxVendedor.SetFocus: Exit Sub
        Else
            LbVendedor = Devolver_Dato(1, TxVendedor, "VENDEDOR", "COD_VEN", False, "Des_Ven")
        End If
    End If
    TxRepresentante.SetFocus
End If
End Sub

Private Sub Mostrar(cC1 As String) 'Muestra los datos
Dim cSqlM As String, cSelM As ADODB.Recordset
If Trim(cC1) = "" Then
    MsgBox "No hay registros para mostrar", vbInformation, "Mensaje"
    Exit Sub
End If
cSqlM = "Select * From VT_CLIENTE Where clientecodigo = '" & cC1 & "'"
Set cSelM = New ADODB.Recordset
cSelM.Open cSqlM, VGCNx, adOpenStatic
If cSelM.RecordCount > 0 Then

 


    TxCodigo = cSelM("clientecodigo")
    If Not IsNull(cSelM("clienterazonsocial")) Then TxRazon = cSelM("clienterazonsocial")
    If Not IsNull(cSelM("clientedireccion")) Then TxFactura = cSelM("clientedireccion")
'    If Not IsNull(cSelM("DIRENT")) Then TxEntrega = cSelM("DIRENT")
    If Not IsNull(cSelM("clientetipopais")) Then TxPais = cSelM("clientetipopais")
    If Not IsNull(cSelM("clienteprovincia")) Then TxProvincia = cSelM("clienteprovincia")
'    If Not IsNull(cSelM("CCODBAN")) Then TxBanco = cSelM("CCODBAN")
'    If Not IsNull(cSelM("CTELEFO")) Then TxTelefono = cSelM("CTELEFO")
    If Not IsNull(cSelM("clientemail")) Then TxEmail = cSelM("clientemail")
'    If Not IsNull(cSelM("NPORDES")) Then TxDscto = cSelM("NPORDES")
'    If Not IsNull(cSelM("CTIPPRE")) Then TxPrecio = cSelM("CTIPPRE")
'    If Not IsNull(cSelM("CVENDE")) Then TxVendedor = cSelM("CVENDE")
    If Not IsNull(cSelM("clienteruc")) Then txRuc = cSelM("clienteruc")
    If Not IsNull(cSelM("clientedepartamento")) Then TxDepartamento = cSelM("clientedepartamento")
    If Not IsNull(cSelM("Clientedistrito")) Then TxDistrito = cSelM("Clientedistrito")
    If Not IsNull(cSelM("negociocodigo")) Then TxGiro = cSelM("negociocodigo")
    If Not IsNull(cSelM("clientepropietario")) Then TxRepresentante = cSelM("clientepropietario")
'    If Not IsNull(cSelM("CNUMCTA")) Then TxCta = cSelM("CNUMCTA")
    If Not IsNull(cSelM("clientefax")) Then TxFax = cSelM("clientefax")
'    If Not IsNull(cSelM("CHOST")) Then TxHost = cSelM("CHOST")
'    If Not IsNull(cSelM("CTIPVTA")) Then TxPago = cSelM("CTIPVTA")
'    If Not IsNull(cSelM("MONCRE")) Then
'       If cSelM("MONCRE") <> "01" Then
'          Combo1.ListIndex = 0
'       Else
'          Combo1.ListIndex = 1
'       End If
'    End If
'    If Not IsNull(cSelM("DFECCRE")) Then LbFecha = Format(cSelM("DFECCRE"), "DD/MM/YYYY")
    
'    If cSelM("MONCRE") <> "01" Then
'        TxLimite = "" & cSelM("LMCRUS")
'    ElseIf cSelM("MONCRE") = "01" Then
'        TxLimite = cSelM("LMCRMN")
'    End If
    
    LbDistrito = Mid(fDis(TxDistrito), 1, 15)
    LbGiro = Mid(fGir(TxGiro), 1, 15)
    LbPrecio = Mid(fPre(TxPrecio), 1, 15)
Else
    MsgBox "No existe registro", vbInformation, "Mensaje"
    CmdSalir2_Click
End If
cSelM.Close
End Sub

Private Sub InhabObj(ntipo As Boolean) ' Habilita e Inhabilita los objetos
TxCodigo.Enabled = ntipo
TxRazon.Enabled = ntipo
TxFactura.Enabled = ntipo
TxEntrega.Enabled = ntipo
TxPais.Enabled = ntipo
TxProvincia.Enabled = ntipo
TxBanco.Enabled = ntipo
TxTelefono.Enabled = ntipo
TxEmail.Enabled = ntipo
TxDscto.Enabled = ntipo
TxPrecio.Enabled = ntipo
TxLimite.Enabled = ntipo
TxVendedor.Enabled = ntipo
txRuc.Enabled = ntipo
TxDepartamento.Enabled = ntipo
TxDistrito.Enabled = ntipo
TxGiro.Enabled = ntipo
TxRepresentante.Enabled = ntipo
TxCta.Enabled = ntipo
TxFax.Enabled = ntipo
TxHost.Enabled = ntipo
TxPago.Enabled = ntipo
Combo1.Enabled = ntipo
End Sub
Private Sub InhTex(ntipo As Boolean)
TxFactura.Enabled = ntipo
TxEntrega.Enabled = ntipo
TxPais.Enabled = ntipo
TxProvincia.Enabled = ntipo
TxBanco.Enabled = ntipo
TxTelefono.Enabled = ntipo
TxEmail.Enabled = ntipo
TxDscto.Enabled = ntipo
TxPrecio.Enabled = ntipo
TxLimite.Enabled = ntipo
TxVendedor.Enabled = ntipo
TxDepartamento.Enabled = ntipo
TxDistrito.Enabled = ntipo
TxGiro.Enabled = ntipo
TxRepresentante.Enabled = ntipo
TxCta.Enabled = ntipo
TxFax.Enabled = ntipo
TxHost.Enabled = ntipo
TxPago.Enabled = ntipo
Combo1.Enabled = ntipo
End Sub
