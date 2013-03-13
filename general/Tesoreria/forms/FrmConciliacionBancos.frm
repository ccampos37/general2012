VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Begin VB.Form FrmConciliacionBancos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conciliacion Bancaria"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13440
   Icon            =   "FrmConciliacionBancos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   13440
   Begin VB.Frame Frame2 
      Caption         =   "Opciones"
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   90
      TabIndex        =   37
      Top             =   8190
      Width           =   8055
      Begin VB.CommandButton Cmdcancelar 
         Caption         =   "Cancelar"
         Height          =   330
         Left            =   1035
         TabIndex        =   44
         Top             =   225
         Width           =   825
      End
      Begin VB.CommandButton cmdaceptar 
         Caption         =   "Grabar"
         Height          =   330
         Left            =   90
         TabIndex        =   43
         Top             =   225
         Width           =   915
      End
      Begin VB.CommandButton Cmdeliminar 
         Caption         =   "Eliminar"
         Height          =   330
         Left            =   1890
         TabIndex        =   42
         Top             =   225
         Width           =   915
      End
      Begin VB.CommandButton cmdmodificar 
         Caption         =   "Modificar"
         Height          =   330
         Left            =   2835
         TabIndex        =   41
         Top             =   225
         Width           =   915
      End
      Begin VB.CommandButton Cmdimprimir 
         Caption         =   "Imp.Conciliados"
         Height          =   330
         Index           =   0
         Left            =   4005
         TabIndex        =   40
         Top             =   225
         Width           =   1275
      End
      Begin VB.CommandButton Cmdimprimir 
         Caption         =   "Imp.Pendientes"
         Height          =   330
         Index           =   1
         Left            =   5310
         TabIndex        =   39
         Top             =   225
         Width           =   1230
      End
      Begin VB.CommandButton Cmdimprimir 
         Caption         =   "Imprimir Todo"
         Height          =   330
         Index           =   2
         Left            =   6570
         TabIndex        =   38
         Top             =   225
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1875
      Left            =   135
      TabIndex        =   4
      Top             =   30
      Width           =   13215
      Begin TextFer.TxFer Txtsaldoini 
         Height          =   315
         Left            =   6165
         TabIndex        =   23
         Top             =   495
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         Object.CausesValidation=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         ColorIlumina    =   13040123
         Valor           =   ""
         TipoDato        =   1
      End
      Begin VB.PictureBox axButton1 
         Height          =   1680
         Left            =   6060
         ScaleHeight     =   1680
         ScaleWidth      =   30
         TabIndex        =   22
         Top             =   135
         Width           =   30
      End
      Begin VB.CheckBox chkconciliado 
         Caption         =   "Doc. Conciliados"
         Height          =   225
         Left            =   2730
         TabIndex        =   15
         Top             =   1515
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker DTPfechaini 
         Height          =   285
         Left            =   1065
         TabIndex        =   13
         Top             =   1485
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   " MMM - yyyy"
         Format          =   50790403
         CurrentDate     =   37513
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaBancoCuenta 
         Height          =   315
         Left            =   1095
         TabIndex        =   16
         Top             =   585
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   556
         XcodMaxLongitud =   25
         xcodwith        =   2000
         NomTabla        =   "te_cuentabancos"
         TituloAyuda     =   "Cuentas Bancarias"
         ListaCampos     =   "cbanco_codigo(1),cbanco_numero(1),monedacodigo(1)"
         XcodCampo       =   "cbanco_codigo"
         XListCampo      =   "cbanco_numero"
         ListaCamposDescrip=   "CodBanco,Cuenta,Moneda"
         ListaCamposText =   "cbanco_codigo,cbanco_numero,monedacodigo"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaBanco 
         Height          =   300
         Left            =   1095
         TabIndex        =   17
         Top             =   270
         Width           =   4860
         _ExtentX        =   8573
         _ExtentY        =   529
         XcodMaxLongitud =   3
         xcodwith        =   800
         NomTabla        =   "gr_banco"
         TituloAyuda     =   "Ayuda de Bancos"
         ListaCampos     =   "bancocodigo(1),bancodescripcion(1)"
         XcodCampo       =   "bancocodigo"
         XListCampo      =   "bancodescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "bancocodigo,bancodescripcion"
      End
      Begin VB.PictureBox axButton2 
         Height          =   195
         Left            =   75
         ScaleHeight     =   135
         ScaleWidth      =   5925
         TabIndex        =   20
         Top             =   915
         Width           =   5985
      End
      Begin TextFer.TxFer TxSaldExtBanc 
         Height          =   315
         Left            =   6150
         TabIndex        =   28
         Top             =   1350
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         Object.CausesValidation=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         ColorIlumina    =   13040123
         Valor           =   ""
         TipoDato        =   1
         NumeroDecimales =   2
      End
      Begin VB.Label Label3 
         Caption         =   "Saldo Extracto Bancario"
         Height          =   420
         Left            =   6195
         TabIndex        =   29
         Top             =   900
         Width           =   1440
      End
      Begin VB.Label LbTotales 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBE0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   5
         Left            =   11190
         TabIndex        =   27
         Top             =   1320
         Width           =   1635
      End
      Begin VB.Label LbTotales 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FBE3D9&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   9030
         TabIndex        =   26
         Top             =   1320
         Width           =   1635
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   8310
         TabIndex        =   25
         Top             =   1380
         Width           =   615
      End
      Begin VB.Label lbMon 
         Caption         =   "Moneda : "
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   105
         TabIndex        =   24
         Top             =   1125
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Saldo Final"
         Height          =   285
         Left            =   6165
         TabIndex        =   21
         Top             =   225
         Width           =   1440
      End
      Begin VB.Label Label5 
         Caption         =   "Cuenta"
         Height          =   285
         Left            =   135
         TabIndex        =   19
         Top             =   675
         Width           =   885
      End
      Begin VB.Label Label4 
         Caption         =   "Banco"
         Height          =   255
         Left            =   135
         TabIndex        =   18
         Top             =   315
         Width           =   885
      End
      Begin VB.Label lbfechini 
         Caption         =   "Periodo"
         Height          =   240
         Left            =   90
         TabIndex        =   14
         Top             =   1515
         Width           =   1065
      End
      Begin VB.Label LbTotales 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBE0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   11190
         TabIndex        =   12
         Top             =   870
         Width           =   1635
      End
      Begin VB.Label LbTotales 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBE0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   11190
         TabIndex        =   11
         Top             =   585
         Width           =   1635
      End
      Begin VB.Label LbTotales 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FBE3D9&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   9045
         TabIndex        =   10
         Top             =   870
         Width           =   1635
      End
      Begin VB.Label LbTotales 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FBE3D9&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   9045
         TabIndex        =   9
         Top             =   585
         Width           =   1635
      End
      Begin VB.Label LeDolares 
         AutoSize        =   -1  'True
         Caption         =   "TOT. DOLARES US$"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   11100
         TabIndex        =   8
         Top             =   315
         Width           =   1800
      End
      Begin VB.Label leSoles 
         AutoSize        =   -1  'True
         Caption         =   "TOT. SOLES S/."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   9030
         TabIndex        =   7
         Top             =   360
         Width           =   1440
      End
      Begin VB.Label leHaber 
         AutoSize        =   -1  'True
         Caption         =   "HABER"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   8295
         TabIndex        =   6
         Top             =   930
         Width           =   645
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         Height          =   1470
         Left            =   8025
         Shape           =   4  'Rounded Rectangle
         Top             =   255
         Width           =   5055
      End
      Begin VB.Label leDebe 
         AutoSize        =   -1  'True
         Caption         =   "DEBE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   8280
         TabIndex        =   5
         Top             =   630
         Width           =   510
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         Height          =   1500
         Left            =   8010
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   5085
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   6045
      Left            =   135
      ScaleHeight     =   5985
      ScaleWidth      =   13140
      TabIndex        =   0
      Top             =   2085
      Width           =   13200
      Begin TrueOleDBGrid70.TDBGrid TDBG_concil 
         Height          =   5400
         Left            =   60
         TabIndex        =   1
         Top             =   60
         Width           =   12945
         _ExtentX        =   22834
         _ExtentY        =   9525
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Nº Recibo"
         Columns(0).DataField=   "cabrec_numrecibo"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Año"
         Columns(1).DataField=   "Anno"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Mes"
         Columns(2).DataField=   "Mes"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "T/D"
         Columns(3).DataField=   "detrec_tipodoc_concepto"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Nº Doc"
         Columns(4).DataField=   "detrec_numdocumento"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Fecha"
         Columns(5).DataField=   "detrec_fechacancela"
         Columns(5).NumberFormat=   "Short Date"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "T/Dc"
         Columns(6).DataField=   "detrec_tdqc"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Nº Doc Cancela"
         Columns(7).DataField=   "detrec_ndqc"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "I/E"
         Columns(8).DataField=   "cabrec_ingsal"
         Columns(8).NumberFormat=   "###,###,###.00"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Importe Soles"
         Columns(9).DataField=   "detrec_importesoles"
         Columns(9).NumberFormat=   "###,###,###.00"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   4
         Columns(10)._MaxComboItems=   5
         Columns(10).ValueItems(0)._DefaultItem=   0
         Columns(10).ValueItems(0).Value=   "1"
         Columns(10).ValueItems(0).Value.vt=   8
         Columns(10).ValueItems(0).DisplayValue=   "1"
         Columns(10).ValueItems(0).DisplayValue.vt=   8
         Columns(10).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(10).ValueItems(1)._DefaultItem=   0
         Columns(10).ValueItems(1).Value=   "0"
         Columns(10).ValueItems(1).Value.vt=   8
         Columns(10).ValueItems(1).DisplayValue=   "0"
         Columns(10).ValueItems(1).DisplayValue.vt=   8
         Columns(10).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
         Columns(10).ValueItems.Count=   2
         Columns(10).Caption=   "CH"
         Columns(10).DataField=   "chkconcil"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "Importe Dolares"
         Columns(11).DataField=   "detrec_importedolares"
         Columns(11).NumberFormat=   "###,###,###.00"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "Observaciones"
         Columns(12).DataField=   "detrec_observacion"
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(13)._VlistStyle=   0
         Columns(13)._MaxComboItems=   5
         Columns(13).Caption=   "Fecha Concil"
         Columns(13).DataField=   "fechconcil"
         Columns(13).NumberFormat=   "dd/mm/yyyy"
         Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   14
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=14"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1931"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1852"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=979"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=900"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8196"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=661"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=582"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=8196"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=635"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=556"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=8196"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=8196"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=1640"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=1561"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=8196"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=847"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=767"
         Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=8196"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(36)=   "Column(7).Width=2672"
         Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=2593"
         Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=8196"
         Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(41)=   "Column(8).Width=582"
         Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=503"
         Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=8194"
         Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(46)=   "Column(9).Width=2196"
         Splits(0)._ColumnProps(47)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(48)=   "Column(9)._WidthInPix=2117"
         Splits(0)._ColumnProps(49)=   "Column(9)._ColStyle=8194"
         Splits(0)._ColumnProps(50)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(51)=   "Column(10).Width=714"
         Splits(0)._ColumnProps(52)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(53)=   "Column(10)._WidthInPix=635"
         Splits(0)._ColumnProps(54)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(55)=   "Column(11).Width=2196"
         Splits(0)._ColumnProps(56)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(57)=   "Column(11)._WidthInPix=2117"
         Splits(0)._ColumnProps(58)=   "Column(11)._ColStyle=8194"
         Splits(0)._ColumnProps(59)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(60)=   "Column(12).Width=3995"
         Splits(0)._ColumnProps(61)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(62)=   "Column(12)._WidthInPix=3916"
         Splits(0)._ColumnProps(63)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(64)=   "Column(13).Width=2725"
         Splits(0)._ColumnProps(65)=   "Column(13).DividerColor=0"
         Splits(0)._ColumnProps(66)=   "Column(13)._WidthInPix=2646"
         Splits(0)._ColumnProps(67)=   "Column(13).Order=14"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         Appearance      =   2
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         MultiSelect     =   2
         AnimateWindow   =   2
         AnimateWindowClose=   2
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
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&H344A87&"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.locked=-1"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.locked=-1"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.locked=-1"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=54,.parent=13,.locked=-1"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=58,.parent=13,.locked=-1"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=62,.parent=13,.locked=-1"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=66,.parent=13,.locked=-1"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=63,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=64,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=65,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=70,.parent=13,.locked=-1"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=67,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=68,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=69,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=74,.parent=13,.alignment=1,.locked=-1"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=71,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=72,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=73,.parent=17"
         _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=78,.parent=13,.alignment=1,.locked=-1"
         _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=75,.parent=14"
         _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=76,.parent=15"
         _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=77,.parent=17"
         _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=50,.parent=13,.bgcolor=&HBFFFAA&"
         _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=47,.parent=14"
         _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=48,.parent=15"
         _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=49,.parent=17"
         _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=82,.parent=13,.alignment=1,.locked=-1"
         _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=79,.parent=14"
         _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=80,.parent=15"
         _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=81,.parent=17"
         _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=86,.parent=13"
         _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=83,.parent=14"
         _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=84,.parent=15"
         _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=85,.parent=17"
         _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=90,.parent=13"
         _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=87,.parent=14"
         _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=88,.parent=15"
         _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=89,.parent=17"
         _StyleDefs(92)  =   "Named:id=33:Normal"
         _StyleDefs(93)  =   ":id=33,.parent=0"
         _StyleDefs(94)  =   "Named:id=34:Heading"
         _StyleDefs(95)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(96)  =   ":id=34,.wraptext=-1"
         _StyleDefs(97)  =   "Named:id=35:Footing"
         _StyleDefs(98)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(99)  =   "Named:id=36:Selected"
         _StyleDefs(100) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(101) =   "Named:id=37:Caption"
         _StyleDefs(102) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(103) =   "Named:id=38:HighlightRow"
         _StyleDefs(104) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(105) =   "Named:id=39:EvenRow"
         _StyleDefs(106) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(107) =   "Named:id=40:OddRow"
         _StyleDefs(108) =   ":id=40,.parent=33"
         _StyleDefs(109) =   "Named:id=41:RecordSelector"
         _StyleDefs(110) =   ":id=41,.parent=34"
         _StyleDefs(111) =   "Named:id=42:FilterBar"
         _StyleDefs(112) =   ":id=42,.parent=33"
      End
      Begin VB.Label lbnreg 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "0 "
         Height          =   255
         Left            =   11955
         TabIndex        =   3
         Top             =   5520
         Width           =   1050
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Registros :"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   10860
         TabIndex        =   2
         Top             =   5550
         Width           =   975
      End
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   11595
      TabIndex        =   36
      Top             =   8280
      Width           =   615
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "HABER CONCIL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   9960
      TabIndex        =   35
      Top             =   8295
      Width           =   1380
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "DEBE CONCIL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   8340
      TabIndex        =   34
      Top             =   8280
      Width           =   1245
   End
   Begin VB.Label lbtot 
      Appearance      =   0  'Flat
      BackColor       =   &H00DDF7F9&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   2
      Left            =   11580
      TabIndex        =   33
      Top             =   8520
      Width           =   1605
   End
   Begin VB.Label lbtot 
      Appearance      =   0  'Flat
      BackColor       =   &H00DDF7F9&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   1
      Left            =   9930
      TabIndex        =   32
      Top             =   8535
      Width           =   1605
   End
   Begin VB.Label Label8 
      Caption         =   "Label6"
      Height          =   270
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   2040
   End
   Begin VB.Label lbtot 
      Appearance      =   0  'Flat
      BackColor       =   &H00DDF7F9&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   8310
      TabIndex        =   30
      Top             =   8535
      Width           =   1605
   End
End
Attribute VB_Name = "FrmConciliacionBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents RsConcil As ADODB.Recordset
Attribute RsConcil.VB_VarHelpID = -1
Dim RsSaldoIni As ADODB.Recordset
Dim tmontosolesDebe As Double, tmontodolaresDebe As Double
Dim tmontosolesHaber As Double, tmontodolaresHaber As Double
Dim montosolesDebe As Double, montodolaresDebe As Double
Dim montosolesHaber As Double, montodolaresHaber As Double
Dim mtsoles As Double, mtdolar As Double

Dim tsoles As Double, tdolar As Double
Dim montoextbanc As Double
Dim mon As String
Dim flagcal As Boolean

Private Sub axBAceptar_MouseEnter()

End Sub

Private Sub chkconciliado_Click()
 If Ctr_Ayudabanco.xclave <> Empty Then
    Call Listar
    Call CalcularTotales(RsConcil)
 End If
End Sub

Private Sub Ctr_AyudaBanco_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    Ctr_AyudaBancoCuenta.Filtro = "cbanco_codigo='" & Trim(ColecCampos("bancocodigo").Value) & "'"
End Sub
Private Sub Ctr_AyudaBancoCuenta_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Dim vardllgen As New dllgeneral.dll_general
    mon = ColecCampos("monedacodigo").Value
        
    lbMon.Caption = IIf(ColecCampos("monedacodigo").Value = "01", "Moneda de origen : Soles", "Moneda de Origen : Dolares ")
    
    Call Listar
    Select Case mon
        Case "01":
            LeDolares.Visible = False
            LbTotales(3).Visible = False
            LbTotales(4).Visible = False
            LbTotales(5).Visible = False
            TDBG_concil.Columns(11).Visible = False
            
            leSoles.Visible = True
            LbTotales(0).Visible = True
            LbTotales(1).Visible = True
            LbTotales(2).Visible = True
            TDBG_concil.Columns(9).Visible = True
            
            
        Case "02"
            leSoles.Visible = False
            LbTotales(0).Visible = False
            LbTotales(1).Visible = False
            LbTotales(2).Visible = False
            TDBG_concil.Columns(11).Visible = True
            
            LeDolares.Visible = True
            LbTotales(3).Visible = True
            LbTotales(4).Visible = True
            LbTotales(5).Visible = True
            TDBG_concil.Columns(9).Visible = False
    End Select
End Sub

Private Sub DTPfechaini_Change()
    Call Listar
End Sub

Private Sub Form_Initialize()
    Ctr_Ayudabanco.SetFocus
End Sub

Private Sub Form_Load()
    'lbfechini.Enabled = False
    'DTPfechaini.Enabled = False
    
    Width = 13530: Height = 9390
    Left = 0: Top = 0
    Call Ctr_Ayudabanco.conexion(VGCNx)
    Call Ctr_AyudaBancoCuenta.conexion(VGCNx)
    DTPfechaini.Value = Date
    TDBG_concil.FetchRowStyle = True
End Sub

Private Sub cmdaceptar_Click()
Dim X As Integer
    RsConcil.Update
    Cmdimprimir(0).Enabled = True
    Cmdimprimir(1).Enabled = True
    Cmdimprimir(2).Enabled = True
    cmdaceptar.Enabled = False
    X = 0
    VGCNx.Execute "Update te_controlasaldos " & _
               " set ctrlsaldo_saldobanco =" & mtsoles & "," & _
               " ctrlsaldo_ingresobanco=" & tmontosolesDebe - montosolesDebe & "," & _
               " ctrlsaldo_egresobanco=" & tmontosolesHaber - montosolesHaber & _
               " Where ctrlsaldo_bancocaja='" & Trim(Ctr_Ayudabanco.xclave) & "' and " & _
               " ctrlsaldo_numectacte='" & Trim(Ctr_AyudaBancoCuenta.xnombre) & "' and " & _
               " ctrlsaldo_año='" & Format(DTPfechaini.Year, "0") & "' and " & _
               " ctrlsaldo_mes='" & Format(DTPfechaini.Month, "00") & "' and " & _
               " ctrlsaldo_mon='01'", X
    If X = 0 Then
        VGCNx.Execute "Insert Into " & _
                   "te_controlasaldos(ctrlsaldo_saldobanco,ctrlsaldo_ingresobanco," & _
                   "ctrlsaldo_egresobanco,ctrlsaldo_bancocaja,ctrlsaldo_numectacte," & _
                   "ctrlsaldo_año,ctrlsaldo_mes,ctrlsaldo_mon,ctrlsaldo_tipobc) Values (" & _
                   mtsoles & "," & tmontosolesDebe - montosolesDebe & "," & _
                   tmontosolesHaber - montosolesHaber & ",'" & Trim(Ctr_Ayudabanco.xclave) & "','" & _
                   Trim(Ctr_AyudaBancoCuenta.xnombre) & "','" & Format(DTPfechaini.Year, "0") & "','" & _
                   Format(DTPfechaini.Month, "00") & "','01','B')"
    End If
    X = 0
    VGCNx.Execute "Update te_controlasaldos " & _
               " set ctrlsaldo_saldobanco =" & mtdolar & "," & _
               " ctrlsaldo_ingresobanco=" & tmontodolaresDebe - montodolaresDebe & "," & _
               " ctrlsaldo_egresobanco=" & tmontodolaresHaber - montodolaresHaber & _
               " Where ctrlsaldo_bancocaja='" & Trim(Ctr_Ayudabanco.xclave) & "' and " & _
               " ctrlsaldo_numectacte='" & Trim(Ctr_AyudaBancoCuenta.xnombre) & "' and " & _
               " ctrlsaldo_año='" & Format(DTPfechaini.Year, "0") & "' and " & _
               " ctrlsaldo_mes='" & Format(DTPfechaini.Month, "00") & "' and " & _
               " ctrlsaldo_mon='02'", X
               
    If X = 0 Then
        VGCNx.Execute "Insert Into " & _
                   "te_controlasaldos(ctrlsaldo_saldobanco,ctrlsaldo_ingresobanco," & _
                   "ctrlsaldo_egresobanco,ctrlsaldo_bancocaja,ctrlsaldo_numectacte," & _
                   "ctrlsaldo_año,ctrlsaldo_mes,ctrlsaldo_mon,ctrlsaldo_tipobc) Values (" & _
                   mtdolar & "," & tmontodolaresDebe - montodolaresDebe & "," & _
                   tmontodolaresHaber - montodolaresHaber & ",'" & Trim(Ctr_Ayudabanco.xclave) & "','" & _
                   Trim(Ctr_AyudaBancoCuenta.xnombre) & "','" & Format(DTPfechaini.Year, "0") & "','" & _
                   Format(DTPfechaini.Month, "00") & "','02','B')"
    End If
    
End Sub

Private Sub CmdCancelar_Click()
    If RsConcil Is Nothing Then
        Unload Me
        Exit Sub
    End If
    
    RsConcil.CancelBatch
    Unload Me
End Sub
Private Sub Listar()
Dim vardllgen As New dllgeneral.dll_general
Dim sqlcad As String
Dim fecha1 As String
    fecha1 = Format(DateSerial(DTPfechaini.Year, DTPfechaini.Month, 1), "dd/mm/yyyy")
    sqlcad = "select " & _
             " A.chkconcil, " & _
             " Anno=year(A.detrec_fechacancela),Mes=month(A.detrec_fechacancela), " & _
             " A.cabrec_numrecibo,A.detrec_emisioncheque,A.detrec_tipodoc_concepto, " & _
             " A.detrec_numdocumento,B.cabrec_ingsal,A.detrec_tipocajabanco, " & _
             " A.detrec_numctacte,A.detrec_monedadocumento, " & _
             " A.detrec_importesoles,A.detrec_importedolares,A.detrec_monedacancela, " & _
             " A.detrec_tdqc,A.detrec_ndqc,A.detrec_fechacancela,B.cabrec_estadoreg, " & _
             " B.cabrec_fechadocumento,A.detrec_observacion,A.fechconcil " & _
             " from te_detallerecibos A " & _
             " Inner join te_cabecerarecibos  B " & _
             " on A.cabrec_numrecibo=B.cabrec_numrecibo " & _
             " Where " & _
             " A.detrec_emisioncheque='B' and " & _
             " A.detrec_tipocajabanco='B' and " & _
             " ltrim(rtrim(Isnull(A.detrec_numctacte,'')))  <>'' and " & _
             " B.cabrec_estadoreg <> 1 and ltrim(rtrim(Isnull(A.detrec_numctacte,'')))='" & Trim(Ctr_AyudaBancoCuenta.xnombre) & "' and  " & _
             " A.detrec_fechacancela < '" & Format(DateSerial(Year(DateAdd("m", 1, DTPfechaini)), Month(DateAdd("m", 1, DTPfechaini)), 1), "dd/mm/yyyy") & "' and " & _
             " ( fechconcil is null or fechconcil >='" & fecha1 & "')"
             
    If chkconciliado.Value = 1 Then
        sqlcad = sqlcad & "  and isnull(chkconcil,0)<>0 "
    End If
    sqlcad = sqlcad & " order by year(A.detrec_fechacancela) desc,month(A.detrec_fechacancela) desc, A.detrec_fechacancela "

    Set RsConcil = New ADODB.Recordset
    RsConcil.Open sqlcad, VGCNx, adOpenDynamic, adLockBatchOptimistic
    
    If RsConcil.RecordCount = 0 Then
       ' lbfechini.Enabled = False
       ' DTPfechaini.Enabled = False
      Else
        lbfechini.Enabled = True
        DTPfechaini.Enabled = True
    End If
    lbnreg.Caption = Format(RsConcil.RecordCount, "0 ")
    Set TDBG_concil.DataSource = RsConcil
    If CLng(lbnreg.Caption) > 0 Then
        Cmdimprimir(0).Enabled = True
        Cmdimprimir(1).Enabled = True
        Cmdimprimir(2).Enabled = True
    Else
        Cmdimprimir(0).Enabled = False
        Cmdimprimir(1).Enabled = False
        Cmdimprimir(2).Enabled = False
    End If
    Call CalcularTotal(RsConcil)
    Call CalcularTotales(RsConcil)
    Set RsSaldoIni = New ADODB.Recordset
    RsSaldoIni.Open "Select * from te_controlasaldos where  ctrlsaldo_bancocaja='" & Trim(Ctr_Ayudabanco.xclave) & "' and " & _
                    " rtrim(ltrim(isnull(ctrlsaldo_numectacte,'')))='" & Trim(Ctr_AyudaBancoCuenta.xnombre) & "'" & _
                    " and ctrlsaldo_año='" & Format(DTPfechaini.Year, "0 ") & "' and ctrlsaldo_mes='" & Format(DTPfechaini.Month, "00") & "' and ctrlsaldo_mon='" & mon & "'", VGCNx, adOpenKeyset, adLockReadOnly
    
    If RsSaldoIni.RecordCount > 0 Then
        Txtsaldoini.valor = vardllgen.ESNULO(RsSaldoIni!ctrlsaldo_saldocontaingre, 0) - vardllgen.ESNULO(RsSaldoIni!ctrlsaldo_saldocontasalida, 0) + vardllgen.ESNULO(RsSaldoIni!ctrlsaldo_saldocontainicial, 0)
        'ctrlsaldo_saldobanco +
        '(ctrlsaldo_ingresobanco - ctrlsaldo_egresobanco)
        TxSaldExtBanc.valor = vardllgen.ESNULO(RsSaldoIni!ctrlsaldo_saldobanco, 0) - (vardllgen.ESNULO(RsSaldoIni!ctrlsaldo_ingresobanco, 0) - vardllgen.ESNULO(RsSaldoIni!ctrlsaldo_egresobanco, 0))
      Else
        Txtsaldoini.valor = "0"
        TxSaldExtBanc.valor = 0
    End If
    TxSaldExtBanc.Text = TxSaldExtBanc.valor
    Txtsaldoini.Text = Txtsaldoini.valor
End Sub

Private Sub CalcularTotales(ByVal rs As Recordset)
Dim rsaux As ADODB.Recordset
Dim vardllgen As New dllgeneral.dll_general
Set rsaux = rs.Clone(adLockReadOnly)


montosolesDebe = 0: montodolaresDebe = 0:
montosolesHaber = 0: montodolaresHaber = 0:
mtsoles = 0: mtdolar = 0

If rsaux.BOF = True Or rsaux.EOF = True Then Exit Sub
Dim Fecha As Double
Fecha = DateSerial(Year(DateAdd("m", 1, DTPfechaini)), Month(DateAdd("m", 1, DTPfechaini)), 1)

rsaux.MoveFirst
    While Not rsaux.EOF
    If rsaux("chkconcil").Value <> 0 And Not (rsaux!fechconcil >= Fecha) Then
        montosolesDebe = montosolesDebe + IIf(rsaux!cabrec_ingsal = "I", vardllgen.ESNULO(rsaux!detrec_importesoles, 0), 0)
        montodolaresDebe = montodolaresDebe + IIf(rsaux!cabrec_ingsal = "I", vardllgen.ESNULO(rsaux!detrec_importedolares, 0), 0)
        montosolesHaber = montosolesHaber + IIf(rsaux!cabrec_ingsal = "E", vardllgen.ESNULO(rsaux!detrec_importesoles, 0), 0)
        montodolaresHaber = montodolaresHaber + IIf(rsaux!cabrec_ingsal = "E", vardllgen.ESNULO(rsaux!detrec_importedolares, 0), 0)
    End If
    rsaux.MoveNext
    Wend
    'Soles
    mtsoles = ((tmontosolesDebe - montosolesDebe) - (tmontosolesHaber - montosolesHaber)) + montoextbanc
    LbTotales(0).Caption = Format(tmontosolesDebe - montosolesDebe, "###,###,###,###.00 ") ' Debe
    LbTotales(1).Caption = Format(tmontosolesHaber - montosolesHaber, "###,###,###,###.00 ") ' Haber
    LbTotales(2).Caption = Format(mtsoles, "###,###,###,###.00 ")   ' Haber
    'Dolares
    mtdolar = ((tmontodolaresDebe - montodolaresDebe) - (tmontodolaresHaber - montodolaresHaber)) + montoextbanc
    LbTotales(3).Caption = Format(tmontodolaresDebe - montodolaresDebe, "###,###,###,###.00 ") ' Debe
    LbTotales(4).Caption = Format(tmontodolaresHaber - montodolaresHaber, "###,###,###,###.00 ") ' Haber
    LbTotales(5).Caption = Format(mtdolar, "###,###,###,###.00 ") ' Haber
    
    If mon = "01" Then
        lbtot(0).Caption = Format(montosolesDebe, "###,###,###,###.00")
        lbtot(1).Caption = Format(montosolesHaber, "###,###,###,###.00")
        lbtot(2).Caption = Format(montosolesDebe - montosolesHaber, "###,###,###,###.00")
      Else
        lbtot(0).Caption = Format(montodolaresDebe, "###,###,###,###.00")
        lbtot(1).Caption = Format(montodolaresHaber, "###,###,###,###.00")
        lbtot(2).Caption = Format(montodolaresDebe - montodolaresHaber, "###,###,###,###.00")
    End If
        
End Sub
Private Sub CalcularTotal(ByVal rs As Recordset)
Dim rsaux As ADODB.Recordset
Dim vardllgen As New dllgeneral.dll_general
Set rsaux = rs.Clone(adLockReadOnly)

    tmontosolesDebe = 0: tmontodolaresDebe = 0:
    tmontosolesHaber = 0: tmontodolaresHaber = 0:
    tsoles = 0: tdolar = 0
    If rsaux.BOF = True Or rsaux.EOF = True Then Exit Sub
    rsaux.MoveFirst
    montoextbanc = CDbl(vardllgen.ESNULO(Espunto(TxSaldExtBanc.valor), 0))
    While Not rsaux.EOF
        tmontosolesDebe = tmontosolesDebe + IIf(rsaux!cabrec_ingsal = "I", vardllgen.ESNULO(rsaux!detrec_importesoles, 0), 0)
        tmontodolaresDebe = tmontodolaresDebe + IIf(rsaux!cabrec_ingsal = "I", vardllgen.ESNULO(rsaux!detrec_importedolares, 0), 0)
        tmontosolesHaber = tmontosolesHaber + IIf(rsaux!cabrec_ingsal = "E", vardllgen.ESNULO(rsaux!detrec_importesoles, 0), 0)
        tmontodolaresHaber = tmontodolaresHaber + IIf(rsaux!cabrec_ingsal = "E", vardllgen.ESNULO(rsaux!detrec_importedolares, 0), 0)
        rsaux.MoveNext
    Wend
    'Soles
    tsoles = tmontosolesDebe - tmontosolesHaber + montoextbanc
    LbTotales(0).Caption = Format(tmontosolesDebe, "###,###,###,###.00 ") ' Debe
    LbTotales(1).Caption = Format(tmontosolesHaber, "###,###,###,###.00 ") ' Haber
    LbTotales(2).Caption = Format(tsoles, "###,###,###,###.00 ")     ' Total
    'Dolares
    tdolar = tmontodolaresDebe - tmontodolaresHaber + montoextbanc
    LbTotales(3).Caption = Format(tmontodolaresDebe, "###,###,###,###.00 ") ' Debe
    LbTotales(4).Caption = Format(tmontodolaresHaber, "###,###,###,###.00 ") ' Haber
    LbTotales(5).Caption = Format(tdolar, "###,###,###,###.00 ") ' Haber
End Sub

Private Sub RsConcil_FieldChangeComplete(ByVal cFields As Long, ByVal Fields As Variant, ByVal pError As ADODB.error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Static Cont As Integer
   ' If flagcal Then Exit Sub
    cmdaceptar.Enabled = True
    Cmdimprimir(0).Enabled = False
    Cmdimprimir(1).Enabled = False
    Cmdimprimir(2).Enabled = False
    If Cont = 1 Then
        Cont = 0
        Exit Sub
    End If
    Call CalcularTotales(RsConcil)
    Cont = 1
    If pRecordset.Fields("chkconcil").Value Then
        RsConcil!fechconcil = DateSerial(DTPfechaini.Year, DTPfechaini.Month, 1)
     Else
        RsConcil!fechconcil = Null
    End If
End Sub

Private Sub cmdimprimir_Click(Index As Integer)
    If RsConcil.RecordCount = 0 Then Exit Sub
Dim valor As String
    Select Case Index
        Case 0: valor = "1"
        Case 1: valor = "2"
        Case 2: valor = "0"
    End Select
    Call Imprimir(valor)
End Sub
Private Sub Imprimir(ValorConci As String)
Dim vardllgen As New dllgeneral.dll_general
Dim arrform(7) As Variant, arrparm(5) As Variant
Dim NombreRep As String, CadOrden As String
Dim Fecha As Double
Dim fecha1 As String
    fecha1 = Format(DateSerial(DTPfechaini.Year, DTPfechaini.Month, 1), "dd/mm/yyyy")
    '@Base,@cuenta,@concil,@Fecharef,@anno,@mes
    arrparm(0) = VGParamSistem.BDEmpresa
    arrparm(1) = Trim(Ctr_AyudaBancoCuenta.xnombre)
    arrparm(2) = ValorConci
    arrparm(3) = Format(DateSerial(Year(DateAdd("m", 1, DTPfechaini)), Month(DateAdd("m", 1, DTPfechaini)), 1), "dd/mm/yyyy")
    arrparm(4) = fecha1
    
    
    Select Case ValorConci
        Case "0": arrform(0) = "Todos"
        Case "1": arrform(0) = "Conciliados"
        Case "2": arrform(0) = "Pendientes"
    End Select
        
    Fecha = DateSerial(Year(DateAdd("m", 1, DTPfechaini)), Month(DateAdd("m", 1, DTPfechaini)), 1)
    
    arrform(1) = "Banco='" & Ctr_Ayudabanco.xnombre & "'"
    arrform(2) = "Cuenta='" & Ctr_AyudaBancoCuenta.xnombre & "'"
    arrform(3) = "mon='" & mon & "'"
    arrform(4) = "Fecha='" & Format(DateAdd("d", -1, Fecha), "dd/mm/yyyy") & "'"
    arrform(5) = "SExtBanc=" & vardllgen.ESNULO(Espunto(TxSaldExtBanc.valor), 0)
    NombreRep = "te_concilbanc.rpt"
    Call ImpresionRptProc(NombreRep, arrform, arrparm, , "Conciliación Bancaria")
End Sub
Private Sub TDBG_concil_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
Dim rsclone As New ADODB.Recordset
    On Error Resume Next
    Set rsclone = RsConcil.Clone(adLockReadOnly)
    If rsclone.RecordCount = 0 Then Exit Sub
    rsclone.Bookmark = Bookmark
    If rsclone!anno = Year(DTPfechaini) And rsclone!mes = Month(DTPfechaini) Then
       RowStyle.BackColor = RGB(254, 251, 218)
       '185,251,210
    End If
    If rsclone!fechconcil > DateSerial(DTPfechaini.Year, DTPfechaini.Month, 1) Then
       RowStyle.BackColor = RGB(200, 250, 100)
    End If
    
    flagcal = True
End Sub

Private Sub TDBG_concil_HeadClick(ByVal ColIndex As Integer)
 With RsConcil
    If .Sort = Empty Then
        .Sort = TDBG_concil.Columns.Item(ColIndex).DataField & " asc"
    ElseIf Right(.Sort, 3) = "asc" Then
        .Sort = TDBG_concil.Columns.Item(ColIndex).DataField & " desc"
    ElseIf Right(.Sort, 4) = "desc" Then
        .Sort = TDBG_concil.Columns.Item(ColIndex).DataField & " asc"
    End If
    TDBG_concil.Refresh
 End With
End Sub

Private Sub TxSaldExtBanc_Change()
  If RsConcil Is Nothing Then Exit Sub
    Call CalcularTotal(RsConcil)
    Call CalcularTotales(RsConcil)
End Sub
