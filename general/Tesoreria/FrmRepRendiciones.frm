VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Begin VB.Form FrmRepRendiciones 
   Caption         =   "Reportes de Rendiciones"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13380
   LinkTopic       =   "Form1"
   ScaleHeight     =   8610
   ScaleWidth      =   13380
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Opciones"
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   135
      TabIndex        =   37
      Top             =   7920
      Width           =   8055
      Begin VB.CommandButton Cmdcancelar 
         Caption         =   "Cancelar"
         Height          =   330
         Left            =   1080
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
         Caption         =   "Imprimir"
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
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.CommandButton Cmdimprimir 
         Caption         =   "Imprimir Todo"
         Height          =   330
         Index           =   2
         Left            =   6570
         TabIndex        =   38
         Top             =   225
         Visible         =   0   'False
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1875
      Left            =   135
      TabIndex        =   8
      Top             =   0
      Width           =   13215
      Begin TextFer.TxFer Txtsaldoini 
         Height          =   315
         Left            =   6165
         TabIndex        =   9
         Top             =   495
         Width           =   1665
         _ExtentX        =   2937
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
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaOficina 
         Height          =   300
         Left            =   1575
         TabIndex        =   1
         Top             =   150
         Width           =   4260
         _ExtentX        =   7514
         _ExtentY        =   529
         XcodMaxLongitud =   3
         xcodwith        =   400
         NomTabla        =   "cp_oficina"
         TituloAyuda     =   "Ayuda de Caja"
         ListaCampos     =   "vendedorcodigo(1),vendedornombres(1)"
         XcodCampo       =   "vendedorcodigo"
         XListCampo      =   "vendedornombres"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "vendedorcodigo,vendedornombres"
      End
      Begin TextFer.TxFer TxSaldofin 
         Height          =   300
         Left            =   6150
         TabIndex        =   10
         Top             =   1230
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   529
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
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuRendicion 
         Height          =   315
         Left            =   1530
         TabIndex        =   4
         Top             =   1410
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         XcodMaxLongitud =   0
         xcodwith        =   1000
         NomTabla        =   "te_rendiciones"
         TituloAyuda     =   "Busqueda de Rendiciones"
         ListaCampos     =   "rendicionnumero(1),monedacodigo(1),rendicionfecha(2)"
         XcodCampo       =   "rendicionnumero"
         XListCampo      =   "monedacodigo"
         ListaCamposDescrip=   "Nro Rendicion,Moneda, fecha rendicion"
         ListaCamposText =   "rendicionnumero,monedacodigo,rendicionfecha"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaCaja 
         Height          =   315
         Left            =   1575
         TabIndex        =   2
         Top             =   585
         Width           =   4260
         _ExtentX        =   7514
         _ExtentY        =   556
         XcodMaxLongitud =   11
         xcodwith        =   400
         NomTabla        =   "te_codigocaja"
         TituloAyuda     =   "Busqueda de Caja"
         ListaCampos     =   "cajacodigo(1),cajadescripcion(1)"
         XcodCampo       =   "cajacodigo"
         XListCampo      =   "cajadescripcion"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "cajacodigo,cajadescripcion"
      End
      Begin MSComCtl2.DTPicker DTPfechaini 
         Height          =   285
         Left            =   4380
         TabIndex        =   34
         Top             =   1440
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   " MMM - yyyy"
         DateIsNull      =   -1  'True
         Format          =   20971521
         CurrentDate     =   37513
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaMoneda 
         Height          =   315
         Left            =   1545
         TabIndex        =   3
         Top             =   990
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         XcodMaxLongitud =   2
         xcodwith        =   300
         NomTabla        =   "gr_moneda"
         TituloAyuda     =   "Busqueda de Moneda"
         ListaCampos     =   "monedacodigo(1),monedadescripcion(1)"
         XcodCampo       =   "monedacodigo"
         XListCampo      =   "monedadescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "monedacodigo,monedadescripcion"
      End
      Begin VB.Label Label5 
         Caption         =   "Moneda : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   135
         TabIndex        =   36
         Top             =   1035
         Width           =   855
      End
      Begin VB.Label lbfechini 
         Caption         =   "Fecha Rendicion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3135
         TabIndex        =   35
         Top             =   1380
         Width           =   990
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         Height          =   1620
         Left            =   8010
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   5085
      End
      Begin VB.Label leDebe 
         AutoSize        =   -1  'True
         Caption         =   "INGRESOS"
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
         TabIndex        =   26
         Top             =   630
         Width           =   975
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         Height          =   1590
         Left            =   8025
         Shape           =   4  'Rounded Rectangle
         Top             =   255
         Width           =   5055
      End
      Begin VB.Label leHaber 
         AutoSize        =   -1  'True
         Caption         =   "EGREOS"
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
         TabIndex        =   25
         Top             =   1050
         Width           =   780
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
         Left            =   9390
         TabIndex        =   24
         Top             =   360
         Width           =   1440
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
         TabIndex        =   23
         Top             =   315
         Width           =   1800
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
         Left            =   9285
         TabIndex        =   22
         Top             =   585
         Width           =   1515
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
         Left            =   9285
         TabIndex        =   21
         Top             =   990
         Width           =   1515
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
         Left            =   11385
         TabIndex        =   20
         Top             =   585
         Width           =   1515
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
         Left            =   11430
         TabIndex        =   19
         Top             =   990
         Width           =   1515
      End
      Begin VB.Label Label4 
         Caption         =   "Oficina"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   18
         Top             =   315
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "Saldo Inicial"
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
         Left            =   6165
         TabIndex        =   17
         Top             =   225
         Width           =   1440
      End
      Begin VB.Label lbMon 
         Caption         =   "Rendicion :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   105
         TabIndex        =   16
         Top             =   1455
         Width           =   960
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
         TabIndex        =   15
         Top             =   1500
         Width           =   615
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
         Left            =   9270
         TabIndex        =   14
         Top             =   1440
         Width           =   1515
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
         Left            =   11430
         TabIndex        =   13
         Top             =   1440
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "Saldo Final"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   6195
         TabIndex        =   12
         Top             =   900
         Width           =   1440
      End
      Begin VB.Label Label4 
         Caption         =   "Cod. Caja"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   690
         Width           =   885
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   5925
      Left            =   135
      ScaleHeight     =   5865
      ScaleWidth      =   13140
      TabIndex        =   0
      Top             =   1935
      Width           =   13200
      Begin TrueOleDBGrid70.TDBGrid TDBG_concil 
         Height          =   5400
         Left            =   60
         TabIndex        =   5
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
         Columns(1).Caption=   "T/D"
         Columns(1).DataField=   "detrec_tipodoc_concepto"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Nº Doc"
         Columns(2).DataField=   "detrec_numdocumento"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Fecha"
         Columns(3).DataField=   "detrec_fechacancela"
         Columns(3).NumberFormat=   "Short Date"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "T/Dc"
         Columns(4).DataField=   "detrec_tdqc"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Nº Doc Cancela"
         Columns(5).DataField=   "detrec_ndqc"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "I/E"
         Columns(6).DataField=   "cabrec_ingsal"
         Columns(6).NumberFormat=   "###,###,###.00"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Importe Soles"
         Columns(7).DataField=   "detrec_importesoles"
         Columns(7).NumberFormat=   "###,###,###.00"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   4
         Columns(8)._MaxComboItems=   5
         Columns(8).ValueItems(0)._DefaultItem=   0
         Columns(8).ValueItems(0).Value=   "1"
         Columns(8).ValueItems(0).Value.vt=   8
         Columns(8).ValueItems(0).DisplayValue=   "1"
         Columns(8).ValueItems(0).DisplayValue.vt=   8
         Columns(8).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(8).ValueItems(1)._DefaultItem=   0
         Columns(8).ValueItems(1).Value=   "0"
         Columns(8).ValueItems(1).Value.vt=   8
         Columns(8).ValueItems(1).DisplayValue=   "0"
         Columns(8).ValueItems(1).DisplayValue.vt=   8
         Columns(8).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
         Columns(8).ValueItems.Count=   2
         Columns(8).Caption=   "CH"
         Columns(8).DataField=   "chkconcil"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Importe Dolares"
         Columns(9).DataField=   "detrec_importedolares"
         Columns(9).NumberFormat=   "###,###,###.00"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "Observaciones"
         Columns(10).DataField=   "detrec_observacion"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "Fecha Concil"
         Columns(11).DataField=   "fechconcil"
         Columns(11).NumberFormat=   "dd/mm/yyyy"
         Columns(11).EditMask=   "ss/mm/aaaa"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "Nro.Rendicion"
         Columns(12).DataField=   "rendicionnumero"
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   13
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=13"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1667"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1588"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=635"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=556"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8196"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2170"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2090"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=8196"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=1640"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1561"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=8196"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=847"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=767"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=8196"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=2672"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2593"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=8196"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=582"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=503"
         Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=8194"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(36)=   "Column(7).Width=2196"
         Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=2117"
         Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=8194"
         Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(41)=   "Column(8).Width=714"
         Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=635"
         Splits(0)._ColumnProps(44)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(45)=   "Column(9).Width=2196"
         Splits(0)._ColumnProps(46)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(47)=   "Column(9)._WidthInPix=2117"
         Splits(0)._ColumnProps(48)=   "Column(9)._ColStyle=8194"
         Splits(0)._ColumnProps(49)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(50)=   "Column(10).Width=3360"
         Splits(0)._ColumnProps(51)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(52)=   "Column(10)._WidthInPix=3281"
         Splits(0)._ColumnProps(53)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(54)=   "Column(11).Width=2725"
         Splits(0)._ColumnProps(55)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(56)=   "Column(11)._WidthInPix=2646"
         Splits(0)._ColumnProps(57)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(58)=   "Column(12).Width=2725"
         Splits(0)._ColumnProps(59)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(60)=   "Column(12)._WidthInPix=2646"
         Splits(0)._ColumnProps(61)=   "Column(12).Order=13"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=System"
         PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=System"
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
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=54,.parent=13,.locked=-1"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=58,.parent=13,.locked=-1"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=62,.parent=13,.locked=-1"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=66,.parent=13,.locked=-1"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=70,.parent=13,.locked=-1"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=74,.parent=13,.alignment=1,.locked=-1"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=71,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=72,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=73,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=78,.parent=13,.alignment=1,.locked=-1"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=75,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=76,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=77,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=50,.parent=13,.bgcolor=&HBFFFAA&"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=47,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=48,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=49,.parent=17"
         _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=82,.parent=13,.alignment=1,.locked=-1"
         _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=79,.parent=14"
         _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=80,.parent=15"
         _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=81,.parent=17"
         _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=86,.parent=13"
         _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=83,.parent=14"
         _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=84,.parent=15"
         _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=85,.parent=17"
         _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=90,.parent=13"
         _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=87,.parent=14"
         _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=88,.parent=15"
         _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=89,.parent=17"
         _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=32,.parent=13"
         _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=29,.parent=14"
         _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=30,.parent=15"
         _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=31,.parent=17"
         _StyleDefs(88)  =   "Named:id=33:Normal"
         _StyleDefs(89)  =   ":id=33,.parent=0"
         _StyleDefs(90)  =   "Named:id=34:Heading"
         _StyleDefs(91)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(92)  =   ":id=34,.wraptext=-1"
         _StyleDefs(93)  =   "Named:id=35:Footing"
         _StyleDefs(94)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(95)  =   "Named:id=36:Selected"
         _StyleDefs(96)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(97)  =   "Named:id=37:Caption"
         _StyleDefs(98)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(99)  =   "Named:id=38:HighlightRow"
         _StyleDefs(100) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(101) =   "Named:id=39:EvenRow"
         _StyleDefs(102) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(103) =   "Named:id=40:OddRow"
         _StyleDefs(104) =   ":id=40,.parent=33"
         _StyleDefs(105) =   "Named:id=41:RecordSelector"
         _StyleDefs(106) =   ":id=41,.parent=34"
         _StyleDefs(107) =   "Named:id=42:FilterBar"
         _StyleDefs(108) =   ":id=42,.parent=33"
      End
      Begin VB.Label lbnreg 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "0 "
         Height          =   255
         Left            =   11955
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   5550
         Width           =   975
      End
   End
   Begin VB.Label lbtot 
      Appearance      =   0  'Flat
      BackColor       =   &H00DDF7F9&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   8310
      TabIndex        =   33
      Top             =   8145
      Width           =   1605
   End
   Begin VB.Label Label8 
      Caption         =   "Label6"
      Height          =   270
      Left            =   0
      TabIndex        =   32
      Top             =   90
      Width           =   2040
   End
   Begin VB.Label lbtot 
      Appearance      =   0  'Flat
      BackColor       =   &H00DDF7F9&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   1
      Left            =   9930
      TabIndex        =   31
      Top             =   8145
      Width           =   1605
   End
   Begin VB.Label lbtot 
      Appearance      =   0  'Flat
      BackColor       =   &H00DDF7F9&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   2
      Left            =   11580
      TabIndex        =   30
      Top             =   8130
      Width           =   1605
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "INGR.  CONCIL"
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
      TabIndex        =   29
      Top             =   7890
      Width           =   1335
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "EGR. CONCIL"
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
      TabIndex        =   28
      Top             =   7905
      Width           =   1200
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
      TabIndex        =   27
      Top             =   7890
      Width           =   615
   End
End
Attribute VB_Name = "FrmRepRendiciones"
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
Dim SQL As String
Dim Fecha As Date
Dim tsoles As Double, tdolar As Double
Dim montoextbanc As Double
Dim mon As String
Dim mon_descripcion As String
Dim flagcal As Boolean

Private Sub Ctr_AyudaCaja_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
 Ctr_AyuRendicion.Filtro = " codigocaja='" & Ctr_AyudaCaja.xclave & "'"
End Sub



Private Sub Ctr_AyudaMoneda_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
 Ctr_AyuRendicion.Filtro = " codigocaja='" & Ctr_AyudaCaja.xclave & "' and monedacodigo='" & Ctr_AyudaMoneda.xclave & "'"


End Sub

Private Sub Ctr_AyuRendicion_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Fecha = ColecCampos("rendicionfecha")
    Call Listar
    CmdAceptar.Enabled = True
    Cmdimprimir(0).Enabled = True
    Cmdimprimir(1).Enabled = True
    Cmdimprimir(2).Enabled = True

End Sub

Private Sub Form_Load()
    'lbfechini.Enabled = False
    'DTPfechaini.Enabled = False
    
    Width = 13530: Height = 9390
    Left = 0: Top = 0
    Call Ctr_AyudaOficina.Conexion(VGCNx)
    Call Ctr_AyudaCaja.Conexion(VGCNx): Ctr_AyudaCaja.Filtro = " cajarendiciones=1 "
    Call Ctr_AyuRendicion.Conexion(VGCNx)
    Call Ctr_AyudaMoneda.Conexion(VGCNx)
'    Set RsConcil = VGcnx.Execute("select * from te_codigocaja where cajarendiciones= 1 ")
    TDBG_concil.FetchRowStyle = True
 '   Ctr_AyudaOficina.SetFocus
End Sub

Private Sub cmdaceptar_Click()
Dim X As Integer
Dim rsql As New ADODB.Recordset
    
Cmdimprimir(0).Enabled = True
Cmdimprimir(1).Enabled = True
Cmdimprimir(2).Enabled = True
CmdAceptar.Enabled = False

SQL = " select * from te_codigocaja where cajacodigo='" & Ctr_AyudaCaja.xclave & "'"
Set rsql = VGCNx.Execute(SQL)
If Ctr_AyudaMoneda.xclave = "01" Then
   TxtNrorendicion.Text = Format(rsql!rendicionnumero01, "000000")
   SQL = " update te_codigocaja set rendicionnumero01='" & Format(TxtNrorendicion.Text + 1, "000000") & "'"
   SQL = SQL & " where cajacodigo='" & Ctr_AyudaCaja.xclave & "'"
 Else
   TxtNrorendicion.Text = Format(rsql!rendicionnumero02, "000000")
   SQL = " update te_codigocaja set rendicionnumero02='" & Format(TxtNrorendicion.Text + 1, "000000") & "'"
   SQL = SQL & " where cajacodigo='" & Ctr_AyudaCaja.xclave & "'"
End If
Set rsql = VGCNx.Execute(SQL)
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
Dim SQL As String
SQL = "select  A.chkconcil,a.rendicionnumero,A.cabrec_numrecibo, "
SQL = SQL & "a.detrec_item,A.detrec_emisioncheque,A.detrec_tipodoc_concepto, "
SQL = SQL & "A.detrec_numdocumento,B.cabrec_ingsal,A.detrec_tipocajabanco,A.detrec_numctacte,"
SQL = SQL & "A.detrec_monedadocumento,A.detrec_importesoles,A.detrec_importedolares,"
SQL = SQL & "A.detrec_monedacancela,A.detrec_tdqc,A.detrec_ndqc,A.detrec_fechacancela,"
SQL = SQL & "B.cabrec_estadoreg,B.cabrec_fechadocumento,A.detrec_observacion,A.fechconcil "
SQL = SQL & " from te_detallerecibos A Inner join te_cabecerarecibos  B on A.cabrec_numrecibo=B.cabrec_numrecibo "
SQL = SQL & " Where a.detrec_cajabanco1='" & Ctr_AyudaCaja.xclave & "' and a.detrec_monedacancela='"
SQL = SQL & Ctr_AyudaMoneda.xclave & "' and A.rendicionnumero='" & Ctr_AyuRendicion.xclave & "'"
SQL = SQL & " and B.cabrec_estadoreg <> 1  and isnull(a.detalle_no_saldos,0)<>1 "
            
SQL = SQL & " order by A.cabrec_numrecibo "

    Set RsConcil = New ADODB.Recordset
    RsConcil.Open (SQL), VGCNx, adOpenDynamic, adLockBatchOptimistic
    
    lbfechini.Enabled = True
    DTPfechaini.Enabled = True
    
    lbnreg.Caption = Format(RsConcil.RecordCount, "0 ")

        
    SQL = " Select rendicionsaldoinicial,rendicioningresos,rendicionegresos,rendicionsaldofinal,"
    SQL = SQL & " rendicionfecha from te_rendiciones "
    SQL = SQL & "where oficinacodigo='" & Ctr_AyudaOficina.xclave & "' and codigocaja='"
    SQL = SQL & Ctr_AyudaCaja.xclave & "' and monedacodigo='" & Ctr_AyudaMoneda.xclave & "'"
    SQL = SQL & " and rendicionnumero='" & Ctr_AyuRendicion.xclave & "'"
    Set rsql = VGCNx.Execute(SQL)
    
    If rsql.RecordCount() > 0 Then
       Txtsaldoini.Text = Round(rsql!rendicionsaldoinicial, 2)
       TxSaldofin.Text = Round(rsql!rendicionsaldoinicial + rsql!rendicioningresos - rsql!rendicionegresos, 2)
       DTPfechaini.Value = rsql!rendicionfecha
    End If
    
    
    Set TDBG_concil.DataSource = RsConcil

    Call CalcularTotal(RsConcil)
    Call CalcularTotales(RsConcil)
    
   
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


rsaux.MoveFirst
    While Not rsaux.EOF
      If rsaux("chkconcil").Value <> 0 Then
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
'    TxSaldofin.Text = Round(CDbl(vardllgen.ESNULO(Espunto(Txtsaldoini.Text), 0)) + CDbl(lbtot(0).Caption) - CDbl(lbtot(1).Caption), 2)
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
    montoextbanc = CDbl(vardllgen.ESNULO(Espunto(TxSaldofin.valor), 0))
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

Private Sub cmdimprimir_Click(Index As Integer)
    If RsConcil.RecordCount = 0 Then Exit Sub
Dim valor As String
    Select Case Index
        Case 0: valor = "0"
        Case 1: valor = "1"
        Case 2: valor = "2"
    End Select
 '   If valor = "0" Then
 '       Call Imprimirgastos
 '    Else
       Call Imprimir(valor)
 '   End If
End Sub
Private Sub Imprimir(ValorConci As String)
Dim vardllgen As New dllgeneral.dll_general
Dim arrform(7) As Variant, arrparm(6) As Variant
Dim NombreRep As String, CadOrden As String
Dim Fecha As String
Dim fecha1 As String
    fecha1 = Format(DateSerial(DTPfechaini.Year, DTPfechaini.Month, 1), "dd/mm/yyyy")
    '@Base,@cuenta,@concil,@Fecharef,@anno,@mes
    arrparm(0) = VGParamSistem.BDEmpresa
    arrparm(1) = Trim(Ctr_AyudaCaja.xclave)
    arrparm(2) = Ctr_AyuRendicion.xnombre
'    Trim (Ctr_AyuRendicion.xclave)
    arrparm(3) = Format(DTPfechaini.Value, "dd/mm/yyyy")
    
    Select Case ValorConci
        Case "0": arrform(0) = "Tipo='" & "Todos'"
        Case "1": arrform(0) = "Tipo='" & "Conciliados'"
        Case "2": arrform(0) = "Tipo='" & "Pendientes'"
    End Select
    arrparm(4) = recibosrendicion(1, RsConcil)
    arrparm(5) = "1"
    arrform(1) = "Oficina='" & Ctr_AyudaOficina.xnombre & "'"
    arrform(2) = "Caja='" & Ctr_AyudaCaja.xnombre & "'"
    arrform(3) = "mon='" & mon_descripcion & "'"
    arrform(4) = "Fecha='" & Format(DTPfechaini.Value, "dd/mm/yyyy") & "'"
    arrform(5) = "Saldoinicial=" & Txtsaldoini.valor
    arrform(6) = "Nrorendicion=" & Format(Ctr_AyuRendicion.xclave, "000000")
    NombreRep = "XX_Rendiciones.rpt"
    Call ImpresionRptProc(NombreRep, arrform, arrparm, , "Rendiciones")
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

Private Sub Imprimirgastos()
Dim vardllgen As New dllgeneral.dll_general
Dim arrform(7) As Variant, arrparm(4) As Variant
Dim NombreRep As String, CadOrden As String
Dim Fecha As String
Dim fecha1 As String
    fecha1 = Format(DateSerial(DTPfechaini.Year, DTPfechaini.Month, 1), "dd/mm/yyyy")
    '@Base,@cuenta,@concil,@Fecharef,@anno,@mes
    arrparm(0) = VGParamSistem.BDEmpresa
    arrparm(1) = Trim(Ctr_AyudaCaja.xclave)
    arrparm(2) = mon
    arrparm(3) = Ctr_AyuRendicion.xclave

    arrform(1) = "Oficina='" & Ctr_AyudaOficina.xnombre & "'"
    arrform(2) = "Caja='" & Ctr_AyudaCaja.xnombre & "'"
    arrform(3) = "mon='" & mon_descripcion & "'"
    arrform(4) = "Nrorendicion=" & Format(Ctr_AyuRendicion.xclave, "000000")
    arrform(5) = "saldoinicial=" & Txtsaldoini.Text
    NombreRep = "te_GastosxRendicion.rpt"
    Call ImpresionRptProc(NombreRep, arrform, arrparm, , "Gastos x Rendiciones")

End Sub
