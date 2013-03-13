VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Begin VB.Form frmMantPlangastos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plan de Gastos"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11325
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   11325
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   1058
      ButtonWidth     =   1667
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Nuevo"
            Object.ToolTipText     =   "Añadir registro al Plan Cuentas"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "E&ditar"
            Object.ToolTipText     =   "Modificar Plan de Cuentas"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Eliminar"
            Object.ToolTipText     =   "Eliminar Plan de Cuentas"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Imprimir"
            Object.ToolTipText     =   "Listar Plan de Cuentas"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salir"
            Object.ToolTipText     =   "Cerrar la Ventana "
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   7305
      Left            =   0
      TabIndex        =   0
      Top             =   -105
      Width           =   4380
      Begin TextFer.TxFer txtBuscar 
         Height          =   300
         Left            =   48
         TabIndex        =   3
         Top             =   768
         Width           =   3720
         _ExtentX        =   6562
         _ExtentY        =   529
         BackColor       =   16777215
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
         Enabled         =   0   'False
         ColorIlumina    =   -2147483624
         Valor           =   ""
         NoCaracteres    =   "0123456789"
         SignoNegativo   =   0   'False
         MarcarTextoAlEnfoque=   -1  'True
         NoRangoCadena   =   -1  'True
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   135
         Top             =   6615
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   13
         ImageHeight     =   13
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantPlanGastos.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantPlanGastos.frx":0108
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   300
         Left            =   3840
         TabIndex        =   2
         Top             =   780
         Width           =   285
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   6216
         Left            =   48
         TabIndex        =   1
         Top             =   1080
         Width           =   4152
         _ExtentX        =   7329
         _ExtentY        =   10954
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   353
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "ImageList1"
         Appearance      =   1
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6585
      Left            =   4320
      TabIndex        =   7
      Top             =   600
      Width           =   6720
      _ExtentX        =   11853
      _ExtentY        =   11615
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Consulta"
      TabPicture(0)   =   "frmMantPlanGastos.frx":0204
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblNumReg"
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(2)=   "TDBGrid1"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmMantPlanGastos.frx":0220
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lbl(9)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lbl(7)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lbl(6)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lbl(5)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lbl(4)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lbl(14)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lbl(11)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lblNivel"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label3"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lbl(8)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "lbl(0)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label2"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "lbl(3)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "lbl(10)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "letipdoc"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Ctr_AyuTipoDoc"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "txt(3)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Ctr_AyuAnalitico"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Ctrayu_cuenta"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "CtrAyu_tipogastos"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "txt(2)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "txt(1)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "txt(0)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "chk(3)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "chk(0)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "cmdDistribucion"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "chk(2)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "chk(1)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "cAcepta"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "cCancela"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "chk(4)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).ControlCount=   31
      TabCaption(2)   =   "Cuentas Distribución"
      TabPicture(2)   =   "frmMantPlanGastos.frx":023C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdAceptar"
      Tab(2).Control(1)=   "cmdSigue"
      Tab(2).Control(2)=   "txtPorcen"
      Tab(2).Control(3)=   "Ctr_Ayuda3"
      Tab(2).Control(4)=   "TDBGrid2"
      Tab(2).Control(5)=   "Ctr_Ayuda2"
      Tab(2).Control(6)=   "Label4"
      Tab(2).Control(7)=   "lbl(12)"
      Tab(2).Control(8)=   "lbl(13)"
      Tab(2).Control(9)=   "Label5"
      Tab(2).Control(10)=   "lblPorcen"
      Tab(2).ControlCount=   11
      Begin VB.CheckBox chk 
         Height          =   240
         Index           =   4
         Left            =   2715
         TabIndex        =   49
         Top             =   3720
         Width           =   210
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "..."
         Height          =   300
         Left            =   -70725
         TabIndex        =   17
         Top             =   2610
         Width           =   270
      End
      Begin VB.CommandButton cmdSigue 
         Caption         =   "&Aceptar"
         Height          =   360
         Left            =   -72360
         TabIndex        =   15
         Top             =   5175
         Width           =   1125
      End
      Begin VB.CommandButton cCancela 
         Caption         =   "&Cancelar"
         Height          =   360
         Left            =   3468
         TabIndex        =   14
         Top             =   5955
         Width           =   1125
      End
      Begin VB.CommandButton cAcepta 
         Caption         =   "&Aceptar"
         Height          =   360
         Left            =   1560
         TabIndex        =   13
         Top             =   5955
         Width           =   1125
      End
      Begin VB.CheckBox chk 
         Enabled         =   0   'False
         Height          =   240
         Index           =   1
         Left            =   2670
         TabIndex        =   12
         Top             =   5280
         Width           =   285
      End
      Begin VB.CheckBox chk 
         Enabled         =   0   'False
         Height          =   240
         Index           =   2
         Left            =   2670
         TabIndex        =   11
         Top             =   5625
         Width           =   210
      End
      Begin VB.CommandButton cmdDistribucion 
         Caption         =   "..."
         Height          =   264
         Left            =   2970
         TabIndex        =   10
         Top             =   5130
         Width           =   255
      End
      Begin VB.CheckBox chk 
         Height          =   240
         Index           =   0
         Left            =   2688
         TabIndex        =   9
         Top             =   2064
         Width           =   195
      End
      Begin VB.CheckBox chk 
         Height          =   240
         Index           =   3
         Left            =   2715
         TabIndex        =   8
         Top             =   4920
         Width           =   210
      End
      Begin TextFer.TxFer txtPorcen 
         Height          =   315
         Left            =   -72075
         TabIndex        =   16
         Top             =   2595
         Width           =   1305
         _ExtentX        =   2302
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
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         TipoDato        =   1
         NumeroDecimales =   2
         SignoNegativo   =   0   'False
         MarcarTextoAlEnfoque=   -1  'True
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   5256
         Left            =   -74964
         TabIndex        =   18
         Top             =   672
         Width           =   7152
         _ExtentX        =   12621
         _ExtentY        =   9260
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Cuenta"
         Columns(0).DataField=   "gastoscodigo"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Descripción"
         Columns(1).DataField=   "gastosdescripcion"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Cta Nivel"
         Columns(2).DataField=   "gastosnivel"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Cuenta Contable"
         Columns(3).DataField=   "cuentacodigo"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Crtl Costos"
         Columns(4).DataField=   "gastosctrlcostos"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Centros de Costos"
         Columns(5).DataField=   "gastoscostos"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Tipo Gastos"
         Columns(6).DataField=   "tipogastoscodigo"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Est. Distrib."
         Columns(7).DataField=   "gastosestadodistribucion"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Estado Gasto"
         Columns(8).DataField=   "gastosestado"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Tipo Analitico"
         Columns(9).DataField=   "tipoanaliticocodigo"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "Detraccion"
         Columns(10).DataField=   "habilitadodetraccion"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "Ctacte"
         Columns(11).DataField=   "gastosgeneractacte"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "Tipo Doc"
         Columns(12).DataField=   "tipodocumentocodigo"
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   13
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=13"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2037"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1958"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=3016"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2937"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=1323"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1244"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=1482"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=1402"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=2514"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2434"
         Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(21)=   "Column(5).Width=3043"
         Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=2963"
         Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(25)=   "Column(6).Width=3043"
         Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=2963"
         Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(29)=   "Column(7).Width=3043"
         Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=2963"
         Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(33)=   "Column(8).Width=3043"
         Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=2963"
         Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(37)=   "Column(9).Width=3043"
         Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=2963"
         Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(41)=   "Column(10).Width=2725"
         Splits(0)._ColumnProps(42)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(10)._WidthInPix=2646"
         Splits(0)._ColumnProps(44)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(45)=   "Column(11).Width=2725"
         Splits(0)._ColumnProps(46)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(47)=   "Column(11)._WidthInPix=2646"
         Splits(0)._ColumnProps(48)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(49)=   "Column(12).Width=2725"
         Splits(0)._ColumnProps(50)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(51)=   "Column(12)._WidthInPix=2646"
         Splits(0)._ColumnProps(52)=   "Column(12).Order=13"
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000014&,.fgcolor=&H80000007&"
         _StyleDefs(7)   =   ":id=1,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=82,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=79,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=80,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=81,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=62,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=59,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=60,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=61,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=54,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=51,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=52,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=53,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=58,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=55,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=56,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=57,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=66,.parent=13"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
         _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=70,.parent=13"
         _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=67,.parent=14"
         _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=68,.parent=15"
         _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=69,.parent=17"
         _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=74,.parent=13"
         _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=71,.parent=14"
         _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=72,.parent=15"
         _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=73,.parent=17"
         _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=78,.parent=13"
         _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=75,.parent=14"
         _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=76,.parent=15"
         _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=77,.parent=17"
         _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=86,.parent=13"
         _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=83,.parent=14"
         _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=84,.parent=15"
         _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=85,.parent=17"
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
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda3 
         Height          =   450
         Left            =   -74850
         TabIndex        =   19
         Top             =   1875
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   794
         XcodMaxLongitud =   0
         xcodwith        =   800
         NomTabla        =   "ct_cuenta"
         ListaCampos     =   "cuentacodigo(1),cuentadescripcion(1)"
         XcodCampo       =   "cuentacodigo"
         XListCampo      =   "cuentadescripcion"
         ListaCamposDescrip=   "Cuenta,Descripción"
         ListaCamposText =   "cuentacodigo,cuentadescripcion"
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
         Height          =   1635
         Left            =   -74970
         TabIndex        =   20
         Top             =   3015
         Width           =   6480
         _ExtentX        =   11430
         _ExtentY        =   2884
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000018&,.bold=0,.fontsize=825"
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
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda2 
         Height          =   345
         Left            =   -74835
         TabIndex        =   21
         Top             =   1155
         Width           =   4350
         _ExtentX        =   7673
         _ExtentY        =   609
         XcodMaxLongitud =   20
         xcodwith        =   800
         NomTabla        =   "ct_cuenta"
         ListaCampos     =   "cuentacodigo(1),cuentadescripcion(1)"
         XcodCampo       =   "cuentacodigo"
         XListCampo      =   "cuentadescripcion"
         ListaCamposDescrip=   "Cuenta,Descripción"
         ListaCamposText =   "cuentacodigo,cuentadescripcion"
      End
      Begin TextFer.TxFer txt 
         Height          =   348
         Index           =   0
         Left            =   2688
         TabIndex        =   22
         Top             =   720
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   609
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
         MaxLength       =   20
         Text            =   ""
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         SignoNegativo   =   0   'False
         MarcarTextoAlEnfoque=   -1  'True
      End
      Begin TextFer.TxFer txt 
         Height          =   348
         Index           =   1
         Left            =   2688
         TabIndex        =   23
         Top             =   1188
         Width           =   3636
         _ExtentX        =   6403
         _ExtentY        =   609
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
         MaxLength       =   35
         Text            =   ""
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         MarcarTextoAlEnfoque=   -1  'True
      End
      Begin TextFer.TxFer txt 
         Height          =   348
         Index           =   2
         Left            =   2640
         TabIndex        =   24
         Top             =   2400
         Width           =   3636
         _ExtentX        =   6403
         _ExtentY        =   609
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
         MaxLength       =   35
         Text            =   ""
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         MarcarTextoAlEnfoque=   -1  'True
      End
      Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_tipogastos 
         Height          =   384
         Left            =   2640
         TabIndex        =   25
         Top             =   2880
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   688
         XcodMaxLongitud =   2
         xcodwith        =   500
         NomTabla        =   "co_tipogastos"
         ListaCampos     =   "tipogastoscodigo(1),tipogastosdescripcion(1)"
         XcodCampo       =   "tipogastoscodigo"
         XListCampo      =   "tipogastosdescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "tipogastoscodigo,tipogastosdescripcion"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctrayu_cuenta 
         Height          =   390
         Left            =   2670
         TabIndex        =   26
         Top             =   1680
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   688
         XcodMaxLongitud =   0
         xcodwith        =   700
         NomTabla        =   "ct_cuenta"
         ListaCampos     =   "cuentacodigo(1),cuentadescripcion(1)"
         XcodCampo       =   "cuentacodigo"
         XListCampo      =   "cuentadescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "cuentacodigo,cuentadescripcion"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuAnalitico 
         Height          =   390
         Left            =   2685
         TabIndex        =   27
         Top             =   3360
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   688
         XcodMaxLongitud =   4
         xcodwith        =   500
         NomTabla        =   "ct_tipoanalitico"
         ListaCampos     =   "tipoanaliticocodigo(1),tipoanaliticodescripcion(1)"
         XcodCampo       =   "tipoanaliticocodigo"
         XListCampo      =   "tipoanaliticodescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "tipoanaliticocodigo,tipoanaliticodescripcion"
      End
      Begin TextFer.TxFer txt 
         Height          =   345
         Index           =   3
         Left            =   2640
         TabIndex        =   28
         Top             =   4440
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   609
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
         MaxLength       =   3
         Text            =   ""
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         SignoNegativo   =   0   'False
         MarcarTextoAlEnfoque=   -1  'True
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuTipoDoc 
         Height          =   315
         Left            =   2640
         TabIndex        =   51
         Top             =   4080
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         XcodMaxLongitud =   2
         NomTabla        =   "cp_tipodocumento"
         TituloAyuda     =   "Busqueda de Tipo de  Documento"
         ListaCampos     =   "tdocumentocodigo(1),tdocumentodescripcion(1),tdocumentotipo(1),documentoretencion(1)"
         XcodCampo       =   "tdocumentocodigo"
         XListCampo      =   "tdocumentodescripcion"
         ListaCamposDescrip=   "Código,Descripción,CargoAbono,Retencion"
         ListaCamposText =   "tdocumentocodigo,tdocumentodescripcion,tdocumentotipo,documentoretencion"
         Requerido       =   0   'False
      End
      Begin VB.Label letipdoc 
         Caption         =   "Tipo Doc.en Cta cte. :"
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   4200
         Visible         =   0   'False
         Width           =   2160
      End
      Begin VB.Label lbl 
         Caption         =   "Genera registro Cta.Cte"
         Height          =   285
         Index           =   10
         Left            =   240
         TabIndex        =   50
         Top             =   3720
         Width           =   1845
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Registros"
         Height          =   276
         Left            =   -70536
         TabIndex        =   48
         Top             =   6144
         Width           =   948
      End
      Begin VB.Label lblNumReg 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblNumReg"
         ForeColor       =   &H0000FFFF&
         Height          =   288
         Left            =   -69408
         TabIndex        =   47
         Top             =   6132
         Width           =   912
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Porcentaje de Distribución (%)"
         ForeColor       =   &H0080FFFF&
         Height          =   300
         Left            =   -74955
         TabIndex        =   46
         Top             =   2610
         Width           =   2895
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Caption         =   "Seleccionar Cuenta al Abono"
         Height          =   255
         Index           =   12
         Left            =   -74835
         TabIndex        =   45
         Top             =   1635
         Width           =   4095
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Caption         =   "Seleccionar Cuenta al Cargo"
         Height          =   255
         Index           =   13
         Left            =   -74790
         TabIndex        =   44
         Top             =   900
         Width           =   4035
      End
      Begin VB.Label Label5 
         Caption         =   "Total (%)"
         Height          =   255
         Left            =   -70440
         TabIndex        =   43
         Top             =   4740
         Width           =   765
      End
      Begin VB.Label lblPorcen 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   -69705
         TabIndex        =   42
         Top             =   4695
         Width           =   1215
      End
      Begin VB.Label lbl 
         Caption         =   "Centros de Costos"
         Height          =   276
         Index           =   3
         Left            =   192
         TabIndex        =   41
         Top             =   2496
         Width           =   1848
      End
      Begin VB.Label Label2 
         Caption         =   "Código Gastos/operacion"
         Height          =   360
         Left            =   195
         TabIndex        =   40
         Top             =   720
         Width           =   1845
      End
      Begin VB.Label lbl 
         Caption         =   "Descripción "
         Height          =   276
         Index           =   0
         Left            =   192
         TabIndex        =   39
         Top             =   1284
         Width           =   1848
      End
      Begin VB.Label lbl 
         Caption         =   "Controla Centro Costo"
         Height          =   288
         Index           =   8
         Left            =   216
         TabIndex        =   38
         Top             =   2124
         Width           =   1848
      End
      Begin VB.Label Label3 
         Caption         =   "Nivel Cuenta"
         Height          =   324
         Left            =   4656
         TabIndex        =   37
         Top             =   780
         Width           =   948
      End
      Begin VB.Label lblNivel 
         BorderStyle     =   1  'Fixed Single
         Height          =   324
         Left            =   5700
         TabIndex        =   36
         Top             =   744
         Width           =   600
      End
      Begin VB.Label lbl 
         Caption         =   "Codigo Inactivo"
         Height          =   285
         Index           =   11
         Left            =   195
         TabIndex        =   35
         Top             =   5655
         Width           =   1845
      End
      Begin VB.Label lbl 
         Caption         =   "Cuenta Contable"
         Height          =   288
         Index           =   14
         Left            =   192
         TabIndex        =   34
         Top             =   1764
         Width           =   1848
      End
      Begin VB.Label lbl 
         Caption         =   "Tipo de Gastos"
         Height          =   288
         Index           =   4
         Left            =   192
         TabIndex        =   33
         Top             =   2928
         Width           =   1848
      End
      Begin VB.Label lbl 
         Caption         =   "Distribucion de Gastos"
         Height          =   285
         Index           =   5
         Left            =   195
         TabIndex        =   32
         Top             =   5310
         Width           =   1845
      End
      Begin VB.Label lbl 
         Caption         =   "Tipo de Analitico"
         Height          =   285
         Index           =   6
         Left            =   240
         TabIndex        =   31
         Top             =   3405
         Width           =   1845
      End
      Begin VB.Label lbl 
         Caption         =   "Activa Detraccion"
         Height          =   285
         Index           =   7
         Left            =   240
         TabIndex        =   30
         Top             =   4920
         Width           =   1845
      End
      Begin VB.Label lbl 
         Caption         =   "Codigo  Equivalente"
         Height          =   285
         Index           =   9
         Left            =   240
         TabIndex        =   29
         Top             =   4560
         Width           =   1845
      End
   End
   Begin VB.Label lbl 
      Caption         =   "Cuenta Contable"
      Height          =   288
      Index           =   2
      Left            =   0
      TabIndex        =   5
      Top             =   84
      Width           =   2328
   End
   Begin VB.Label lbl 
      Caption         =   "Cuenta Contable"
      Height          =   288
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   84
      Width           =   2328
   End
End
Attribute VB_Name = "frmMantPlangastos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim modoinsert As Boolean
Dim modoedit As Boolean
Dim i_filaorigen As Integer
Dim rs As New ADODB.Recordset
Dim rsDist As New ADODB.Recordset
Dim cCta As String
Dim dCta As String
Dim xgastos As String
Dim xdllgen As New dll_general
Dim FLAGMOVIMIENTODISTRI As Boolean
Dim FLAGDISTRIBUCION As Boolean
Dim l_error As String

Private Sub Ctr_AyuAnalitico_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
  If modoedit = True Then
     cAcepta.Enabled = True
  End If
End Sub



Private Sub CtrAyu_tipogastos_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
  
  If modoedit = True Then
     cAcepta.Enabled = True
  End If

End Sub

Private Sub Form_Load()
  Screen.MousePointer = 11
  l_error = Empty
  Call ConfiguraForm
  Call MuestraDatos(Empty)
  Call Arbol(txtbuscar.Text)
  Set xdllgen = New dll_general
  If Len(l_error) > 0 Then
     MsgBox "Error inesperado: " & err.Number & " " & err.Description, Caption
     Resume Next
    Screen.MousePointer = 1
  
  End If

  Screen.MousePointer = 1
  TDBGrid1.FetchRowStyle = True
  xgastos = "%"
End Sub

Sub ConfiguraForm()
  SSTab1.Tab = 0
  SSTab1.TabEnabled(1) = False
  SSTab1.TabEnabled(2) = False
  Call CtrAyu_Cuenta.conexion(VGcnxCT):
  CtrAyu_Cuenta.Filtro = "(cuentanivel=" & VGnumniveles & " and cuentacodigo <>'00') "
  If VGParametros.sistemamultiempresas Then CtrAyu_Cuenta.NomTabla = "ct_cuenta"
  Call CtrAyu_tipogastos.conexion(VGCNx)
  Ctr_Ayuda2.Filtro = "gastosnivel=" & VGnumnivgas
  Ctr_Ayuda3.Filtro = "gastosnivel=" & VGnumnivgas
  Call Ctr_AyuAnalitico.conexion(VGcnxCT)
  Call Ctr_AyuTipoDoc.conexion(VGCNx)
  cAcepta.Enabled = False
  lblNumReg.Caption = Empty
  frmMantPlangastos.Width = 11310
  frmMantPlangastos.Height = 7650
  FLAGMOVIMIENTODISTRI = False
  FLAGDISTRIBUCION = False
  Call IniciaGridDist
End Sub

Public Function MuestraDatos(xCta As String)
  Dim SQL As String
   SQL = "SELECT *  "
   SQL = SQL & "FROM co_gastos "
   SQL = SQL & "WHERE co_gastos.gastoscodigo<>'00'"
   If xCta <> Empty Then
     SQL = SQL & "AND co_gastos.gastoscodigo like '" & Trim(xCta) & "%' "
   End If
   SQL = SQL & "ORDER BY 1"
   Set rs = VGCNx.Execute(SQL)
   Set TDBGrid1.DataSource = rs
   Call ConfiguraTdbgrid
   lblNumReg.Caption = rs.RecordCount
   SSTab1.Tab = 0
End Function

Private Sub txtBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
      Call CmdBuscar_Click
    End If
End Sub

Private Sub CmdBuscar_Click()
  Set VGvardllgen = New dllgeneral.dll_general
  txtbuscar.Text = VGvardllgen.ESNULO(txtbuscar.Text, "%")
  Call MuestraDatos(txtbuscar.Text)
  Call Arbol(txtbuscar.Text)
End Sub

Sub EditarValores()
 Dim I As Integer
  With TDBGrid1
    txt(0).Text = Trim(xdllgen.ESNULO(.Columns(0).Text, Empty))
    txt(1).Text = Trim(xdllgen.ESNULO(.Columns(1).Text, Empty))
    lblNivel.Caption = Trim(xdllgen.ESNULO(.Columns(2).Text, Empty))
    CtrAyu_Cuenta.xclave = Trim(xdllgen.ESNULO(.Columns(3).Text, Empty)): CtrAyu_Cuenta.Ejecutar
    chk(0).Value = IIf(Trim(.Columns(4).Text) = -1, 1, 0)
    txt(2).Text = Trim(xdllgen.ESNULO(.Columns(5).Text, Empty))
    CtrAyu_tipogastos.xclave = Trim(xdllgen.ESNULO(.Columns(6).Text, Empty)): CtrAyu_tipogastos.Ejecutar
    chk(1).Value = IIf(Trim(.Columns(7).Text) = -1, 1, 0)
    chk(2).Value = IIf(Trim(.Columns(8).Text) = -1, 1, 0)
    chk(3).Value = IIf(Trim(.Columns(10).Text) = 1, 1, 0)
    chk(4).Value = IIf(Trim(.Columns(11).Text) = -1, 1, 0)
    Ctr_AyuAnalitico.xclave = Trim(xdllgen.ESNULO(.Columns(9).Text, Empty)): Ctr_AyuAnalitico.Ejecutar
    Ctr_AyuTipoDoc.xclave = Trim(xdllgen.ESNULO(.Columns(11).Text, Empty)): Ctr_AyuTipoDoc.Ejecutar
  End With
  Call ConfiguraModoEdicion
End Sub

Sub ConfiguraModoEdicion()
    If lblNivel.Caption = Empty Then
       MsgBox "Debe registrar el Código de gastos Contable", vbInformation, Caption
       Call ModoNormal  'Deshabilitar todos los objetos de ingreso
       txt(0).SetFocus
     Else
       If lblNivel.Caption = VGnumnivgas Then
          Call ModoEdicion(True)
        Else
          Call ModoEdicion(False)
       End If
    End If
End Sub

Public Function LimpiarValores()
 Dim I As Integer
 For I = 0 To 2
    txt(I).Text = Empty
  Next
  For I = 0 To 2
    chk(I).Value = 0
  Next
  lblNivel.Caption = Empty
  
End Function

Private Sub cAcepta_Click()
 If ValidarData = True Then
   Call GrabarData
 End If
End Sub

Private Sub cCancela_Click()
  SSTab1.TabEnabled(0) = True
  SSTab1.Tab = 0
  SSTab1.SetFocus
  Toolbar1.Visible = True
  TreeView1.Enabled = True
  modoinsert = False
  modoedit = False
  i_filaorigen = -1
  FLAGDISTRIBUCION = False
End Sub

Sub GrabarData()
  Dim SQL As String
  On Error GoTo X
  
  SSTab1.TabEnabled(0) = True
   
  If modoinsert = True Then

    VGCNx.BeginTrans
    SQL = GrabarPlangastos(0)
    VGCNx.Execute (SQL)
    
    If CInt(lblNivel.Caption) = VGnumnivgas And FLAGDISTRIBUCION = True Then
       Call GrabargastosDistribucion
       Call GrabarTablaSaldos
    End If
    
    VGCNx.CommitTrans
    Call Arbol(txtbuscar.Text)
                  
  ElseIf modoedit = True Then
    VGCNx.BeginTrans
    SQL = GrabarPlangastos(1)
    VGCNx.Execute (SQL)
    
    If CInt(lblNivel.Caption) = VGnumnivgas And FLAGDISTRIBUCION = True Then
        Call GrabargastosDistribucion
    End If
    VGCNx.CommitTrans
  End If
  Call MuestraDatos(Trim(txt(0).Text))
  Toolbar1.Visible = True: TreeView1.Enabled = True: txt(0).Enabled = True
  modoinsert = False: modoedit = False
  i_filaorigen = -1
    FLAGDISTRIBUCION = False
  FLAGMOVIMIENTODISTRI = False
  Set rsDist = Nothing
  Exit Sub

X:
  If err.Number = -2147217873 Then
    MsgBox "Esta intentando registrar uno Existente " & err.Description, vbInformation, Caption
  Else
    MsgBox "Error inesperado: " & err.Number & " " & err.Description, Caption
  Resume
  End If
  VGCNx.RollbackTrans

End Sub

Function ValidarData() As Boolean
 Dim I As Integer
 Dim SQL As String
  If lblNivel.Caption = Empty Then
    MsgBox "No se ha podido registrar el Número de Nivel de la gastos Contable", vbInformation, Caption
    ValidarData = False
    txt(0).SetFocus
    Exit Function
  End If
  
  If VG_gNIVELES(lblNivel.Caption - 1) <> CLng(Len(txt(0).Text)) Then
    MsgBox "La gastos a registrar no corresponde con el Nivel de gastos", vbInformation, Caption
    ValidarData = False
    txt(0).SetFocus
    Exit Function
  End If
  
  If CLng(lblNivel.Caption) > 1 Then
   SQL = "SELECT gastoscodigo FROM co_gastos WHERE gastoscodigo='" & Left(txt(0).Text, VG_gNIVELES(lblNivel.Caption - 2)) & "'"
   If xdllgen.VerificaDatoExistente(VGCNx, SQL) <= 0 Then
     MsgBox "La gastos a registrar no tiene la gastos Superior Correspondiente ", vbInformation, Caption
     ValidarData = False
     txt(0).SetFocus
     Exit Function
   End If
  End If
  
  SQL = "SELECT gastoscodigo FROM co_gastos WHERE gastoscodigo='" & txt(0).Text & "'"
  If modoinsert = True And xdllgen.VerificaDatoExistente(VGCNx, SQL) > 0 Then
    MsgBox "La gastos de Gastos se encuentra registrada en la Base Datos, Debe registrar otra", vbInformation, Caption
    ValidarData = False
    txt(0).SetFocus
    Exit Function
  End If
  
  If CtrAyu_tipogastos.xclave = Empty And lblNivel.Caption = VGnumnivgas Then
     MsgBox "No existe Código de Tipo de gastos en el registro editado", vbInformation, Caption
     ValidarData = False
     Exit Function
  End If
  
  If chk(1).Value = 1 And FLAGDISTRIBUCION = False Then
      MsgBox "No Existe Porcentaje de Distribución para esta gastos, Deshabilitar el check", vbInformation, Caption
      ValidarData = False
      chk(3).SetFocus
      Exit Function
  End If
   
  ValidarData = True
End Function

Private Sub chk_Click(Index As Integer)
  Select Case Index
    Case 1
       If chk(1).Value = 1 Then
         If FLAGMOVIMIENTODISTRI = False Then
            cmdDistribucion.Visible = True
            SSTab1.TabEnabled(2) = True
            SSTab1.Tab = 2
            FLAGMOVIMIENTODISTRI = False
            Call LlenarPorcentajes
         End If
       Else
         cmdDistribucion.Visible = False
         SSTab1.TabEnabled(2) = False
       End If
    Case 4
         letipdoc.Visible = True
         Ctr_AyuTipoDoc.Visible = True
    End Select
  
  If modoedit = True Then
     cAcepta.Enabled = True
  End If
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set rs = Nothing
  Set rsDist = Nothing
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
  SSTab1.TabEnabled(PreviousTab) = False
End Sub

Private Sub TDBGrid1_DblClick()
  If rs.RecordCount > 0 And (modoedit = False And modoinsert = False) Then Call Mantenimiento(1)
End Sub

Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
 Dim rsX As ADODB.Recordset
 Set rsX = rs.Clone(adLockReadOnly)
 rsX.Bookmark = Bookmark
 If rsX!gastosnivel = 1 Then
   RowStyle.BackColor = &H80000018
 End If
End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next
  If rs.Sort = Empty Then
      rs.Sort = TDBGrid1.Columns.Item(ColIndex).DataField & " asc"
   ElseIf Right(rs.Sort, 3) = "asc" Then
      rs.Sort = TDBGrid1.Columns.Item(ColIndex).DataField & " desc"
   ElseIf Right(rs.Sort, 4) = "desc" Then
      rs.Sort = TDBGrid1.Columns.Item(ColIndex).DataField & " asc"
  End If
  Call ConfiguraTdbgrid
  TDBGrid1.Refresh
End Sub

Private Sub ConfiguraTdbgrid()
    With TDBGrid1
        .Columns(0).Width = 900
        .Columns(1).Width = 2100
      '  .Columns(2).Width = 800
      ' .Columns(3).Width = 300
      '  .Columns(4).Width = 1200
    End With

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  Call Mantenimiento(Button.Index - 1)
End Sub

Sub Mantenimiento(indice As Integer)
  Dim j As Integer
  Dim spos As Integer
  Dim SQL As String

  On Error GoTo X
  SSTab1.TabEnabled(1) = True

  Select Case indice
     Case 0   'nuevo
        SSTab1.Tab = 1
        SSTab1.TabEnabled(2) = False
        Call LimpiarValores
        Call ModoNormal
        Toolbar1.Visible = False
        TreeView1.Enabled = False
        modoinsert = True
        FLAGDISTRIBUCION = False

     Case 1   'modificar
        If TDBGrid1.Row < 0 Then
          Exit Sub
        End If
        Call EditarValores
        modoedit = True
        SSTab1.Tab = 1
        Toolbar1.Visible = False
        TreeView1.Enabled = False
        i_filaorigen = TDBGrid1.Row
        txt(0).Enabled = False
        cAcepta.Enabled = False
        FLAGDISTRIBUCION = False

     Case 2   'eliminar
       If MsgBox("Desea eliminar el registro?", vbYesNo + vbDefaultButton2, "AVISO") = vbYes Then
          Dim rs As ADODB.Recordset
          Set rs = New ADODB.Recordset
          SQL = "Select isnull(count(*),0) from co_gastos where gastoscodigo like '" & Trim(TDBGrid1.Columns(0).Value) & "%'"
          Set rs = VGCNx.Execute(SQL)
          If rs(0) > 1 Then
             MsgBox "La gastos a Eliminar tiene gastos Dependientes al Nivel Inferior" & Chr(10) & Chr(13) & "Deben Eliminarse primero las gastos de Nivel Inferior", vbInformation, Caption
             Exit Sub
          End If
          SQL = "DELETE FROM co_gastos WHERE gastoscodigo = '" & Trim(TDBGrid1.Columns(0).Value) & "'"
          VGCNx.Execute (SQL)
          Call MuestraDatos(Trim(TDBGrid1.Columns(0).Value))
       End If

     Case 3   'Imprimir
       With MDIPrincipal
          .cryRpt.Destination = crptToWindow
          .cryRpt.WindowState = crptMaximized
          .cryRpt.StoredProcParam(0) = VGParamSistem.BDEmpresa
          .cryRpt.StoredProcParam(1) = VGParamSistem.BDEmpresaCT
          .cryRpt.StoredProcParam(2) = Trim(xgastos) & "%"
          .cryRpt.formulas(0) = "@Empresa='" & VGParametros.NomEmpresa & "'"
          .cryRpt.ReportFileName = VGParamSistem.RutaReport & "co_Plangastos.rpt"
          If VGsql = 1 Then
             .cryRpt.Connect = "Provider=SQLOLEDB;Password=" & VGParamSistem.PwdGEN & ";Persist Security Info=True;User ID=" & VGParamSistem.UsuarioGEN & ";Initial Catalog=" & VGParamSistem.BDEmpresaGEN & ";SERVER=" & VGParamSistem.ServidorGEN
           Else
            .cryRpt.Connect = vgCADENAREPORT2
          End If
          .cryRpt.DiscardSavedData = True
          .cryRpt.Action = 1
       End With

     Case 4  ' salir
       Unload Me
  End Select
  Exit Sub

X:
  If indice = 2 And err.Number = -2147217873 Then
    MsgBox "Registro no podrá Eliminarse mientras exista Información en las Tablas Relacionadas", vbInformation, Caption
  Else
    MsgBox "Error inesperado: " & err.Description & "  " & err.Number, vbInformation, Caption
    On Error Resume Next
    Exit Sub
    Resume
  End If
End Sub

Private Sub txt_Change(Index As Integer)
  cAcepta.Enabled = IIf(xdllgen.ESNULO(lblNivel.Caption, 0) = VGnumnivgas, ValidargastosUltimoNivel(), ValidargastosNivel())
End Sub

  Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
          If KeyAscii = 13 And Index = 1 And cAcepta.Value = True Then
    cAcepta.SetFocus
    Call cAcepta_Click
  End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
 Dim I As Integer
 txt(Index).Text = UCase(txt(Index).Text)
 If modoinsert = True And Index = 0 Then
    For I = 1 To VGnumnivgas
        If VG_gNIVELES(I - 1) = Len(Trim(txt(0).Text)) Then
           lblNivel.Caption = I
           Call ConfiguraModoEdicion
           Exit For
         Else
           lblNivel.Caption = Empty
        End If
     Next
 End If
 If Index = 1 Then Call ConfiguraModoEdicion
  
End Sub

Private Sub Ctr_Ayuda4_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
   cAcepta.Enabled = IIf(xdllgen.ESNULO(lblNivel.Caption, 0) = VGnumnivgas, ValidargastosUltimoNivel(), ValidargastosNivel())
End Sub

Private Sub cboTipoAjuste_Click()
  cAcepta.Enabled = IIf(xdllgen.ESNULO(lblNivel.Caption, 0) = VGnumnivgas, ValidargastosUltimoNivel(), ValidargastosNivel())
End Sub

Function ValidargastosUltimoNivel() As Boolean
  ValidargastosUltimoNivel = True
End Function

Function ValidargastosNivel() As Boolean
  ValidargastosNivel = True
End Function

Sub ModoEdicion(flagULTIMONIVEL As Boolean) 'True: Ultimo Nivel  False:Otros Niveles
    CtrAyu_Cuenta.Enabled = flagULTIMONIVEL
    chk(0).Enabled = flagULTIMONIVEL
    txt(2).Enabled = flagULTIMONIVEL
    CtrAyu_tipogastos.Enabled = flagULTIMONIVEL
    chk(1).Enabled = flagULTIMONIVEL
    chk(2).Enabled = flagULTIMONIVEL
    chk(3).Enabled = flagULTIMONIVEL
    chk(4).Enabled = flagULTIMONIVEL
    cmdDistribucion.Visible = flagULTIMONIVEL
    Ctr_AyuTipoDoc.Enabled = flagULTIMONIVEL
End Sub

Sub ModoNormal()
 Dim I As Integer
  For I = 0 To 2
     txt(I).Enabled = True
  Next
  For I = 0 To 2
     chk(I).Enabled = True
  Next
End Sub

Function GrabarPlangastos(tipooperacion As Integer) As String
 Dim SQL As String
Select Case tipooperacion
Case 0
     SQL = "INSERT INTO co_gastos (gastoscodigo, gastosdescripcion, gastosnivel,"
     SQL = SQL & "cuentacodigo, gastosctrlcostos, gastoscostos, tipogastoscodigo, "
     SQL = SQL & "gastosestadodistribucion,gastosestado,habilitadodetraccion,usuariocodigo,fechaact,"
     SQL = SQL & "tipoanaliticocodigo,gastosequivalente,gastosgeneractacte,tipodocumentocodigo) "
     SQL = SQL & " VALUES ('" & RTrim(txt(0).Text) & "', '" & RTrim(txt(1).Text) & "', "
     SQL = SQL & "'" & lblNivel.Caption & "', '" & RTrim(CtrAyu_Cuenta.xclave) & "', "
     SQL = SQL & "'" & chk(0).Value & "', '" & RTrim(txt(2).Text) & "','" & RTrim(CtrAyu_tipogastos.xclave) & "', "
     SQL = SQL & "'" & chk(1).Value & "', '" & chk(2).Value & "','" & chk(3).Value & "',"
     SQL = SQL & "'" & VGUsuario & "', '" & Format(Now, "dd/mm/yyyy") & "',"
     SQL = SQL & "'" & RTrim(Ctr_AyuAnalitico.xclave) & "','" & RTrim(txt(3).Text) & "'," & chk(4).Value & ","
     SQL = SQL & "'" & Ctr_AyuTipoDoc.xclave & "')"
   Case 1
     SQL = "UPDATE co_gastos SET "
     SQL = SQL & "gastosdescripcion='" & RTrim(txt(1).Text) & "', "
     SQL = SQL & "gastosnivel='" & lblNivel.Caption & "', "
     SQL = SQL & "cuentacodigo='" & RTrim(CtrAyu_Cuenta.xclave) & "', "
     SQL = SQL & "gastoscostos='" & RTrim(txt(2).Text) & "', "
     SQL = SQL & "gastosctrlcostos='" & chk(0).Value & "', "
     SQL = SQL & "tipogastoscodigo='" & RTrim(CtrAyu_tipogastos.xclave) & "', "
     SQL = SQL & "gastosestadodistribucion='" & chk(1).Value & "', "
     SQL = SQL & "habilitadodetraccion='" & chk(3).Value & "', "
     SQL = SQL & "gastosestado='" & chk(2).Value & "', "
     SQL = SQL & "usuariocodigo='" & VGUsuario & "', "
     SQL = SQL & "fechaact='" & Format(Now, "dd/mm/yyyy hh:mm:ss") & "', "
     SQL = SQL & "tipoanaliticocodigo='" & RTrim(Ctr_AyuAnalitico.xclave) & "', "
     SQL = SQL & "gastosequivalente='" & RTrim(txt(3).Text) & "',"
     SQL = SQL & "gastosgeneractacte='" & chk(4).Value & "',"
     SQL = SQL & "tipodocumentocodigo='" & Ctr_AyuTipoDoc.xclave & "'"
     SQL = SQL & " WHERE gastoscodigo='" & RTrim(txt(0).Text) & "'"

 End Select
 GrabarPlangastos = SQL

End Function

Function GrabargastosDistribucion()
 Dim I As Long
 Dim SQL As String
 
 If rsDist.RecordCount > 0 Then
   SQL = "DELETE FROM ct_distribucion WHERE gastoscodigo='" & txt(0).Text & "'"
   VGCNx.Execute (SQL)
   rsDist.MoveFirst
   For I = 0 To rsDist.RecordCount - 1
     SQL = "INSERT ct_distribucion (gastoscodigo,distribucioncargo,distribucionabono,distribucionporcen,usuariocodigo,fechaact) VALUES "
     SQL = SQL & "('" & rsDist(0) & "','" & rsDist(1) & "','" & rsDist(2) & "'," & rsDist(3) & ",'" & VGUsuario & "','" & Date & "')"
     VGCNx.Execute (SQL)
     rsDist.MoveNext
   Next
 End If

End Function

Function GrabarTablaSaldos()
 Dim SQL As String
 Dim NombreTabla As String
    NombreTabla = "Co_SALDOS" & VGParamSistem.Anoproceso
    SQL = "INSERT " & NombreTabla & "(gastoscodigo,usuariocodigo,fechaact)"
    SQL = SQL & "VALUES ('" & txt(0).Text & "','" & VGUsuario & "','" & Date & "')"
    VGCNx.Execute (SQL)

End Function

Private Sub cmdSigue_Click()
 If rsDist.RecordCount > 0 And CDbl(lblPorcen.Caption) <> 100 Then
   MsgBox "El Total de % Distribución no se ha completado al 100%", vbInformation, Caption
 Else
   cAcepta.Enabled = IIf(rsDist.RecordCount > 0 And FLAGMOVIMIENTODISTRI = True, True, False)
   chk(3).Value = IIf(rsDist.RecordCount > 0, 1, 0)
   SSTab1.Tab = 1
   SSTab1.TabEnabled(1) = True
 End If

End Sub

Private Sub cmdDistribucion_Click()
  SSTab1.TabEnabled(2) = True
  SSTab1.Tab = 2
  FLAGMOVIMIENTODISTRI = False
  Call LlenarPorcentajes
End Sub

Sub LlenarPorcentajes()
  Set rsDist = Nothing
  Call IniciaGridDist
  Call CargaGridDist
  Call ConfigGridDist
  FLAGDISTRIBUCION = True
End Sub

Private Function IniciaGridDist()
  Call rsDist.Fields.Append("gastos", adVarChar, 20)
  Call rsDist.Fields.Append("gastos ", adVarChar, 20)
  Call rsDist.Fields.Append("Porcentaje", adDouble)
  Call rsDist.Fields.Append("Item", adInteger)
  rsDist.Open
End Function

Private Sub CargaGridDist()
  Dim xRs As ADODB.Recordset
  Dim nConta As Long
  Dim SQL As String
  Set xRs = New ADODB.Recordset
  SQL = "SELECT gastoscodigo,distribucioncargo,distribucionabono,distribucionporcen "
  SQL = SQL & "FROM co_distribucion WHERE gastoscodigo='" & txt(0).Text & "'"
  Set xRs = VGCNx.Execute(SQL)
  nConta = 1
  If xRs.RecordCount > 0 Then
     xRs.MoveFirst
     Do While Not xRs.EOF
       rsDist.AddNew
       rsDist.Fields(0) = xRs(0).Value
       rsDist.Fields(1) = xRs(1).Value
       rsDist.Fields(2) = xRs(2).Value
       rsDist.Fields(3) = xRs(3).Value
       rsDist.Fields(4) = nConta
       nConta = nConta + 1
       rsDist.Update
       xRs.MoveNext
     Loop
  End If
  Set xRs = Nothing
End Sub

Private Function ConfigGridDist()
 Dim I As Integer
  Set TDBGrid2.DataSource = rsDist
  With TDBGrid2
    For I = 0 To 4
      .Columns(I).AllowSizing = False
    Next
    .Columns(0).Visible = False
    .Columns(0).Caption = "gastos"
    .Columns(1).Width = 1700
    .Columns(1).Caption = "gastos Cargo"
    .Columns(2).Width = 1700
    .Columns(2).Caption = "gastos Abono"
    .Columns(3).Width = 1900
    .Columns(3).Caption = "Distribución(%)"
    .Columns(4).Width = 800
    .Columns(4).Caption = "Item"
  End With
  lblPorcen.Caption = DevuelveTotPor()
  TDBGrid2.Refresh
End Function

Private Function ActualizaGridDist()
 Dim nReg As Long
 nReg = DevuelveNumReg() + 1
 With rsDist
   .AddNew
   .Fields(0) = txt(0).Text
   .Fields(1) = Ctr_Ayuda2.xclave
   .Fields(2) = Ctr_Ayuda3.xclave
   .Fields(3) = CDbl(txtPorcen.Text)
   .Fields(4) = CLng(nReg)
   .Update
 End With
End Function

Private Sub txtbuscargastos_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
   Call CmdBuscar_Click
 End If
End Sub


Private Sub txtPorcen_Change()
 If Ctr_Ayuda2.xclave <> Empty And Ctr_Ayuda3.xclave <> Empty Then
   If txtPorcen.Text <> Empty Then
     cmdAceptar.Enabled = True
     Exit Sub
   End If
 End If
 cmdAceptar.Enabled = False
End Sub

Function ValidarGridDist() As Boolean
 
 If DevuelveTotPor() + Val(txtPorcen.Text) > 100 Then
   MsgBox "La Sumatoria de Porcentaje excede al 100%", vbInformation, Caption
   txtPorcen.SetFocus
   SendKeys "{HOME}+{END}"
   ValidarGridDist = False
   Exit Function
 End If
  
 If Val(txtPorcen.Text) = 0 Then
   MsgBox "Valor de Porcentaje  no permitido", vbInformation, Caption
   txtPorcen.SetFocus
   SendKeys "{HOME}+{END}"
   ValidarGridDist = False
   Exit Function
 End If
  
 ValidarGridDist = True
End Function

Function DevuelveTotPor() As Double
  Dim I As Integer
  Dim nSum As Double
  nSum = 0
  If rsDist.RecordCount > 0 Then
    rsDist.MoveFirst
    Do While Not rsDist.EOF
      nSum = nSum + rsDist.Fields(3).Value
      rsDist.MoveNext
    Loop
  End If
  DevuelveTotPor = Format(nSum, "###0.#0")
End Function

Private Sub TDBGrid2_Click()
  TDBGrid2.SetFocus
End Sub

Private Sub TDBGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim nvalor As String
  If rsDist.RecordCount > 0 Then
    If KeyCode = 46 Then
      nvalor = TDBGrid2.Columns(4).Text
      If rsDist.RecordCount > 0 Then
         rsDist.MoveFirst
         Do Until rsDist.EOF
           If rsDist.Fields(4) = nvalor Then
             rsDist.Delete adAffectCurrent
             rsDist.Update
             Exit Do
           End If
           rsDist.MoveNext
         Loop
         ConfigGridDist
         lblPorcen.Caption = DevuelveTotPor
      End If
    End If
  End If
End Sub

Function DevuelveNumReg() As Long
 Dim nUlt As Long
 If rsDist.RecordCount > 0 Then
   rsDist.MoveFirst
   Do While Not rsDist.EOF
     nUlt = rsDist.Fields(4)
     rsDist.MoveNext
   Loop
   DevuelveNumReg = nUlt
 Else
   DevuelveNumReg = 0
 End If
End Function

Private Sub Arbol(xCta As String)
On Error GoTo xx
  Dim rs As ADODB.Recordset
  Dim SQL As String
  Dim CodCan As String
  Dim I As Long
  Dim k As Integer
  Dim nodX As NODE
  
  Set VGvardllgen = New dllgeneral.dll_general
  xCta = VGvardllgen.ESNULO(xCta, "%")
  
  SQL = "SELECT A.gastoscodigo,A.gastosdescripcion, A.gastosnivel FROM co_gastos A "
  SQL = SQL & "WHERE A.gastoscodigo<>'00' "
  SQL = SQL & "ORDER BY 1"
  Set rs = New ADODB.Recordset
  Set rs = VGCNx.Execute(SQL)
  
  If (rs.EOF Or rs.BOF) Then
    Exit Sub
  End If
  TreeView1.Nodes.Clear
  Set nodX = TreeView1.Nodes.Add(, , "P", "Plan de gastoss", 1)
  nodX.EnsureVisible
  
  For I = 1 To rs.RecordCount
    cCta = "P" & Trim(rs(0))
    If rs(2) > 1 Then
      Set nodX = TreeView1.Nodes.Add("P" & Trim(Mid(cCta, 2, VG_gNIVELES(rs(2) - 2))), tvwChild, cCta, rs(0), 1)
    Else
      Set nodX = TreeView1.Nodes.Add("P", tvwChild, cCta, rs(0), 1)
      nodX.EnsureVisible
    End If
    rs.MoveNext
   Next I
   rs.MoveFirst
   Exit Sub
xx:
  MsgBox "Error de Inconsistencia en Base de Datos: " & "No existe Plan de gastoss superior para la gastos " & rs(0) & " - " & rs(1), vbInformation, Caption
  l_error = l_error & "No existe gastos superior para la gastos " & rs(0) & " - " & rs(1) & Chr(13)
  Resume Next
End Sub

Private Sub TREEVIEW1_EXPAND(ByVal NODE As MSComctlLib.NODE)
  NODE.Image = 2
End Sub

Private Sub TREEVIEW1_COLLAPSE(ByVal NODE As MSComctlLib.NODE)
  NODE.Image = 1
End Sub

Private Sub TreeView1_NodeClick(ByVal NODE As MSComctlLib.NODE)
  If NODE.Key <> Empty Then
    xgastos = Right(Trim(NODE.Key), Len(Trim(NODE.Key)) - 1)
    MuestraDatos (Right(Trim(NODE.Key), Len(Trim(NODE.Key)) - 1))
  End If
End Sub

Private Sub cmdaceptar_Click()
 If ValidarGridDist = True Then
   Call ActualizaGridDist
   Call ConfigGridDist
   Ctr_Ayuda2.xclave = Empty
   Ctr_Ayuda3.xclave = Empty
   Ctr_Ayuda2.xnombre = Empty
   Ctr_Ayuda3.xnombre = Empty
   txtPorcen.Text = Empty
   FLAGMOVIMIENTODISTRI = True
 End If
End Sub

Function DevuelveTipogastos()
 Dim rsX As New ADODB.Recordset
 Dim SQL As String
 
 Set rsX = New ADODB.Recordset
 SQL = "Select tipogastoscodigo from co_gastos where gastoscodigo=left('" & Trim(txt(0).Text) & "',2)"
 Set rsX = VGCNx.Execute(SQL)
 If rsX.RecordCount > 0 Then
   DevuelveTipogastos = rsX(0)
 Else
   DevuelveTipogastos = 0
 End If
 Set rsX = Nothing

End Function

