VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmConciliacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conciliacion Bancaria"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15210
   Icon            =   "FrmConciliacion.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   15210
   Begin VB.Frame Frame1 
      Height          =   1230
      Left            =   135
      TabIndex        =   7
      Top             =   30
      Width           =   14940
      Begin VB.CheckBox chkconciliado 
         Caption         =   "Doc. Conciliados"
         Height          =   225
         Left            =   7110
         TabIndex        =   4
         Top             =   840
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker DTPfechaini 
         Height          =   285
         Left            =   1125
         TabIndex        =   19
         Top             =   825
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   16449537
         CurrentDate     =   37513
      End
      Begin MSComCtl2.DTPicker DTPfechafin 
         Height          =   285
         Left            =   3645
         TabIndex        =   20
         Top             =   810
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   16449537
         CurrentDate     =   37513
      End
      Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_Cuenta 
         Height          =   315
         Left            =   1995
         TabIndex        =   8
         Top             =   300
         Width           =   7440
         _ExtentX        =   13123
         _ExtentY        =   556
         XcodMaxLongitud =   20
         xcodwith        =   1000
         NomTabla        =   "ct_cuenta"
         TituloAyuda     =   "Busqueda de Cuenta"
         ListaCampos     =   "cuentacodigo(1),cuentadescripcion(1),cuentaestadoccostos(2),cuentaestadoanalitico(2),cuentadocumento(2),tipoanaliticocodigo(1)"
         XcodCampo       =   "cuentacodigo"
         XListCampo      =   "cuentadescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "cuentacodigo,cuentadescripcion,cuentaestadoccostos,cuentaestadoanalitico,cuentadocumento,tipoanaliticocodigo"
         Requerido       =   0   'False
      End
      Begin VB.Label lbfechafin 
         Caption         =   "Fecha fin :"
         Height          =   240
         Left            =   2715
         TabIndex        =   23
         Top             =   855
         Width           =   780
      End
      Begin VB.Label lbfechini 
         Caption         =   "Fecha inicio :"
         Height          =   240
         Left            =   135
         TabIndex        =   22
         Top             =   855
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
         Left            =   12825
         TabIndex        =   18
         Top             =   765
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
         Left            =   12825
         TabIndex        =   17
         Top             =   480
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
         Left            =   10560
         TabIndex        =   16
         Top             =   780
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
         Left            =   10560
         TabIndex        =   15
         Top             =   495
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
         Left            =   12735
         TabIndex        =   14
         Top             =   210
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
         Left            =   10545
         TabIndex        =   13
         Top             =   270
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
         Left            =   9810
         TabIndex        =   12
         Top             =   840
         Width           =   645
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         Height          =   960
         Left            =   9690
         Shape           =   4  'Rounded Rectangle
         Top             =   165
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
         Left            =   9810
         TabIndex        =   11
         Top             =   540
         Width           =   510
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         Height          =   990
         Left            =   9675
         Shape           =   4  'Rounded Rectangle
         Top             =   150
         Width           =   5085
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta Contable :"
         Height          =   195
         Left            =   135
         TabIndex        =   9
         Top             =   375
         Width           =   1275
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   7065
      Left            =   135
      ScaleHeight     =   7005
      ScaleWidth      =   14865
      TabIndex        =   0
      Top             =   1350
      Width           =   14925
      Begin TrueOleDBGrid70.TDBGrid TDBG_concil 
         Height          =   6480
         Left            =   60
         TabIndex        =   1
         Top             =   60
         Width           =   14685
         _ExtentX        =   25903
         _ExtentY        =   11430
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   4
         Columns(0)._MaxComboItems=   5
         Columns(0).ValueItems(0)._DefaultItem=   0
         Columns(0).ValueItems(0).Value=   "1"
         Columns(0).ValueItems(0).Value.vt=   8
         Columns(0).ValueItems(0).DisplayValue=   "1"
         Columns(0).ValueItems(0).DisplayValue.vt=   8
         Columns(0).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(0).ValueItems(1)._DefaultItem=   0
         Columns(0).ValueItems(1).Value=   "0"
         Columns(0).ValueItems(1).Value.vt=   8
         Columns(0).ValueItems(1).DisplayValue=   "0"
         Columns(0).ValueItems(1).DisplayValue.vt=   8
         Columns(0).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
         Columns(0).ValueItems.Count=   2
         Columns(0).Caption=   "CH"
         Columns(0).DataField=   "detcomprobconci"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Nº Comprob"
         Columns(1).DataField=   "cabcomprobnumero"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Asiento"
         Columns(2).DataField=   "asientocodigo"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "SubAsiento"
         Columns(3).DataField=   "subasientocodigo"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "T/D"
         Columns(4).DataField=   "tipdocref"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Nº"
         Columns(5).DataField=   "detcomprobnumref"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Fecha"
         Columns(6).DataField=   "detcomprobfechaemision"
         Columns(6).NumberFormat=   "Short Date"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Codigo Analisis"
         Columns(7).DataField=   "analiticocodigo"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Razon Social"
         Columns(8).DataField=   "entidadrazonsocial"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Debe Soles"
         Columns(9).DataField=   "detcomprobdebe"
         Columns(9).NumberFormat=   "###,###,###.00"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "Haber Soles"
         Columns(10).DataField=   "detcomprobhaber"
         Columns(10).NumberFormat=   "###,###,###.00"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "Debe USS"
         Columns(11).DataField=   "detcomprobussdebe"
         Columns(11).NumberFormat=   "###,###,###.00"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "Haber USS"
         Columns(12).DataField=   "detcomprobusshaber"
         Columns(12).NumberFormat=   "###,###,###.00"
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(13)._VlistStyle=   0
         Columns(13)._MaxComboItems=   5
         Columns(13).Caption=   "T/C"
         Columns(13).DataField=   "detcomprobtipocambio"
         Columns(13).NumberFormat=   "#.000"
         Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   14
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=14"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=714"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=635"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2143"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2064"
         Splits(0)._ColumnProps(8)=   "Column(1)._ColStyle=8196"
         Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(10)=   "Column(2).Width=979"
         Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=900"
         Splits(0)._ColumnProps(13)=   "Column(2)._ColStyle=8196"
         Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(15)=   "Column(3).Width=1244"
         Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=1164"
         Splits(0)._ColumnProps(18)=   "Column(3)._ColStyle=8196"
         Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(20)=   "Column(4).Width=873"
         Splits(0)._ColumnProps(21)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(22)=   "Column(4)._WidthInPix=794"
         Splits(0)._ColumnProps(23)=   "Column(4)._ColStyle=8196"
         Splits(0)._ColumnProps(24)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(25)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(26)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(28)=   "Column(5)._ColStyle=8196"
         Splits(0)._ColumnProps(29)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(30)=   "Column(6).Width=2090"
         Splits(0)._ColumnProps(31)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(32)=   "Column(6)._WidthInPix=2011"
         Splits(0)._ColumnProps(33)=   "Column(6)._ColStyle=8196"
         Splits(0)._ColumnProps(34)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(35)=   "Column(7).Width=2090"
         Splits(0)._ColumnProps(36)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(37)=   "Column(7)._WidthInPix=2011"
         Splits(0)._ColumnProps(38)=   "Column(7)._ColStyle=8196"
         Splits(0)._ColumnProps(39)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(40)=   "Column(8).Width=2196"
         Splits(0)._ColumnProps(41)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(42)=   "Column(8)._WidthInPix=2117"
         Splits(0)._ColumnProps(43)=   "Column(8)._ColStyle=8196"
         Splits(0)._ColumnProps(44)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(45)=   "Column(9).Width=2275"
         Splits(0)._ColumnProps(46)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(47)=   "Column(9)._WidthInPix=2196"
         Splits(0)._ColumnProps(48)=   "Column(9)._ColStyle=8194"
         Splits(0)._ColumnProps(49)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(50)=   "Column(10).Width=2196"
         Splits(0)._ColumnProps(51)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(52)=   "Column(10)._WidthInPix=2117"
         Splits(0)._ColumnProps(53)=   "Column(10)._ColStyle=8194"
         Splits(0)._ColumnProps(54)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(55)=   "Column(11).Width=2196"
         Splits(0)._ColumnProps(56)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(57)=   "Column(11)._WidthInPix=2117"
         Splits(0)._ColumnProps(58)=   "Column(11)._ColStyle=8194"
         Splits(0)._ColumnProps(59)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(60)=   "Column(12).Width=2170"
         Splits(0)._ColumnProps(61)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(62)=   "Column(12)._WidthInPix=2090"
         Splits(0)._ColumnProps(63)=   "Column(12)._ColStyle=8194"
         Splits(0)._ColumnProps(64)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(65)=   "Column(13).Width=1217"
         Splits(0)._ColumnProps(66)=   "Column(13).DividerColor=0"
         Splits(0)._ColumnProps(67)=   "Column(13)._WidthInPix=1138"
         Splits(0)._ColumnProps(68)=   "Column(13)._ColStyle=8194"
         Splits(0)._ColumnProps(69)=   "Column(13).Order=14"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=50,.parent=13,.bgcolor=&HBFFFAA&"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.locked=-1"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.locked=-1"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.locked=-1"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.locked=-1"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.locked=-1"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.locked=-1"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13,.locked=-1"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=70,.parent=13,.locked=-1"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
         _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=74,.parent=13,.alignment=1,.locked=-1"
         _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
         _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
         _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
         _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=78,.parent=13,.alignment=1,.locked=-1"
         _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=14"
         _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=15"
         _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=17"
         _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=82,.parent=13,.alignment=1,.locked=-1"
         _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=79,.parent=14"
         _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=80,.parent=15"
         _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=81,.parent=17"
         _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=86,.parent=13,.alignment=1,.locked=-1"
         _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=83,.parent=14"
         _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=84,.parent=15"
         _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=85,.parent=17"
         _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=90,.parent=13,.alignment=1,.locked=-1"
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
         Left            =   13695
         TabIndex        =   3
         Top             =   6660
         Width           =   1050
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Registros :"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   12600
         TabIndex        =   2
         Top             =   6690
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmConciliacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents RsConcil As ADODB.Recordset
Attribute RsConcil.VB_VarHelpID = -1

Private Sub chkconciliado_Click()
 If CtrAyu_Cuenta.xclave <> Empty Then
    Call Listar
    Call CalcularTotales(RsConcil)
 End If
End Sub

Private Sub Form_Load()
    lbfechini.Enabled = False
    DTPfechaini.Enabled = False
    lbfechafin.Enabled = False
    DTPfechafin.Enabled = False
    Width = 15330: Height = 9390
    Left = 0: Top = 0
    Call CtrAyu_Cuenta.conexion(VGcnx)
    CtrAyu_Cuenta.Filtro = "left(cuentacodigo," & VG_aNIVELES(2) & ") like '104%' and " & _
                           "cuentanivel =" & VGnumnivelescuenta
End Sub

Private Sub axBAceptar_Click()
    RsConcil.UpdateBatch
    axBImprimir(0).Enabled = True
    axBImprimir(1).Enabled = True
    axbAceptar.Enabled = False
End Sub

Private Sub axBCancelar_Click()
    If RsConcil Is Nothing Then
        Unload Me
        Exit Sub
    End If
    
    RsConcil.CancelBatch
    Unload Me
End Sub

Private Sub CtrAyu_Cuenta_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    Call Listar
    Call CalcularTotales(RsConcil)
End Sub

Private Sub Listar()
    Dim sqlcad As String
    
    sqlcad = "select A.cabcomprobnumero,A.asientocodigo,A.subasientocodigo,A.monedacodigo," & _
                  "tipdocref=isnull(A.documentocodigo,''),detcomprobnumref=isnull(A.detcomprobnumdocumento,''),A.detcomprobfechaemision," & _
                  "A.analiticocodigo,C.entidadrazonsocial,A.detcomprobdebe,A.detcomprobhaber,A.detcomprobusshaber," & _
                  "A.detcomprobussdebe , A.detcomprobtipocambio,A.detcomprobconci " & _
                  "From ct_detcomprob" & VGParamSistem.Anoproceso & " A,dbo.v_analiticoentidad C " & _
                  "Where A.analiticocodigo = C.analiticocodigo and " & _
                  "A.cuentacodigo='" & CtrAyu_Cuenta.xclave & "' and " & _
                  "A.cabcomprobmes<=" & CInt(VGParamSistem.Mesproceso) & " and "
    If chkconciliado.Value = 1 Then
        sqlcad = sqlcad & "A.detcomprobconci like '-1' "
    Else
        sqlcad = sqlcad & "A.detcomprobconci like '%' "
    End If
    sqlcad = sqlcad & "ORDER BY A.detcomprobfechaemision"

    Set RsConcil = New ADODB.Recordset
    RsConcil.Open sqlcad, VGcnx, adOpenDynamic, adLockBatchOptimistic
    
    If RsConcil.RecordCount = 0 Then
        lbfechini.Enabled = False
        DTPfechaini.Enabled = False
        lbfechafin.Enabled = False
        DTPfechafin.Enabled = False
      Else
        lbfechini.Enabled = True
        DTPfechaini.Enabled = True
        lbfechafin.Enabled = True
        DTPfechafin.Enabled = True
    End If
    lbnreg.Caption = Format(RsConcil.RecordCount, "0 ")
    Set TDBG_concil.DataSource = RsConcil
    If CLng(lbnreg.Caption) > 0 Then
        axBImprimir(0).Enabled = True
        axBImprimir(1).Enabled = True
    Else
        axBImprimir(0).Enabled = False
        axBImprimir(1).Enabled = False
    End If
End Sub

Private Sub CalcularTotales(ByVal rs As Recordset)
Dim rsaux As ADODB.Recordset
Dim vardllgen As New dllgeneral.dll_general
Set rsaux = rs.Clone(adLockReadOnly)

Dim montosolesDebe As Double, montodolaresDebe As Double
Dim montosolesHaber As Double, montodolaresHaber As Double

montosolesDebe = 0: montodolaresDebe = 0:
montosolesHaber = 0: montodolaresHaber = 0:
If rsaux.BOF = True Or rsaux.EOF = True Then Exit Sub

rsaux.MoveFirst
    While Not rsaux.EOF
    If rsaux("detcomprobconci").Value <> 0 Then
        montosolesDebe = montosolesDebe + vardllgen.ESNULO(rsaux!detcomprobdebe, 0)
        montodolaresDebe = montodolaresDebe + vardllgen.ESNULO(rsaux!detcomprobussdebe, 0)
        montosolesHaber = montosolesHaber + vardllgen.ESNULO(rsaux!detcomprobhaber, 0)
        montodolaresHaber = montodolaresHaber + vardllgen.ESNULO(rsaux!detcomprobusshaber, 0)
    End If
    rsaux.MoveNext
    Wend
    'Soles
    LbTotales(0).Caption = Format(montosolesDebe, "###,###,###,###.00 ") ' Debe
    LbTotales(1).Caption = Format(montosolesHaber, "###,###,###,###.00 ") ' Haber
    'Dolares
    LbTotales(3).Caption = Format(montodolaresDebe, "###,###,###,###.00 ") ' Debe
    LbTotales(4).Caption = Format(montodolaresHaber, "###,###,###,###.00 ") ' Haber
End Sub

Private Sub RsConcil_FieldChangeComplete(ByVal cFields As Long, ByVal Fields As Variant, ByVal pError As ADODB.error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    axbAceptar.Enabled = True
    axBImprimir(0).Enabled = False
    axBImprimir(1).Enabled = False
    Call CalcularTotales(RsConcil)
End Sub

Private Sub axBImprimir_Click(Index As Integer)
    If RsConcil.RecordCount > 0 Then Call imprimir(IIf(Index = 0, -1, 0))
End Sub

Private Sub imprimir(ValorConci As Integer)
Dim arrform(1) As Variant, arrparm(6) As Variant
Dim NombreRep As String, CadOrden As String
Dim mon As String
    arrparm(0) = VGParamSistem.BDEmpresa
    arrparm(1) = VGParamSistem.Anoproceso
    arrparm(2) = Trim(VGParamSistem.Mesproceso)
    arrparm(3) = Trim(CtrAyu_Cuenta.xclave)
    arrparm(4) = ValorConci
    arrparm(5) = Format(CInt((Trim(VGParamSistem.Mesproceso)) - 1), "00")
    
    If ValorConci = -1 Then
        arrform(0) = "@TituloReporte='" & "Conciliación Bancaria - Documentos: Conciliados" & "'"
    Else
        arrform(0) = "@TituloReporte='" & "Conciliación Bancaria - Documentos: Pendientes" & "'"
    End If
    NombreRep = "rptConciliacion.rpt"
    CadOrden = "+{ct_conciliacion_rpt.detcomprobfechaemision},"
    Call ImpresionRptProc(NombreRep, arrform, arrparm, CadOrden, "Conciliación Bancaria")
End Sub

Private Sub TDBG_concil_DblClick()
  If RsConcil.RecordCount > 0 Then
     frmConsultaDetComprobConciliacion.NumeroComprobante = TDBG_concil.Columns(1).Value
     frmConsultaDetComprobConciliacion.Show vbModal
  End If
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
