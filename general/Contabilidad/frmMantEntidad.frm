VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Begin VB.Form frmMantEntidad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entidad (Analìtico)"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   8250
   StartUpPosition =   2  'CenterScreen
   Begin TextFer.TxFer txtbuscar 
      Height          =   300
      Left            =   3525
      TabIndex        =   26
      Top             =   15
      Width           =   4185
      _ExtentX        =   7382
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
      Valor           =   ""
   End
   Begin VB.Frame frmbotones 
      Height          =   555
      Left            =   1028
      TabIndex        =   8
      Top             =   5235
      Width           =   5715
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Nuevo"
         Height          =   330
         Index           =   0
         Left            =   60
         TabIndex        =   13
         Top             =   165
         Width           =   1080
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "E&ditar"
         Height          =   330
         Index           =   1
         Left            =   1185
         TabIndex        =   12
         Top             =   165
         Width           =   1080
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Eliminar"
         Height          =   330
         Index           =   2
         Left            =   2310
         TabIndex        =   11
         Top             =   165
         Width           =   1080
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Salir"
         Height          =   330
         Index           =   4
         Left            =   4560
         TabIndex        =   10
         Top             =   165
         Width           =   1080
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Imprimir"
         Height          =   330
         Index           =   3
         Left            =   3435
         TabIndex        =   9
         Top             =   165
         Width           =   1080
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   -15
      TabIndex        =   14
      Top             =   45
      Width           =   8070
      _ExtentX        =   14235
      _ExtentY        =   9128
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Consulta"
      TabPicture(0)   =   "frmMantEntidad.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblNumReg"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "TDBGrid1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Mantenimiento"
      TabPicture(1)   =   "frmMantEntidad.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "cCancela"
      Tab(1).Control(2)=   "cAcepta"
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame1 
         Height          =   4710
         Left            =   -75000
         TabIndex        =   15
         Top             =   330
         Width           =   6555
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Caption         =   "Cerrado/Anulado/Suspendido"
            Height          =   495
            Left            =   120
            TabIndex        =   27
            Top             =   2280
            Width           =   2775
         End
         Begin VB.ComboBox cboTipoCont 
            Height          =   315
            ItemData        =   "frmMantEntidad.frx":0038
            Left            =   2730
            List            =   "frmMantEntidad.frx":0045
            TabIndex        =   5
            Top             =   1845
            Width           =   3795
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1590
            Left            =   30
            TabIndex        =   24
            Top             =   2955
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   2805
            View            =   2
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin TextFer.TxFer txt 
            Height          =   300
            Index           =   0
            Left            =   2715
            TabIndex        =   0
            Top             =   255
            Width           =   3765
            _ExtentX        =   6641
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
            MaxLength       =   11
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            SignoNegativo   =   0   'False
            MarcarTextoAlEnfoque=   -1  'True
         End
         Begin TextFer.TxFer txt 
            Height          =   300
            Index           =   1
            Left            =   2715
            TabIndex        =   2
            Top             =   885
            Width           =   3780
            _ExtentX        =   6668
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
            MaxLength       =   40
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            NoCaracteres    =   "',"
            MarcarTextoAlEnfoque=   -1  'True
         End
         Begin TextFer.TxFer txt 
            Height          =   300
            Index           =   2
            Left            =   2715
            TabIndex        =   3
            Top             =   1215
            Width           =   3780
            _ExtentX        =   6668
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
            MaxLength       =   25
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            NoCaracteres    =   "',"
            MarcarTextoAlEnfoque=   -1  'True
         End
         Begin TextFer.TxFer txt 
            Height          =   300
            Index           =   3
            Left            =   2715
            TabIndex        =   4
            Top             =   1530
            Width           =   3780
            _ExtentX        =   6668
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
            MaxLength       =   15
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            NoCaracteres    =   "',"
            MarcarTextoAlEnfoque=   -1  'True
         End
         Begin TextFer.TxFer txt 
            Height          =   300
            Index           =   4
            Left            =   2715
            TabIndex        =   1
            Top             =   570
            Width           =   3780
            _ExtentX        =   6668
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
            MaxLength       =   11
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            TipoDato        =   1
            SignoNegativo   =   0   'False
            MarcarTextoAlEnfoque=   -1  'True
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo Contribuyente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   25
            Top             =   1890
            Width           =   2310
         End
         Begin VB.Label lbl 
            Caption         =   "Razón Social"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   19
            Top             =   930
            Width           =   2310
         End
         Begin VB.Label lbl 
            Caption         =   "Dirección"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   18
            Top             =   1230
            Width           =   2310
         End
         Begin VB.Label lbl 
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
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   17
            Top             =   1530
            Width           =   2310
         End
         Begin VB.Label lbl 
            Caption         =   "RUC"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   16
            Left            =   135
            TabIndex        =   16
            Top             =   630
            Width           =   2385
         End
         Begin VB.Label lbl 
            Caption         =   "Código"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   20
            Top             =   315
            Width           =   840
         End
      End
      Begin VB.CommandButton cCancela 
         Caption         =   "&Cancelar"
         Height          =   345
         Left            =   -68280
         TabIndex        =   7
         Top             =   3840
         Width           =   1140
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   4110
         Left            =   45
         TabIndex        =   21
         Top             =   360
         Width           =   7605
         _ExtentX        =   13414
         _ExtentY        =   7250
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
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=64,.bold=0,.fontsize=825,.italic=0"
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
      Begin VB.CommandButton cAcepta 
         Caption         =   "&Aceptar"
         Height          =   345
         Left            =   -68280
         TabIndex        =   6
         Top             =   3240
         Width           =   1140
      End
      Begin VB.Label lblNumReg 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblNumReg"
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   5685
         TabIndex        =   23
         Top             =   5460
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Registros"
         Height          =   270
         Left            =   4740
         TabIndex        =   22
         Top             =   5475
         Width           =   900
      End
   End
End
Attribute VB_Name = "frmMantEntidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Dim modoinsert As Boolean
Dim modoedit As Boolean
Dim i_filaorigen As Integer
Dim rs As New ADODB.Recordset
Dim FLAG_CHECK As Boolean

Private Sub Form_Load()
  Call ConfiguraForm
  Call MuestraDatos("%")
  Call CargaLista
End Sub

Sub ConfiguraForm()
  SSTab1.Tab = 0
  SSTab1.TabEnabled(1) = False
  cAcepta.Enabled = False
  lblNumReg.Caption = Empty
  FLAG_CHECK = False
End Sub

'FIXIT: Declare 'MuestraDatos' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Private Function MuestraDatos(xCod As String)
 Dim SQL As String
  SQL = "SELECT entidadcodigo,entidadrazonsocial,entidaddireccion,entidadtelefono,entidadruc,entidadtipocontri,proyectocierre From ct_entidad "
'  SQL = SQL & "  AND WHERE entidadcodigo<>'00' "
  
  If IsNumeric(xCod) = True Then
    SQL = SQL & " WHERE entidadcodigo LIKE '" & xCod & "%'"
  Else
    SQL = SQL & " WHERE  entidadrazonsocial LIKE '" & xCod & "%'"
  End If
  Set rs = VGCNx.Execute(SQL)
  Set TDBGrid1.DataSource = rs
  Call ConfiguraTdbgrid
  lblNumReg.Caption = rs.RecordCount
  SSTab1.Tab = 0
End Function

Sub CargaLista()
   ' Declara una variable para agregar objetos ListItem.
  Dim rsX As ADODB.Recordset
  Dim itmX As ListItem
  Dim SQL As String
  Set rsX = New ADODB.Recordset
  
   ListView1.ColumnHeaders.Clear
   ListView1.ListItems.Clear
   ListView1.ColumnHeaders.Add , , "Codigo Tipo Analítico", ListView1.Width / 3
   ListView1.ColumnHeaders.Add , , "Descripción Tipo Analítico", ListView1.Width * 2 / 3, lvwColumnCenter
   ListView1.View = lvwReport
   
   SQL = "select tipoanaliticocodigo,tipoanaliticodescripcion FROM ct_tipoanalitico WHERE tipoanaliticocodigo<>'00'"
   Set rsX = VGCNx.Execute(SQL)
   While Not rsX.EOF
     Set itmX = ListView1.ListItems.Add(, , rsX(0))
     itmX.SubItems(1) = rsX(1)
     rsX.MoveNext
   Wend
   Set rsX = Nothing

End Sub

Private Sub cCancela_Click()
  SSTab1.TabEnabled(0) = True
  SSTab1.Tab = 0
  SSTab1.SetFocus
  frmbotones.Visible = True
  modoinsert = False
  modoedit = False
  i_filaorigen = -1

End Sub

Private Sub cmdBotones_Click(Index As Integer)
  Dim J As Integer
  Dim spos As Integer
  Dim SQL As String
  
  On Error GoTo x
  SSTab1.TabEnabled(1) = True
  
  Select Case Index
     Case 0   'nuevo
        modoinsert = True
        frmbotones.Visible = False
        SSTab1.Tab = 1
        Call LimpiarValores
        Call CargaLista
        txt(0).Enabled = True
        txt(0).SetFocus
        
     Case 1   'editar
        If TDBGrid1.Row < 0 Then
          Exit Sub
        End If
        modoedit = True
        frmbotones.Visible = False
        SSTab1.Tab = 1
        Call EditarValores
        Call MuestraCheckTipoAnalitico
        cAcepta.Enabled = False
        txt(0).Enabled = False
      
     Case 2   'eliminar
       If MsgBox("Desea eliminar el registro de Código Entidad " & Trim$(TDBGrid1.Columns(0).Value), vbYesNo + vbDefaultButton2, "AVISO") = vbYes Then
          VGCNx.BeginTrans
          SQL = "DELETE FROM ct_analitico WHERE entidadcodigo='" & Trim$(TDBGrid1.Columns(0).Value) & "'"
          VGCNx.Execute (SQL)
          SQL = "DELETE FROM ct_entidad WHERE entidadcodigo='" & Trim$(TDBGrid1.Columns(0).Value) & "'"
          VGCNx.Execute (SQL)
          VGCNx.CommitTrans
          Call MuestraDatos(txtBuscar.Text)
       End If
        
     Case 3   'imprimir
       Call Impresion("rptEntidad.rpt")
     
     Case 4  ' salir
       Unload Me
  End Select
  
  Exit Sub
   
x:
  If Index = 2 And err.Number = -2147217873 Then
    MsgBox "Registro no podrá Eliminarse mientras exista Información en la Tablas Relacionadas", vbInformation, Caption
    VGCNx.RollbackTrans
  Else
    MsgBox "Error inesperado: " & err.Description & "  " & err.Number, vbInformation, Caption
  End If
   
End Sub

Sub EditarValores()
 Dim i As Integer
  With TDBGrid1
    For i = 0 To 4
      txt(i).Text = Trim$(.Columns(i).Text)
    Next

    If Trim$(.Columns(5).Text) <> Empty Then
      cboTipoCont.Text = cboTipoCont.List(CInt(.Columns(5).Text) - 1)
    Else
      cboTipoCont.Text = Empty
    End If
    Check1.Value = .Columns(6).Value
  End With
End Sub

'FIXIT: Declare 'LimpiarValores' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Public Function LimpiarValores()
 Dim i As Integer
  For i = 0 To 4
    txt(i).Text = Empty
  Next
  cboTipoCont.Text = Empty
End Function

Private Sub cAcepta_Click()
 If ValidaData() = True Then
    Call GrabaData
 End If
     
End Sub

Function ValidaData() As Boolean
 Dim rsX As ADODB.Recordset
 Dim SQL As String
 Dim i As Integer
   
   
    If txt(0).Text = Empty Then
        MsgBox "Debe Registrar el Código de Analítico", vbInformation, Caption
        ValidaData = False
        txt(4).SetFocus
        Exit Function
    End If
   
'    If Len(Trim$(txt(0).Text)) <> 11 And Trim$(txt(0).Text) <> "00" Then
'        MsgBox "El Código de Analítico debe tener 11 caracteres", vbInformation, Caption
'        ValidaData = False
'        txt(4).SetFocus
'        Exit Function
'    End If
   
'    If txt(4).Text = Empty Then
'        MsgBox "Debe Registrar el Nº de RUC", vbInformation, Caption
'        ValidaData = False
'        txt(4).SetFocus
'        Exit Function
'    End If
   
   For i = 1 To ListView1.ListItems.Count
      If ListView1.ListItems.Item(i).Checked = True Then
         ValidaData = True
         Exit For
      Else
         ValidaData = False
         If i = ListView1.ListItems.Count Then
           MsgBox "Falta Completar el Tipo de Analítico", vbInformation, Caption
           Exit Function
         End If
      End If
   Next
   
 '  If txt(4).Text <> Empty And Len(Trim$(txt(4).Text)) <> 11 Then
 '    MsgBox "El número de RUC tiene 11 dígitos", vbInformation, Caption
 '    ValidaData = False
 '    txt(4).SetFocus
 '    Exit Function
 '  End If
   
   If txt(4).Text <> Empty Then
      SQL = "SELECT count(entidadruc) FROM ct_entidad WHERE entidadruc='" & txt(4).Text & "'"
      Set rsX = New ADODB.Recordset
      Set rsX = VGCNx.Execute(SQL)
      Set VGvardllgen = New dllgeneral.dll_general
      
      If modoedit = True And VGvardllgen.ESNULO(rsX(0), 0) = 1 Then
        ValidaData = True
        SQL = "SELECT count(entidadruc) FROM ct_entidad WHERE entidadruc='" & txt(4).Text & "'"
        SQL = SQL & " AND entidadcodigo<>'" & txt(0).Text & "'"
        Set rsX = VGCNx.Execute(SQL)
        If VGvardllgen.ESNULO(rsX(0), 0) > 0 Then
           MsgBox "Esta intentando registrar un numero de RUC existente", vbInformation, Caption
           ValidaData = False
           txt(4).SetFocus
           Exit Function
        End If
      Else
        If VGvardllgen.ESNULO(rsX(0), 0) > 0 Then
           MsgBox "Esta intentando registrar un número de RUC existente", vbInformation, Caption
           ValidaData = False
           txt(4).SetFocus
           Exit Function
        Else
          ValidaData = True
        End If
      End If
   End If

  ValidaData = True
End Function

Sub GrabaData()
  Dim xVarCbo As String
  Dim SQL As String
  On Error GoTo x
  
  SSTab1.TabEnabled(0) = True
  
  xVarCbo = Trim$(Left(cboTipoCont.List(cboTipoCont.ListIndex), 2))
  
  If modoinsert = True Then
    SQL = "INSERT CT_ENTIDAD(entidadcodigo,entidadrazonsocial,entidaddireccion,entidadtelefono,entidadruc,entidadtipocontri,usuariocodigo,fechaact) "
    SQL = SQL & "VALUES ('" & UCase$(txt(0).Text) & "','" & UCase$(txt(1).Text) & "','" & UCase$(txt(2).Text) & "','" & txt(3).Text & "','" & txt(4).Text & "','" & xVarCbo & "','" & VGUsuario & "','" & Date & "')"
    VGCNx.BeginTrans
    VGCNx.Execute (SQL)
    Call GrabaCheckTipoAnalitico
    VGCNx.CommitTrans
                  
  ElseIf modoedit = True Then
    SQL = "UPDATE CT_ENTIDAD SET entidadrazonsocial='" & Trim$(UCase$(txt(1).Text)) & "',"
    SQL = SQL & "entidaddireccion='" & Trim$(UCase$(txt(2).Text)) & "',"
    SQL = SQL & "entidadtelefono='" & txt(3).Text & "',"
    SQL = SQL & "entidadruc='" & txt(4).Text & "',"
    SQL = SQL & "entidadtipocontri='" & xVarCbo & "',"
    SQL = SQL & "proyectocierre='" & Check1.Value & "',"
    SQL = SQL & "usuariocodigo='" & VGUsuario & "',fechaact='" & Format(Date, "dd/mm/yyyy") & "' "
    SQL = SQL & "WHERE entidadcodigo='" & txt(0).Text & "'"
    VGCNx.BeginTrans
    VGCNx.Execute (SQL)
    
    If FLAG_CHECK = True Then
      'Call DeleteCheckTipoAnalitico
      Call GrabaCheckTipoAnalitico
    End If
    
    VGCNx.CommitTrans
    
  End If
  
  Call MuestraDatos(txtBuscar.Text)
  frmbotones.Visible = True
  modoinsert = False: modoedit = False: FLAG_CHECK = False
  i_filaorigen = -1
  Exit Sub

x:
  If err.Number = -2147217873 Then
    MsgBox "Esta intentando registrar Código Analítico Existente " & err.Description, vbInformation, Caption
  Else
    MsgBox "Error inesperado: " & err.Number & " " & err.Description
  End If
  VGCNx.RollbackTrans

End Sub

Sub MuestraCheckTipoAnalitico()
 Dim rsX As ADODB.Recordset
 Dim i As Long
 Dim SQL As String
 SQL = "select tipoanaliticocodigo,analiticocodigo FROM ct_analitico WHERE entidadcodigo='" & txt(4).Text & "'"
 Set rsX = VGCNx.Execute(SQL)
 
 While Not rsX.EOF
   For i = 1 To ListView1.ListItems.Count
     If ListView1.ListItems.Item(i).Text = rsX(0) Then
       ListView1.ListItems.Item(i).Checked = True
     Else
       ListView1.ListItems.Item(i).Checked = False
     End If
   Next
   rsX.MoveNext
 Wend
 Set rsX = Nothing

End Sub

Sub DeleteCheckTipoAnalitico()
 Dim SQL As String
  SQL = "DELETE FROM ct_analitico WHERE entidadcodigo='" & txt(0).Text & "'"
  VGCNx.Execute (SQL)
End Sub

Sub GrabaCheckTipoAnalitico()
 Dim SQL As String
 Dim i As Long
 Dim xCodAnalitico As String
 Dim rsX As New ADODB.Recordset
   Set rsX = New ADODB.Recordset
   For i = 1 To ListView1.ListItems.Count
      If ListView1.ListItems.Item(i).Checked = True Then
         xCodAnalitico = Trim$(txt(0).Text) & Trim$(ListView1.ListItems.Item(i).Text)
         SQL = "select count(*) from ct_analitico where analiticocodigo='" & xCodAnalitico & "'"
         Set rsX = VGCNx.Execute(SQL)
         If rsX(0) = 0 Then
           SQL = "INSERT ct_analitico (analiticocodigo,entidadcodigo,tipoanaliticocodigo,usuariocodigo,fechaact) "
           SQL = SQL & "VALUES ('" & xCodAnalitico & "','" & Trim$(txt(0).Text) & "','" & Trim$(ListView1.ListItems.Item(i).Text) & "','" & VGUsuario & "','" & Date & "')"
           VGCNx.Execute (SQL)
         End If
     End If
   Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set rs = Nothing
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ColumnHeader)
  ListView1.SortKey = ColumnHeader.Index - 1
  ListView1.Sorted = True
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
  FLAG_CHECK = True
  cAcepta.Enabled = ValidaDataIngreso()
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
  SSTab1.TabEnabled(PreviousTab) = False
  cAcepta.Enabled = False
  If PreviousTab = 0 Then
        txtBuscar.Visible = False
  Else
        txtBuscar.Visible = True
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

Private Sub TDBGrid1_DblClick()
 If rs.RecordCount > 0 And (modoedit = False And modoinsert = False) Then
   Call cmdBotones_Click(1)
 End If
End Sub

Private Sub ConfiguraTdbgrid()
  TDBGrid1.Columns(0).Width = 1100
  TDBGrid1.Columns(1).Width = 3500
  TDBGrid1.Columns(2).Width = 1800
  TDBGrid1.Columns(3).Width = 900
  TDBGrid1.Columns(4).Width = 1100

End Sub

Function ValidaDataIngreso() As Boolean
 Dim i As Integer
  For i = 0 To 1
   If txt(i).Text = Empty Then
     ValidaDataIngreso = False
     Exit Function
   End If
  Next

  ValidaDataIngreso = True
End Function

Private Sub txt_Change(Index As Integer)
  cAcepta.Enabled = ValidaDataIngreso()
End Sub

Private Sub txt_LostFocus(Index As Integer)
  txt(Index).Text = UCase$(txt(Index).Text)
  If Index = 0 And modoedit = False Then txt(4).Text = txt(0).Text
End Sub

Private Sub cboTipoCont_Click()
  cAcepta.Enabled = ValidaDataIngreso()
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 And Index = 15 Then
    cAcepta.SetFocus
    Call cAcepta_Click
  End If
End Sub

Private Sub txtBuscar_Change()
  Call MuestraDatos(txtBuscar.Text)
End Sub
