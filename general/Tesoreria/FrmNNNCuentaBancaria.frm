VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "CONTROLAYUDA.OCX"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TEXTFER.OCX"
Begin VB.Form FrmNNNCuentaBancaria 
   Caption         =   "Cuentas Bancarias"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6075
   ScaleWidth      =   7830
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmbotones 
      Height          =   555
      Left            =   1185
      TabIndex        =   9
      Top             =   5445
      Width           =   5715
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Imprimir"
         Height          =   330
         Index           =   3
         Left            =   3435
         TabIndex        =   14
         Top             =   165
         Width           =   1080
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Salir"
         Height          =   330
         Index           =   4
         Left            =   4560
         TabIndex        =   13
         Top             =   165
         Width           =   1080
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Eliminar"
         Height          =   330
         Index           =   2
         Left            =   2310
         TabIndex        =   12
         Top             =   165
         Width           =   1080
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "E&ditar"
         Height          =   330
         Index           =   1
         Left            =   1185
         TabIndex        =   11
         Top             =   165
         Width           =   1080
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Nuevo"
         Height          =   330
         Index           =   0
         Left            =   60
         TabIndex        =   10
         Top             =   165
         Width           =   1080
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5010
      Left            =   165
      TabIndex        =   7
      Top             =   270
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   8837
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      WordWrap        =   0   'False
      TabCaption(0)   =   "Consulta"
      TabPicture(0)   =   "FrmNNNCuentaBancaria.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "TDBGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Mantenimiento"
      TabPicture(1)   =   "FrmNNNCuentaBancaria.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Ctr_Mon"
      Tab(1).Control(1)=   "Txt(0)"
      Tab(1).Control(2)=   "Ctr_Banco"
      Tab(1).Control(3)=   "cAcepta"
      Tab(1).Control(4)=   "cCancela"
      Tab(1).Control(5)=   "Txt(1)"
      Tab(1).Control(6)=   "Txt(2)"
      Tab(1).Control(7)=   "Txt(3)"
      Tab(1).Control(8)=   "Txt(4)"
      Tab(1).Control(9)=   "Label6"
      Tab(1).Control(10)=   "Label5"
      Tab(1).Control(11)=   "Label4"
      Tab(1).Control(12)=   "Label3"
      Tab(1).Control(13)=   "Label2"
      Tab(1).Control(14)=   "Label1(1)"
      Tab(1).Control(15)=   "Label1(0)"
      Tab(1).ControlCount=   16
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Mon 
         Height          =   300
         Left            =   -72495
         TabIndex        =   1
         Top             =   1200
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   529
         XcodMaxLongitud =   0
         xcodwith        =   600
         NomTabla        =   "gr_moneda"
         ListaCampos     =   "monedacodigo(1),monedadescripcion(1)"
         XcodCampo       =   "monedacodigo"
         XListCampo      =   "monedadescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "monedacodigo,monedadescripcion"
         Requerido       =   0   'False
      End
      Begin TextFer.TxFer Txt 
         Height          =   300
         Index           =   0
         Left            =   -72495
         TabIndex        =   2
         Top             =   1695
         Width           =   4230
         _ExtentX        =   7461
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
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Banco 
         Height          =   315
         Left            =   -72495
         TabIndex        =   0
         Top             =   795
         Width           =   4230
         _ExtentX        =   7461
         _ExtentY        =   556
         XcodMaxLongitud =   0
         xcodwith        =   600
         NomTabla        =   "gr_banco"
         ListaCampos     =   "bancocodigo(1),bancodescripcion(1)"
         XcodCampo       =   "bancocodigo"
         XListCampo      =   "bancodescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "bancocodigo,bancodescripcion"
         Requerido       =   0   'False
      End
      Begin VB.CommandButton cAcepta 
         Caption         =   "&Aceptar"
         Height          =   360
         Left            =   -72780
         TabIndex        =   16
         Top             =   4425
         Width           =   1245
      End
      Begin VB.CommandButton cCancela 
         Caption         =   "&Cancelar"
         Height          =   360
         Left            =   -71325
         TabIndex        =   15
         Top             =   4425
         Width           =   1245
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   4350
         Left            =   105
         TabIndex        =   8
         Top             =   510
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   7673
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
      Begin TextFer.TxFer Txt 
         Height          =   300
         Index           =   1
         Left            =   -72495
         TabIndex        =   3
         Top             =   2085
         Width           =   4230
         _ExtentX        =   7461
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
      Begin TextFer.TxFer Txt 
         Height          =   300
         Index           =   2
         Left            =   -72495
         TabIndex        =   4
         Top             =   2520
         Width           =   4230
         _ExtentX        =   7461
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
      Begin TextFer.TxFer Txt 
         Height          =   300
         Index           =   3
         Left            =   -72495
         TabIndex        =   5
         Top             =   2955
         Width           =   4230
         _ExtentX        =   7461
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
      Begin TextFer.TxFer Txt 
         Height          =   300
         Index           =   4
         Left            =   -72495
         TabIndex        =   6
         Top             =   3345
         Width           =   4230
         _ExtentX        =   7461
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
      Begin VB.Label Label6 
         Caption         =   "Cuenta Analitica"
         Height          =   225
         Left            =   -74670
         TabIndex        =   23
         Top             =   3375
         Width           =   1860
      End
      Begin VB.Label Label5 
         Caption         =   "Cuenta Contable"
         Height          =   270
         Left            =   -74670
         TabIndex        =   22
         Top             =   2955
         Width           =   1860
      End
      Begin VB.Label Label4 
         Caption         =   "Número de Cheque"
         Height          =   285
         Left            =   -74670
         TabIndex        =   21
         Top             =   2580
         Width           =   1860
      End
      Begin VB.Label Label3 
         Caption         =   "Referencia"
         Height          =   300
         Left            =   -74670
         TabIndex        =   20
         Top             =   2145
         Width           =   1860
      End
      Begin VB.Label Label2 
         Caption         =   "Cuenta"
         Height          =   285
         Left            =   -74670
         TabIndex        =   19
         Top             =   1725
         Width           =   1860
      End
      Begin VB.Label Label1 
         Caption         =   "Moneda"
         Height          =   300
         Index           =   1
         Left            =   -74670
         TabIndex        =   18
         Top             =   1215
         Width           =   1860
      End
      Begin VB.Label Label1 
         Caption         =   "Banco"
         Height          =   300
         Index           =   0
         Left            =   -74670
         TabIndex        =   17
         Top             =   855
         Width           =   1860
      End
   End
End
Attribute VB_Name = "FrmNNNCuentaBancaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim modoinsert As Boolean
Dim modoedit As Boolean
Dim i_filaorigen As Integer
Dim rs As New ADODB.Recordset
Dim FLAG_CHECK As Boolean

Private Sub Form_Load()
  Call Ctr_Banco.Conexion(cn)
  Call Ctr_Mon.Conexion(cn)
  Call ConfiguraForm
  Call MuestraDatos
 
End Sub

Sub ConfiguraForm()
  SSTab1.Tab = 0
  SSTab1.TabEnabled(1) = False
  cAcepta.Enabled = False
  FLAG_CHECK = False
  Me.Width = 7860
  Me.Height = 6795

End Sub

Private Function MuestraDatos()
 Dim SQL As String
  SQL = "select cbanco_codigo as Cod_Banco,monedacodigo as Moneda,cbanco_numero as Cuenta,cbanco_referenciacta as Referencia,cbanco_nrocheque,cbanco_cuenta,cbanco_analitico from te_cuentabancos"
  Set rs = cn.Execute(SQL)
  Set TDBGrid1.DataSource = rs
  Call ConfiguraTdbgrid
  SSTab1.Tab = 0
End Function

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
  Dim j As Integer
  Dim spos As Integer
  Dim SQL As String
  
  On Error GoTo X
  SSTab1.TabEnabled(1) = True
  
  Select Case Index
     Case 0   'nuevo
        modoinsert = True
        frmbotones.Visible = False
        SSTab1.Tab = 1
        Call LimpiarValores
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
        cAcepta.Enabled = True
      
     Case 2   'eliminar
       If MsgBox("Desea eliminar el Registro de la Cuenta Bancaria " & Trim(TDBGrid1.Columns(2).Value), vbYesNo + vbDefaultButton2, "AVISO") = vbYes Then
          Call EliminaCuenta
          Call MuestraDatos
       End If
        
     Case 3   'imprimir
       'Call Impresion("RepvtMantBanc.rpt")
     
     Case 4  ' salir
       Unload Me
  End Select
  
  Exit Sub
   
X:
  If Index = 2 And Err.Number = -2147217873 Then
    MsgBox "Registro no podrá Eliminarse mientras exista Información en la Tablas Relacionadas", vbInformation, Caption
    cn.RollbackTrans
  Else
    MsgBox "Error inesperado: " & Err.Description & "  " & Err.Number, vbInformation, Caption
  End If
   
End Sub

Sub EditarValores()
 Dim I As Integer
  With TDBGrid1
    Ctr_Banco.xclave = Trim(.Columns(0).Text): Ctr_Banco.Ejecutar
    Ctr_Mon.xclave = Trim(.Columns(1).Text): Ctr_Mon.Ejecutar
    For I = 0 To 4
      txt(I).Text = Escadena((.Columns(2 + I).Text))
    Next
  End With
End Sub

Public Function LimpiarValores()
 Dim I As Integer
  
  Ctr_Banco.xclave = Empty: Ctr_Banco.Ejecutar
  Ctr_Mon.xclave = Empty: Ctr_Mon.Ejecutar
  For I = 0 To 4
    txt(I).Text = Empty
  Next
  
End Function

Private Sub cAcepta_Click()
 If ValidaData() = True Then
    Call GrabaData
 End If
     
End Sub

Function ValidaData() As Boolean
 Dim rsX As ADODB.Recordset
 Dim SQL As String
 Dim I As Integer
   
    If Ctr_Banco.xclave = Empty Then
       MsgBox "Debe Registrar el Código de Banco", vbInformation, Caption
       ValidaData = False
       Ctr_Banco.SetFocus
       Exit Function
    End If
   
    If Ctr_Mon.xclave = Empty Then
       MsgBox "Debe Registrar la Moneda", vbInformation, Caption
       ValidaData = False
       Ctr_Mon.SetFocus
       Exit Function
    End If
   
    If txt(0).Text = Empty Then
        MsgBox "Debe Registrar el Nº de Cuenta", vbInformation, Caption
        ValidaData = False
        txt(0).SetFocus
        Exit Function
    End If

  ValidaData = True
End Function

Sub GrabaData()
  Dim xVarCbo As String
  Dim SQL As String
  On Error GoTo X
  
  SSTab1.TabEnabled(0) = True
  
  If modoinsert = True Then
     SQL = "INSERT  te_cuentabancos (cbanco_codigo,monedacodigo,cbanco_numero,cbanco_referenciacta,cbanco_nrocheque,cbanco_cuenta,cbanco_analitico,usuariocodigo,fechaact) "
     SQL = SQL & "VALUES ('" & Trim(Ctr_Banco.xclave) & "','" & Trim(Ctr_Mon.xclave) & "','" & Trim(txt(0).Text) & "','" & txt(1).Text & "','" & txt(2).Text & "','" & txt(3).Text & "','" & txt(4).Text & "','" & VGParamSistem.Usuario & "'," & Format(Now, "dd/mm/yyyy") & ")"
     cn.BeginTrans
     cn.Execute (SQL)
     cn.CommitTrans
  ElseIf modoedit = True Then
     SQL = "UPDATE te_cuentabancos SET cbanco_codigo='" & Trim(Ctr_Banco.xclave) & "',"
     SQL = SQL & "monedacodigo='" & Trim(Ctr_Mon.xclave) & "',cbanco_numero='" & Trim(txt(0).Text) & "',"
     SQL = SQL & "cbanco_referenciacta='" & txt(1).Text & "',cbanco_nrocheque='" & txt(2).Text & "',"
     SQL = SQL & "cbanco_cuenta='" & txt(3).Text & "',cbanco_analitico='" & txt(4).Text & "',"
     SQL = SQL & "usuariocodigo='" & VGParamSistem.Usuario & "',fechaact=getdate() "
     SQL = SQL & "WHERE cbanco_codigo='" & Trim(Ctr_Banco.xclave) & "' and "
     SQL = SQL & "monedacodigo='" & Trim(Ctr_Mon.xclave) & "' AND cbanco_numero='" & Trim(txt(0).Text) & "'"
     cn.BeginTrans
     cn.Execute (SQL)
     cn.CommitTrans
  End If

  Call MuestraDatos
  frmbotones.Visible = True
  modoinsert = False: modoedit = False: FLAG_CHECK = False
  i_filaorigen = -1
  Exit Sub

X:
  If Err.Number = -2147217873 Then
    MsgBox "Esta intentando registrar Cuenta Bancaria Existente " & Err.Description, vbInformation, Caption
  Else
    MsgBox "Error inesperado: " & Err.Number & " " & Err.Description
  End If
  cn.RollbackTrans

End Sub

Sub EliminaCuenta()
 Dim SQL As String
   SQL = "DELETE FROM te_cuentabancos "
   SQL = SQL & "WHERE cbanco_codigo='" & Trim(Ctr_Banco.xclave) & "' and "
   SQL = SQL & "monedacodigo='" & Trim(Ctr_Mon.xclave) & "' AND cbanco_numero='" & Trim(txt(0).Text) & "'"
   cn.Execute (SQL)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set rs = Nothing
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
  With TDBGrid1
    .Columns(0).Width = 1000
    .Columns(1).Width = 1000
    .Columns(2).Width = 2700
    .Columns(3).Width = 2200
    .Columns(4).Visible = False
    .Columns(5).Visible = False
    .Columns(6).Visible = False
  End With
  
End Sub

Function ValidaDataIngreso() As Boolean
 Dim I As Integer
  For I = 0 To 4
   If txt(I).Text = Empty Then
     ValidaDataIngreso = False
     Exit Function
   End If
  Next

  ValidaDataIngreso = True
End Function

Private Sub txt_Change(Index As Integer)
  cAcepta.Enabled = ValidaDataIngreso()
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 And Index = 15 Then
    cAcepta.SetFocus
    Call cAcepta_Click
  End If
End Sub
