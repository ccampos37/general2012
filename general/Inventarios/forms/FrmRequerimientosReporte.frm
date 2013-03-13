VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmRequerimientosReportes 
   Caption         =   "Reportes de requerimientos"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5640
   ScaleWidth      =   10665
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   1005
      Left            =   5880
      TabIndex        =   14
      Top             =   240
      Width           =   3120
      Begin VB.CheckBox ChkFech 
         Caption         =   "Rango de Fechas"
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Top             =   -45
         Width           =   1620
      End
      Begin MSComCtl2.DTPicker DTPFechaIni 
         Height          =   300
         Left            =   1260
         TabIndex        =   16
         Top             =   315
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   110034945
         CurrentDate     =   37623.1285069444
      End
      Begin MSComCtl2.DTPicker DTPFechaFin 
         Height          =   300
         Left            =   1260
         TabIndex        =   17
         Top             =   675
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   110034945
         CurrentDate     =   37623.1264351852
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Inicio :"
         Height          =   210
         Index           =   1
         Left            =   150
         TabIndex        =   19
         Top             =   360
         Width           =   1110
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Fin :"
         Height          =   210
         Left            =   315
         TabIndex        =   18
         Top             =   735
         Width           =   810
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Tipo de Reporte"
      Height          =   855
      Left            =   6000
      TabIndex        =   8
      Top             =   1320
      Width           =   3015
      Begin VB.OptionButton Optresumido 
         Caption         =   "Resumido"
         Height          =   375
         Left            =   1440
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Optdetallado 
         Caption         =   "Detallado"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   9240
      TabIndex        =   7
      Top             =   120
      Width           =   1335
      Begin VB.CommandButton cmdreporte 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   240
         Picture         =   "FrmRequerimientosReporte.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   825
      End
      Begin VB.CommandButton Cmdbotones 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   675
         Index           =   12
         Left            =   240
         Picture         =   "FrmRequerimientosReporte.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1200
         Width           =   825
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtro Por"
      Height          =   2055
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   5295
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1560
         TabIndex        =   13
         Text            =   "Combo1"
         Top             =   1440
         Width           =   3615
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_ayutipoorden 
         Height          =   270
         Left            =   1200
         TabIndex        =   2
         Top             =   840
         Width           =   3990
         _ExtentX        =   7038
         _ExtentY        =   476
         XcodMaxLongitud =   11
         xcodwith        =   500
         NomTabla        =   "co_tipodeorden"
         TituloAyuda     =   "Busqueda de Tipo de Orden"
         ListaCampos     =   "tipoordencodigo(1),tipoordendescripcion(1),tipoordennumeracion(2)"
         XcodCampo       =   "tipoordencodigo"
         XListCampo      =   "tipoordendescripcion"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "tipoordencodigo,tipoordendescripcion,tipoordennumeracion"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayusolicitante 
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Top             =   240
         Width           =   3990
         _ExtentX        =   7038
         _ExtentY        =   556
         XcodMaxLongitud =   3
         xcodwith        =   500
         NomTabla        =   "co_solicitantes"
         TituloAyuda     =   "Busqueda de Solicitante"
         ListaCampos     =   "solicitantecodigo(1),solicitantenombre(1)"
         XcodCampo       =   "solicitantecodigo"
         XListCampo      =   "solicitantenombre"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "solicitantecodigo,solicitantenombre"
         Requerido       =   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Estado orden    :"
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   1485
         Width           =   1170
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Solicitante     :"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   285
         Width           =   1005
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Tipo Orden     :"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1035
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   3135
      Left            =   240
      TabIndex        =   0
      Top             =   2280
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   5530
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
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   14933984
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
      DeadAreaBackColor=   14933984
      RowDividerColor =   14933984
      RowSubDividerColor=   14933984
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=15,.bold=0,.fontsize=825,.italic=0"
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
End
Attribute VB_Name = "FrmRequerimientosReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adodc1 As New ADODB.Recordset
Private Sub CtrAyu_solicitante_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Call Mostrar
End Sub


Private Sub ChkFech_Click()
If ChkFech.Value = 1 Then
    DTPFechaIni.Enabled = True
    DTPFechaFin.Enabled = True
  Else
    DTPFechaIni.Enabled = False
    DTPFechaFin.Enabled = False
End If

End Sub

Private Sub cmdBotones_Click(Index As Integer)
Unload Me
End Sub

Private Sub cmdreporte_Click()

Dim arrform(3) As Variant, arrparm(7) As Variant
On Error GoTo Imprime
Screen.MousePointer = 11
arrparm(0) = VGCNx.DefaultDatabase
If Ctr_Ayusolicitante.xclave = "" Then
   arrparm(1) = "%%"
   arrform(0) = "Solicitante='Todos'"
 Else
   arrparm(1) = Ctr_Ayusolicitante.xclave
   arrform(0) = "Solicitante='" & Ctr_Ayusolicitante.xnombre & "'"
End If
If Ctr_AyutipoOrden.xclave = "" Then
   arrparm(2) = "%%"
  Else
   arrparm(2) = Ctr_AyutipoOrden.xclave
End If
If Combo1.ListIndex = 0 Then
   arrparm(3) = "%%"
 Else
   arrparm(3) = Combo1.ListIndex
End If
arrform(2) = "estado='" & Right(Combo1.text, Len(Combo1.text) - 1) & "'"
arrparm(4) = ChkFech.Value
arrparm(5) = DTPFechaIni.Value
arrparm(6) = DTPFechaFin.Value

If ChkFech.Value = 1 Then
   arrform(1) = "Fecha='" & " DEL  " & DTPFechaIni.Value & "  AL  " & DTPFechaFin.Value & "'"

Else
   arrform(1) = "Fecha=''"
End If
If Optdetallado.Value = True Then
   Call ImpresionRptProc("al_relacionrequerimientos.rpt", arrform, arrparm, , "Listado de Requerimientos")
Else
   Call ImpresionRptProc("al_relacionrequerimientosResumen.rpt", arrform, arrparm, , "Listado de Requerimientos")
End If
Screen.MousePointer = 1
Exit Sub
Imprime:
Screen.MousePointer = 1
MsgBox Err.Description
End Sub
Private Sub Combo1_Click()
Call Mostrar
End Sub

Private Sub Ctr_Ayusolicitante_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
  Call Mostrar
End Sub

Private Sub Ctr_AyutipoOrden_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
If Ctr_AyutipoOrden.xclave <> "" Then
  Call Mostrar
End If
End Sub

Private Sub Form_Load()
Call Ctr_Ayusolicitante.conexion(VGCNx)
Call Ctr_AyutipoOrden.conexion(VGCNx)
DTPFechaIni = Date - 60
DTPFechaFin = Date
Optdetallado.Value = True
Call carga_combo
'Combo1.ListIndex = 2

Call Mostrar
End Sub
Private Sub carga_combo()
Dim RSQL As New ADODB.Recordset
SQL = " Select * from co_nivelrequerimiento"
Set RSQL = VGCNx.Execute(SQL)
If RSQL.EOF Then Exit Sub
 If RSQL.BOF Then Exit Sub
 Combo1.AddItem ("0 Todos")
 RSQL.MoveFirst
 Do While Not RSQL.EOF
        SQL = "" & RSQL!nivelrequerimientocodigo & "  " & RSQL!nivelrequerimientodescripcion & ""
        Combo1.AddItem SQL
      RSQL.MoveNext
      If RSQL.EOF Then Exit Do
 Loop
 RSQL.MoveFirst
 Combo1.ListIndex = 0
End Sub

Private Sub Mostrar()
Set VGvardllgen = New dllgeneral.dll_general
SQL = "SELECT a.tipoordencodigo,OC_CNUMORD,OC_DFECDOC,OC_CCODPRO,OC_CRAZSOC,"
SQL = SQL & "OC_DFECENT,estadoocdescripcion , tipoordendescripcion"
SQL = SQL & " FROM co_cabordcompra a inner join co_estadorequerimiento b on a.estadooccodigo= b.estadooccodigo"
SQL = SQL & " inner join co_tipodeorden c on a.tipoordencodigo=c.tipoordencodigo where estadoocatendido<>1 "
SQL = SQL & " and flagrequerimientosPedidos=1 and b.estadoocatendido<>1 "

If Ctr_Ayusolicitante.xclave <> "" Then
   SQL = SQL & " AND OC_CSOLICT='" & Ctr_Ayusolicitante.xclave & "'"
End If
If Ctr_AyutipoOrden.xclave <> "" Then
   SQL = SQL & " and a.TIPOORDENCODIGO='" & Ctr_AyutipoOrden.xclave & "'"
End If
If Left(Combo1.text, 1) <> "0" Then
   SQL = SQL & " and b.nivelrequerimientoCODIGO ='" & Left(Combo1.ListIndex, 1) & "'"
End If
SQL = SQL & " ORDER BY oc_cnumord "
Set adodc1 = VGCNx.Execute(SQL)
Set TDBGrid1.DataSource = adodc1

End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
 With adodc1
    If .Sort = Empty Then
        .Sort = TDBGrid1.Columns.item(ColIndex).DataField & " asc"
    ElseIf Right(.Sort, 3) = "asc" Then
        .Sort = TDBGrid1.Columns.item(ColIndex).DataField & " desc"
    ElseIf Right(.Sort, 4) = "desc" Then
        .Sort = TDBGrid1.Columns.item(ColIndex).DataField & " asc"
    End If
    TDBGrid1.Refresh
 End With
End Sub
