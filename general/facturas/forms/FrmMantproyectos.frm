VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmMntProyectos 
   ClientHeight    =   6405
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8385
   LinkTopic       =   "Form2"
   ScaleHeight     =   6405
   ScaleWidth      =   8385
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Framegraba 
      Height          =   1215
      Left            =   5760
      TabIndex        =   26
      Top             =   5040
      Width           =   2415
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&Grabar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   240
         Width           =   870
      End
      Begin VB.CommandButton Cmdsalirgraba 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   1335
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   240
         Width           =   870
      End
   End
   Begin VB.Frame Framemodifica 
      Height          =   1215
      Left            =   360
      TabIndex        =   12
      Top             =   5040
      Width           =   7815
      Begin VB.CommandButton CmdModi 
         Caption         =   "&Modificar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   1185
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   870
      End
      Begin VB.CommandButton CmdEli 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   2250
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   870
      End
      Begin VB.CommandButton CmdFicha 
         Caption         =   "&imprimir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   870
      End
      Begin VB.CommandButton CmdIng 
         Caption         =   "&Crear"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   870
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   870
      End
   End
   Begin VB.Frame FrameGrupo 
      Caption         =   "Grupo"
      Height          =   1695
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin VB.CheckBox CheckTodos 
         Caption         =   "Todos"
         Height          =   195
         Left            =   6600
         TabIndex        =   32
         Top             =   1440
         Width           =   855
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayufamilia 
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   661
         XcodMaxLongitud =   10
         xcodwith        =   450
         NomTabla        =   "familia"
         ListaCampos     =   "FAM_CODIGO(1),FAM_NOMBRE(1)"
         XcodCampo       =   "FAM_CODIGO"
         XListCampo      =   "FAM_NOMBRE"
         ListaCamposDescrip=   "codigo, descripcion"
         ListaCamposText =   "FAM_CODIGO,FAM_NOMBRE"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayulinea 
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   720
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   661
         XcodMaxLongitud =   10
         xcodwith        =   450
         NomTabla        =   "lineas"
         ListaCampos     =   "lin_codigo(1), lin_nombre(1)"
         XcodCampo       =   "lin_codigo"
         XListCampo      =   "lin_nombre"
         ListaCamposDescrip=   "codigo, descripcion"
         ListaCamposText =   "lin_codigo, lin_nombre"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayugrupo 
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   1200
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   661
         XcodMaxLongitud =   10
         xcodwith        =   450
         NomTabla        =   "grupo"
         ListaCampos     =   "gru_codigo(1),gru_nombre(1),gru_nemotecnico(1)"
         XcodCampo       =   "gru_codigo"
         XListCampo      =   "gru_nombre"
         ListaCamposDescrip=   "codigo, descripcion,nemotecnico"
         ListaCamposText =   "gru_codigo,gru_nombre,gru_nemotecnico"
         Requerido       =   0   'False
      End
      Begin VB.Label Label4 
         Caption         =   "Grupo"
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Linea"
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Familia"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame Framedatos 
      Caption         =   "registro de proyectos"
      Height          =   3015
      Left            =   360
      TabIndex        =   4
      Top             =   1920
      Width           =   7575
      Begin VB.CheckBox CheckCierre 
         Alignment       =   1  'Right Justify
         Caption         =   "Cierre de Proyecto"
         Height          =   495
         Left            =   120
         TabIndex        =   31
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox Textproyecto 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         Height          =   375
         Left            =   1080
         MaxLength       =   11
         TabIndex        =   29
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox Textimporte 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Height          =   375
         Left            =   5640
         TabIndex        =   8
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox Textdescripcion 
         Height          =   375
         Left            =   1080
         MaxLength       =   40
         TabIndex        =   9
         Top             =   2040
         Width           =   6015
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaMoneda 
         Height          =   375
         Left            =   1080
         TabIndex        =   7
         Top             =   1560
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
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
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   5760
         TabIndex        =   6
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   108068865
         CurrentDate     =   41229
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayucliente 
         Height          =   375
         Left            =   1080
         TabIndex        =   5
         Top             =   1080
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   661
         XcodMaxLongitud =   11
         xcodwith        =   800
         NomTabla        =   "vt_cliente"
         TituloAyuda     =   "Ayuda de Clientes"
         ListaCampos     =   "clientecodigo(1),clienterazonsocial(1)"
         XcodCampo       =   "clientecodigo"
         XListCampo      =   "clienterazonsocial"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "clientecodigo,clienterazonsocial"
         Requerido       =   0   'False
      End
      Begin VB.Label Label10 
         Caption         =   "Codigo Proyecto"
         Height          =   495
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Cliente"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Importe"
         Height          =   255
         Left            =   4680
         TabIndex        =   22
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha Inicio"
         Height          =   375
         Left            =   4800
         TabIndex        =   21
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "descripcion"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Moneda"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1560
         Width           =   855
      End
   End
   Begin VB.Frame FrameListado 
      Height          =   3135
      Left            =   360
      TabIndex        =   24
      Top             =   1920
      Width           =   7815
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2295
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   4048
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmMntProyectos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim graba As Integer
Dim nemotecnico As String
Dim correlativo As Integer

Private Sub CheckTodos_Click()
Call consulta
End Sub

Private Sub CmdEli_Click()
If ok = 0 Then
   Exit Sub
End If
graba = 3
Call modifica
If MsgBox("esta seguro de eliminar ( S/N )", vbYesNo) = vbYes Then
   Call grabar(graba)
End If
End Sub

Private Sub CmdFicha_Click()
Dim aparam(6) As Variant
Dim aform(1) As Variant
Dim Cadorden As String
Dim reporte As String
aparam(0) = VGCNx.DefaultDatabase
aparam(1) = VGParametros.empresacodigo
aparam(2) = IIf(Ctr_Ayufamilia.xclave = "", "%%", Ctr_Ayufamilia.xclave)
aparam(3) = IIf(Ctr_Ayulinea.xclave = "", "%%", Ctr_Ayulinea.xclave)
aparam(4) = IIf(Ctr_Ayugrupo.xclave = "", "%%", Ctr_Ayugrupo.xclave)
aparam(5) = CheckTodos.Value

aform(0) = "titulo='" & Ctr_Ayufamilia.xnombre & "'"

Cadorden = ""
reporte = "gr_listaproyectos.rpt"
Call ImpresionRptProc(reporte, aform, aparam, Cadorden, "Lista proyectos ")
'Call ImpresionRptProc(NombreRep, arrform, arrparm, Cadorden, "Registro de Compras ")
 
End Sub

Private Sub CmdGrabar_Click()
Call grabar(graba)
Textproyecto.Enabled = True
End Sub

Private Sub CmdIng_Click()
If ok = 0 Then
      Exit Sub
End If
If numerocorrelativo = 1 Then
   graba = 1

   Textproyecto.Enabled = False
   Framedatos.Visible = True
   Call frame(False)
   limpia
End If
End Sub
Private Function ok()
ok = 0
If Ctr_Ayulinea.xclave = "" Then
   MsgBox (" Digite codigo de linea ")
   Exit Function
End If
If Ctr_Ayulinea.xclave = "" Then
   MsgBox (" Digite codigo de Grupoa ")
   Exit Function
End If
ok = 1
End Function

Private Sub CmdModi_Click()
If ok = 0 Then
   Exit Sub
End If
graba = 2
Call modifica
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Cmdsalirgraba_Click()
Call frame(True)
End Sub

Private Sub Ctr_Ayufamilia_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Ctr_Ayulinea.Filtro = " FAM_CODIGO='" & Ctr_Ayufamilia.xclave & "'"
End Sub

Private Sub Ctr_Ayugrupo_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Dim xx As String
FrameListado.Visible = True
Framedatos.Visible = False
nemotecnico = ColecCampos("gru_nemotecnico")
Call consulta
Call frame(True)
End Sub
Private Sub consulta()
Dim acmd As New ADODB.Command
Set acmd.ActiveConnection = VGgeneral
acmd.CommandType = adCmdStoredProc
acmd.CommandText = "gr_proyectos_pro"
acmd.Parameters.Refresh
With acmd
    .Parameters("@base") = VGCNx.DefaultDatabase
    .Parameters("@empresa") = VGParametros.empresacodigo
    .Parameters("@familia") = Ctr_Ayufamilia.xclave
    .Parameters("@linea") = Ctr_Ayulinea.xclave
    .Parameters("@grupo") = Ctr_Ayugrupo.xclave
    .Parameters("@tipo") = 4
    .Parameters("@todos") = CheckTodos.Value
    Set RSQL = .Execute
End With
Set DataGrid1.DataSource = RSQL
DataGrid1.Refresh
End Sub
Private Sub Ctr_Ayulinea_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Ctr_Ayugrupo.Filtro = " FAM_CODIGO='" & Ctr_Ayufamilia.xclave & "' and lin_CODIGO='" & Ctr_Ayulinea.xclave & "'"
End Sub

Private Sub Form_Load()
Call Ctr_Ayufamilia.conexion(VGCNx)
Call Ctr_Ayulinea.conexion(VGCNx)
Call Ctr_Ayugrupo.conexion(VGCNx)
Call Ctr_Ayucliente.conexion(VGCNx)
Call Ctr_AyudaMoneda.conexion(VGCNx)
Ctr_Ayufamilia.xclave = VGParamSistem.familiaproyectos: Ctr_Ayufamilia.Ejecutar
Ctr_Ayufamilia.Enabled = False
Ctr_AyudaMoneda.Filtro = " monedacodigo <>'00' "
DTPicker1.Value = VGParamSistem.FechaTrabajo
FrameGrupo.Enabled = True
FrameListado.Visible = False
Framedatos.Visible = False
Framegraba.Visible = False
CheckTodos.Value = 0

CmdIng.Picture = MDIPrincipal.ImageList2.ListImages.Item("Insertar").Picture
CmdModi.Picture = MDIPrincipal.ImageList2.ListImages.Item("Modificar").Picture
CmdEli.Picture = MDIPrincipal.ImageList2.ListImages.Item("Eliminar").Picture
CmdFicha.Picture = MDIPrincipal.ImageList2.ListImages.Item("Nuevo").Picture
CmdGrabar.Picture = MDIPrincipal.ImageList2.ListImages.Item("Grabar").Picture
CmdSalir.Picture = MDIPrincipal.ImageList2.ListImages.Item("Retornar").Picture
Cmdsalirgraba.Picture = MDIPrincipal.ImageList2.ListImages.Item("Retornar").Picture

End Sub

Private Sub frame(valor As Boolean)
Framemodifica.Visible = valor
Framegraba.Visible = Not valor
Framedatos.Visible = Not valor
FrameListado.Visible = valor
End Sub
Private Sub grabar(dato As Integer)
Dim acmd As New ADODB.Command
Set acmd.ActiveConnection = VGgeneral
acmd.CommandType = adCmdStoredProc
acmd.CommandText = "gr_proyectos_pro"
acmd.Parameters.Refresh
With acmd
    .Parameters("@base") = VGCNx.DefaultDatabase
    .Parameters("@empresa") = VGParametros.empresacodigo
    .Parameters("@familia") = Ctr_Ayufamilia.xclave
    .Parameters("@linea") = Ctr_Ayulinea.xclave
    .Parameters("@grupo") = Ctr_Ayugrupo.xclave
    .Parameters("@tipo") = dato
    .Parameters("@proyectocodigo") = Textproyecto.Text
    .Parameters("@proyectodescripcion") = Textdescripcion
    .Parameters("@clientecodigo") = Ctr_Ayucliente.xclave
    .Parameters("@monedacodigo") = Ctr_AyudaMoneda.xclave
    .Parameters("@proyectoimporte") = Textimporte.Text
    .Parameters("@proyectofechainicio") = DTPicker1.Value
    .Parameters("@tipoanalitico") = VGParamSistem.tipoanaliticocodigo
    .Parameters("@proyectotipo") = Left(nemotecnico, 1)
    .Parameters("@periodo") = VGParamSistem.AnoProceso
    .Parameters("@correlativo") = correlativo
    .Parameters("@cierre") = CheckCierre.Value
    .Parameters("@usuariocodigo") = g_usuario
    .Execute
End With
If graba = 3 Then
   Call entidadElimina(graba)
 Else
   Call entidadAdiciona(dato)
End If
Call consulta
Call frame(True)
graba = 0
End Sub

Private Sub modifica()
Call frame(False)
Textproyecto = RSQL!proyectocodigo
Textproyecto.Enabled = False
Ctr_Ayufamilia.xclave = RSQL!fam_codigo
Ctr_Ayulinea.xclave = RSQL!lin_codigo
Ctr_Ayugrupo.xclave = RSQL!gru_codigo
Textdescripcion = RSQL!proyectodescripcion
Ctr_Ayucliente.xclave = RSQL!clientecodigo: Ctr_Ayucliente.Ejecutar
Ctr_AyudaMoneda.xclave = RSQL!monedacodigo: Ctr_AyudaMoneda.Ejecutar
Textimporte.Text = RSQL!proyectoimporte
DTPicker1.Value = RSQL!proyectofechainicio
CheckCierre.Value = RSQL!proyectocierre
End Sub

Private Sub entidadAdiciona(dato As Integer)
Dim r1sql As New ADODB.Recordset
Set r1sql = VGCNx.Execute("select * from ct_entidad where entidadcodigo='" & Textproyecto & "'")
If r1sql.RecordCount = 0 Then
   SQL = "INSERT CT_ENTIDAD(entidadcodigo,entidadrazonsocial,entidaddireccion,entidadruc,usuariocodigo) "
   SQL = SQL & " VALUES ('" & Textproyecto & "','" & Textdescripcion & "','" & Left(Ctr_Ayucliente.xnombre, 25) & "','" & Textproyecto & "','" & g_usuario & "')"
 Else
    SQL = "UPDATE CT_ENTIDAD SET entidadrazonsocial='" & Textdescripcion.Text & "',entidaddireccion='" & Left(Ctr_Ayucliente.xnombre, 25) & "',"
    SQL = SQL & " usuariocodigo='" & VGUsuario & "',fechaact=getdate() , proyectocierre=" & CheckCierre.Value & ""
    SQL = SQL & " WHERE entidadcodigo='" & Textproyecto & "'"
End If
VGCNx.BeginTrans
VGCNx.Execute (SQL)
VGCNx.CommitTrans

Set r1sql = VGCNx.Execute("select * from ct_analitico where analiticocodigo='" & (Textproyecto) & VGParamSistem.tipoanaliticocodigo & "'")
If r1sql.RecordCount = 0 Then
   SQL = "INSERT CT_analitico(analiticocodigo,tipoanaliticocodigo,entidadcodigo,usuariocodigo) "
   SQL = SQL & " VALUES ('" & (Textproyecto) & VGParamSistem.tipoanaliticocodigo & "','" & VGParamSistem.tipoanaliticocodigo & "','" & Textproyecto & "','" & g_usuario & "')"
   VGCNx.BeginTrans
   VGCNx.Execute (SQL)
   VGCNx.CommitTrans
End If

Call frame(False)
End Sub
Private Sub entidadElimina(dato As Integer)
SQL = "DELETE ct_analitico WHERE analiticocodigo='" & Textproyecto & Ctr_Ayufamilia.xclave & "'"
VGCNx.BeginTrans
VGCNx.Execute (SQL)
VGCNx.CommitTrans

SQL = "DELETE ct_entidad WHERE entidadcodigo='" & Textproyecto & "'"

VGCNx.BeginTrans
VGCNx.Execute (SQL)
VGCNx.CommitTrans
End Sub
Private Sub Textproyecto_KeyPress(KeyAscii As Integer)
Dim rsql1 As New ADODB.Recordset
If KeyAscii = 13 Then
   If graba = 1 Then
      Set rsql1 = VGCNx.Execute(" select * from gr_proyectos where proyectocodigo='" & Textproyecto.Text & "'")
      If rsql1.RecordCount > 0 Then
         MsgBox ("proyecto existe en : " & rsql1!fam_codigo & " - " & rsql1!lin_codigo & " - " & rsql1!gru_codigo & "")
         Textproyecto.SetFocus
         Exit Sub
      End If
      SendKeys "{tab}"
   End If
End If
End Sub
Private Sub Textproyecto_LostFocus()
   Call Textproyecto_KeyPress(13)
End Sub
Private Sub limpia()
Textdescripcion = ""
Ctr_Ayucliente.xclave = ""
Ctr_AyudaMoneda.xclave = ""
Textimporte.Text = 0
CheckCierre.Value = 0
End Sub
Private Function numerocorrelativo()
numerocorrelativo = 0
SQL = "select ultimo=isnull(max(proyectocorrelativo),0)+1 from gr_proyectos where proyectoperiodo=" & VGParamSistem.AnoProceso & ""
SQL = SQL & " and proyectotipo='" & Left(nemotecnico, 1) & "' "
Set RSQL = VGCNx.Execute(SQL)
If RSQL!ultimo = 1 Then
   If MsgBox(" Es proyecto nuevo para este periodo : " & VGParamSistem.AnoProceso & "", vbYesNo) = vbNo Then
      Exit Function
   End If
End If
Textproyecto = nemotecnico + Right(VGParamSistem.AnoProceso, 2) + Format(RSQL!ultimo, "000")
correlativo = RSQL!ultimo
numerocorrelativo = 1
End Function
