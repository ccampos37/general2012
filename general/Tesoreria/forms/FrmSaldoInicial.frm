VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Begin VB.Form Frmsaldoinicial 
   Caption         =   "Saldos Iniciales"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   7845
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   495
      TabIndex        =   10
      Top             =   1845
      Width           =   5280
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayumonedacuenta 
         Height          =   315
         Left            =   90
         TabIndex        =   3
         Top             =   405
         Width           =   5070
         _ExtentX        =   8943
         _ExtentY        =   556
         XcodMaxLongitud =   2
         xcodwith        =   200
         Requerido       =   0   'False
      End
      Begin TextFer.TxFer TxSalDisp 
         Height          =   345
         Left            =   45
         TabIndex        =   4
         Top             =   1260
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   609
         Alignment       =   1
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
         ColorIlumina    =   14546937
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         TipoDato        =   1
         NumeroDecimales =   2
         SignoNegativo   =   0   'False
         MarcarTextoAlEnfoque=   -1  'True
      End
      Begin VB.Label Label5 
         Caption         =   "Saldo Inicial Disponible :"
         Height          =   165
         Left            =   375
         TabIndex        =   12
         Top             =   945
         Width           =   1845
      End
      Begin VB.Label Lb2 
         Caption         =   "Numero de Cuenta :"
         Height          =   315
         Left            =   1125
         TabIndex        =   11
         Top             =   180
         Width           =   1860
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2790
      Left            =   6075
      TabIndex        =   9
      Top             =   540
      Width           =   1500
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   765
         Left            =   360
         Picture         =   "FrmSaldoInicial.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   945
      End
      Begin VB.CommandButton CmdCancelar 
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
         Height          =   765
         Left            =   360
         Picture         =   "FrmSaldoInicial.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1620
         Width           =   945
      End
   End
   Begin VB.ComboBox CmbOper 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "FrmSaldoInicial.frx":0884
      Left            =   1350
      List            =   "FrmSaldoInicial.frx":088E
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   135
      Width           =   3000
   End
   Begin VB.Frame Frame1 
      Height          =   3285
      Left            =   180
      TabIndex        =   0
      Top             =   540
      Width           =   5745
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayucodcajabanco 
         Height          =   345
         Left            =   1380
         TabIndex        =   2
         Top             =   720
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   609
         XcodMaxLongitud =   2
         xcodwith        =   200
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuempresa 
         Height          =   315
         Left            =   1455
         TabIndex        =   13
         Top             =   180
         Width           =   4050
         _ExtentX        =   7144
         _ExtentY        =   556
         XcodMaxLongitud =   3
         xcodwith        =   300
         NomTabla        =   "co_multiempresas"
         TituloAyuda     =   "Busqueda de Empresas"
         ListaCampos     =   "empresacodigo(1),empresadescripcion(1)"
         XcodCampo       =   "empresacodigo"
         XListCampo      =   "empresadescripcion"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "empresacodigo,empresadescripcion"
      End
      Begin VB.Label Lblempresa 
         AutoSize        =   -1  'True
         Caption         =   "Empresa :"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   240
         Width           =   705
      End
      Begin VB.Label LbCod 
         Caption         =   "Codigo de Banco :"
         Height          =   405
         Left            =   120
         TabIndex        =   7
         Top             =   645
         Width           =   1170
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Operacion :"
      Height          =   270
      Left            =   195
      TabIndex        =   8
      Top             =   165
      Width           =   1005
   End
End
Attribute VB_Name = "Frmsaldoinicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rssaldo As ADODB.Recordset
Dim IMant As Integer
Dim VlMon As String

Private Sub CmbOper_Click()
    If CmbOper.ListIndex = 0 Then
        LbCod.Caption = "Codigo de Caja :"
        Lb2.Caption = "Moneda de la Caja : "
        
        Ctr_Ayucodcajabanco.NomTabla = "te_codigocaja"
        Ctr_Ayucodcajabanco.ListaCampos = "cajacodigo(1),cajadescripcion(1)"
        Ctr_Ayucodcajabanco.ListaCamposText = "cajacodigo,cajadescripcion"
        Ctr_Ayucodcajabanco.ListaCamposDescrip = "Codigo,Descripcion "
        Ctr_Ayucodcajabanco.TituloAyuda = "Busqueda de Caja"
        Ctr_Ayucodcajabanco.XcodCampo = "cajacodigo"
        Ctr_Ayucodcajabanco.XListCampo = "cajadescripcion"
                
        Ctr_Ayumonedacuenta.NomTabla = "gr_moneda"
        Ctr_Ayumonedacuenta.ListaCampos = "monedacodigo(1),monedadescripcion(1)"
        Ctr_Ayumonedacuenta.ListaCamposText = "monedacodigo,monedadescripcion"
        Ctr_Ayumonedacuenta.ListaCamposDescrip = "Codigo,Descripcion"
        Ctr_Ayumonedacuenta.TituloAyuda = "Busqueda de Moneda"
        Ctr_Ayumonedacuenta.XcodCampo = "monedacodigo"
        Ctr_Ayumonedacuenta.XListCampo = "monedadescripcion"
        Ctr_Ayumonedacuenta.XcodMaxLongitud = 2
        Ctr_Ayumonedacuenta.xcodwith = 200
       Else
        Ctr_Ayucodcajabanco.NomTabla = "gr_banco"
        Ctr_Ayucodcajabanco.ListaCampos = "bancocodigo(1),bancodescripcion(1)"
        Ctr_Ayucodcajabanco.ListaCamposText = "bancocodigo,bancodescripcion"
        Ctr_Ayucodcajabanco.ListaCamposDescrip = "Codigo,Descripcion"
        Ctr_Ayucodcajabanco.TituloAyuda = "Busqueda de Banco"
        Ctr_Ayucodcajabanco.XcodCampo = "bancocodigo"
        Ctr_Ayucodcajabanco.XListCampo = "bancodescripcion"
                
        Ctr_Ayumonedacuenta.NomTabla = "te_cuentabancos"
        Ctr_Ayumonedacuenta.ListaCampos = "cbanco_numero(1),cbanco_referenciacta(1),cbanco_codigo(1),monedacodigo(1)"
        Ctr_Ayumonedacuenta.ListaCamposText = "cbanco_numero,cbanco_referenciacta,cbanco_codigo,monedacodigo"
        Ctr_Ayumonedacuenta.ListaCamposDescrip = "Codigo,Descripcion"
        Ctr_Ayumonedacuenta.TituloAyuda = "Busqueda de Cuentas de Banco"
        Ctr_Ayumonedacuenta.XcodCampo = "cbanco_numero"
        Ctr_Ayumonedacuenta.XListCampo = "cbanco_referenciacta"
        LbCod.Caption = "Codigo de Banco :"
        Lb2.Caption = "Numero de Cuenta : "
        Ctr_Ayumonedacuenta.XcodMaxLongitud = 30
        Ctr_Ayumonedacuenta.xcodwith = 3000
    End If
'    Call Mostrar
End Sub

Private Sub cmdaceptar_Click()
If Ctr_Ayucodcajabanco.xclave = "" Then
   MsgBox "Ingrese Codigo de Caja o Banco", vbInformation, MsgTitle
       Ctr_Ayucodcajabanco.SetFocus
       Exit Sub
End If
If Ctr_Ayumonedacuenta.xclave = "" Then
   MsgBox "Ingrese Codigo de Moneda o Cuenta Bancaria", vbInformation, MsgTitle
       Ctr_Ayumonedacuenta.SetFocus
       Exit Sub
End If
If IMant = 1 Then Call GrabaSaldo
If IMant = 2 Then Call actualizasaldo
Ctr_Ayucodcajabanco.xclave = "": Ctr_Ayucodcajabanco.Ejecutar
Ctr_Ayumonedacuenta.xclave = "": Ctr_Ayumonedacuenta.Ejecutar
TxSalDisp.Text = 0
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub Ctr_Ayucodcajabanco_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Dim rssaldo As New ADODB.Recordset
If CmbOper.ListIndex = 1 Then
        
        Ctr_Ayumonedacuenta.Filtro = "cbanco_codigo='" & ColecCampos("bancocodigo").Value & "'"
        SQL = "Select * From te_cuentabancos Where cbanco_codigo='" & ColecCampos("bancocodigo").Value & "' "
     
    Set rssaldo = VGCNx.Execute(SQL)
    If rssaldo.RecordCount = 0 Then
        Ctr_Ayucodcajabanco.SetFocus
        Frame2.Visible = False
    Else
        Frame2.Visible = True
    End If
End If

End Sub

Private Sub Ctr_Ayumonedacuenta_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    If CmbOper.ListIndex = 1 Then
        VlMon = ColecCampos("monedacodigo").Value
    End If
    Call Mostrar

End Sub

Private Sub Form_Load()
     Call Ctr_Ayucodcajabanco.Conexion(VGCNx)
    Call Ctr_Ayumonedacuenta.Conexion(VGCNx)
    Call Ctr_Ayuempresa.Conexion(VGCNx)
    If VGParametros.sistemamultiempresas Then
       Ctr_Ayuempresa.Enabled = True
     Else
       Ctr_Ayuempresa.xclave = "01"
       Ctr_Ayuempresa.Enabled = False
    End If
    Set rssaldo = New ADODB.Recordset
End Sub
Private Sub Mostrar()
    '@ano    varchar (2),@mes    varchar (4), @Oper
Dim oper As String
    If CmbOper.ListIndex = 0 Then
        oper = "C"
      Else
        oper = "B"
    End If
    SQL = "Select * From te_saldosmensuales Where  empresacodigo='" & Ctr_Ayuempresa.xclave & "'and  tipocajabanco='" & oper & "'"
    SQL = SQL & " and CajaBancoCodigo='" & Ctr_Ayucodcajabanco.xclave & "' "
    SQL = SQL & " and MonedaCuenta='" & Ctr_Ayumonedacuenta.xclave & "'"
    SQL = SQL & " and mesproceso='" & VGParamSistem.AnoProceso & Format(VGParamSistem.MesProceso, "00") & "'"
    Set rssaldo = VGCNx.Execute(SQL)
    If rssaldo.RecordCount = 0 Then
        IMant = 1
      Else
        IMant = 2
        TxSalDisp.Text = rssaldo!Saldoinicial
    End If
End Sub

Private Sub actualizasaldo()
Dim oper As String
    If CmbOper.ListIndex = 0 Then
        oper = "C"
      Else
        oper = "B"
    End If
    SQL = " update te_saldosmensuales set saldoinicial=" & CDbl(TxSalDisp.valor) & ", "
    SQL = SQL & "fechaact=" & Format(Now, "dd/mm/yyyy") & ",usuariocodigo='" & Trim(VGParamSistem.Usuario) & "'"
    SQL = SQL & " Where  empresacodigo='" & Ctr_Ayuempresa.xclave & "' and tipocajabanco='" & oper & "'"
    SQL = SQL & " and CajaBancoCodigo='" & Ctr_Ayucodcajabanco.xclave & "' "
    SQL = SQL & " and MonedaCuenta='" & Ctr_Ayumonedacuenta.xclave & "'"
    SQL = SQL & " and mesproceso='" & VGParamSistem.AnoProceso & Format(VGParamSistem.MesProceso, "00") & "'"
    Set rssaldo = VGCNx.Execute(SQL)

End Sub
Private Sub GrabaSaldo()
    Dim Comando As ADODB.Command
    Set Comando = New ADODB.Command
    Comando.ActiveConnection = VGgeneral
    Comando.CommandType = adCmdStoredProc
    Comando.CommandText = "te_actsaldo_pro"
    Comando.Parameters.Refresh
    With Comando
        .Parameters("@base").Value = VGParamSistem.BDEmpresa
        .Parameters("@op").Value = IMant
        .Parameters("@empresa") = Ctr_Ayuempresa.xclave
        .Parameters("@Oper").Value = Trim(IIf(CmbOper.ListIndex = 0, "C", "B"))
        .Parameters("@CodCajaBanco").Value = Ctr_Ayucodcajabanco.xclave
        .Parameters("@CtaBanco").Value = Trim(Ctr_Ayumonedacuenta.xclave)
        .Parameters("@aaaamm").Value = VGParamSistem.AnoProceso + Format(VGParamSistem.MesProceso, "00")
        .Parameters("@CodMon").Value = VlMon
        .Parameters("@SaldoDisp").Value = CDbl(TxSalDisp.valor)
        .Parameters("@usuariocodigo").Value = Trim(VGParamSistem.Usuario)
        .Parameters("@fechaact").Value = Now
        .Execute
    End With
End Sub

