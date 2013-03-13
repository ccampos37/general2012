VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form frmParametros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parámetros Generales"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   9945
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   7380
      Left            =   30
      TabIndex        =   16
      Top             =   -45
      Width           =   9588
      Begin VB.CheckBox chk 
         Height          =   240
         Index           =   4
         Left            =   2448
         TabIndex        =   33
         Top             =   1800
         Width           =   288
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_moneda 
         Height          =   336
         Left            =   2388
         TabIndex        =   11
         Top             =   3552
         Width           =   4476
         _ExtentX        =   7885
         _ExtentY        =   582
         XcodMaxLongitud =   0
         xcodwith        =   500
         NomTabla        =   "gr_moneda"
         ListaCampos     =   "monedacodigo(1),monedadescripcion(1)"
         XcodCampo       =   "monedacodigo"
         XListCampo      =   "monedadescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "monedacodigo,monedadescripcion"
         Requerido       =   0   'False
      End
      Begin VB.CheckBox chk 
         Height          =   240
         Index           =   3
         Left            =   8340
         TabIndex        =   14
         Top             =   1800
         Visible         =   0   'False
         Width           =   645
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaCuentaAjuste 
         Height          =   348
         Index           =   0
         Left            =   2388
         TabIndex        =   9
         Top             =   2892
         Width           =   4440
         _ExtentX        =   7832
         _ExtentY        =   609
         XcodMaxLongitud =   0
         xcodwith        =   800
         NomTabla        =   "ct_cuenta"
         ListaCampos     =   "cuentacodigo(1),cuentadescripcion(1)"
         XcodCampo       =   "cuentacodigo"
         XListCampo      =   "cuentadescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "cuentacodigo,cuentadescripcion"
      End
      Begin VB.CheckBox chk 
         Height          =   240
         Index           =   2
         Left            =   5745
         TabIndex        =   6
         Top             =   1800
         Width           =   288
      End
      Begin VB.CheckBox chk 
         Height          =   240
         Index           =   1
         Left            =   8412
         TabIndex        =   5
         Top             =   1356
         Width           =   645
      End
      Begin VB.CheckBox chk 
         Height          =   240
         Index           =   0
         Left            =   2412
         TabIndex        =   2
         Top             =   795
         Width           =   360
      End
      Begin TextFer.TxFer txt 
         Height          =   300
         Index           =   0
         Left            =   2412
         TabIndex        =   0
         Top             =   180
         Width           =   5508
         _ExtentX        =   9710
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
         MaxLength       =   40
         Text            =   ""
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         MarcarTextoAlEnfoque=   -1  'True
      End
      Begin TextFer.TxFer txt 
         Height          =   300
         Index           =   1
         Left            =   2412
         TabIndex        =   1
         Top             =   480
         Width           =   3408
         _ExtentX        =   6006
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
         MaxLength       =   15
         Text            =   ""
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         MarcarTextoAlEnfoque=   -1  'True
      End
      Begin TextFer.TxFer txt 
         Height          =   300
         Index           =   3
         Left            =   2400
         TabIndex        =   4
         Top             =   1332
         Width           =   2112
         _ExtentX        =   3731
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
         MaxLength       =   11
         Text            =   ""
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         TipoDato        =   1
         SignoNegativo   =   0   'False
         MarcarTextoAlEnfoque=   -1  'True
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2880
         Left            =   72
         TabIndex        =   13
         Top             =   4320
         Width           =   2856
         _ExtentX        =   5027
         _ExtentY        =   5080
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin TextFer.TxFer txt 
         Height          =   300
         Index           =   2
         Left            =   2400
         TabIndex        =   3
         Top             =   1032
         Width           =   5520
         _ExtentX        =   9737
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
         MaxLength       =   40
         Text            =   ""
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         MarcarTextoAlEnfoque=   -1  'True
      End
      Begin TextFer.TxFer txt 
         Height          =   300
         Index           =   4
         Left            =   6432
         TabIndex        =   12
         Top             =   768
         Width           =   1368
         _ExtentX        =   2408
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
         MaxLength       =   11
         Text            =   ""
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         TipoDato        =   1
         NumeroDecimales =   2
         SignoNegativo   =   0   'False
         MarcarTextoAlEnfoque=   -1  'True
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaCuentaAjuste 
         Height          =   348
         Index           =   1
         Left            =   2388
         TabIndex        =   10
         Top             =   3228
         Width           =   4452
         _ExtentX        =   7858
         _ExtentY        =   609
         XcodMaxLongitud =   0
         xcodwith        =   800
         NomTabla        =   "ct_cuenta"
         ListaCampos     =   "cuentacodigo(1),cuentadescripcion(1)"
         XcodCampo       =   "cuentacodigo"
         XListCampo      =   "cuentadescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "cuentacodigo,cuentadescripcion"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Codigo 
         Height          =   336
         Index           =   0
         Left            =   2388
         TabIndex        =   7
         Top             =   2112
         Visible         =   0   'False
         Width           =   3828
         _ExtentX        =   6747
         _ExtentY        =   582
         XcodMaxLongitud =   3
         xcodwith        =   600
         NomTabla        =   "ct_asiento"
         ListaCampos     =   "asientocodigo(1),asientodescripcion(1)"
         XcodCampo       =   "asientocodigo"
         XListCampo      =   "asientodescripcion"
         ListaCamposDescrip=   "Código,Descripcion"
         ListaCamposText =   "asientocodigo,asientodescripcion"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Codigo 
         Height          =   336
         Index           =   1
         Left            =   2388
         TabIndex        =   8
         Top             =   2472
         Visible         =   0   'False
         Width           =   3816
         _ExtentX        =   6720
         _ExtentY        =   582
         XcodMaxLongitud =   4
         xcodwith        =   600
         NomTabla        =   "ct_subasiento"
         ListaCampos     =   "subasientocodigo(1),subasientodescripcion(1)"
         XcodCampo       =   "subasientocodigo"
         XListCampo      =   "subasientodescripcion"
         ListaCamposDescrip=   "Código,Descripcion"
         ListaCamposText =   "subasientocodigo,subasientodescripcion"
         Requerido       =   0   'False
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2880
         Left            =   4080
         TabIndex        =   35
         Top             =   4380
         Width           =   2856
         _ExtentX        =   5027
         _ExtentY        =   5080
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lbl 
         Caption         =   "Niveles Centro de Costos"
         Height          =   216
         Index           =   13
         Left            =   4104
         TabIndex        =   36
         Top             =   4080
         Width           =   1848
      End
      Begin VB.Label lbl 
         Caption         =   "Contabilidad Monista"
         Height          =   216
         Index           =   12
         Left            =   120
         TabIndex        =   34
         Top             =   1800
         Width           =   1848
      End
      Begin VB.Label Label4 
         Caption         =   "Asiento Descuadrado"
         Height          =   240
         Left            =   108
         TabIndex        =   32
         Top             =   2196
         Visible         =   0   'False
         Width           =   2808
      End
      Begin VB.Label Label3 
         Caption         =   "SubAsiento Descuadrado"
         Height          =   180
         Left            =   108
         TabIndex        =   31
         Top             =   2532
         Visible         =   0   'False
         Width           =   2808
      End
      Begin VB.Label lbl 
         Caption         =   "Cuenta Ajuste Haber"
         Height          =   216
         Index           =   11
         Left            =   120
         TabIndex        =   30
         Top             =   3276
         Width           =   2808
      End
      Begin VB.Label lbl 
         Caption         =   "Cuenta Ajuste  Debe"
         Height          =   216
         Index           =   10
         Left            =   120
         TabIndex        =   29
         Top             =   2952
         Width           =   2808
      End
      Begin VB.Label Label1 
         Caption         =   "Ajuste en Línea"
         Height          =   210
         Left            =   6585
         TabIndex        =   28
         Top             =   1830
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lbl 
         Caption         =   "Valor IGV (%)"
         Height          =   216
         Index           =   9
         Left            =   3828
         TabIndex        =   27
         Top             =   816
         Width           =   2808
      End
      Begin VB.Label lbl 
         Caption         =   "Niveles Plan de Cuentas"
         Height          =   216
         Index           =   8
         Left            =   96
         TabIndex        =   26
         Top             =   4020
         Width           =   1848
      End
      Begin VB.Label lbl 
         Caption         =   "Moneda Base"
         Height          =   216
         Index           =   7
         Left            =   108
         TabIndex        =   25
         Top             =   3576
         Width           =   2808
      End
      Begin VB.Label lbl 
         Caption         =   "Impresión de Asiento (Comprobante)"
         Height          =   210
         Index           =   6
         Left            =   2955
         TabIndex        =   24
         Top             =   1815
         Width           =   2580
      End
      Begin VB.Label lbl 
         Caption         =   "Cuadre de Asiento (Comprobante)"
         Height          =   216
         Index           =   5
         Left            =   5376
         TabIndex        =   23
         Top             =   1356
         Width           =   2808
      End
      Begin VB.Label lbl 
         Caption         =   "RUC Empresa"
         Height          =   210
         Index           =   4
         Left            =   105
         TabIndex        =   22
         Top             =   1410
         Width           =   2805
      End
      Begin VB.Label lbl 
         Caption         =   "Usar Descri. Larga en Reportes"
         Height          =   210
         Index           =   3
         Left            =   90
         TabIndex        =   21
         Top             =   840
         Width           =   2805
      End
      Begin VB.Label lbl 
         Caption         =   "Dirección Empresa"
         Height          =   210
         Index           =   2
         Left            =   105
         TabIndex        =   20
         Top             =   1125
         Width           =   2805
      End
      Begin VB.Label lbl 
         Caption         =   "Descripción Empresa (Abrev.)"
         Height          =   210
         Index           =   1
         Left            =   105
         TabIndex        =   19
         Top             =   525
         Width           =   2805
      End
      Begin VB.Label lbl 
         Caption         =   "Descripción Empresa (Larga)"
         Height          =   210
         Index           =   0
         Left            =   105
         TabIndex        =   18
         Top             =   225
         Width           =   2805
      End
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Cancelar"
      Height          =   360
      Index           =   1
      Left            =   4560
      TabIndex        =   17
      Top             =   7515
      Width           =   1605
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Aceptar"
      Height          =   360
      Index           =   0
      Left            =   2610
      TabIndex        =   15
      Top             =   7515
      Width           =   1605
   End
End
Attribute VB_Name = "frmParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As New ADODB.Recordset
Dim FlagNUEVO As Boolean

Private Sub Ctr_AyudaCuentaAjuste_AlDevolverDato(Index As Integer, ByVal ColecCampos As ADODB.Fields)
  cmdBotones(0).Enabled = ValidaBoton()
End Sub

Private Sub Ctr_Codigo_AlDevolverDato(Index As Integer, ByVal ColecCampos As ADODB.Fields)
   Ctr_Codigo(1).Filtro = "asientocodigo='" & ColecCampos(0).Value & "'"
End Sub

Private Sub Form_Load()
  Set rs = New ADODB.Recordset
  Ctr_moneda.Enabled = True
  Ctr_moneda.conexion VGCNx
  Ctr_AyudaCuentaAjuste(0).conexion VGCNx
  Ctr_AyudaCuentaAjuste(1).conexion VGCNx
  Ctr_AyudaCuentaAjuste(0).Filtro = "cuentanivel=" & VGnumnivelescuenta
  Ctr_AyudaCuentaAjuste(1).Filtro = "cuentanivel=" & VGnumnivelescuenta
  Ctr_Codigo(0).conexion VGCNx
  Ctr_Codigo(1).conexion VGCNx
  Call CargarData
  Me.Width = 8910
  Me.Height = 8445
  cmdBotones(0).Enabled = False
End Sub

Sub CargarData()
 Dim SQL As String
'SQL = "SELECT sistemadescripcionempresa,sistemadescrcortaempresa,sistemaesttipodescrempresa,sistemadireccionempresa,sistemaempresaruc,sistemaestcuadreasiento,sistemaestimpresionasiento,sistemaconfiguracuenta,monedacodigo,sistemavalorigv,sistemaajustelinea,isnull(sistemactaajustedeb,''),isnull(sistemactaajustehab,''),sistemaasientocodigo,sistemasubasientocodigo,sistemamonista FROM ct_sistema"
 SQL = "SELECT * FROM ct_sistema"
  Set rs = VGCNx.Execute(SQL)
  If rs.RecordCount = 0 Then
    FlagNUEVO = True
    Call LlenarLista
    Exit Sub
  End If
  
  Call MuestraData
  Call LlenarLista
  Call MarcarLista

End Sub

Sub MuestraData()
 Dim i As Integer
   For i = 0 To 3
     txt(i).Text = Trim$(IIf(i > 1, rs!sistemadescrcortaempresa, rs!sistemadescripcionempresa))
   Next
   chk(0).Value = IIf(rs!sistemaesttipodescrempresa = 0, 0, 1)
   chk(1).Value = IIf(rs!sistemaestcuadreasiento = 0, 0, 1)
   chk(2).Value = IIf(rs!sistemaestimpresionasiento = 0, 0, 1)
   chk(4).Value = IIf(rs!sistemamonista = 0, 0, 1)
   
'SQL = 1,2,3,sistemadireccionempresa,sistemaempresaruc,5,6,
'sistemaconfiguracuenta,8 ,9,10,11,12  ,13 ,14 , FROM ct_sistema"
   
   
   Ctr_moneda.xclave = Trim$(rs!monedacodigo): Ctr_moneda.Ejecutar
   txt(4).Text = rs!sistemavalorigv
   chk(3).Value = IIf(rs!sistemaajustelinea = 0, 0, 1)

   Ctr_AyudaCuentaAjuste(0).xclave = rs!sistemactaajustedeb: Ctr_AyudaCuentaAjuste(0).Ejecutar
   Ctr_AyudaCuentaAjuste(1).xclave = rs!sistemactaajustehab: Ctr_AyudaCuentaAjuste(1).Ejecutar
   
   Ctr_Codigo(0).xclave = rs!sistemaasientocodigo: Ctr_Codigo(0).Ejecutar
   Ctr_Codigo(1).xclave = rs!sistemasubasientocodigo: Ctr_Codigo(1).Ejecutar

End Sub

Sub LlenarLista()
 Dim i As Integer
 Dim itmX As ListItem
 
   ListView1.ColumnHeaders.Clear
   ListView1.ListItems.Clear
   ListView1.ColumnHeaders.Add , , "Número Dígitos", ListView1.Width / 1
   ListView1.View = lvwReport
   
   ListView2.ColumnHeaders.Clear
   ListView2.ListItems.Clear
   ListView2.ColumnHeaders.Add , , "Número Dígitos", ListView2.Width / 1
   ListView2.View = lvwReport
   For i = 1 To 9
     Set itmX = ListView1.ListItems.Add(, , i)
     Set itmX = ListView2.ListItems.Add(, , i)
   Next

End Sub

Sub MarcarLista()
 Dim i As Integer
 Dim J As Integer
      
   Call ParametroCuenta(0)
   For i = 1 To VGnumnivelescuenta
     For J = 1 To 9
       If ListView1.ListItems.Item(J).Text = VG_aNIVELES(i - 1) Then
          ListView1.ListItems.Item(J).Checked = True
       End If
     Next
   Next
   
''**
If VGnumnivelescentrocosto > 0 Then
    For i = 1 To VGnumnivelescentrocosto
     For J = 1 To 9
       If ListView2.ListItems.Item(J).Text = VG_cNIVELES(i - 1) Then
          ListView2.ListItems.Item(J).Checked = True
       End If
     Next
   Next
End If

End Sub

Private Sub cmdBotones_Click(Index As Integer)
  Select Case Index
    Case 0:
     If ValidarData() = True Then
       Call GrabarData
       Call CargarParametrosContabilidad
     End If
    
    Case 1: Unload Me
  End Select
  
End Sub

Function ValidarData() As Boolean
 Dim i As Integer
 Dim flagList As Boolean
 Dim nC As Integer
 Dim SQL As String
 Dim rsX As ADODB.Recordset
  For i = 1 To 9
      If ListView1.ListItems.Item(i).Checked = True Then
        ValidarData = True
        Exit For
      Else
        ValidarData = False
        If i = 9 Then
          MsgBox "Falta Seleccionar los Niveles para el Plan de Cuentas", vbInformation, Caption
          Exit Function
        End If
      End If
  Next

  For i = 1 To 9
    If ListView1.ListItems.Item(i).Checked = True Then
      nC = nC + 1
    End If
  Next
  If nC < 3 Then
    MsgBox "Debe Seleccionar al menos 3 Niveles para el Plan de Cuentas", vbInformation, Caption
    ValidarData = False
    Exit Function
  End If
  nC = 0
  Dim J As Integer
      For J = 1 To 9
        If ListView1.ListItems.Item(J).Checked = True Then
           For i = 1 To VGnumnivelescuenta
             If ListView1.ListItems.Item(J).Text = VG_aNIVELES(i - 1) Then
               nC = nC + 1
             End If
           Next
        End If
      Next
  
  Set rsX = New ADODB.Recordset
'  SQL = "SELECT count(*) FROM ct_cuenta WHERE cuentacodigo<>'00' "
'  Set rsX = VGcnx.Execute(SQL)
'  If rsX(0) > 0 And VGnumnivelescuenta <> nC Then
'    MsgBox "Existe información en el Plan de Cuentas", vbInformation, Caption
'    ValidarData = False
'    Exit Function
'  End If
  
  Set VGvardllgen = New dllgeneral.dll_general
  If VGvardllgen.ESNULO(txt(4).Text, 0) = 0 Then
    MsgBox "Debe registrar Valor(%) para el IGV", vbInformation, Caption
    ValidarData = False
    txt(4).SetFocus
    Exit Function
  End If
  
  If chk(3).Value = 0 And (Ctr_AyudaCuentaAjuste(0).xclave = Empty Or Ctr_AyudaCuentaAjuste(1).xclave = Empty) Then
    MsgBox "El Check de Ajuste en Linea esta inactivo y Cuentas de Ajuste estan sin Datos", vbInformation, Caption
    ValidarData = False
    Ctr_AyudaCuentaAjuste(0).SetFocus
    Exit Function
  End If

  ValidarData = True
End Function

Sub GrabarData()
On Error GoTo X
 Dim SQL  As String
 
 Dim ValorMoneda As String
 ValorMoneda = Ctr_moneda.xclave
 Set VGvardllgen = New dllgeneral.dll_general
  
  strvalor = NivelCuenta(0)
  strvalor1 = NivelCuenta(1)
    Call ParametroCuenta(1)
    
  If FlagNUEVO = True Then
    SQL = "INSERT INTO ct_sistema (sistemadescripcionempresa, sistemadescrcortaempresa,"
    SQL = SQL & "sistemaesttipodescrempresa, sistemadireccionempresa,sistemaempresaruc,"
    SQL = SQL & "sistemaestcuadreasiento,sistemaestimpresionasiento, sistemaconfiguracuenta,"
    SQL = SQL & "monedacodigo,sistemaultimonivel,sistemavalorigv,sistemaajustelinea,"
    SQL = SQL & "sistemactaajustedeb,sistemactaajustehab,usuariocodigo,fechaact,sistemaasientocodigo,"
    SQL = SQL & "sistemasubasientocodigo,sistemamonista,sistemaconfiguracentrocostos,sistemaultimonivelcostos) "
    SQL = SQL & "VALUES ('" & Trim$(txt(0).Text) & "','" & Trim$(txt(1).Text) & "'," & chk(0).Value & ""
    SQL = SQL & ",'" & Trim$(txt(2).Text) & "','" & Trim$(txt(3).Text) & "'," & chk(1).Value & "," & chk(2).Value
    SQL = SQL & ",'" & strvalor & "','" & ValorMoneda & "'," & VGnumnivelescuenta & "," & txt(4).Text
    SQL = SQL & "," & chk(4).Value & ",'" & Ctr_AyudaCuentaAjuste(0).xclave & "'"
    SQL = SQL & ",'" & Ctr_AyudaCuentaAjuste(1).xclave & "','" & VGusuario & "','" & Date & "'"
    SQL = SQL & ",'" & Trim$(Ctr_Codigo(0).xclave) & "','" & Trim$(Ctr_Codigo(1).xclave) & "'"
    SQL = SQL & "," & chk(4).Value & ",'" & strvalor1 & "'," & VGnumnivelescentrocosto & ")"
  Else
    SQL = "Update ct_sistema "
    SQL = SQL & "SET sistemadescripcionempresa='" & Trim$(txt(0).Text) & "'"
    SQL = SQL & ",sistemadescrcortaempresa='" & Trim$(txt(1).Text) & "'"
    SQL = SQL & ",sistemaesttipodescrempresa='" & chk(0).Value & "'"
    SQL = SQL & ",sistemadireccionempresa='" & Trim$(txt(2).Text) & "',sistemaempresaruc='" & txt(3).Text & "'"
    SQL = SQL & ",sistemaestcuadreasiento=" & chk(1).Value & ",sistemaestimpresionasiento=" & chk(2).Value
    SQL = SQL & ",sistemaconfiguracuenta='" & strvalor & "'"
    SQL = SQL & ",monedacodigo='" & ValorMoneda & "',sistemaultimonivel=" & VGnumnivelescuenta
    SQL = SQL & ",sistemavalorigv=" & VGvardllgen.ESNULO(txt(4).Text, 0)
    SQL = SQL & ",sistemaajustelinea=" & chk(3).Value & ",sistemactaajustedeb='" & Ctr_AyudaCuentaAjuste(0).xclave & "',sistemactaajustehab='" & Ctr_AyudaCuentaAjuste(1).xclave & "'"
    SQL = SQL & ",usuariocodigo='" & VGusuario & "',fechaact='" & Date & "',"
    SQL = SQL & "sistemaasientocodigo='" & Ctr_Codigo(0).xclave & "',"
    SQL = SQL & "sistemasubasientocodigo='" & Ctr_Codigo(1).xclave & "',"
    SQL = SQL & "sistemamonista=" & chk(4).Value & ""
    SQL = SQL & ",sistemaconfiguracentrocostos='" & strvalor1 & "'"
    SQL = SQL & ",sistemaultimonivelcostos=" & VGnumnivelescentrocosto
    
    
  End If
  VGCNx.Execute (SQL)
  cmdBotones(0).Enabled = False
  
  Exit Sub

X:
  MsgBox "Error inesperado: " & err.Description & "  " & err.Number
End Sub

Function NivelCuenta(Index As Integer) As String
 Dim i As Integer
 Dim valor As String
 valor = Empty
 Select Case Index
    Case 0
         For i = 1 To 9
            If ListView1.ListItems.Item(i).Checked = True Then
               valor = valor & ListView1.ListItems.Item(i).Text & "*"
            End If
         Next
    Case 1
         For i = 1 To 9
            If ListView2.ListItems.Item(i).Checked = True Then
               valor = valor & ListView2.ListItems.Item(i).Text & "*"
            End If
         Next
End Select
NivelCuenta = Left(valor, Len(valor) - 1)
End Function

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
  cmdBotones(0).Enabled = ValidaBoton()
End Sub

Private Sub txt_Change(Index As Integer)
  cmdBotones(0).Enabled = ValidaBoton()
End Sub

Private Sub chk_Click(Index As Integer)
  cmdBotones(0).Enabled = ValidaBoton()
End Sub

Function ValidaBoton() As Boolean
 Dim i As Integer
   For i = 0 To 3
     If txt(i).Text = Empty Then
       ValidaBoton = False
       Exit Function
     End If
   Next
   ValidaBoton = True
       
End Function

Private Sub txt_LostFocus(Index As Integer)
  txt(Index).Text = UCase$(txt(Index).Text)
End Sub
