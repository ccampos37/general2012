VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmCreacionSal 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   2685
   ClientLeft      =   1650
   ClientTop       =   7215
   ClientWidth     =   11880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   6  'Inside Solid
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command7 
      Caption         =   "&Salir"
      Height          =   615
      Left            =   7170
      Picture         =   "FrmCreacionSal.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1950
      Width           =   840
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Limpiar"
      Height          =   615
      Left            =   5850
      Picture         =   "FrmCreacionSal.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1950
      Width           =   840
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Enviar"
      Height          =   615
      Left            =   4560
      Picture         =   "FrmCreacionSal.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1950
      Width           =   840
   End
   Begin VB.Frame Frame1 
      Height          =   2595
      Left            =   0
      TabIndex        =   13
      Top             =   15
      Width           =   12045
      Begin VB.TextBox TxtCanref 
         Height          =   285
         Left            =   1890
         TabIndex        =   5
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txEquip 
         Height          =   285
         Left            =   1875
         MaxLength       =   10
         TabIndex        =   9
         Top             =   2055
         Visible         =   0   'False
         Width           =   1572
      End
      Begin VB.TextBox TxordFab 
         Height          =   285
         Left            =   9420
         MaxLength       =   10
         TabIndex        =   8
         Top             =   960
         Visible         =   0   'False
         Width           =   1488
      End
      Begin VB.TextBox txccosto 
         Height          =   285
         Left            =   4905
         MaxLength       =   10
         TabIndex        =   7
         Top             =   1320
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.TextBox TxDescri 
         Height          =   285
         Left            =   6990
         TabIndex        =   1
         Text            =   "TxDescri"
         Top             =   2115
         Visible         =   0   'False
         Width           =   4845
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   8130
         TabIndex        =   4
         Text            =   "Text3"
         Top             =   570
         Width           =   1215
      End
      Begin VB.TextBox TxtCantidad 
         Height          =   285
         Left            =   1890
         TabIndex        =   6
         Text            =   "Text5"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   5400
         TabIndex        =   3
         Top             =   570
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox Text6 
         Height          =   315
         Left            =   2040
         MaxLength       =   20
         TabIndex        =   2
         Top             =   600
         Width           =   2085
      End
      Begin VB.TextBox TxtArticulo 
         Height          =   285
         Left            =   8250
         MaxLength       =   20
         TabIndex        =   0
         Top             =   1740
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1875
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   540
         Visible         =   0   'False
         Width           =   2130
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuart 
         Height          =   375
         Left            =   1920
         TabIndex        =   35
         Top             =   120
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   661
         XcodMaxLongitud =   20
         xcodwith        =   1500
         NomTabla        =   "v_saldosxAlmacen"
         ListaCampos     =   "acodigo(1),adescri(1),acodigo2(2),aunidad(2)"
         XcodCampo       =   "acodigo"
         XListCampo      =   "adescri"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "acodigo,adescri,acodigo2,aunidad"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuAnalitico 
         Height          =   315
         Left            =   7320
         TabIndex        =   37
         Top             =   1320
         Visible         =   0   'False
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   556
         XcodMaxLongitud =   11
         xcodwith        =   900
         NomTabla        =   "gr_proyectos"
         TituloAyuda     =   "Busqueda de Proyectos"
         ListaCampos     =   "proyectocodigo(1),proyectodescripcion(1)"
         XcodCampo       =   "proyectocodigo"
         XListCampo      =   "proyectodescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "proyectocodigo,proyectodescripcion"
         Requerido       =   0   'False
      End
      Begin VB.Label Lblanalitico 
         AutoSize        =   -1  'True
         Caption         =   "Analitico"
         Height          =   195
         Left            =   6600
         TabIndex        =   36
         Top             =   1365
         Width           =   600
      End
      Begin VB.Label Label68 
         Caption         =   "Cant. Referencial"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   1350
         Width           =   1395
      End
      Begin VB.Label lblMaq 
         Caption         =   "Equipos/Maquinas"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   2085
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Label lblordfab 
         Alignment       =   2  'Center
         Caption         =   "Orden Fabricación"
         Height          =   240
         Left            =   7995
         TabIndex        =   32
         Top             =   990
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblccosto 
         Caption         =   "Centro de Costo"
         Height          =   255
         Left            =   3630
         TabIndex        =   31
         Top             =   1380
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Label Label7 
         Caption         =   "%"
         Height          =   255
         Left            =   9405
         TabIndex        =   30
         Top             =   645
         Width           =   255
      End
      Begin VB.Label lblPreciofin 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblPreciofin"
         Height          =   285
         Left            =   10755
         TabIndex        =   29
         Top             =   600
         Width           =   1200
      End
      Begin VB.Label Label6 
         Caption         =   "Precio Vta."
         Height          =   255
         Index           =   0
         Left            =   9870
         TabIndex        =   28
         Top             =   630
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lbcantstk 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbcantstk"
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
         Left            =   1890
         TabIndex        =   25
         Top             =   945
         Width           =   1575
      End
      Begin VB.Label Label16 
         Caption         =   "Nro Serie \ Lote"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   570
         Width           =   1575
      End
      Begin VB.Label lbEtiNum 
         Caption         =   "Num de Item:"
         Height          =   255
         Left            =   9495
         TabIndex        =   23
         Top             =   210
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10755
         TabIndex        =   22
         Top             =   210
         Width           =   495
      End
      Begin VB.Label Label15 
         Caption         =   "Cantidad en Stock"
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   1020
         Width           =   1365
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label14"
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
         Left            =   6930
         TabIndex        =   20
         Top             =   960
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Valor Vta."
         Height          =   255
         Left            =   4170
         TabIndex        =   19
         Top             =   615
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Unidad Estandar"
         Height          =   375
         Left            =   4140
         TabIndex        =   18
         Top             =   1005
         Width           =   1230
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo"
         Height          =   255
         Left            =   270
         TabIndex        =   17
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label lblUniEst 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblUniEst"
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
         Left            =   5400
         TabIndex        =   16
         Top             =   945
         Width           =   1485
      End
      Begin VB.Label Label12 
         Caption         =   "Descuento"
         Height          =   255
         Left            =   7140
         TabIndex        =   15
         Top             =   630
         Width           =   975
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Salida 
      Height          =   2535
      Left            =   180
      TabIndex        =   26
      Top             =   60
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   4471
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmCreacionSal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cant As Double
Dim I As Integer
Dim fin As Integer
Dim FACTOR As Double
Dim flagserie As String * 1
Dim flaglote As String * 1
Dim xserie As String * 1
Dim array_stldis() As Integer
Dim array_fecven() As Date
Dim rs1 As New ADODB.Recordset
Dim Rs2 As New ADODB.Recordset
'Dim db As Database
Dim dato_invalido As Boolean
Dim serie_lote As String

'Dim frm As Form

Private Sub Combo1_Click()
 If flagserie = "S" Then
    'flaglote = "S"
    ' lbcantstk = Str(array_stldis(Combo1.ListIndex + 1))  revisar
    'MaskEdBox1 = array_fecven(Combo1.ListIndex + 1)
     Command1.Enabled = True
     Command1.SetFocus
 End If
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   Text4.SetFocus
   Combo1_Validate (False)
End If
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
If Cancel Then
   Call Combo1_KeyDown(13, 1)
End If
End Sub

Public Sub Command1_Click()
'Enviar
Dim criterio As String
Dim ncombo As Integer
Dim kflag, J As Integer

    kflag = 0
    For J = 1 To FrmGuiaSal.MSFlexGrid1.Rows - 1
        If Trim(FrmGuiaSal.MSFlexGrid1.TextMatrix(J, 0)) = Trim(TxtArticulo) Then
           kflag = 1
           Exit For
        End If
    Next
    If kflag = 1 And VGSeleccion = 1 Then
       MsgBox "Ya existe el articulos...Verifique!!!", vbInformation, "AVISO"
       Exit Sub
    Else
      TxtArticulo = Trim(TxtArticulo)
    End If

    If (flagserie = "S") And Combo1.text = "" And Combo1.Visible Then
      MsgBox "El articulo no tiene serie para descargar", vbInformation, "Aviso"
      Exit Sub
    End If
    If (flaglote = "S") And Text6 = "" Then
      MsgBox "El articulo no tiene Lote para descargar", vbInformation, "Aviso"
      Exit Sub
    End If
    If Not IsNumeric(TxtCantidad.text) Then
           MsgBox "Ingrese cantidad respectiva", vbOKOnly + vbExclamation, "Error"
           TxtCantidad.SetFocus
           TxtCantidad.SelStart = 0: TxtCantidad.SelLength = Len(TxtCantidad)
           Exit Sub
    End If
    If Ctr_AyuAnalitico.Visible = True And Ctr_AyuAnalitico.xclave = "" Then
           MsgBox " Ingrese codigo de proyecto/equipo ", vbOKOnly + vbExclamation, "Error"
           Ctr_AyuAnalitico.SetFocus
           Exit Sub
    End If
    If Val(lbcantstk) < Val(TxtCantidad) Then  'And (VGRegEnt <> 1)
           MsgBox "La cantidad no puede ser mayor al stock", vbOKOnly + vbExclamation, "Error"
           TxtCantidad.Enabled = True: TxtCantidad.SetFocus
           Exit Sub
    End If
    If flaglote = "S" And (Text6 = "") Then 'And Not Combo1.Enabled
           MsgBox "Ingrese el Número de Lote", vbOKOnly + vbExclamation, "Error"
           Text6.Enabled = True: Text6.SetFocus
           Exit Sub
    End If
    If (flagserie = "S") And Combo1.Visible Then
       If FrmGuiaSal.MSFlexGrid1.Rows <> 1 Then
          For ncombo = 1 To FrmGuiaSal.MSFlexGrid1.Rows - 1
            If UCase(Combo1.text) = UCase(FrmGuiaSal.MSFlexGrid1.TextMatrix(ncombo, 2)) Then
              MsgBox "Ya se ingreso la serie", vbInformation, "Aviso"
              Combo1.SetFocus
              Exit Sub
            End If
          Next ncombo
        End If
      End If
      If flaglote = "S" Then
           If Not existe_lote1(Text6) Then
                 Text6.SetFocus
                  Exit Sub
           End If
      End If
        'carga ' revisar  verifica la conversion de unidades
        If Not dato_invalido Then Exit Sub
        'verifico si actualizo
        GuiaSalida
        limpia
        Combo1.Visible = False
        Text6.Visible = True
        TxtArticulo.Enabled = True
        I = I + 1             'contador de item
        Label11.Visible = True
        Label11 = I
        If I < fin Then
           DisplayDisp         'funcion de llenar los datos
           If flagserie = "S" Or flaglote = "S" Then
                  If flagserie = "S" Then
                     Combo1.Visible = True
                     Text6.Visible = False
                     Command1.Enabled = True
                     Command1.SetFocus
                  Else
                     If Text6.Visible And Text6.Enabled Then
                        Text6.SetFocus
                        Command1.Enabled = False
                     End If
                  End If
            Else
                  TxtCantidad.SetFocus
                  Command1.Enabled = False
           End If
        Else
           Command1.Enabled = False
           Label11.Visible = False
           Ctr_Ayuart.Enabled = True
           Ctr_Ayuart.SetFocus
           Ctr_Ayuart.xclave = "": Ctr_Ayuart.Ejecutar
           'SendKeys "{tab}"
        End If
        If VGSeleccion = 2 Then Unload Me  'Cuando es modificar
End Sub

Private Sub Command3_Click()
   limpia
   If TxtArticulo.Enabled And TxtArticulo.Visible Then TxtArticulo.SetFocus
End Sub

Private Sub Command7_Click()
    VGSeleccion = 1
    Unload Me
End Sub

Private Sub Ctr_AyuAnalitico_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
txEquip = Ctr_AyuAnalitico.xclave
End Sub

Private Sub Ctr_AyuAnalitico_AlNoDevolverNada()
txEquip = ""
End Sub

Private Sub Ctr_Ayuart_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Dim xsql As New ADODB.Recordset
TxtArticulo.text = Ctr_Ayuart.xclave
TxDescri = Ctr_Ayuart.xnombre
Set xsql = VGCNx.Execute(" select stskdis from stkart where stalma='" & FrmGuiaSal.TxtAlmacen & "' and stcodigo='" & Ctr_Ayuart.xclave & "'")
lbcantstk = ESNULO(xsql!STSKDIS, 0)
End Sub

Private Sub Form_Activate()
  
    If VGSeleccion = 2 Then
        'Data1.RecordSource = "SELECT ACODIGO FROM MaeArt "     'modificar
        Set rs1 = VGCNx.Execute("SELECT ACODIGO FROM MaeArt ")
        Text6.Enabled = True
        TxtCantidad.Enabled = True
        Command3.Enabled = False

        If Trim(TxtArticulo) = "" Then
           modifica_guia_salida
           colocastk
        End If
        Text6.Enabled = False
  End If
End Sub

Private Sub Form_Load()
  Dim criterio As String
   FACTOR = 1
   'Data1.DatabaseName = cRuta2
   'Data2.DatabaseName = cRuta2
  ' central FrmCreacionSal
   Me.Left = 100
   Me.Top = 5800
   Command1.Enabled = False
   deshabilitartx5_tx3 (False)
   Text6.Enabled = False
   'Label11.Visible = False
   lbEtiNum.Visible = False
   dato_invalido = True
   VGForm1 = 3     '*******************el codigo a la ayuda de art
   Text4.Visible = True
   limpia
  Call Ctr_Ayuart.conexion(VGCNx)
    If VGRegEnt = 2 Then
     Ctr_Ayuart.filtro = " stalma='" & VGAlma & "' and stskdis> 0  "
  Else
  Ctr_Ayuart.filtro = ""
  End If
  Call Ctr_AyuAnalitico.conexion(VGCNx)
  Ctr_AyuAnalitico.filtro = " tipoanaliticocodigo='" & VGParamSistem.tipoanaliticocodigo & "' and  isnull(proyectocierre,0)=0 "
  Call Ctr_Ayuart.conexion(VGCNx)

End Sub

Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
   Text6_DblClick
End If
End Sub

Private Sub txccosto_DblClick()
  Dim Adodc3 As ADODB.Recordset   'Centro de Costos
  Set Adodc3 = New ADODB.Recordset
  If UCase(Dir$(cRuta4)) = UCase(cNomBd4) Then
        Adodc3.Open "SELECT cencost_codigo,cencost_descripcion FROM centro_costos where  len(cencost_codigo) = '6' ", VGcnxCT, adOpenStatic, adLockOptimistic
  Else
        Adodc3.Open "SELECT cencost_codigo,cencost_descripcion FROM centro_costos ", VGCNx, adOpenStatic, adLockOptimistic
  End If
        frmReferencia.Conectar Adodc3, "SELECT cencost_codigo,cencost_descripcion FROM centro_costos  "
        frmReferencia.Label1.Caption = "Centro de Costos"
        frmReferencia.Show vbModal
        Adodc3.Close
        If vGUtil(1) <> "" Then
                 txccosto = vGUtil(1)
                 'LblCC = vGUtil(2)
        End If
        If txccosto.text <> "" Then txccosto_KeyPress (13)

End Sub

Private Sub txccosto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then txccosto_DblClick
End Sub

Private Sub txccosto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Existe(IIf(UCase(Dir$(cRuta4)) = UCase(cNomBd4), 3, 1), txccosto, "CENTRO_COSTOS", "cencost_codigo", False) = False Then
            MsgBox "La Cuenta no existe", vbInformation, "Mensaje"
            txccosto = ""
            txccosto.SetFocus: Exit Sub
        Else
            Tabula (KeyAscii)
        End If
    End If

End Sub

Private Sub TxDescri_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{TAB}"
End If
End Sub

Private Sub TxtArticulo_DblClick()
 Dim rs As Recordset
 Dim rql As String
   cant = 0
   I = 1                  'indica el numero del item
   'Load (FormAyuArt)
   If VGRegEnt <> 1 Then
        rql = "select  p.ACODIGO, p.ADESCRI,p.AUNIDAD, n.STSKDIS, p.AFSERIE, p.AFLOTE from MaeArt p, StkArt n where p.ACODIGO =  n.STCODIGO and n.STALMA = '" & VGAlma & "'  and  n.STSKDIS >0 "
        'Set RS = VGBaseDatos.OpenRecordset(rql, dbOpenSnapshot)
        Set rs = VGCNx.Execute(rql)
        If rs.RecordCount = 0 Then
             MsgBox "No hay articulos disponibles en el almacen", vbInformation, "Aviso"
             Exit Sub
        End If
        rs.Close
        Set rs = Nothing
   End If
   FormAyuArt.Show 1
   fin = Salida.Rows
   If Salida.Rows = 1 Then Exit Sub
   Label11.Visible = True
   lbEtiNum.Visible = False
   Label11 = I
   DisplayDisp 'Muestra los datosen pantalla
   If flagserie = "S" Or flaglote = "S" Then
        If flaglote = "S" Then
          xserie = "N"
          VGcod = TxtArticulo
          Text6.Visible = True
          Text6.Enabled = True
          Text6.SetFocus
          TxtCantidad.Enabled = True
        Else
          xserie = "S"
          Combo1.SetFocus
          TxtCantidad.Enabled = False
        End If
   Else
        xserie = "X"
        TxtCantidad.Enabled = True
        TxtCantidad.SetFocus
        txtcanref.SetFocus
   End If
End Sub
Private Sub TxtArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
   TxtArticulo_DblClick
ElseIf KeyCode = 46 Then
   TxDescri = ""
End If
End Sub

Private Sub TxtArticulo_KeyPress(KeyAscii As Integer)
  Dim rpta As Integer
  Dim criterio As String
  If KeyAscii = 13 Then
         TxtArticulo = UCase(TxtArticulo)
         If Trim(TxtArticulo) = "TEXTO" Then
            TxtCantidad = 0
            Label14 = ""
            lbcantstk = ""
            SendKeys "{TAB}"
            Exit Sub
         End If
         criterio = "ACODIGO = " & "'" + TxtArticulo.text + "'"
        ' Data1.Recordset.FindFirst criterio
         Set rs1 = VGCNx.Execute("SELECT * FROM MaeArt  WHERE " & criterio)
         If rs1.RecordCount > 0 Then
            TxDescri = rs1.Fields("ADESCRI")
            Label14.Caption = "" & rs1.Fields("AUNIDAD")
            VGabrev = Label14
            lblUniEst = Nombre_Unidad(VGabrev)
            flagserie = IIf(Not IsNull(rs1.Fields("AFSERIE")), rs1.Fields("AFSERIE"), "N")
            flaglote = IIf(Not IsNull(rs1.Fields("AFLOTE")), rs1.Fields("AFLOTE"), "N")
            cant = 0
            Label15.Visible = True
          
          
            criterio = " STCODIGO = '" & TxtArticulo.text & "'"
            criterio = criterio + "and  STALMA = '" & VGAlma & "'"
            Set Rs2 = VGCNx.Execute("SELECT * FROM STKART WHERE " & criterio)
            If Rs2.RecordCount > 0 Then
               If stockcomp Then
                  cant = Rs2.Fields("STSKDIS") - Rs2.Fields("STSKcom")
                Else
                  cant = Rs2.Fields("STSKDIS")
               End If
            End If
            Rs2.Close
            Set Rs2 = Nothing
            lbcantstk = cant
            TxtArticulo.Enabled = False
            ver_serie_lote
            If flagserie = "S" Or flaglote = "S" Then    ' crear funcion
                If flaglote = "S" Then
                   xserie = "N"
                   Text6.Visible = True
                   Text6.SetFocus
                Else
                   xserie = "S"
                End If
                VGcod = TxtArticulo
            Else
                xserie = "X"
                TxtCantidad.SetFocus
                txtcanref.SetFocus
            End If
         Else
             If Val(TxtArticulo) = 0 Then
                TxtArticulo_DblClick
             Else
                If Trim(TxtArticulo) <> "" Then
                   MsgBox "El Código de Articulo no existe ", vbInformation, "Aviso"
                End If
                TxtArticulo.SetFocus
                txtcanref.SetFocus
             End If
         End If
 Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
 End If
End Sub


Private Sub Text3_Change()
Dim a As Double
If Text4 <> "" Then
   If Not IsNumeric(Text3.text) And Text3 <> "" Then
        MsgBox "Ingrese un Numero", vbOKOnly + vbExclamation, "Error"
        Text3.SetFocus
   ElseIf Not IsNumeric(Text4.text) And TxtArticulo <> "" Then
        MsgBox "Ingrese un Numero", vbOKOnly + vbExclamation, "Error"
   Else
        Text4 = Val(Text4) * (100 - Val(Text3)) / 100
        lblPreciofin = Format(Val(Text4), "###0.0000")
   End If
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
       If Not IsNumeric(Text3) And Text3 <> "" Then
          MsgBox "Ingrese un valor numerico", vbInformation, mensaje1
          Exit Sub
       Else
         SendKeys "{tab}"
         KeyAscii = 0
       End If
  End If
End Sub

Private Sub Text4_Change()
   If Not IsNumeric(Text4.text) And TxtArticulo <> "" Then
       MsgBox "Ingrese un Numero", vbOKOnly + vbExclamation, "Error"
       Text4.SetFocus
   End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
   If IsNumeric(Text4) Then
     Text3.SetFocus
   Else
     SendKeys "{tab}"
   End If
 End If
End Sub

Private Sub TxtCanref_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If Len(Trim(txtcanref)) = 0 Then txtcanref = 0
    SendKeys "{tab}"
  End If
End Sub

Private Sub TxtCantidad_Change()
   If Not IsNumeric(TxtCantidad.text) And TxtArticulo <> "" Then
      MsgBox "Ingrese un Numero", vbOKOnly + vbExclamation, "Error"
      TxtCantidad.SetFocus
   Else
      Command1.Enabled = True
   End If
End Sub

Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Not IsNumeric(TxtCantidad.text) Then
           MsgBox "Ingrese un Numero", vbOKOnly + vbExclamation, "Error"
           TxtCantidad.SetFocus
      Else
           Command1.Enabled = True
           Tabula (KeyAscii)
      End If
    Else
      If ((Chr$(KeyAscii) < "0" Or Chr(KeyAscii) > "9")) And KeyAscii <> 8 And Chr(KeyAscii) <> "." Then KeyAscii = 0
    End If
   
End Sub

Public Sub DisplayDisp()
  'funcion de llenar los datos de formulario utilizando los datos MSflexGrid
  Dim criterio As String
   TxtArticulo = Salida.TextMatrix(I, 0)   'codigo
   TxDescri = Salida.TextMatrix(I, 1)  'descripcion
   VGabrev = Salida.TextMatrix(I, 2)  'UNIDAD
   flagserie = Salida.TextMatrix(I, 3) 'serie
   flaglote = Salida.TextMatrix(I, 4) 'serie
   criterio = " STCODIGO ='" & TxtArticulo.text & "'"
   criterio = criterio + " and STALMA = '" & VGAlma & "'"
   
   Set Rs2 = VGCNx.Execute("SELECT * FROM STKART WHERE " & criterio)
   'Data2.Recordset.FindFirst criterio
   cant = 0
   If Rs2.RecordCount > 0 Then
      If stockcomp Then
         cant = Rs2.Fields("STSKDIS") - Rs2.Fields("STSKcom")
       Else
          cant = Rs2.Fields("STSKDIS")
      End If
   End If
   Rs2.Close
   Set Rs2 = Nothing
   
   Label14 = VGabrev  ' label14 variable auxiliar
   lblUniEst = Nombre_Unidad(VGabrev)
   lbcantstk = cant
   TxtArticulo.Enabled = True
   ver_serie_lote
   lbEtiNum.Visible = True
   
End Sub
Private Sub ver_serie_lote()
    
  If flagserie = "S" Or flaglote = "S" Then
     Text6.Enabled = True
     If flagserie = "S" Then  'con guia de salida
       agregar_combo
       Combo1.Visible = True
       Text6.Visible = False
     End If
     If flaglote = "S" Then
       VGcod = TxtArticulo
     Else
       TxtCantidad = 1
       TxtCantidad.Enabled = False
       Command1.Enabled = True
     End If
   Else
     Text6.Visible = True
     Text6.Enabled = False
     Text6.BackColor = &H8000000F
   End If
   TxtCantidad.Enabled = True
End Sub


Function existe_lote1(text As TextBox) As Boolean
Dim rs As New ADODB.Recordset
Dim RSQL As String
   RSQL = "select STSLOTE from STKLOTE where STSALMA ='" & VGAlma & "' and STSCODIGO = '" & TxtArticulo & "' and STSLOTE = '" & text & "'"
   'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
   Set rs = VGCNx.Execute(RSQL)
   If rs.RecordCount > 0 Then
     'If Not graba Then MsgBox "Lote Registrado en Almacen", vbInformation, "Aviso"
     existe_lote1 = True
   Else
     MsgBox "Lote  No Registrado en Almacen", vbInformation, "Aviso"
     existe_lote1 = False
   End If
   rs.Close
   Set rs = Nothing
End Function
Function coduso(dato As String) As String
   Dim RSQL As String
   Dim rs As New ADODB.Recordset
   RSQL = "select UM_ABREV from TabUniMed where UM_NOMBRE ='" & dato & "'"
   'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
   Set rs = VGCNx.Execute(RSQL)
   If rs.RecordCount = 0 Then
    coduso = ""
   Else
    coduso = rs(0)
   End If
   rs.Close
   Set rs = Nothing
End Function

Function Nombre_Unidad(dato As String) As String
   Dim RSQL As String
   Dim rs As New ADODB.Recordset
   RSQL = "select UM_NOMBRE from TabUniMed where UM_ABREV ='" & dato & "'" '
   'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
   Set rs = VGCNx.Execute(RSQL)
   If rs.RecordCount = 0 Then
     Nombre_Unidad = ""
   Else
     Nombre_Unidad = rs(0)
   End If
   rs.Close
   Set rs = Nothing
End Function

Function preciovta(Cod As String) As Double
  Dim RSQL As String
  Dim rs As New ADODB.Recordset
  RSQL = "select APRECIO from maeart where ACODIGO='" & Trim(TxtArticulo) & "'"
  'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
  Set rs = VGCNx.Execute(RSQL)
  If rs.RecordCount = 0 Then
    preciovta = 0
  Else
    preciovta = rs(0)
  End If
  rs.Close
  Set rs = Nothing
End Function

Private Sub limpia()
   Label11 = ""
   TxDescri = ""
   lblUniEst = ""
   lblPreciofin = ""
   TxtArticulo.text = ""
   Text4.text = ""
   lbcantstk = ""
   Text3.text = ""
   TxtCantidad.Enabled = True
   TxtCantidad.text = ""
   TxtArticulo.Enabled = True
   Text6.Enabled = True
   Text6 = ""
   Text6.Enabled = False
   'Label11.Visible = False
   lbEtiNum.Visible = False
   Command1.Enabled = False
   'txEquip = ""
   'txccosto = ""
   'TxordFab = ""
   Combo1.Clear
   
End Sub

Private Sub agregar_combo()
  Dim rs As New ADODB.Recordset
  Dim RSQL As String
  Dim contador1 As Integer
  contador1 = 1
  If flagserie = "S" Then
     RSQL = "select stsserie from stkseri where  STSALMA='" & VGAlma & "' and STSCODIGO='" & TxtArticulo & "' and STSSKDIS<> 0"
  Else
     Exit Sub
  End If
  Combo1.Clear
  'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
  Set rs = VGCNx.Execute(RSQL)
  If rs.RecordCount = 0 Then Exit Sub
    rs.MoveLast
    rs.MoveFirst
    While Not rs.EOF
       Combo1.AddItem (rs(0))
       contador1 = contador1 + 1
       rs.MoveNext
    Wend
  rs.Close
  Set rs = Nothing
  Combo1.ListIndex = 0
End Sub

Private Sub Text6_DblClick()
  VGForm1 = 3
  If flaglote = "S" Then
    FormAyuLote.Show 1
    If Text6.text <> "" Then TxtCantidad.SetFocus
  End If
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 And Text6 <> "" Then
   If flaglote = "S" Then
     existe_lote Text6
   ElseIf flagserie = "S" Then
     TxtCantidad = "1"
     TxtCantidad.Enabled = False
     Command1.SetFocus
   Else
     SendKeys "{tab}"
   End If
 End If
End Sub
              
Private Sub deshabilitartx5_tx3(flag As Boolean)
   TxtCantidad.Enabled = flag
End Sub
 
Private Sub existe_lote(text As TextBox)
Dim rs As New ADODB.Recordset
Dim RSQL As String
   RSQL = "select STSLOTE, STSLKDIS,STSFECVEN from STKLOTE where  STSALMA ='" & VGAlma & "' and STSCODIGO = '" & TxtArticulo & "' and STSLOTE = '" & text & "'"
   'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
   Set rs = VGCNx.Execute(RSQL)
   If rs.RecordCount > 0 Then
     MsgBox "Lote Registrado en Almacen", vbInformation, "Aviso"
     lbcantstk = rs(1)
     If Val(lbcantstk) = 0 Then
         MsgBox "Ingrese otro Nro de Lote ", vbInformation, "Aviso"
         Text6.SetFocus
     Else
         TxtCantidad.SetFocus
     End If
   Else
     MsgBox "Lote no Registrado en Almacen", vbInformation, "Aviso"
     Text6.SetFocus
   End If
   rs.Close
   Set rs = Nothing
End Sub

Private Sub GuiaSalida()
      If flagserie = "S" Then
           serie_lote = IIf(VGSeleccion = 2, Text6, Combo1.text)
      Else
           serie_lote = Text6
      End If
      If VGSeleccion = 2 Then
         FrmGuiaSal.MSFlexGrid1.TextMatrix(FrmGuiaSal.MSFlexGrid1.Row, 2) = serie_lote  'serie
         FrmGuiaSal.MSFlexGrid1.TextMatrix(FrmGuiaSal.MSFlexGrid1.Row, 3) = Format(Val(TxtCantidad.text), "##0.00") 'çantidad
         FrmGuiaSal.MSFlexGrid1.TextMatrix(FrmGuiaSal.MSFlexGrid1.Row, 4) = VGabrev   '   unidad ref verificar ojo
         FrmGuiaSal.MSFlexGrid1.TextMatrix(FrmRegistro.MSFlexGrid1.Row, 5) = Val(Text4.text)   'Precio
         FrmGuiaSal.MSFlexGrid1.TextMatrix(FrmGuiaSal.MSFlexGrid1.Row, 6) = Format(Val(lblPreciofin), "##0.0000")  'Caption  'unidad principal
         FrmGuiaSal.MSFlexGrid1.TextMatrix(FrmGuiaSal.MSFlexGrid1.Row, 7) = xserie   'Caption  'unidad principal
       ' FrmRegistro.MSFlexGrid1.AddItem (TxtArticulo.Text & vbTab & txdescri & vbTab & TxtCantidad.Text & vbTab & Text9.Text& vbTab & label14)
         FrmGuiaSal.MSFlexGrid1.TextMatrix(FrmGuiaSal.MSFlexGrid1.Row, 8) = txccosto
         FrmGuiaSal.MSFlexGrid1.TextMatrix(FrmGuiaSal.MSFlexGrid1.Row, 9) = TxordFab
         FrmGuiaSal.MSFlexGrid1.TextMatrix(FrmGuiaSal.MSFlexGrid1.Row, 10) = txEquip
       Else
        pro_xserie                                       '     0                  1                2 SERI                3                    4                5              6
        FrmGuiaSal.MSFlexGrid1.AddItem (TxtArticulo.text & vbTab & TxDescri & vbTab & serie_lote & vbTab & Format(TxtCantidad.text, "###0.00") & vbTab & VGabrev & vbTab & Text4 & vbTab & Val(lblPreciofin) & vbTab & xserie & vbTab & txccosto & vbTab & TxordFab & vbTab & txEquip & vbTab & txtcanref)
       End If
End Sub

Private Sub pro_xserie()
  If flagserie = "S" Then
        xserie = "S"
        Exit Sub
  End If
  If flaglote = "S" Then
        xserie = "N"
        Exit Sub
  End If
  xserie = "X"
End Sub

Private Sub modifica_guia_salida()
      
      TxtArticulo.text = FrmGuiaSal.MSFlexGrid1.TextMatrix(FrmGuiaSal.MSFlexGrid1.Row, 0)
      Text6 = FrmGuiaSal.MSFlexGrid1.TextMatrix(FrmGuiaSal.MSFlexGrid1.Row, 2)
      TxtCantidad.text = FrmGuiaSal.MSFlexGrid1.TextMatrix(FrmGuiaSal.MSFlexGrid1.Row, 3)
      Ctr_Ayuart.xclave = TxtArticulo.text: Ctr_Ayuart.Ejecutar
      'DisplayDisp
      Label14 = FrmGuiaSal.MSFlexGrid1.TextMatrix(FrmGuiaSal.MSFlexGrid1.Row, 4)
      TxDescri = FrmGuiaSal.MSFlexGrid1.TextMatrix(FrmGuiaSal.MSFlexGrid1.Row, 1)
      TxDescri.Enabled = False
      lblUniEst = Nombre_Unidad(Label14)
      'Text3 = coduso(Label14)
      'lbcantstk = cant
      TxtArticulo.Enabled = False
      TxtArticulo.TabStop = False
      txccosto = FrmGuiaSal.MSFlexGrid1.TextMatrix(FrmGuiaSal.MSFlexGrid1.Row, 8)
      TxordFab = FrmGuiaSal.MSFlexGrid1.TextMatrix(FrmGuiaSal.MSFlexGrid1.Row, 9)
      txEquip = FrmGuiaSal.MSFlexGrid1.TextMatrix(FrmGuiaSal.MSFlexGrid1.Row, 10)
      
      If FrmGuiaSal.MSFlexGrid1.TextMatrix(FrmGuiaSal.MSFlexGrid1.Row, 7) = "X" Then
         Text6.Enabled = False
         Text6.BackColor = &H8000000F
         TxtCantidad.SetFocus
      Else
         Text6.SetFocus
         If FrmGuiaSal.MSFlexGrid1.TextMatrix(FrmGuiaSal.MSFlexGrid1.Row, 7) = "S" Then
              TxtCantidad.Enabled = False
         End If
         Text6.Enabled = True
      End If
End Sub

Private Sub colocastk()
  Dim cadena As String
   cadena = FrmGuiaSal.MSFlexGrid1.TextMatrix(FrmGuiaSal.MSFlexGrid1.Row, 0)
   If FrmGuiaSal.MSFlexGrid1.TextMatrix(FrmGuiaSal.MSFlexGrid1.Row, 7) = "S" Then
        seriestk
   ElseIf FrmGuiaSal.MSFlexGrid1.TextMatrix(FrmGuiaSal.MSFlexGrid1.Row, 7) = "N" Then
        lotestk
   Else
        llenastk
  End If
End Sub

Private Sub llenastk()
Dim RSQL As String
Dim rs As New ADODB.Recordset
  
   RSQL = "select  stskdis, stskmin,stskmax,stpunrep from stkart  WHERE STALMA='" & VGAlma & "' and  stcodigo ='" & TxtArticulo & "'"
   'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
   Set rs = VGCNx.Execute(RSQL)
   If rs.RecordCount > 0 Then
     lbcantstk = rs(0)
   Else
     lbcantstk = 0
   End If
   rs.Close
End Sub

Private Sub lotestk()
Dim rs As New ADODB.Recordset
Dim RSQL As String
Dim Lote As String
   Lote = FrmGuiaSal.MSFlexGrid1.TextMatrix(FrmGuiaSal.MSFlexGrid1.Row, 2)
   RSQL = "select  STSLKDIS from STKLOTE where STSALMA ='" & VGAlma & "' and STSCODIGO = '" & TxtArticulo & "' and STSLOTE = '" & Lote & "'"
   'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
   Set rs = VGCNx.Execute(RSQL)
   If rs.RecordCount > 0 Then
        lbcantstk = rs(0)
   Else
        lbcantstk = 0
   End If
   rs.Close
   Set rs = Nothing
   
End Sub

Private Sub seriestk()
Dim rs As Recordset
Dim RSQL As String
Dim Serie As String
   Serie = FrmGuiaSal.MSFlexGrid1.TextMatrix(FrmGuiaSal.MSFlexGrid1.Row, 2)
   RSQL = "select STSSKDIS from STKSERI where  STSALMA ='" & VGAlma & "' and STSCODIGO = '" & TxtArticulo & "' and STSSERIE = '" & Serie & "'"
   'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
   Set rs = VGCNx.Execute(RSQL)
   If rs.RecordCount > 0 Then
      lbcantstk = rs(0)
   Else
      lbcantstk = 0
   End If
   rs.Close
   Set rs = Nothing
   
End Sub
