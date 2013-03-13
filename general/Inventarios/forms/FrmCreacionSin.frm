VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmCreacionSin 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   3765
   ClientLeft      =   1650
   ClientTop       =   7215
   ClientWidth     =   11880
   ControlBox      =   0   'False
   DrawStyle       =   6  'Inside Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "&Enviar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   4110
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2370
      Width           =   1170
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Limpiar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   5370
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2370
      Width           =   1170
   End
   Begin VB.CommandButton Command7 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   6630
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2370
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      Height          =   3600
      Left            =   30
      TabIndex        =   12
      Top             =   -60
      Width           =   11895
      Begin VB.TextBox txtcanref 
         Height          =   375
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1680
         Width           =   1332
      End
      Begin VB.TextBox txccosto 
         Height          =   375
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   8
         Top             =   2040
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.TextBox TxordFab 
         Height          =   285
         Left            =   4890
         MaxLength       =   10
         TabIndex        =   7
         Top             =   1710
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.TextBox txEquip 
         Height          =   285
         Left            =   10200
         MaxLength       =   10
         TabIndex        =   19
         Top             =   2040
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.CheckBox chkserie 
         Caption         =   "Por Cantidades"
         Height          =   255
         Left            =   4080
         TabIndex        =   15
         Top             =   570
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox TxtArticulo 
         Height          =   285
         Left            =   960
         MaxLength       =   20
         TabIndex        =   0
         Top             =   3000
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.TextBox TxtCantidad 
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   4890
         TabIndex        =   5
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Text6 
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   1
         Top             =   600
         Width           =   2070
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   285
         Left            =   10380
         TabIndex        =   4
         Top             =   570
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   -2147483634
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   285
         Left            =   7350
         TabIndex        =   3
         Top             =   570
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   -2147483634
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1830
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   570
         Visible         =   0   'False
         Width           =   2100
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   7335
         TabIndex        =   20
         Top             =   570
         Visible         =   0   'False
         Width           =   1335
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuart 
         Height          =   375
         Left            =   1800
         TabIndex        =   39
         Top             =   240
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   661
         XcodMaxLongitud =   20
         xcodwith        =   1500
         NomTabla        =   "v_SaldosXAlmacen"
         ListaCampos     =   "acodigo(1),adescri(1),acodigo2(2),aunidad(2)"
         XcodCampo       =   "acodigo"
         XListCampo      =   "adescri"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "acodigo,adescri,acodigo2,aunidad"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuAnalitico 
         Height          =   315
         Left            =   7440
         TabIndex        =   41
         Top             =   1680
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
         Left            =   6720
         TabIndex        =   40
         Top             =   1725
         Width           =   600
      End
      Begin VB.Label LblPrecio 
         Height          =   345
         Left            =   1830
         TabIndex        =   38
         Top             =   2550
         Width           =   1725
      End
      Begin VB.Label lblccosto 
         AutoSize        =   -1  'True
         Caption         =   "Centro de Costo"
         Height          =   195
         Left            =   240
         TabIndex        =   37
         Top             =   2130
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Label lblccosto3 
         AutoSize        =   -1  'True
         Caption         =   "Merma"
         Height          =   195
         Left            =   240
         TabIndex        =   36
         Top             =   1680
         Width           =   480
      End
      Begin VB.Label lblordfab 
         Caption         =   "Orden Fabricación"
         Height          =   255
         Left            =   3480
         TabIndex        =   35
         Top             =   1740
         Visible         =   0   'False
         Width           =   1380
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4890
         TabIndex        =   17
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codigo"
         Height          =   195
         Left            =   240
         TabIndex        =   34
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Unidad Estandar"
         Height          =   195
         Left            =   3480
         TabIndex        =   33
         Top             =   1020
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Vcto."
         Height          =   255
         Left            =   6300
         TabIndex        =   32
         Top             =   600
         Width           =   1245
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad"
         Height          =   195
         Left            =   240
         TabIndex        =   31
         Top             =   1350
         Width           =   630
      End
      Begin VB.Label Label10 
         Caption         =   "Fecha Fab."
         Height          =   255
         Left            =   9420
         TabIndex        =   30
         Top             =   600
         Width           =   945
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label13"
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
         Left            =   7440
         TabIndex        =   14
         Top             =   3000
         Visible         =   0   'False
         Width           =   4485
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
         Height          =   255
         Left            =   6720
         TabIndex        =   18
         Top             =   960
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad en Stock"
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   1050
         Width           =   1320
      End
      Begin VB.Label Label6 
         Caption         =   "Unidad referencial"
         Height          =   255
         Left            =   3480
         TabIndex        =   28
         Top             =   1350
         Width           =   1335
      End
      Begin VB.Label label7 
         Caption         =   "Label7"
         Height          =   195
         Left            =   6720
         TabIndex        =   27
         Top             =   1350
         Width           =   675
      End
      Begin VB.Label LblCantidad 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label8"
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
         Left            =   7500
         TabIndex        =   13
         Top             =   1290
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8970
         TabIndex        =   26
         Top             =   1320
         Width           =   1980
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
         Left            =   10440
         TabIndex        =   25
         Top             =   270
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lbEtiNum 
         Caption         =   "Num de Item:"
         Height          =   255
         Left            =   9270
         TabIndex        =   24
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Nro Serie \ Lote"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   630
         Width           =   1140
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
         Height          =   375
         Left            =   1800
         TabIndex        =   16
         Top             =   960
         Width           =   1335
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Salida 
      Height          =   2535
      Left            =   360
      TabIndex        =   22
      Top             =   360
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   4471
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmCreacionSin"
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
Dim serie_lote As String
'Dim db As Database
Dim dato_invalido As Boolean
Dim nuevodet As Boolean
Dim ya_grabo_det As Boolean
Dim graba As Boolean
Dim hubo_error As Boolean
Dim estadocosto As Integer
'***********************************
Dim rsSTKART As New ADODB.Recordset
'************************************
Dim codigo As String
Dim varform As Form
Dim CANTIDAD As Double
Dim cantidadini As Double          ' contiene el valor el valor inicial de la cantidad a modificar
Dim VGDllGeneral As New dll_general

Private Sub Combo1_Click()
   Command1.Enabled = True
   Command1.SetFocus
End Sub
'Enviar

Private Sub Combo1_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
   Command1.Enabled = True
   Command1.SetFocus
 End If
End Sub

Public Sub Command1_Click()
Dim criterio As String
Dim dato1 As String
Dim ncombo As Integer
Dim kflag, J As Integer
graba = True

kflag = 0
For J = 1 To FrmRegistro.MSFlexGrid1.Rows - 1
    If FrmRegistro.CENTROCOSTO = 1 Then
       criterio = Trim(FrmRegistro.MSFlexGrid1.TextMatrix(J, 0)) + Trim(FrmRegistro.MSFlexGrid1.TextMatrix(J, 2)) + Trim(FrmRegistro.MSFlexGrid1.TextMatrix(J, 11))
       dato1 = Trim(TxtArticulo) + Trim(Text6) + Trim(txccosto)
     Else
       criterio = Trim(FrmRegistro.MSFlexGrid1.TextMatrix(J, 0)) + Trim(FrmRegistro.MSFlexGrid1.TextMatrix(J, 2))
       dato1 = Trim(TxtArticulo) + Trim(Text6)
    End If
    If criterio = dato1 Then
       kflag = 1
       Exit For
    End If
Next
If kflag = 1 Then
   If Trim(Text6) <> "" Then
      MsgBox "Ya existe el lote para el articulo...Verifique!!!", vbInformation, "AVISO"
    ElseIf Trim(txccosto) <> "" Then
            MsgBox "Ya existe el articulo + centro de costos ...Verifique!!!", vbInformation, "AVISO"
         Else
           MsgBox "Ya existe el articulo...Verifique!!!", vbInformation, "AVISO"
   End If
   Exit Sub
Else
  TxtArticulo = Trim(TxtArticulo)
End If

If Not IsNumeric(TxtCantidad.text) Then
       MsgBox "Ingrese cantidad respectiva", vbOKOnly + vbExclamation, "Error"
       TxtCantidad.SetFocus
       TxtCantidad.SelStart = 0: TxtCantidad.SelLength = Len(TxtCantidad)
       Exit Sub
End If

If Ctr_AyuAnalitico.Enabled = True And Ctr_AyuAnalitico.xclave = "" Then
           MsgBox " Ingrese codigo de proyecto/equipo ", vbOKOnly + vbExclamation, "Error"
           Ctr_AyuAnalitico.SetFocus
           Exit Sub
End If
If Val(lbcantstk) < Val(TxtCantidad) And (VGRegEnt <> 1) Then
    MsgBox "La cantidad no puede ser mayor al stock", vbOKOnly + vbExclamation, "Error"
    If TxtCantidad.Enabled Then TxtCantidad.SetFocus
    Exit Sub
End If
If flagserie = "S" And (VGRegEnt = 1) And Text6 = "" Then 'And Not Combo1.Enabled
     MsgBox "Ingrese el Número de serie", vbOKOnly + vbExclamation, "Error"
     Text6.SetFocus
     Exit Sub
End If
If flaglote = "S" And (Text6 = "") Then 'And Not Combo1.Enabled
     MsgBox "Ingrese el Número de Lote", vbOKOnly + vbExclamation, "Error"
     Text6.SetFocus
     Exit Sub
  End If
If (flagserie = "S") Then
    If FrmRegistro.MSFlexGrid1.Rows <> 1 Then
        For ncombo = 1 To FrmRegistro.MSFlexGrid1.Rows - 1
          If Combo1.Visible Then
            If UCase(Combo1.text) = UCase(FrmRegistro.MSFlexGrid1.TextMatrix(ncombo, 2)) Then
              MsgBox "Ya se ingreso la serie", vbInformation, "Aviso"
              Combo1.SetFocus
              Exit Sub
            End If
          ElseIf Text6 <> "" Then
            If UCase(Text6.text) = UCase(FrmRegistro.MSFlexGrid1.TextMatrix(ncombo, 2)) Then
              MsgBox "Ya se ingreso la serie", vbInformation, "Aviso"
              Text6.SetFocus
              Exit Sub
            End If
          End If
        Next ncombo
    End If
End If
If (flagserie = "S") And (VGRegEnt = 1) Then
   If existe_serie(Text6) Then Exit Sub
End If
If flaglote = "S" And (VGRegEnt <> 1) Then
   If Not existe_lote1(Text6) Then
        Text6.SetFocus
        Exit Sub
   End If
End If
If flagserie = "N" And VGSeleccion <> 2 Then Carga ' revisar  verifica la conversion de unidades
If Not dato_invalido Then Exit Sub
If VGForm <> 6 Or VGEstadomodi Or VGtipocreacion = 2 Then
    'verifico si actualizo
                    'GuiaSalida
    ' Else
                   'ingreso o salida
     CANTIDAD = Val(TxtCantidad)
     ingreso_salida
     If hubo_error Then
         MsgBox "No pudo registrar", vbInformation, "Aviso"
     End If
     If VGEstadomodi Then
          Unload Me
          Exit Sub
     End If
End If
limpia
'*************************
'FrmRegistro.buscar_trans
'*************************
Combo1.Visible = False
Text6.Visible = True
TxtArticulo.Enabled = True
'Entra a las multiples opciones

 
 If I <> 0 Then         'solo entra cuando hay  dato en temporal de salida
   I = I + 1             'contador de item
   Label11.Visible = True
   Label11 = I
   
   If I < fin Then
                DisplayDisp         'funcion de llenar los datos
                If flagserie = "S" Or flaglote = "S" Then
                       If (VGRegEnt <> 1) And (flagserie = "S") Then
                          Combo1.Visible = True
                          Combo1.SetFocus
                          Command1.Enabled = True
                          TxtCantidad.Enabled = False
                          Text6.Visible = False
                       ElseIf flagserie = "S" Then
                           MaskEdBox1.BackColor = &H8000000F
                           MaskEdBox2.BackColor = &H80000009
                           MaskEdBox1.Enabled = False
                           MaskEdBox2.Enabled = True
                           Text6.SetFocus
                       Else
                           MaskEdBox1.BackColor = &H80000009
                           MaskEdBox2.BackColor = &H80000009
                           MaskEdBox1.Enabled = True
                           MaskEdBox2.Enabled = True
                           Text6.SetFocus
                       End If
                 Else
                        MaskEdBox1.BackColor = &H8000000F
                        MaskEdBox2.BackColor = &H8000000F
                       TxtCantidad.Enabled = True
                       TxtCantidad.SetFocus
                End If
      Else
                Command7.SetFocus
      End If
End If
Command1.Enabled = False
Ctr_Ayuart.SetFocus
Ctr_Ayuart.xclave = "": Ctr_Ayuart.Ejecutar
graba = False
If VGSeleccion = 2 Or VGSeleccion = 3 Then
    Unload Me
End If
End Sub

Private Sub Command3_Click()
limpia
FrmRegistro.buscar_trans
TxtArticulo.SetFocus
End Sub

Private Sub Command7_Click()
Label9.Caption = ""
lbcantstk = ""
If VGtipocreacion = 2 And TxtArticulo <> "" And VGSeleccion = 2 Then
    If Not ya_grabo_det Then
      reactualizastk (TxtArticulo)
    End If
End If
Unload Me
End Sub

Private Sub Chkserie_Click()
 If chkserie.Value Then
        formIngSerie.Show 1
 End If
End Sub

Private Sub Ctr_AyuAnalitico_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
txEquip = Ctr_AyuAnalitico.xclave
End Sub

Private Sub Ctr_Ayuart_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Dim xsql As New ADODB.Recordset
TxtArticulo.text = Ctr_Ayuart.xclave
Label13 = Ctr_Ayuart.xnombre
Set xsql = VGCNx.Execute(" select stskdis from stkart where stalma='" & FrmRegistro.Text11.text & "' and stcodigo='" & Ctr_Ayuart.xclave & "'")
lbcantstk = ESNULO(xsql!STSKDIS, 0)
End Sub

Private Sub Form_Load()
Dim criterio As String
If VGSeleccion = 2 Then
    'Inicializa para modificar
    codigo = varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 0)
    Ctr_Ayuart.xclave = codigo: Ctr_Ayuart.Ejecutar
    Text6.Enabled = True
    deshabilitartx5_tx3 (True)
    Command3.Enabled = False
    If TxtArticulo.text = "" Then
       modifica_ingreso_salida
    End If
    colocastk
    Text6.Enabled = False
Else
    Label15.Visible = True '
End If
If VGEstadomodi Then             'viene del formulario modificar
    Text6.Enabled = False        'SE puede modificar la serie y sus demas datos
    MaskEdBox1.Enabled = False   ' DEPENDE SI TIENE SERIE POR MEJORAR
    MaskEdBox2.Enabled = False
    VGtipocreacion = 2
    
End If
Set rsSTKART = VGCNx.Execute("Select * from STKART WHERE STALMA='" & VGAlma & "'")
Call Ctr_AyuAnalitico.conexion(VGCNx)
Ctr_AyuAnalitico.filtro = " tipoanaliticocodigo='" & VGParamSistem.tipoanaliticocodigo & "' and  isnull(proyectocierre,0)=0 "
Call Ctr_Ayuart.conexion(VGCNx)
If VGRegEnt = 1 Then
   Ctr_Ayuart.NomTabla = "maeart"
 ElseIf VGRegEnt = 0 Then
   Ctr_Ayuart.NomTabla = "v_SaldosXAlmacen"
   Ctr_Ayuart.filtro = " stalma='" & VGAlma & "' and  isnull(stskdis,0)> 0 "
End If
Left = (Screen.Width - Me.Width) / 2
Top = Screen.Height - Me.Height
graba = False
FACTOR = 1

central FrmCreacionSin
nuevodet = False
ya_grabo_det = False
Command1.Enabled = False
deshabilitartx5_tx3 (False)
Text6.Enabled = False
Label11.Visible = False
lbEtiNum.Visible = False
dato_invalido = True
VGForm1 = 2
limpia
  
Select Case VGtipocreacion
     Case 1
            Set varform = FrmRegistro
     Case 2
            Set varform = FrmModificar
End Select
   'revisar cuando viene de modificar
   'VGRegEnt = 1  en cualquier formulario  significa entrada
If VGRegEnt = 1 Then
     Label7.Caption = "Cantidad a Entrar "
Else
     Label7.Caption = "Cantidad a Salir"
     Label10.Visible = False
     MaskEdBox2.Visible = False
     Label6.Visible = False
     Text3.Visible = False
End If

Command1.Picture = MDIPrincipal.ImageList2.ListImages.item("Insertar").Picture
Command3.Picture = MDIPrincipal.ImageList2.ListImages.item("Sacar").Picture
Command7.Picture = MDIPrincipal.ImageList2.ListImages.item("Retornar").Picture

End Sub


Private Sub MaskEdBox1_GotFocus()
  MaskEdBox1.SelStart = 0: MaskEdBox1.SelLength = Len(MaskEdBox1)
End Sub

Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If MaskEdBox1 = "__/__/____" Then
        ' MsgBox "Ingrese  Fecha de Vencimiento ", vbExclamation + vbOKOnly, "Advertencia"
         'MaskEdBox1.SetFocus
         'Exit Sub                             'Cambios de la version 6
    End If
    If MaskEdBox2.Visible Then
         MaskEdBox2.SetFocus
    Else
        TxtCantidad.SetFocus
'         SendKeys "{tab}"
'         KeyAscii = 0
    End If
End If
End Sub

Private Sub MaskEdBox1_LostFocus()
Dim cValor As String

If MaskEdBox1 = "__/__/____" Then
Else
    cValor = ValidFecha(MaskEdBox1)
    If cValor = "" Then
       MsgBox "Ingrese la Fecha Correctamente", vbExclamation + vbOKOnly, "Advertencia"
       MaskEdBox1 = "__/__/____"
       MaskEdBox1.SetFocus
       Exit Sub
    Else
      MaskEdBox1 = cValor
    End If
    If CDate(MaskEdBox1) < Date Then
        MsgBox "El articulo ya vencio", vbExclamation + vbOKOnly, "Error"
        MaskEdBox1.SetFocus
    End If
End If
End Sub

Private Sub MaskEdBox2_GotFocus()
  MaskEdBox1.SelStart = 0: MaskEdBox1.SelLength = Len(MaskEdBox1)
End Sub

Private Sub MaskEdBox2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
      If MaskEdBox2 = "__/__/____" And VGRegEnt = 1 Then
        MsgBox "Ingrese  Fecha  de Fabricación ", vbExclamation + vbOKOnly, "Advertencia"
        MaskEdBox2.SetFocus
        'Exit Sub                  'Cambios de la version 6
      End If
      ' TxtCantidad.SetFocus
      SendKeys "{tab}"
      KeyAscii = 0
End If
End Sub

Private Sub MaskEdBox2_LostFocus()
Dim cValor As String

If MaskEdBox2 = "__/__/____" Then
Else
    cValor = ValidFecha(MaskEdBox2)
    If cValor = "" Then
        MsgBox "Ingrese la Fecha Correctamente", vbExclamation + vbOKOnly, "Advertencia"
        MaskEdBox2 = "__/__/____"
        MaskEdBox2.SetFocus
    ElseIf CDate(cValor) > Date Then
'        MsgBox "fecha de fab. es mayor que la fecha actual", vbExclamation + vbOKOnly, "Error"
'        MaskEdBox2.SetFocus
'    ElseIf MaskEdBox2 < CDate(cValor) Then
'        MsgBox "Ingrese fecha Valida", vbExclamation + vbOKOnly, "Error"
'        MaskEdBox2.SetFocus
    Else
        MaskEdBox2 = cValor
    End If
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     SendKeys "{tab}"
     KeyAscii = 0
  End If
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
If Cancel Then
   Text6_KeyPress (13)
End If
End Sub

Private Sub txccosto_DblClick()
  Dim Adodc3 As ADODB.Recordset   'Centro de Costos
  Set Adodc3 = New ADODB.Recordset
 Adodc3.Open "SELECT centrocostocodigo,centrocostodescripcion FROM ct_centrocosto where empresacodigo='" & VGParametros.empresacodigo & "' and  centrocostonivel = 3", VGcnxCT, adOpenStatic, adLockOptimistic
 frmReferencia.Conectar Adodc3, "SELECT centrocostocodigo,centrocostodescripcion FROM ct_centrocosto where empresacodigo='" & VGParametros.empresacodigo & "' and  centrocostonivel = 3"
 frmReferencia.Label1.Caption = "Centro de Costos"
 frmReferencia.Show vbModal

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
        If Existe(3, txccosto.text, "ct_CENTROCOSTO", "centrocostocodigo", False) = False Then
            MsgBox "Centro de Costo no existe", vbInformation, "Mensaje"
            txccosto = ""
            txccosto.SetFocus: Exit Sub
        Else
            Tabula (KeyAscii)
        End If
    End If

End Sub

Private Sub txEquip_KeyPress(KeyAscii As Integer)
Tabula (KeyAscii)
End Sub

Private Sub TxordFab_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then frm_manten_ordfabri.Show 1
End Sub

Private Sub TxordFab_KeyPress(KeyAscii As Integer)
Tabula (KeyAscii)
End Sub

Private Sub TxtArticulo_DblClick()
On Error Resume Next
cant = 0
I = 1
'Load (FormAyuArt)

VGForm1 = 2
FormAyuArt.Show 1
fin = Salida.Rows
If Salida.Rows = 1 Then Exit Sub
Label11.Visible = True
lbEtiNum.Visible = False
Label11 = I
DisplayDisp
If flagserie = "S" Or flaglote = "S" Then
          If flaglote = "S" Then
                MaskEdBox1.Enabled = True
                MaskEdBox2.Enabled = True
                xserie = "N"
                VGcod = TxtArticulo
                Text6.Visible = True
                Text6.Enabled = True
                Text6.SetFocus
                TxtCantidad.Enabled = True
          Else
                xserie = "S"
                chkserie.Enabled = True
                TxtCantidad.Enabled = False
                TxtCantidad = "1"
                Command1.Enabled = True
                If VGRegEnt <> 1 Then
                        Combo1.Visible = True
                        Text6.Visible = False
                        Combo1.SetFocus
                Else
                        If Text6.Enabled = True Then Text6.SetFocus
                End If
          End If
Else
          xserie = "X"
          TxtCantidad.Enabled = True
          TxtCantidad.SetFocus
'          txtcanref.SetFocus
End If
'Text3.Enabled = True

End Sub
Private Sub TxtArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 112 Then
       If VGRegEnt = 1 Then
          VGAlma = "" & Trim(FrmRegistro.Text11)
       Else
          VGAlma = "" & Trim(FrmRegistro.Text11)
       End If
       TxtArticulo_DblClick
ElseIf KeyCode = 46 Then
       Label13 = ""
End If
End Sub

Private Sub TxtArticulo_KeyPress(KeyAscii As Integer)
 Dim rpta As Integer
 Dim criterio As String
 Dim rsa, RSB As New ADODB.Recordset
  
If KeyAscii = 13 Then
         'If Len(TxtArticulo.text) = VGLongCodigo Then
        If Not Validadato(TxtArticulo) Then
          MsgBox "CODIGO NO VALIDO....!!", vbInformation, "AVISO"
          Call VGDllGeneral.Enfoquetexto(TxtArticulo)
          Exit Sub
        
        End If
        criterio = " where a.ACODIGO = " & "'" + TxtArticulo.text + "'"
        'Data1.Recordset.FindFirst criterio
        'Set rsa = VGCNx.Execute("Select * from MAEART WHERE " & criterio)
        Set rsa = VGCNx.Execute("Select A.*,B.PRODUCTOPRECVTA from MAEART A INNER JOIN LISTAPRE1 B ON A.ACODIGO=B.PRODUCTOCODIGO" & criterio)
        If rsa.RecordCount > 0 Then
           Set rsSTKART = VGCNx.Execute("Select * from STKART WHERE STALMA='" & VGAlma & "'")
           
           Label13.Caption = rsa.Fields("ADESCRI")
           Label14.Caption = "" & rsa.Fields("AUNIDAD")
           LblPrecio.Caption = ESNULO(rsa.Fields("productoprecvta"), 0)
           VGabrev = Label14
           lblUniEst = Nombre_Unidad(VGabrev)
           flagserie = IIf(Not IsNull(rsa.Fields("AFSERIE")), rsa.Fields("AFSERIE"), "N")
           flaglote = IIf(Not IsNull(rsa.Fields("AFLOTE")), rsa.Fields("AFLOTE"), "N")
           Label15.Visible = True
           'esto esta dentro un procedimiento
           criterio = " STCODIGO ='" & TxtArticulo.text & "' and  STALMA ='" & VGAlma & "'"
           'Data2.Recordset.FindFirst criterio
           'RMM ****************************************************
           rsSTKART.Filter = criterio
           'RMM ****************************************************
           If Not rsSTKART.EOF Then
               If stockcomp Then
                 cant = numero(rsSTKART("STSKDIS")) - numero(rsSTKART("STSKcom"))
                 Else
                 cant = numero(rsSTKART("STSKDIS"))
              End If
           Else
               cant = 0
           End If
           
           lbcantstk = cant
           TxtArticulo.Enabled = False
           ver_serie_lote
           If flagserie = "S" Or flaglote = "S" Then    ' crear funcion
                If flaglote = "S" Then
                        MaskEdBox1.Enabled = True
                        MaskEdBox2.Enabled = True
                        xserie = "N"
                        VGcod = TxtArticulo
                        SendKeys "{tab}"
                        KeyAscii = 0
                Else
                        xserie = "S"
                        chkserie.Enabled = True
                        If VGRegEnt <> 1 Then
                           Combo1.Visible = True
                           Text6.Visible = False
                           agregar_combo
                           Combo1.SetFocus
                        Else
                            Text6.SetFocus
                        End If
                End If
           Else
                xserie = "X"
                TxtCantidad.SetFocus
                'txtcanref.SetFocus
           End If
        Else
        
             If Val(TxtArticulo) = 0 Then
                TxtArticulo_DblClick
             ElseIf TxtArticulo <> "" Then
                    MsgBox "El Código de Articulo no existe ", vbExclamation, mensaje1
               End If
               'Cambios de la version 6
'            rpta = MsgBox("Desea registrar un nuevo articulo", vbYesNo + vbQuestion + vbDefaultButton2, "Crear un nuevo Articulo")
'            If rpta = vbYes Then
'               VGcrea = True
'               FormArticulos.show 1
'               VGcrea = False
'               Text3.Enabled = True
'               TxtCantidad.Enabled = True
'            Else
             TxtArticulo.SetFocus
             txtcanref.SetFocus
'          End If
        End If
Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
I = 0        ' para establecer que no hay nada seleccionado
End Sub

Private Sub Text3_DblClick()
'Dim db As Database
Dim rs As New ADODB.Recordset
Dim RSQL As String
Dim FACTOR As Double
If VGRegEnt = 1 Then
    Frmayuunidades.Show 1
    If TxtCantidad <> "" Then Command1.Enabled = True
    Carga
    If Not dato_invalido Then Exit Sub
End If
End Sub

Private Sub TxtCanref_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Len(Trim(txtcanref)) = 0 Then txtcanref = 0
      SendKeys "{tab}"
   End If
End Sub

Private Sub TxtCantidad_Change()
If TxtCantidad <> "" Then
    If TxtCantidad.Enabled Then
        If Not IsNumeric(TxtCantidad.text) And TxtArticulo <> "" Then
            MsgBox "Ingrese un Numero", vbOKOnly + vbExclamation, "Error"
            'If TxtCantidad.Visible And TxtCantidad.Enabled Then TxtCantidad.SetFocus
            'MOMENTO DE MODIFICAR
        Else
            If IsNumeric(TxtCantidad) Then
                LblCantidad = Val(TxtCantidad) * FACTOR             'entra siempre al momento de editar
                If Label9 = "" Then Label9 = lblUniEst
                Command1.Enabled = True
            End If
        End If
    End If
End If
End Sub

Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Not IsNumeric(TxtCantidad.text) And TxtArticulo <> "" Then
        MsgBox "Ingrese la cantidad", vbOKOnly + vbExclamation, "Error"
        Tabula (KeyAscii)
    Else
        'debe revisar solo cuando tenga tipo de unidad
        If VGtipocreacion = 1 Then Carga '   devuelve dato_invalido=false  cuando se produjo error
        If Not dato_invalido Then Exit Sub
        If Label13 <> "" Then
            Command1.Enabled = True
            Tabula (KeyAscii)
       End If
   End If
 Else
        If Chr(KeyAscii) = "." And IsNumeric(TxtCantidad) Then Exit Sub
        If ((Chr$(KeyAscii) < "0" Or Chr(KeyAscii) > "9")) And KeyAscii <> 8 Then KeyAscii = 0
 End If
End Sub

Public Sub DisplayDisp()
  'funcion de llenar los datos de formulario utilizando los datos MSflexGrid
Dim criterio As String
Dim RSQL As New ADODB.Recordset
On Error GoTo Err
If I > Salida.Rows - 1 Then Exit Sub
TxtArticulo = Salida.TextMatrix(I, 0)   'codigo
If TxtArticulo.text = "" Then Exit Sub
Label13 = Salida.TextMatrix(I, 1)  'descripcion
VGabrev = Salida.TextMatrix(I, 2)  'UNIDAD
flagserie = Salida.TextMatrix(I, 3) 'serie
flaglote = Salida.TextMatrix(I, 4) 'serie
criterio = "STCODIGO ='" & TxtArticulo.text & "' and STALMA ='" & VGAlma & "'"
   'RMM ****************************************************
criterio = "select * from stkart where " & criterio
   'RMM ****************************************************
   Set RSQL = VGCNx.Execute(criterio)
      
   If RSQL.RecordCount() > 0 Then
      'Data2.Recordset.FindFirst criterio
      If stockcomp Then
         cant = ESNULO(RSQL!STSKDIS, 0) - ESNULO(RSQL!STSKcom, 0)
       Else
         cant = ESNULO(RSQL!STSKDIS, 0)
      End If
   Else
      cant = 0
   End If

Label14 = VGabrev  ' label14 variable auxiliar
lblUniEst = Nombre_Unidad(VGabrev)
lbcantstk = cant
Text3.text = lblUniEst
 'TxtArticulo.Locked = True
ver_serie_lote
lbEtiNum.Visible = True
Exit Sub
Err:
   MsgBox Err.Description
  Resume
End Sub

Public Sub Carga()
Dim criterio1 As String
''Dim db As Database
Dim rs As New ADODB.Recordset
Dim RSQL As String
Dim FACTOR As Double
FACTOR = 1
If Trim(Label14) <> Trim(VGabrev) Then                          'CONSULTA POR DEFECTO MODIFICAR
        RSQL = "select  p.EQCANTEQUI from TabEqui p where p.EQUNIPRI = '" & VGabrev & "'   and p.EQUNIEQUI = '" & Label14.Caption & "'"
        Set rs = VGCNx.Execute(RSQL)
        If rs.RecordCount = 0 Then
            MsgBox "la unidad de referencia no tiene unidad equivalente"
            lblUniEst = Nombre_Unidad(Label14)
            Exit Sub
        End If
        rs.MoveFirst
        FACTOR = rs.Fields("EQCANTEQUI")
        rs.Close
  Else
        FACTOR = 1
  End If
cant = Val(lbcantstk)
  
  If cant < Val(TxtCantidad.text) * FACTOR And VGRegEnt = 0 Then  ' revisar si validar en creacion
        MsgBox "No hay stock suficente", 48, "Aviso"
        TxtCantidad.SetFocus
        dato_invalido = False
        Exit Sub
  End If
  dato_invalido = True
  LblCantidad = Val(TxtCantidad.text) * FACTOR 'VGcant
  lblUniEst = Nombre_Unidad(Label14)
End Sub

Function coduso(dato As String) As String
Dim RSQL As String
Dim rs As New ADODB.Recordset
RSQL = "select UM_ABREV from TabUniMed where UM_NOMBRE ='" & dato & "'"
'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
Set rs = VGCNx.Execute(RSQL)
If rs.EOF Then
    coduso = ""
Else
    coduso = rs(0)
End If
rs.Close
End Function

Function Nombre_Unidad(dato As String) As String
Dim RSQL As String
Dim rs As New ADODB.Recordset
RSQL = "select UM_NOMBRE from TabUniMed where UM_ABREV ='" & dato & "'" '   AND UM_ESTADO ='A'"
'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
Set rs = VGCNx.Execute(RSQL)
If rs.EOF Then
        Nombre_Unidad = ""
Else
        Nombre_Unidad = rs(0)
End If
rs.Close
End Function

Function preciovta(Cod As String) As Double
Dim RSQL As String
Dim rs As Recordset
RSQL = "select APRECIO from maeart where ACODIGO='" & TxtArticulo & "'"
'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
Set rs = VGCNx.Execute(RSQL)
If rs.EOF Then
    preciovta = 0
Else
    preciovta = rs(0)
End If
rs.Close
End Function

Private Sub limpia()
   Label11 = ""
   Label13.Caption = ""
   Label14.Caption = ""
   lblUniEst = ""
   LblCantidad.Caption = ""
   Label9.Caption = ""
   TxtArticulo.text = ""
   Text4.text = ""
   lbcantstk = ""
   Text3.text = ""
   MaskEdBox1 = "__/__/____"
   Text6.Enabled = True
   Text6.text = ""
   Text6.Enabled = False
   TxtCantidad.Enabled = True
   TxtCantidad.text = ""
   MaskEdBox2 = "__/__/____"
   MaskEdBox1.BackColor = &H80000009
   Text6.BackColor = &H80000009
   MaskEdBox2.BackColor = &H80000009
   TxtArticulo.Enabled = True
   MaskEdBox1.Enabled = False
   Text6.Enabled = False
   MaskEdBox2.Enabled = False
   Label11.Visible = False
   lbEtiNum.Visible = False
   Command1.Enabled = False
   Combo1.Clear
   Combo1.Visible = False
   Text6.Visible = True
   chkserie.Enabled = False
End Sub

Private Sub ver_serie_lote()

    If flagserie = "S" Or flaglote = "S" Then
             Text6.Enabled = True
             Text6.Visible = True
             If (VGRegEnt <> 1) And flagserie = "S" And VGtipocreacion = 1 Then 'con guia de salida
                        agregar_combo
                        Combo1.Visible = True
                        Text6.Visible = False
             End If
             If flaglote = "S" Then
                          MaskEdBox1.BackColor = &H80000009
                          MaskEdBox2.BackColor = &H80000009
                          MaskEdBox1.Enabled = True
                          MaskEdBox2.Enabled = True
                          VGcod = TxtArticulo
                          Text6.Visible = True
              Else
                         TxtCantidad = "1"
                         TxtCantidad.Enabled = False
                         MaskEdBox1.BackColor = &H8000000F
                         MaskEdBox2.BackColor = &H80000009
                         MaskEdBox1.Enabled = False
                         MaskEdBox2.Enabled = True
             End If
   Else
     'Text6.Visible = False
     Text6.Enabled = False
     Text6.BackColor = &H8000000F
     MaskEdBox1.BackColor = &H8000000F
     MaskEdBox2.BackColor = &H8000000F
     MaskEdBox1.Enabled = False
     MaskEdBox2.Enabled = False
     Text3.Enabled = True
     TxtCantidad.Enabled = True
   End If
    

End Sub

Private Sub agregar_combo()
  Dim rs As New ADODB.Recordset
  Dim RSQL As String
  If flagserie = "S" Then
     RSQL = "select stsserie from stkseri where  STSALMA='" & VGAlma & "' and STSCODIGO='" & TxtArticulo & "' and STSSKDIS<>0 "
  End If
  'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
  Set rs = VGCNx.Execute(RSQL)
  If rs.EOF Then Exit Sub                     'revisar porque no entra al bucle
  Combo1.Clear
  While Not rs.EOF
     Combo1.AddItem (rs(0))
     rs.MoveNext
  Wend
  rs.Close
  Combo1.ListIndex = 0
  Command1.Enabled = True
End Sub

Private Sub Text6_Change()
 If Trim(Text6) <> "" Then Command1.Enabled = True
End Sub

Private Sub Text6_DblClick()
  If flaglote = "S" Then
    FormAyuLote.Show 1
  End If
End Sub

Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 112 Then
     Text6_DblClick
 ElseIf KeyCode = vbKeyTab Then
     Text6_KeyPress (13)
 End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
 Dim ncombo As Integer
 If KeyAscii = 13 Then
    Text6 = Trim(Text6)
    If Text6 <> "" Then
                If flaglote = "S" Then
                         existe_lote1 Text6
                         MaskEdBox1.SetFocus
                Else
                         If flagserie = "S" Then
                            If FrmRegistro.MSFlexGrid1.Rows <> 1 Then
                                For ncombo = 1 To FrmRegistro.MSFlexGrid1.Rows - 1
                                  If Combo1.Visible Then
                                      If UCase(Combo1.text) = UCase(FrmRegistro.MSFlexGrid1.TextMatrix(ncombo, 2)) Then
                                        MsgBox "Ya se ingreso la serie", vbInformation, "Aviso"
                                        Combo1.SetFocus
                                        Exit Sub
                                      End If
                                  ElseIf Text6 <> "" Then
                                      If UCase(Text6.text) = UCase(FrmRegistro.MSFlexGrid1.TextMatrix(ncombo, 2)) Then
                                        MsgBox "Ya se ingreso la serie", vbInformation, "Aviso"
                                        Text6.SetFocus
                                        Exit Sub
                                      End If
                                  End If
                                Next ncombo
                            End If
                            If existe_serie(Text6) Then Exit Sub
                            TxtCantidad = "1"
                            TxtCantidad.Enabled = False
                            Text3.Enabled = False
                            Command1.Enabled = True
                           
                         End If
                         SendKeys "{tab}"
                         KeyAscii = 0
                         Text6_Validate (False)
                End If
      Else
             If flaglote = "S" Then
                          MsgBox "Ingrese el número de Lote", vbInformation, "Aviso"
             Else
                          MsgBox "Ingrese el número de Serie", vbInformation, "Aviso"
             End If
             Text6.SetFocus
      End If
 End If
End Sub
              
Private Sub deshabilitartx5_tx3(flag As Boolean)
    TxtCantidad.Enabled = flag
    Text3.Enabled = flag
End Sub
 
Private Sub existe_lote(text As TextBox)
Dim rs As New ADODB.Recordset
Dim RSQL As String
   RSQL = "select STSLOTE, STSLKDIS,STSFECVEN,STSFECFAB from STKLOTE where  STSALMA ='" & VGAlma & "' and STSCODIGO = '" & TxtArticulo & "' and STSLOTE = '" & text & "'"
   'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
   Set rs = VGCNx.Execute(RSQL)
   If Not rs.EOF Then
     MsgBox "Lote Registrado en Almacen", vbInformation, "Aviso"
     lbcantstk = rs(1)
     MaskEdBox1 = IIf(IsNull(rs(2)), "__/__/____", rs(2))
     MaskEdBox2 = IIf(IsNull(rs(3)), "__/__/____", rs(3))
   End If
End Sub

Function existe_lote1(text As TextBox) As Boolean
Dim rs As New ADODB.Recordset
Dim RSQL As String
   RSQL = "select STSLOTE,STSFECVEN,STSFECFAB from STKLOTE where  STSALMA ='" & VGAlma & "' and STSCODIGO = '" & TxtArticulo & "' and STSLOTE = '" & Text6 & "'"
   'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
   Set rs = VGCNx.Execute(RSQL)
   If Not rs.EOF Then
     If Not graba Then MsgBox "Lote Registrado en Almacen", vbInformation, "Aviso"
     existe_lote1 = True
     If Not IsNull(rs(1)) Then
           MaskEdBox1 = rs(1)
     End If
     If Not IsNull(rs(2)) Then
           MaskEdBox2 = rs(2)
     End If
   Else
     MsgBox "Lote  No Registrado en Almacen", vbInformation, "Aviso"
     existe_lote1 = False
   End If
   rs.Close
End Function

Function existe_serie(text As TextBox) As Boolean
Dim rs As New ADODB.Recordset
Dim RSQL As String
   RSQL = "select STSSERIE,STSFECVEN from STKSERI where  STSALMA ='" & VGAlma & "' and STSCODIGO = '" & TxtArticulo & "' and STSSERIE = '" & Text6 & "'"
   'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
   Set rs = VGCNx.Execute(RSQL)
   If Not rs.EOF Then  'Not graba
     If True Then MsgBox "Serie Registrada en Almacen", vbInformation, "Aviso"
     If Not IsNull(rs(1)) Then
        MaskEdBox1 = rs(1)
     End If
     existe_serie = True
   Else
     If Not graba Then MsgBox "Serie  No Registrada en Almacen", vbInformation, "Aviso"
     existe_serie = False
   End If
   rs.Close
End Function

Private Sub ingreso_salida()
       If VGtipocreacion = 2 Then
           hubo_error = False
           grabadetalle
           If hubo_error Then Exit Sub
'           MsgBox "Se grabo sastifactoriamente", vbInformation, "Aviso"
       End If
       If (VGRegEnt <> 1) And (flagserie = "S") And VGtipocreacion = 1 Then
           serie_lote = Combo1.text
       Else
           serie_lote = Trim(Text6)
       End If
       If VGSeleccion = 2 Then
        LblCantidad = IIf(IsNumeric(LblCantidad), LblCantidad, TxtCantidad)
        varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 2) = serie_lote  'serie
        varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 3) = TxtCantidad.text  'çantidad
        varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 4) = VGabrev
        '**********************************************************************
        varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 11) = txccosto
        varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 12) = TxordFab
        varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 13) = txEquip
      
        '**********************************************************************
       'varform.MsFlexgrid1.TextMatrix(varform.MsFlexgrid1.Row, 4) = Text9.Text   'Precio
        If VGtipocreacion = 1 Then
             varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 6) = Val(LblCantidad)   'Cantidad informada
        Else
             varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 6) = Val(FrmModificar.numitem)    'Cantidad informada
        End If
        If xserie = "S" Then
           varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 10) = "S"
           varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 8) = MaskEdBox1
           varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 9) = MaskEdBox1
           xserie = "S"
        ElseIf xserie = "N" Then
           varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 10) = "N"
        Else
           varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 10) = "X"
        End If
      Else
      '                                       0                 1                 2                    3                       4             5               6                   7                      8                9               10
        pro_xserie
        LblCantidad = IIf(IsNumeric(LblCantidad), LblCantidad, TxtCantidad)
        If VGtipocreacion = 1 Then
           FrmRegistro.MSFlexGrid1.AddItem (TxtArticulo.text & vbTab & Label13 & vbTab & serie_lote & vbTab & Format(TxtCantidad.text, "##0.00") & vbTab & VGabrev & vbTab & "" & vbTab & Val(LblCantidad) & vbTab & "" & vbTab & MaskEdBox1 & vbTab & MaskEdBox2 & vbTab & xserie & vbTab & txccosto & vbTab & TxordFab & vbTab & txEquip & vbTab & txtcanref)
        Else
         FrmModificar.MSFlexGrid1.AddItem (TxtArticulo.text & vbTab & Label13 & vbTab & serie_lote & vbTab & Format(TxtCantidad.text, "###0.00") & vbTab & VGabrev & vbTab & "" & vbTab & Val(FrmModificar.numitem - 1) & vbTab & "" & vbTab & MaskEdBox1 & vbTab & MaskEdBox2 & vbTab & xserie & txccosto & vbTab & TxordFab & vbTab & txEquip)
        'FrmModificar.MSFlexGrid1.AddItem (TxtArticulo.text & vbTab & Label13 & vbTab & serie_lote & vbTab & Format(TxtCantidad.text, "###0.00") & vbTab & VGabrev & vbTab & "" & vbTab & Val(FrmModificar.numitem - 1) & vbTab & "" & vbTab & MaskEdBox1 & vbTab & MaskEdBox2 & vbTab & xserie)
      End If
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

Private Sub modifica_ingreso_salida()

      MaskEdBox1.Enabled = True
      MaskEdBox2.Enabled = True
      If varform.MSFlexGrid1.Row <> 0 Then
        TxtArticulo.text = varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 0)
        TxtCantidad.text = varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 3)
        cantidadini = CDbl(TxtCantidad)
        Label14 = varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 4)
        Label13 = varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 1)
        xserie = varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 10)
        
        txccosto = varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 11)
        TxordFab = varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 12)
        txEquip = varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 13)
        
        VGabrev = Label14
        If xserie = "S" Then
             Text6 = varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 2)
             deshabilitartx5_tx3 (False)
        ElseIf xserie = "N" Then
             Text6 = varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 2)
             If varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 9) <> "" Then MaskEdBox2 = varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 9)
             If varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 8) <> "" Then MaskEdBox1 = varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 8)
        Else
             flagserie = "N"
             flaglote = "N"
            ' ver_serie_lote
        End If
        lblUniEst = Nombre_Unidad(Label14)
        Text3 = lblUniEst
        If VGRegEnt <> 1 Then Text3.Enabled = False
        lbcantstk = cant
        TxtArticulo.Enabled = False
        TxtArticulo.TabStop = False
        'formato de pantalla
        If varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 10) = "X" Then
                flaglote = "N"
                flagserie = "N"
                Text6.Enabled = False
                Text6.BackColor = &H8000000F
                MaskEdBox1.BackColor = &H8000000F
                MaskEdBox2.BackColor = &H8000000F
                MaskEdBox1.Enabled = False
                MaskEdBox2.Enabled = False
                TxtCantidad.Enabled = True
                TxtCantidad.SetFocus
        Else
                Text6.Enabled = True
                If varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 10) = "S" Then
                     flagserie = "S"
                    flaglote = "N"
                     MaskEdBox1.BackColor = &H8000000F
                     MaskEdBox2.BackColor = &H80000009
                     MaskEdBox1.Enabled = False
                     MaskEdBox2.Enabled = True
                     TxtCantidad = "1"
                     TxtCantidad.Enabled = False
                Else
                     flaglote = "S"
                     flagserie = "N"
                     MaskEdBox1.BackColor = &H80000009
                     MaskEdBox2.BackColor = &H80000009
                     MaskEdBox1.Enabled = True
                     MaskEdBox2.Enabled = True
                End If
               Text6.SetFocus
        End If
        'TxtCantidad.SetFocus
    End If
End Sub
Private Sub colocastk()
  Dim cadena As String
   cadena = varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 0)
   If varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 10) = "S" Then
        seriestk
   ElseIf varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 10) = "N" Then
        lotestk
   Else
        llenastk
  End If
End Sub

Private Sub llenastk()
Dim RSQL As String
Dim rs As New ADODB.Recordset
  
   RSQL = "select  stskdis, stskmin,stskmax,stpunrep from stkart  WHERE STALMA='" & VGAlma & "' and  stcodigo ='" & codigo & "'"
   'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
   Set rs = VGCNx.Execute(RSQL)
   If Not rs.EOF Then
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
   Lote = varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 2)
   RSQL = "select  STSLKDIS from STKLOTE where STSALMA ='" & VGAlma & "' and STSCODIGO = '" & codigo & "' and STSLOTE = '" & Lote & "'"
   'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
   Set rs = VGCNx.Execute(RSQL)
   If Not rs.EOF Then
        lbcantstk = rs(0)
     Else
        lbcantstk = 0
     End If
   
   rs.Close
End Sub

Private Sub seriestk()
Dim rs As New ADODB.Recordset
Dim RSQL As String
Dim Serie As String
   Serie = varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 2)
   RSQL = "select STSSKDIS from STKSERI where  STSALMA ='" & VGAlma & "' and STSCODIGO = '" & codigo & "' and STSSERIE = '" & Serie & "'"
   Set rs = VGCNx.Execute(RSQL)
   If Not rs.EOF Then
      lbcantstk = rs(0)
   Else
      lbcantstk = 0
   End If
   rs.Close
End Sub

Private Sub grabadetalle()
 Dim Adoreg1 As ADODB.Recordset
 Dim AdoReg2 As ADODB.Recordset
 Dim Rsql1 As String
 Dim criterio As String
 Dim item As Integer
' On Error GoTo GrabErr
  TxtArticulo = Trim(TxtArticulo)
  If VGSeleccion = 2 Then   'indica es el dato se modifica
     Rsql1 = "select * from movalmdet where DETD= '" & FrmModificar.TxDoc & "' AND DENUMDOC= '" & FrmModificar.Lblnumdoc & "'  AND DEALMA= '" & VGAlma & "' and  DEITEM = " & FrmModificar.contador & " and DECODIGO = '" & TxtArticulo & "'"
  Else
     Rsql1 = "select * from movalmdet where DETD='TT'"
  End If
  Set Adoreg1 = New ADODB.Recordset
  Adoreg1.Open Rsql1, VGCNx, adOpenDynamic, adLockOptimistic
  ' Si es nuevo adicciono los datos primary key
  If Adoreg1.RecordCount = 0 Then
        Adoreg1.AddNew
        Adoreg1("dealma") = VGAlma
Retor:
        Rsql1 = "select * from movalmdet where DETD= '" & FrmModificar.TxDoc & "' AND DENUMDOC= '" & FrmModificar.Lblnumdoc & "'  AND DEALMA= '" & VGAlma & "' and  DEITEM = " & FrmModificar.numitem & ""
        Set AdoReg2 = New ADODB.Recordset
        AdoReg2.Open Rsql1, VGCNx, adOpenDynamic, adLockOptimistic
        If Not AdoReg2.EOF Then
           FrmModificar.numitem = AdoReg2("deitem") + 1
           FrmModificar.contador = FrmModificar.numitem
           FormCreacion.Label11 = FrmModificar.numitem
           GoTo Retor
        End If
        AdoReg2.Close
        Adoreg1("deitem") = FrmModificar.numitem
        Adoreg1("DECODIGO") = Trim(TxtArticulo)   ' Format(MSFlexGrid1.TextMatrix(contador, 0), "00000000")
        Adoreg1("DEDESCRI") = Label13
        Adoreg1("detd") = FrmModificar.TxDoc
        Adoreg1("denumdoc") = FrmModificar.Lblnumdoc
        
        FrmModificar.numitem = FrmModificar.numitem + 1
  End If
  ' adicciono la nueva cantidad, serie y lote
     Adoreg1("decantid") = Val(TxtCantidad)        '
     CANTIDAD = Val(TxtCantidad)
     If xserie = "S" Then
          actserie
          Adoreg1("DESERIE") = Trim(Text6)
     ElseIf xserie = "N" Then
         grabalote
         Adoreg1("DELOTE") = Trim(Text6)
     End If
     '***********************************
        Adoreg1("DECENCOS") = txccosto
        Adoreg1("DEORDFAB") = TxordFab
        Adoreg1("DEQUIPO") = txEquip
     '***********************************
    
    Adoreg1.Update
    Set Adoreg1 = Nothing
    nuevodet = True   'para que no actualice dos veces
    actualizastk TxtArticulo  'actualizando dos veces serie lote
    ya_grabo_det = True
  
    Exit Sub
GrabErr:
    MsgBox Err.Description
    hubo_error = True
End Sub

Private Sub actualizastk(codigo As String)

Dim criterio As String
Dim canttemp As Double
Dim adoreg As ADODB.Recordset
Dim RSQL As String
  RSQL = "select * from STKART where  STCODIGO= '" & TxtArticulo & "' and STALMA = '" & VGAlma & "'  "
   Set adoreg = New ADODB.Recordset
   adoreg.Open RSQL, VGCNx, adOpenDynamic, adLockOptimistic
   If adoreg.RecordCount = 0 Then
     adoreg.AddNew
     adoreg("stalma") = VGAlma
     adoreg("stcodigo") = TxtArticulo
     adoreg.Update
   Else
     canttemp = adoreg("stskdis")
   End If
   adoreg.Close
   Set adoreg = Nothing
         RSQL = "Update STKART set stskdis= " & IIf(FrmModificar.tipo = "NI", canttemp + Val(TxtCantidad), canttemp - Val(TxtCantidad)) & " where  STCODIGO= '" & TxtArticulo & "' and STALMA = '" & VGAlma & "'"
         VGCNx.Execute RSQL
   ValMes
   nuevodet = False
End Sub

Private Sub reactualizastk(codigo As String)
Dim criterio As String
Dim canttemp As Double
Dim RSQL As String
Dim adoreg As ADODB.Recordset
'On Error GoTo ERR
   RSQL = "select * from STKART where  STCODIGO= '" & TxtArticulo & "' and STALMA = '" & VGAlma & "'  "
   Set adoreg = New ADODB.Recordset
   adoreg.Open RSQL, VGCNx, adOpenDynamic, adLockOptimistic
   If adoreg.RecordCount = 0 Then
     adoreg.AddNew
     adoreg("stalma") = VGAlma
     adoreg("stcodigo") = TxtArticulo
     adoreg.Update
   Else
    canttemp = adoreg("stskdis")
   End If
   
   RSQL = "Update STKART set stskdis=" & IIf(FrmModificar.tipo = "NI", canttemp + cantidadini, canttemp - cantidadini) & " where  STCODIGO= '" & TxtArticulo & "' and STALMA = '" & VGAlma & "'"
   VGCNx.Execute RSQL
   CANTIDAD = cantidadini
   If Not nuevodet Then  ' si no es nuevo tiene que actualizar la serie y lote
        If Text6 <> "" Then
          If xserie = "S" Then actserie  'solo descarga   o carga dependiendo el tipo
          If xserie = "N" Then actlote TxtArticulo  'solo desCarga
        End If
        ValMes             'reactualiza  el moremes
   End If
   adoreg.Close
   nuevodet = False
   Exit Sub
Err:
   MsgBox Err.Description
End Sub

Private Sub grabalote()
Dim uSql As String
Dim Lote As String
Dim nuevo_stk As Double
Dim RSQL As String
Dim rs As New ADODB.Recordset
Dim fecfab As Date
Dim fecven As Date
    On Error GoTo GrabErr
    
    RSQL = "select STSLKDIS FROM STKLOTE where  STSALMA= '" & VGAlma & "' and  STSCODIGO= '" & TxtArticulo & "' and STSLOTE= '" & Text6 & "'" '
    'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
    Set rs = VGCNx.Execute(RSQL)
    If Not rs.EOF Then
       nuevo_stk = IIf(FrmModificar.tipo = "NI", rs(0) + CANTIDAD, rs(0) - CANTIDAD)
       uSql = "Update STKLOTE set STSLKDIS = " & nuevo_stk & " WHERE  STSALMA='" & VGAlma & "' and STSCODIGO='" & codigo & "' AND STSLOTE='" & Text6 & "'"
    Else
        If MaskEdBox1 <> "__/__/____" And (MaskEdBox2 = "__/__/____") Then
            uSql = "insert into STKLOTE (STSALMA,STSCODIGO,STSLOTE,STSLKDIS,STSFECFAB)  VALUES ('" & VGAlma & "','" & TxtArticulo & "','" & Text6 & "' ," & Val(TxtCantidad) & ",'" & DateSQL(MaskEdBox2) & "') "
        ElseIf MaskEdBox1 = "__/__/____" And MaskEdBox2 <> "__/__/____" Then
            uSql = "insert into STKLOTE  (STSALMA,STSCODIGO,STSLOTE,STSLKDIS,STSFECVEN)VALUES ('" & VGAlma & "','" & TxtArticulo & "','" & Text6 & "' ," & Val(TxtCantidad) & " ,' ','" & DateSQL(MaskEdBox1) & "') "  'SIN FECFAB
        ElseIf MaskEdBox1 <> "__/__/____" And MaskEdBox2 <> "__/__/____" Then
            uSql = "insert into STKLOTE  (STSALMA,STSCODIGO,STSLOTE,STSLKDIS,STSFECFAB,STSFECVEN)  VALUES ('" & VGAlma & "','" & TxtArticulo & "','" & Text6 & "' ," & Val(TxtCantidad) & " ,'" & DateSQL(MaskEdBox2) & "','" & DateSQL(MaskEdBox1) & "') "
        Else
            uSql = "insert into STKLOTE  (STSALMA,STSCODIGO,STSLOTE,STSLKDIS)  VALUES ('" & VGAlma & "','" & TxtArticulo & "','" & Text6 & "' ," & Val(TxtCantidad) & "') "
        End If
    
    End If
    rs.Close
    Set rs = Nothing
    VGCNx.Execute uSql
    Exit Sub
GrabErr:
    MsgBox Err.Description
    hubo_error = False
    
End Sub

Private Sub actserie()
Dim uSql As String
Dim Serie As String
Dim valor As Integer
Dim rs As New ADODB.Recordset
Dim RSQL As String
Dim fecfab As Date
Dim fecven As Date
   
On Error GoTo Err
    If Combo1.Visible Then
          Text6 = Combo1.text
    End If
    
    RSQL = "select STSSKDIS FROM STKSERI where   STSALMA= '" & VGAlma & "' and  STSCODIGO= '" & codigo & "' and STSSERIE= '" & Text6 & "'" '
    'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
    Set rs = VGCNx.Execute(RSQL)
    If Not rs.EOF Then
       valor = IIf(FrmModificar.tipo = "NI", 1, 0)
       uSql = "update STKSERI set STSSKDIS = " & valor & " WHERE  STSALMA='" & VGAlma & "' and STSCODIGO='" & codigo & "'AND STSSERIE='" & Text6 & "'"
    Else
       uSql = "insert into STKSERI (STSALMA,STSCODIGO,STSSERIE,STSSKDIS) VALUES ('" & VGAlma & "','" & codigo & "','" & Text6 & "' ,1) "
    End If
    rs.Close
    Set rs = Nothing
    VGCNx.Execute uSql
    Exit Sub
Err:
   MsgBox Err.Description, vbExclamation, "Aviso"
End Sub

Private Sub actlote(codigo As String)
Dim uSql As String
Dim nuevo_stk As Double
Dim RSQL As String
Dim rs As New ADODB.Recordset

    RSQL = "select STSLKDIS FROM STKLOTE where   STSALMA= '" & VGAlma & "' and  STSCODIGO= '" & TxtArticulo & "' and STSLOTE= '" & Text6 & "'" '
    'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
    Set rs = VGCNx.Execute(RSQL)
    If Not rs.EOF Then
         nuevo_stk = IIf(FrmModificar.tipo = "NI", rs(0) + CANTIDAD, rs(0) - CANTIDAD)
         uSql = "Update STKLOTE set STSLKDIS = " & nuevo_stk & " WHERE  STSALMA= '" & VGAlma & "' and STSCODIGO='" & codigo & "'AND STSLOTE='" & Text6 & "'"
          VGCNx.Execute uSql
    End If
     rs.Close
End Sub

Private Sub actvalmes()
 
  Dim criterio As String
  Dim Adoreg1 As ADODB.Recordset
  Dim RSQL As String
  Dim Cantent As Double
  Dim Cantsal As Double
  Dim mespro As String
  Dim uSql As String
  On Error GoTo Err
   mespro = Year(FrmModificar.DTPicker1) & Format(Month(FrmModificar.DTPicker1), "00")
   CANTIDAD = Val(TxtCantidad)
   RSQL = "select SMCANENT,SMCANSAL FROM MoResMes where  SMALMA = '" & VGAlma & "' and SMMESPRO = '" & mespro & "' AND  SMCODIGO= '" & Trim(TxtArticulo) & "'" '
   Set Adoreg1 = New ADODB.Recordset
   Adoreg1.Open RSQL, VGCNx, adOpenDynamic, adLockOptimistic
   If Adoreg1.RecordCount <> 0 Then
      If FrmModificar.tipo = "NI" Then
        Cantent = Adoreg1(0) + CANTIDAD
        uSql = "Update MoResMes set SMCANENT = " & Cantent & " where SMALMA='" & VGAlma & "'  and  SMCODIGO ='" & TxtArticulo & "' AND SMMESPRO ='" & mespro & "' "
       Else
        Cantsal = Adoreg1(1) - CANTIDAD
        uSql = "Update MoResMes set SMCANSAL = " & Cantsal & " where SMALMA='" & VGAlma & "' and   SMCODIGO ='" & TxtArticulo & "' AND SMMESPRO ='" & mespro & "' "
       End If
   Else
      If FrmModificar.tipo = "NI" Then
        Cantent = CANTIDAD
        Cantsal = 0
      Else
        Cantsal = CANTIDAD
        Cantent = 0
      End If
       uSql = "insert into MoResMes (SMALMA,SMCODIGO,SMMESPRO,SMCANENT,SMCANSAL,SMSALDOINI) VALUES ('" & VGAlma & "','" & TxtArticulo & "','" & mespro & "' ," & Cantent & "," & Cantsal & ",0) "
   End If
   VGCNx.Execute uSql
   Adoreg1.Close
   Exit Sub
Err:
    MsgBox Err.Description
End Sub

Private Sub ValMes()
 
  Dim criterio As String
 
  Dim Adoreg1 As ADODB.Recordset
  Dim RSQL As String
  Dim Cantent As Double
  Dim Cantsal As Double
  Dim mespro As String
  Dim uSql As String
  
   mespro = Year(FrmModificar.DTPicker1) & Format(Month(FrmModificar.DTPicker1), "00")
   RSQL = "select SMCANENT,SMCANSAL FROM MoResMes where  SMALMA = '" & VGAlma & "' and SMMESPRO = '" & mespro & "' AND  SMCODIGO= '" & TxtArticulo & "'"  '
   
   Set Adoreg1 = New ADODB.Recordset
   Adoreg1.Open RSQL, VGCNx, adOpenDynamic, adLockOptimistic
   If Adoreg1.RecordCount <> 0 Then
      If FrmModificar.tipo = "NI" Then
        Cantent = Adoreg1(0) + CANTIDAD        '
        uSql = "Update MoResMes set SMCANENT = " & Cantent & "  where SMALMA='" & VGAlma & "'  and  SMCODIGO ='" & TxtArticulo & "' AND SMMESPRO ='" & mespro & "' "
       Else
        Cantsal = Adoreg1(1) + CANTIDAD   'INICIALMENTE DESCARGA, ACTUALIZA LA NUEVA CANTIDAD
        uSql = "Update MoResMes set SMCANSAL = " & Cantsal & "  where SMALMA='" & VGAlma & "' and  SMCODIGO ='" & TxtArticulo & "' AND SMMESPRO ='" & mespro & "' "
       End If
   Else
      If FrmModificar.tipo = "NI" Then
        Cantent = CANTIDAD
        Cantsal = 0
      Else
        Cantsal = CANTIDAD
        Cantent = 0
      End If
        uSql = "insert into MoResMes (SMALMA,SMCODIGO,SMMESPRO,SMCANENT,SMCANSAL,SMSALDOINI) VALUES ('" & VGAlma & "','" & TxtArticulo & "','" & mespro & "' ," & Cantent & "," & Cantsal & ",0)"
   End If
   VGCNx.Execute uSql
   Adoreg1.Close
End Sub

Public Function Validadato(pvalor) As Boolean
    Dim k As Integer
    Dim l As Integer
    Dim txt As String
    Dim compara As String
    
    Validadato = True
    
    txt = UCase(Trim(CStr(pvalor)))
    l = Len(Trim(txt))
    compara = "[?%$',#@" & Chr(34) & "*-+{}!¿¡]"
    For k = 1 To l
      If Mid(txt, k, 1) Like compara Then
          Validadato = False
          Exit For
      End If
    Next k

End Function
