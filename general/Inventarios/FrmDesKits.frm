VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmDesKits 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DesArmado de Kids"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   8910
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Txtalma 
      Height          =   285
      Left            =   6720
      MaxLength       =   2
      TabIndex        =   39
      Top             =   2880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame5 
      Height          =   1164
      Left            =   2385
      TabIndex        =   36
      Top             =   3960
      Visible         =   0   'False
      Width           =   4512
      Begin MSComctlLib.ProgressBar PBar 
         Height          =   228
         Left            =   324
         TabIndex        =   37
         Top             =   720
         Width           =   3900
         _ExtentX        =   6879
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label lbmsg 
         Caption         =   "Procesando...."
         Height          =   270
         Left            =   315
         TabIndex        =   38
         Top             =   405
         Width           =   3900
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlexGrid1 
      Height          =   2490
      Left            =   45
      TabIndex        =   12
      Top             =   3285
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   4392
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   248
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   8070
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancelar"
      Height          =   735
      Left            =   3675
      Picture         =   "FrmDesKits.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6015
      Width           =   735
   End
   Begin VB.TextBox TxCantidad 
      Height          =   375
      Left            =   4275
      TabIndex        =   29
      Text            =   "TxCantidad"
      Top             =   3555
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Limpiar"
      Height          =   735
      Left            =   4785
      Picture         =   "FrmDesKits.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6015
      Width           =   735
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Height          =   735
      Left            =   2550
      Picture         =   "FrmDesKits.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6030
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Generales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2724
      Left            =   72
      TabIndex        =   18
      Top             =   540
      Width           =   8760
      Begin VB.TextBox TxSaldo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   288
         Left            =   5076
         MaxLength       =   10
         TabIndex        =   34
         Top             =   2340
         Width           =   876
      End
      Begin VB.TextBox TxCodKid 
         Height          =   288
         Left            =   1656
         TabIndex        =   10
         Top             =   2016
         Width           =   1095
      End
      Begin VB.TextBox Txcant 
         Alignment       =   1  'Right Justify
         Height          =   288
         Left            =   1656
         MaxLength       =   10
         TabIndex        =   11
         Top             =   2376
         Width           =   876
      End
      Begin VB.TextBox TxCambio 
         Height          =   285
         Left            =   7215
         TabIndex        =   9
         Top             =   1920
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox CmbMoneda 
         Height          =   288
         ItemData        =   "FrmDesKits.frx":0CC6
         Left            =   1692
         List            =   "FrmDesKits.frx":0CD0
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   1680
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox TxTransa 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   3
         Text            =   "69"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   5052
         MaxLength       =   11
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   4
         Top             =   960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   6
         Top             =   1320
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   5052
         MaxLength       =   3
         TabIndex        =   7
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   5052
         MaxLength       =   10
         TabIndex        =   5
         Top             =   960
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   108331009
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   36161
      End
      Begin VB.Label Label11 
         Caption         =   "Saldo Stock"
         Height          =   252
         Left            =   3852
         TabIndex        =   35
         Top             =   2376
         Width           =   1212
      End
      Begin VB.Label Label7 
         Caption         =   "Codigo de Kit's"
         Height          =   252
         Left            =   216
         TabIndex        =   33
         Top             =   2052
         Width           =   1692
      End
      Begin VB.Label lblnomkits 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   288
         Left            =   2820
         TabIndex        =   32
         Top             =   2010
         Width           =   4065
      End
      Begin VB.Label Label10 
         Caption         =   "Cantidad"
         Height          =   228
         Left            =   216
         TabIndex        =   31
         Top             =   2412
         Width           =   1200
      End
      Begin VB.Label Label1 
         Caption         =   "Moneda "
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   1680
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Doc."
         ForeColor       =   &H80000006&
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   270
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Transaccion"
         ForeColor       =   &H80000006&
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   630
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Num. Doc"
         ForeColor       =   &H80000006&
         Height          =   252
         Left            =   3852
         TabIndex        =   26
         Top             =   276
         Width           =   852
      End
      Begin VB.Label Label6 
         Caption         =   "Tip Doc Ref"
         ForeColor       =   &H80000006&
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   990
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Orden Compra"
         ForeColor       =   &H80000006&
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   1320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Autorizacion"
         ForeColor       =   &H80000006&
         Height          =   252
         Left            =   3852
         TabIndex        =   23
         Top             =   1332
         Width           =   1092
      End
      Begin VB.Label Label14 
         Caption         =   "Num. Ref"
         ForeColor       =   &H80000006&
         Height          =   252
         Left            =   3852
         TabIndex        =   22
         Top             =   996
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label lbltrans 
         Caption         =   "SALIDA POR DESARMADO DE KITS"
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
         Left            =   2160
         TabIndex        =   21
         Top             =   630
         Width           =   5055
      End
      Begin VB.Label lbltipref 
         Caption         =   "lbltipref"
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
         Left            =   2160
         TabIndex        =   20
         Top             =   990
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label lblauto 
         Caption         =   "lblauto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   5652
         TabIndex        =   19
         Top             =   1320
         Width           =   1932
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   5910
      Picture         =   "FrmDesKits.frx":0CDC
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6015
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   2175
      Left            =   225
      TabIndex        =   13
      Top             =   3420
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   3836
      _Version        =   393216
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuALMACEN 
      Height          =   330
      Left            =   1080
      TabIndex        =   0
      Top             =   135
      Width           =   3780
      _ExtentX        =   6668
      _ExtentY        =   582
      XcodMaxLongitud =   0
      xcodwith        =   100
      NomTabla        =   "tabalm"
      TituloAyuda     =   "Almacenes"
      ListaCampos     =   "TAALMA(1),TADESCRI(1)"
      XcodCampo       =   "TAALMA"
      XListCampo      =   "TADESCRI"
      ListaCamposDescrip=   "Codigo,Descripcion"
      ListaCamposText =   "TAALMA,TADESCRI"
   End
   Begin VB.Label lblalmacen 
      Caption         =   "lblalmacen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7350
      TabIndex        =   41
      Top             =   2925
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label12 
      Caption         =   "Almacen"
      ForeColor       =   &H80000006&
      Height          =   270
      Left            =   240
      TabIndex        =   40
      Top             =   150
      Width           =   1200
   End
End
Attribute VB_Name = "FrmDesKits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adodc1 As ADODB.Recordset
Dim Adodc2 As ADODB.Recordset
Dim Adodc22 As ADODB.Recordset
Dim Adodc3 As ADODB.Recordset
Dim ContaSalida As String
Dim RSQL As String
'Dim db As Database
Dim nument As Long
Dim precioprom As Double
Dim CANTIDAD As Double
Dim canttemp As Double
Dim Campo As String * 2
Dim contador As Integer
Dim num As Integer
Dim TT_CONTADOR As Integer
Dim estadocosto As Integer
Dim cadena As String
Dim alma As String
Dim tipo As String * 2
Dim dato As String
Dim NumDoc As String
Dim Codigo2 As String
Dim ndato As String

Private Sub CmbMoneda_Click()
If CmbMoneda.ListIndex = 0 Then
    VGSoles = True
    TxCambio.Enabled = False
    VGTipCamb = 1
Else
    VGSoles = False
    TxCambio.Enabled = True
    VGTipCamb = TxCambio
End If
End Sub

Private Sub CmbMoneda_KeyPress(KeyAscii As Integer)
Tabula (KeyAscii)
End Sub

'Eliminar
Private Sub Command1_Click()
LoadReceta (TxCodKid)
Txcant = 0
End Sub

'Limpiar
Private Sub Command3_Click()
'TxTransa = ""
TxCodKid = "": lblnomkits = "": Txcant = ""
Text6 = "": Text8 = "": Text4 = ""
Text1 = "": Text9 = "": TxCantidad = "": TxCambio = "0": TxCambio.Enabled = False
MSFlexGrid1.Clear
MSFlexGrid1.Rows = 2
inicializaFG
End Sub

'****************************** Graba LA GUIA****************
Private Sub CmdGrabar_Click()
Dim criterio As String, cadena As String, cadena1 As String, cadena2 As String
Dim rpta As Integer, FACTOR As Double, uSql As String, nIt As Integer
Dim cSel1 As New ADODB.Recordset
Dim n As Long
On Error GoTo GrabErr


If Trim(Txtalma) = "" Then
    MsgBox "Debe Ingresar el xod. de almacen", vbInformation, "Mensaje"
    Txtalma.SetFocus
    Exit Sub
  Else
    VGAlma = Txtalma.text
End If

If Trim(TxTransa) = "" Then
    MsgBox "Debe Ingresar el Movimiento", vbInformation, "Mensaje"
    TxTransa.SetFocus
    Exit Sub
End If

If Not VerificarSTOCK Then
   Exit Sub
End If

If Val(Txcant) = 0 Then
   MsgBox "La Cantidad a desarmar No puede tener Valor   0 ", vbInformation, "Mensaje"
   Txcant.SetFocus:    Exit Sub
End If


CANTIDAD = 0
If MSFlexGrid1.Rows = 1 Then
      MsgBox "No se puede grabar,debe adicionar registro", vbInformation, mensaje1
      Exit Sub
End If



'**************
'Numeracion
VGCNx.BeginTrans
cSel1.Open "select * from  num_documentos where ctncodigo='TR'", VGCNx, adOpenDynamic, adLockOptimistic

If cSel1.RecordCount > 0 Then
    ndato = Right("00000000000" & Trim(CStr(ESNULO(cSel1!ctnnumero, 0))), 11)                   'nro pedido"
Else
   MsgBox " No existe Registro de contador de transacciones...Verifique!!", vbInformation, "AVISO"
   cSel1.Close
   Set cSel1 = Nothing
   VGCNx.RollbackTrans
   Exit Sub
End If
VGCNx.Execute "update num_documentos set ctnnumero=ctnnumero+1  where ctncodigo='TR'"
cSel1.Close
Set cSel1 = Nothing

cSel1.Open "select * from tabalm where taalma='" & Ctr_ayuAlmacen.xclave & "'", VGCNx, adOpenDynamic, adLockOptimistic
If cSel1.RecordCount > 0 Then
   Text4 = Right("00000000000" & Trim(CStr(cSel1!tanument)), 11)                      'nro pedido"
   ContaSalida = Right("00000000000" & Trim(CStr(cSel1!tanumsal)), 11)                      'nro pedido"
End If
VGCNx.Execute "update tabalm set tanument=tanument+1 , tanumsal=tanumsal+1 where taalma='" & Ctr_ayuAlmacen.xclave & "'"
cSel1.Close
Set cSel1 = Nothing
VGCNx.CommitTrans
tipo = "NS"
If existe_numdoc(ContaSalida) Then
   MsgBox "El Numero de Documento de Salida " & ContaSalida & " ya Existe... Verifique el Contador de Salidas del Almacen", vbExclamation, "Aviso"
   Exit Sub
End If

Screen.MousePointer = 11

Frame5.Visible = True
lbmsg.Caption = "Guardando  Datos .......! "

'***
'Verificar I/S
Call grabacabecera("S", "NS", ContaSalida, 0) 'Salida
Call grabacabecera("I", "NI", Text4, 1) 'Ingreso
'***
adodc1.Close
FACTOR = 1
contador = 1
'Graba detalle
NumDoc = Text4
 CANTIDAD = Txcant
 If (IIf(VGRegEnt = 1, True, True)) Then      'verificastk
  cadena = TxCodKid
  criterio = " deNUMDOC = '" & ContaSalida & "'"
  criterio = criterio + " and  deALMA = '" & VGAlma & "'"
  criterio = criterio + " and  deTD = '" & tipo & "'"
  RSQL = "select * from movalmdet where " & criterio
 Adodc2.Open RSQL, VGCNx, adOpenDynamic, adLockOptimistic

   
   If Not VGActualizar Then
          Adodc2.AddNew
   Else
          criterio = "DECODIGO = '" & cadena & "'"
          criterio = criterio + " and  DEALMA = '" & VGAlma & "'"
          If Not Adodc2.EOF Then Adodc2.Filter = criterio
   End If
   Adodc2("DEALMA") = VGAlma
   Adodc2("DETD") = "NS"
   Adodc2("DENUMDOC") = ContaSalida
   Adodc2("DEITEM") = contador
   Adodc2("DECODIGO") = cadena   ' Format(MSFlexGrid1.TextMatrix(contador, 0), "00000000")
   Adodc2("DEDESCRI") = lblnomkits
   Adodc2("DECANTID") = CANTIDAD
   If Trim(MSFlexGrid1.TextMatrix(contador, 2)) <> "" Then 'Cantidad Ingresada
        Call grabastk(cadena, 2, 0)
        If Trim(MSFlexGrid1.TextMatrix(contador, 3)) <> "" Then    'si tiene precio
            Adodc2("DEPRECIO") = 0 '* VGTipCamb '******el precio
            Adodc2("DETIPCAM") = VGTipCamb
        Else
            Adodc2("DEPRECIO") = 0
        End If
        alma = VGAlma
   End If
   Adodc2.Update
   Adodc2.Filter = ""
   Adodc2.Close
 End If
'*************************************************************************
PBar.Min = 0
PBar.Max = (MSFlexGrid1.Rows - 1) * 100
'*************************************************************************
  tipo = "NI"
  criterio = " deNUMDOC = '" & Text4 & "'"
  criterio = criterio + " and  deALMA = '" & VGAlma & "'"
  criterio = criterio + " and  deTD = '" & tipo & "'"
  RSQL = "select * from movalmdet where " & criterio
 Adodc22.Open RSQL, VGCNx, adOpenDynamic, adLockOptimistic

For n = 1 To MSFlexGrid1.Rows - 1
    nIt = nIt + 1
    PBar.Value = n * 80
    Frame5.Refresh
    lbmsg.Caption = "Guardando  Datos .......! "
    Call GrabarDetKit(MSFlexGrid1.TextMatrix(n, 0), "NI", nIt, Devolver_Dato(1, MSFlexGrid1.TextMatrix(n, 0), "MaeArt", "Acodigo", False, "Adescri"), Val(MSFlexGrid1.TextMatrix(n, 3)), 0)
    cadena = MSFlexGrid1.TextMatrix(n, 0)
    CANTIDAD = MSFlexGrid1.TextMatrix(n, 3)
    Call grabastk(cadena, 1, 0)
   ' Call ClsTock.CalcularStock(VGAlma, MSFlexGrid1.TextMatrix(n, 0), DTPicker1.Value)
Next
'*************************************************************************
'Adodc2.Requery
Adodc22.Close
Frame5.Visible = False

If MsgBox("¿Desea Imprimir Guias de Salida?", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    imprimir ("NS")
End If
If MsgBox("¿Desea Imprimir Guias de Ingreso?", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    imprimir ("NI")
End If
MSFlexGrid1.Clear
MSFlexGrid1.Rows = 2
inicializar
VGSoles = True
VGTipCamb = 1
Screen.MousePointer = 1
Exit Sub
GrabErr:
  MsgBox Err.Description, vbExclamation, "Error"
  Screen.MousePointer = 1
  Frame5.Visible = False
  Exit Sub
  Resume

End Sub

Private Sub Command7_Click()
Unload Me
End Sub

Private Sub Ctr_ayuAlmacen_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Txtalma.text = Ctr_ayuAlmacen.xclave
lblalmacen.Caption = Ctr_ayuAlmacen.xnombre
VGAlma = Txtalma.text
TxCodKid.text = Empty
lblnomkits.Caption = Empty
TxSaldo.text = 0
Txcant.text = Empty
'MSFlexGrid1.Rows = 1
Call muestra
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
DTPicker1.Value = UltimoCierreFech(DTPicker1.Value)
If KeyCode = 13 Then
    SendKeys "{tab}"
End If
End Sub

Private Sub Form_Activate()
If VGAutomatico Then Text4.Enabled = False
End Sub

Private Sub Form_Load()
central FrmDesKits
Txtalma.text = VGAlma
Set adodc1 = New ADODB.Recordset
Set Adodc2 = New ADODB.Recordset
Set Adodc22 = New ADODB.Recordset
Set Adodc3 = New ADODB.Recordset

Call Ctr_ayuAlmacen.conexion(VGCNx)
Ctr_ayuAlmacen.filtro = "empresacodigo='" & VGparametros.empresacodigo & "' and puntovtacodigo='" & VGparametros.puntovta & "' "

VGActualizar = False
VGSoles = True
lbltipref = "": lblauto = ""
DTPicker1.MaxDate = VGParamSistem.FechaTrabajo
DTPicker1.Value = UltimoCierreFech(CDate(Format(VGParamSistem.FechaTrabajo, "dd/MM/yyyy")))

dato = "S"
tipo = "NS"
Codigo2 = "NOTA DE SALIDA"

MSFlexGrid1.Clear
MSFlexGrid1.Rows = 2
inicializar
inicializaFG

End Sub

Private Sub MSFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 112 Then MSFlexGrid1_DblClick
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
Alinear

If MSFlexGrid1.Col = 2 Or MSFlexGrid1.Col = 3 Then
    TxCantidad.Visible = True
    TxCantidad.SetFocus

    If KeyAscii <> 13 And KeyAscii <> 27 And KeyAscii <> 9 And KeyAscii <> 8 Then
        TxCantidad.text = TxCantidad.text & Chr(KeyAscii)
    End If
    
    TxCantidad.SelStart = Len(TxCantidad.text)
    TxCantidad.SelLength = 0
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'************************** NUM REF
If KeyAscii = 13 And Text1.text <> "" Then
   If Not IsNumeric(Text1) And (Text6 = "BV") Then
      MsgBox "Ingrese el Numero de  la Boleta", vbOKOnly, "Aviso"
      Exit Sub
   Else
      Text8.SetFocus
   End If
Else
 If Text6 = "BV" Then
   If (Chr$(KeyAscii) > "9" Or Chr$(KeyAscii) < "0") And KeyAscii <> 8 Then KeyAscii = 0
   Tabula (KeyAscii)
 End If
End If
End Sub

Private Sub Text9_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 112 Then Text9_DblClick
End Sub
Private Sub TxCambio_KeyPress(KeyAscii As Integer)
If NumPto(KeyAscii) Then
    If KeyAscii = 13 Then SendKeys "{tab}"
Else
    KeyAscii = 0
End If
End Sub

Private Sub Txcant_Change()
   If Not IsNumeric(Txcant) And Trim(Txcant) <> "" Then
      MsgBox "La Cantidad Tiene que ser Numerica....", vbInformation, "Verifique .."
      Txcant.SetFocus
   End If
End Sub

Private Sub Txcant_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     Call VerificarSTOCK
     Tabula (KeyAscii)
  End If
End Sub

Private Sub TxCodKid_DblClick()
VGRegEnt = 0: VGForm1 = 5
FormAyuArtKid.Show 1
inicializaFG
TxSaldo = ClsTock.SaldoArti(VGAlma, TxCodKid, VGCNx)
If TxCodKid <> "" Then Tabula (13)
End Sub

Private Sub TxCodKid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then TxCodKid_DblClick
End Sub

Private Sub TxCodKid_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Txcant.SetFocus
End Sub

Private Sub TxTransa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    TxTransa_DblClick
ElseIf KeyCode = 46 Then
    lbltrans = ""
End If
  
End Sub

Private Sub TxTransa_KeyPress(KeyAscii As Integer)
'****************** TRANSACCIONES
 If KeyAscii = 13 And Len(TxTransa.text) = 2 Then
       buscar_trans
       lbltrans = Mid(lbltrans, 1, 30)
       If lbltrans = "" Then
            Enfoque Text6
       End If
       Exit Sub
 Else
     If KeyAscii = 8 Then
         lbltrans = ""
         LIMPIACABECERA
    End If
End If
End Sub

Private Sub TxTransa_DblClick()
Dim Adodc3 As ADODB.Recordset
Set Adodc3 = New ADODB.Recordset
Adodc3.Open "SELECT TT_CODMOV,TT_DESCRI FROM Tabtransa where  TT_tipmov = '" & IIf(VGRegEnt = 1, "I", "S") & "'", VGCNx, adOpenStatic, adLockOptimistic
frmReferencia.Conectar Adodc3, "SELECT TT_CODMOV,TT_DESCRI FROM Tabtransa where  TT_tipmov = '" & IIf(VGRegEnt = 1, "I", "S") & "'"
frmReferencia.Label1.Caption = "Transacciones"
frmReferencia.Show vbModal
Adodc3.Close
If vGUtil(1) <> "" Then
    If vGUtil(1) <> "SK" Then
        MsgBox "Debe ser Transferencia por Elaboración de Kids", vbInformation, "Mensaje"
        TxTransa.SetFocus: Exit Sub
    End If
    TxTransa = vGUtil(1)
    buscar_trans
    lbltrans = Mid(vGUtil(2), 1, 30)
End If
If TxTransa.text <> "" Then TxTransa_KeyPress (13)
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text4 <> "" Then
   Text4 = Format(Text4, "00000000000")
   Tabula (KeyAscii)
End If
If KeyAscii = 13 Then Tabula (KeyAscii)
End Sub

'**************** num ref *********************
Private Sub Text6_DblClick()
Dim Adodc3 As ADODB.Recordset
Set Adodc3 = New ADODB.Recordset
Adodc3.Open "SELECT TDO_TIPDOC,TDO_DESCRI  FROM TIPO_DOCU", VGCNx, adOpenStatic, adLockOptimistic
frmReferencia.Conectar Adodc3, "SELECT TDO_TIPDOC,TDO_DESCRI  FROM TIPO_DOCU"
frmReferencia.Label1.Caption = "Tipo de Documentos"
frmReferencia.Show vbModal
Adodc3.Close
If vGUtil(1) <> "" Then Text6 = (vGUtil(1))
If vGUtil(2) <> "" Then lbltipref = (vGUtil(2))
If Text6 <> "" Then
   Text1.SetFocus
Else
    Text6.SetFocus
End If
End Sub
Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 112 Then
    Text6_DblClick
ElseIf KeyCode = 46 Then
     lbltipref = ""
 End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If Len(Text6) = 2 Then
         Text6 = UCase(Text6)
         lbltipref = Mid(ValidarDoc(Text6), 1, 15)
         If lbltipref = "" Then
            Enfoque Text6
            Exit Sub
         End If
         Text1.SetFocus
  Else
         Text6 = ""
         SendKeys "{tab}"
         KeyAscii = 0
  End If
End If
If KeyAscii = 8 Then lbltipref = ""
End Sub

 '***** Orden de compra
Private Sub Text8_KeyPress(KeyAscii As Integer)

Dim criterio As String
If KeyAscii = 13 Then
      Text8 = Trim(Text8)
      If Text8 <> "" Then
            'criterio = "CANUMORD = " & Chr$(34) + Text8.text + Chr$(34) & "AND  CACODPRO = " & Chr$(34) + TxtProveedor.text + Chr$(34)
            'Data1.Recordset.FindFirst criterio
            'If Not Data1.Recordset.NoMatch Then
            '  MsgBox "La Orden de Compra ya ha sido registrada !", vbExclamation, mensaje1
            '  Exit Sub
            'End If
      End If
      Text9.SetFocus
End If
End Sub

Private Sub Text9_DblClick()
FormAyuda.Show 1
End Sub
'Autorizado
Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
      If Trim(Text9) <> "" Then
              'muestra
              If Trim(validarautorizado(Text9)) = "" Then
                      MsgBox "No existe el Autorizado", vbInformation, "Mensaje"
                      If Text9.Enabled And Text9.Visible Then Text9.SetFocus
                      Exit Sub
              End If
              lblauto = Mid(validarautorizado(Text9), 1, 10)
              Tabula (KeyAscii)
      Else
          Tabula (KeyAscii)
      End If
 End If
End Sub
'Numeracion
Private Sub muestra()
Dim nument As Long, numsal As String
Dim rs As Recordset, RSQL As String
If Trim(VGAlma) <> "" Then
    RSQL = "select  TANUMENT, TANUMSAL from TabAlm  WHERE TAALMA='" & VGAlma & "' "
    'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
    Set rs = VGCNx.Execute(RSQL)
    nument = IIf(IsNull(rs(0)), 1, rs(0))
    numsal = IIf(IsNull(rs(1)), 1, rs(1))
    If VGRegEnt = 1 Then
        Text4.text = Format(Val(nument) + 1, "00000000000")
        ContaSalida = Format(Val(numsal) + 1, "00000000000")
    Else
        Text4.text = Format(Val(nument) + 1, "00000000000")
        ContaSalida = Format(Val(numsal) + 1, "00000000000")
    End If
    Command1.Visible = True
    Cmdgrabar.Visible = True
    Command3.Visible = True
    Command7.Visible = True
Else
   MsgBox "No ningún Almacen Activo", vbInformation, "Información"
End If
End Sub

Public Sub grabastk(Art As String, RegEs As Integer, PreUni As Double)
Dim cadena As String, criterio As String
Dim entrada As Boolean
On Error GoTo GrabErr
cadena = Art
criterio = " STCODIGO = '" & cadena & "'"
criterio = criterio + "  and  STALMA = '" & VGAlma & "'"
cadena = "select * from stkart where " & criterio
Adodc3.Open cadena, VGCNx, adOpenStatic, adLockOptimistic
If Not Adodc3.EOF Then Adodc3.Filter = criterio
If Not Adodc3.EOF Then      'si existe el articulo
    canttemp = Adodc3("STSKDIS")     ' revisar si validar en creacion
    If RegEs = 1 Then 'Entrada
        Adodc3("STKFECULT") = DTPicker1.Value
        Adodc3("STSKDIS") = Adodc3("STSKDIS") + CANTIDAD
        'aqui actualiza
        If Not IsNull(Adodc3("STKPREPRO")) Then
            precioprom = Adodc3("STKPREPRO")
            If PreUni <> 0 Then
                   Adodc3("STKPREULT") = PreUni * VGTipCamb 'el precio
                   If PreUni <> 0 And (canttemp + CANTIDAD) <> 0 Then 'Valorizacion
                      Adodc3("STKPREPRO") = Round((precioprom * canttemp + CANTIDAD * Val(Val(PreUni) * VGTipCamb)) / (canttemp + CANTIDAD), 6)
                   End If
              End If
         Else
              precioprom = 0
              If PreUni <> 0 Then
                 Adodc3("STKPREPRO") = Round(Val(PreUni) * VGTipCamb, 6) 'el precio
                 If PreUni <> 0 Then 'Valorzacion
                   Adodc3("STKPREULT") = Round(Val(PreUni) * VGTipCamb) 'el precio
                   Adodc3("STKFECULT") = DTPicker1.Value
                 End If
              End If
           End If
    Else 'para la salida
         Adodc3("STSKDIS") = Adodc3("STSKDIS") - CANTIDAD
         'aqui actualiza
         If Not IsNull(Adodc3("STKPREPRO")) Then
            precioprom = Round(Adodc3("STKPREPRO"), 6)
         Else
            precioprom = 0
         End If
    End If
Else
       Adodc3.AddNew                  'existe
       Adodc3("STALMA") = VGAlma   '"01"
       Adodc3("STCODIGO") = Art
       Adodc3("STKFECULT") = DTPicker1.Value
       If RegEs = 1 Then 'entrda
           Adodc3("STSKDIS") = CANTIDAD
           Adodc3("STKPREULT") = Val(PreUni) * VGTipCamb    'el costo de ingreso
           If PreUni <> 0 Then
                 Adodc3("STKPREPRO") = Round(Val(PreUni), 6) '******el  costo = costo prom
           End If
       End If
End If
Adodc3.Update
Adodc3.Filter = ""
entrada = IIf(RegEs = 1, True, False)
'Call ValMes(VGAlma, entrada) 'para la valorizacion
Adodc3.Requery
Adodc3.Close
Exit Sub
GrabErr:
 MsgBox Err.Description
End Sub

Private Sub buscar_trans()
Dim criterio As String
Dim rs As Recordset
Dim RSQL As String
TxTransa = UCase(LTrim(TxTransa))
If TxTransa = "TD" And VGRegEnt Then
  MsgBox "El tipo de transaccion no puede ser usado para registrar !", vbOKOnly, "Error"
  lbltrans = ""
  TxTransa.SetFocus
  Exit Sub
End If
'Busco la transaccion
RSQL = "select  *  from TabTransa  where TT_CODMOV ='" & TxTransa.text & "' and TT_TIPMOV ='" & dato & "'"
'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
Set rs = VGCNx.Execute(RSQL)

If rs.EOF Then
   MsgBox "El tipo de transaccion no existe !", vbOKOnly, "Error"
   LIMPIACABECERA
   TxTransa.SetFocus
   Exit Sub
End If
lbltrans = Mid(rs("TT_DESCRI"), 1, 30)
       
If Not IsNull(rs("TT_CONT")) Then
    TT_CONTADOR = rs("TT_CONT")
Else
    MsgBox "El tipo de transacción no esta inicializara !" & Chr(13) & "Para inicializarla ir a la tabla de Transacción", vbOKOnly + vbExclamation, "Error"
    Exit Sub
End If
End Sub

Private Sub grabacabecera(Dat As String, Tip As String, num As String, RegEs As Integer)
Dim criterio As String, cadena As String
Dim FACTOR As Double, uSql As String

On Error GoTo GrabErr
'Desea grabar el registro

If num <> "" Then
    If Not VGActualizar Then
      adodc1.AddNew
      adodc1("CAALMA") = VGAlma     '"01"
      adodc1("CANUMDOC") = Mid$(UCase$(num), 1, 11)
    Else
      criterio = " CANUMDOC = '" & num & "' "
      criterio = criterio + " and  CAALMA = '" & VGAlma & "'"
      If Not adodc1.EOF Then adodc1.Filter = criterio
    End If
    adodc1("CATIPMOV") = Dat
    adodc1("CATD") = Tip
    adodc1("CAHORA") = Format(Time, "hh:mm:ss")
    adodc1("CAFECDOC") = Format(DTPicker1.Value, "dd/mm/yyyy")           ' CDate(Text2.text)
   
 
      adodc1("CACODMOV") = IIf(Dat = "S", Mid$(UCase$(TxTransa.text), 1, 2), "29")
   
     If Trim(Text8.text) <> "" And RegEs = 1 Then
      adodc1("CANUMORD") = Mid$(UCase$(Text8.text), 1, 11)
    Else
      adodc1("CANUMORD") = " "
    End If
    If Text9.Visible And Trim(Text9) <> "" Then
      adodc1("CASOLI") = Mid$(UCase$(Text9.text), 1, 3)
    Else
      adodc1("CASOLI") = " "
    End If
    adodc1("CAUSUARI") = UCase(VGUsuario)
    adodc1("CACODMON") = VGCodMon
    adodc1("CASITGUI") = "V"
    'Data1.Recordset("CASITUA") = "V"
    adodc1("CAESTIMP") = "V"
    adodc1("catipotransf") = "TR"
    adodc1("canrotransf") = ndato
    adodc1("cafecact") = Now
    adodc1("empresacodigo") = VGparametros.empresacodigo
    
    adodc1.Update
    adodc1.Filter = ""
    adodc1.Requery
End If
Exit Sub
GrabErr:
    MsgBox Err.Description
End Sub
Function ValidarDoc(txt As TextBox) As String
Dim rs As Recordset, RSQL As String
RSQL = "select TDO_DESCRI  from TIPO_DOCU  where TDO_TIPDOC='" & txt.text & "'"
'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
Set rs = VGCNx.OpenRecordset(RSQL)
If rs.EOF Then
    MsgBox "El tipo de codigo no existe !", vbOKOnly, "Error"
    ValidarDoc = ""
    txt.SetFocus
    Exit Function
End If
ValidarDoc = rs(0)
rs.Close
End Function

Private Sub LIMPIACABECERA()
'TxTransa = "":lbltrans = "":
Text6 = "": Text8 = "": Text4 = "": Text1 = "": Text9 = ""
lbltipref = "": lblauto = "": TxCantidad = "": TxCambio.Enabled = False: TxCambio = "0"
CmbMoneda.ListIndex = 0
End Sub

Private Sub inicializar()
'TxTransa = ""
TxCodKid = "": lblnomkits = "": Txcant = "": TxSaldo = 0
Text6 = "": Text8 = "": Text4 = "": Text1 = "": Text9 = "": TxCantidad = "": TxCambio = "0": TxCambio.Enabled = False
CmbMoneda.ListIndex = 0
inicializaFG
Command1.Visible = True
Command3.Visible = True
Command7.Visible = True
Cmdgrabar.Visible = True
LIMPIACABECERA
End Sub

Private Sub ValMes(almacen As String, entrada As Boolean)
  Dim cadena As String
  Dim criterio As String
  Dim adoreg As ADODB.Recordset
  Dim RSQL As String
  Dim Cantent As Double
  Dim Cantsal As Double
  Dim mespro As String
  Dim uSql As String
  On Error GoTo Err
  mespro = Year(DTPicker1) & Format(Month(DTPicker1), "00")
  cadena = TxCodKid
  RSQL = "select SMCANENT,SMCANSAL FROM MoResMes where  SMALMA = '" & almacen & "' and SMMESPRO = '" & mespro & "' AND  SMCODIGO= '" & cadena & "'"  '
  Set adoreg = New ADODB.Recordset
  adoreg.Open RSQL, VGCNx, adOpenDynamic, adLockOptimistic
  If Not adoreg.EOF Then 'existe
      If entrada Then
          Cantent = adoreg(0) + CANTIDAD
          uSql = "Update MoResMes set SMCANENT = " & Cantent & " where SMALMA='" & almacen & "'  and  SMCODIGO ='" & cadena & "' AND SMMESPRO ='" & mespro & "' "
      Else
          Cantsal = adoreg(1) - CANTIDAD
          uSql = "Update MoResMes set SMCANSAL = " & Cantsal & " where SMALMA='" & almacen & "' and   SMCODIGO ='" & cadena & "' AND SMMESPRO ='" & mespro & "' "
      End If
  Else
      If entrada Then
          Cantent = CANTIDAD
          Cantsal = 0
      Else
          Cantsal = CANTIDAD
          Cantent = 0
      End If
      uSql = "insert into MoResMes (SMALMA,SMCODIGO,SMMESPRO,SMCANENT,SMCANSAL,SMSALDOINI) VALUES ('" & almacen & "','" & cadena & "','" & mespro & "' ," & Cantent & "," & Cantsal & ",0) "
   End If
   VGCNx.Execute uSql
  Exit Sub
Err:
   MsgBox Err.Description
End Sub

Private Sub inicializaFG()
 MSFlexGrid1.FormatString = "^ Codigo               |<Descripción                                                                        |>Cant.Reg.  |>Cant.Desarm.|>Cant.Dispon."
End Sub
Function existe_numdoc(text As String) As Boolean
Dim criterio As String
  criterio = " CANUMDOC = '" & Format(text, String(11, "0")) & "'"
  criterio = criterio + " and  CAALMA = '" & VGAlma & "'"
  criterio = criterio + " and  CATD = '" & tipo & "'"
  RSQL = "select * from movalmcab where " & criterio
 adodc1.Open RSQL, VGCNx, adOpenDynamic, adLockOptimistic
  If Not adodc1.EOF Then
         MsgBox "El Número del documento ya ha sido registrado: " & Format(text, String(10, "0")) & " !", vbExclamation, "Error"
         existe_numdoc = True
  Else
         existe_numdoc = False
  End If
  adodc1.Filter = ""
End Function

Function validarautorizado(text As TextBox) As String
  Dim RSQL As String
  Dim rs As Recordset
  Dim codayu As String
  codayu = 12
  RSQL = "Select TCLAVE,TDESCRI from TABAYU  where TCOD= '" & codayu & "' and  Tclave ='" & Trim(text) & "'"
 ' Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
  Set rs = VGCNx.Execute(RSQL)
   If Not rs.EOF Then 'existe
     validarautorizado = rs(1)
   Else
     validarautorizado = ""
  End If
  rs.Close
End Function

Private Sub imprimir(tipo As String)
Dim cadena As String
On Error GoTo ErrImp
 CrystalReport1.Reset
  CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "REPNOTAING.rpt"
                            
 CrystalReport1.Connect = VGcadenareport2
 CrystalReport1.StoredProcParam(0) = VGParamSistem.BDEmpresa
 CrystalReport1.StoredProcParam(1) = VGAlma
 CrystalReport1.DiscardSavedData = True
 CrystalReport1.Destination = crptToWindow
 CrystalReport1.formulas(0) = "fecha='" & DTPicker1.Value & "'"
 CrystalReport1.formulas(1) = "xtrans = '" & lbltrans.Caption & "' "
 CrystalReport1.formulas(2) = "xtd = '" & Trim(tipo) & "' "
 CrystalReport1.formulas(3) = "xndoc = '" & Text4.text & "' "
  If tipo = "NI" Then
     CrystalReport1.WindowTitle = "RepNotaIng -- Impresion de Notas de Ingreso"
     CrystalReport1.formulas(4) = "Xnalma = '" & VGAlma & "' "
     CrystalReport1.formulas(5) = "Dalma = '" & lblalmacen & "' "
     CrystalReport1.formulas(6) = "AlmaDes = '" & VGAlma & "' "
     CrystalReport1.formulas(7) = "Dalmades = '" & lblalmacen.Caption & "' "
     CrystalReport1.StoredProcParam(2) = tipo
     CrystalReport1.StoredProcParam(3) = Text4.text

  Else
     CrystalReport1.StoredProcParam(2) = tipo
     CrystalReport1.StoredProcParam(3) = ContaSalida
     CrystalReport1.WindowTitle = "RepNotaIng -- Impresion de Notas de Almacen"
     CrystalReport1.PrintFileLinesPerPage = 66
     CrystalReport1.formulas(4) = "Xnalma = '" & VGAlma & "' "
     CrystalReport1.formulas(5) = "Dalma = '" & lblalmacen.Caption & "' "
     CrystalReport1.formulas(6) = "AlmaDes = '" & VGAlma & "' "
     CrystalReport1.formulas(7) = "Dalmades = '" & lblalmacen & "' "
 End If
 CrystalReport1.formulas(8) = "NRef = '" & Text1.text & "' "
 CrystalReport1.formulas(9) = "DocRef = '" & Text6.text & "' "
 CrystalReport1.formulas(10) = "TTrans = '" & TxTransa.text & "' "
 CrystalReport1.formulas(11) = "emp = '" & VGparametros.RucEmpresa & "'"
 CrystalReport1.WindowShowPrintBtn = True
 CrystalReport1.WindowShowRefreshBtn = True
 CrystalReport1.WindowShowSearchBtn = True
 CrystalReport1.WindowShowPrintSetupBtn = True
 CrystalReport1.WindowState = crptMaximized
  If CrystalReport1.Status <> 2 Then
     CrystalReport1.Action = 1
      VGCNx.Execute "Update MovAlmCab Set CaEstImp = 'I' Where CATD = '" & tipo & "' and CANUMDOC = '" & Text4.text & "'"
 End If
 Exit Sub
ErrImp:
     MsgBox Err.Description
     Resume Next
End Sub

Private Sub GrabarDetKit(Art As String, Tip As String, Cont As Integer, Descr As String, cant As Double, Preu As Double)
Dim criterio As String
If Not VGActualizar Then
    Adodc22.AddNew
Else
    criterio = "DECODIGO = '" & Art & "' "
    criterio = criterio + " and  DEALMA = '" & VGAlma & "' "
    If Not Adodc22.EOF Then Adodc22.Filter = criterio
End If
Adodc22("DEALMA") = VGAlma
Adodc22("DETD") = tipo
Adodc22("DENUMDOC") = Text4.text
Adodc22("DEITEM") = Cont
Adodc22("DECODIGO") = Art
Adodc22("DEDESCRI") = Descr
Adodc22("DECANTID") = cant
If Preu <> 0 Then    'si tiene precio
    Adodc22("DEPRECIO") = Val(Preu) '* VGTipCamb '******el precio
    Adodc22("DETIPCAM") = VGTipCamb
Else
    Adodc22("DEPRECIO") = 0
End If
alma = VGAlma

Adodc22.Update
Adodc22.Filter = ""
Adodc22.Requery
End Sub

Sub Alinear()
TxCantidad.Width = MSFlexGrid1.CellWidth
TxCantidad.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
TxCantidad.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
TxCantidad.Height = MSFlexGrid1.CellHeight
End Sub

Private Sub TxCantidad_KeyPress(KeyAscii As Integer)
If NumPto(KeyAscii) Then
    Select Case KeyAscii
      Case Is = 13
         MSFlexGrid1.text = TxCantidad.text
         TxCantidad.Visible = False
         TxCantidad.text = ""
      Case Is = 27
         TxCantidad.Visible = False
         TxCantidad.text = ""
    End Select
Else
    KeyAscii = 0
End If
End Sub


Function VerificarSTOCK() As Boolean
Dim Xstock As Double
Dim Negativos As Boolean
Dim n As Long

VerificarSTOCK = True

TxCantidad.Visible = False
If MSFlexGrid1.TextMatrix(1, 1) = "" Then
   Txcant = 0
   Txcant.SetFocus
   Exit Function
Else
   LoadReceta (TxCodKid)
End If
Negativos = True

TxSaldo = Val(TxSaldo) - Val(Txcant)
If TxSaldo < 0 Then
   Negativos = False
End If

For n = 1 To MSFlexGrid1.Rows - 1
    MSFlexGrid1.TextMatrix(n, 3) = Val(MSFlexGrid1.TextMatrix(n, 2)) * Val(Txcant)
    MSFlexGrid1.TextMatrix(n, 4) = Val(MSFlexGrid1.TextMatrix(n, 4)) + Val(MSFlexGrid1.TextMatrix(n, 3))
Next

If Negativos = False Then 'El negativo es el Cantidad contra el saldo
   MsgBox "El Kit No Dispone de Stock Suficiente para ser Desarmado ", vbInformation, "Verifique ....."
   LoadReceta (TxCodKid)
   VerificarSTOCK = False
End If
End Function
Sub LoadReceta(ByVal arCodKit As String)
   Dim rs As New ADODB.Recordset
   Dim SQL As String
   SQL = "SELECT KITS.CODART, MAEART.ADESCRI, KITS.CANART, 0 AS Expr1, STKART.STSKDIS FROM (KITS INNER JOIN STKART ON KITS.CODKIT = STKART.STCODIGO) LEFT JOIN MAEART ON KITS.CODART =MAEART.ACODIGO where STALMA='" & Txtalma.text & "' AND  KITS.CODkit='" & arCodKit & "'"
   rs.Open SQL, VGCNx, adOpenForwardOnly, adLockReadOnly
   If Not rs.EOF Then
      'Set MSFlexGrid1.DataSource = rS
       MSFlexGrid1.Clear
       MSFlexGrid1.Rows = 2
       'varform.MSFlexGrid1.Row = 1
       MSFlexGrid1.AddItem rs!codart & Chr(9) & rs!ADESCRI & Chr(9) & rs!CANART & Chr(9) & 0 & Chr(9) & ClsTock.SaldoArti(Txtalma.text, rs!codart, VGCNx), 1
       rs.MoveNext
       Do While Not rs.EOF
          MSFlexGrid1.AddItem rs!codart & Chr(9) & rs!ADESCRI & Chr(9) & rs!CANART & Chr(9) & 0 & Chr(9) & ClsTock.SaldoArti(Txtalma.text, rs!codart, VGCNx), 1
          rs.MoveNext
       Loop
       MSFlexGrid1.Rows = MSFlexGrid1.Rows - 1

      
   End If
   MSFlexGrid1.FormatString = "^ Codigo                 |<Descripción                                                                      |>Cant.Reg.  |>Cant.Desarm.|>Cant.Dispon."
   rs.Close
   TxSaldo = ClsTock.SaldoArti(Txtalma.text, arCodKit, VGCNx)
End Sub

Function VerificaIngresos() As Boolean
Dim n, fila As Long
Dim nFactor As Double
VerificaIngresos = True

For n = 1 To MSFlexGrid1.Rows - 1
    If Val(MSFlexGrid1.TextMatrix(n, 3)) <= 0 Then Exit Function
Next

If Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)) <> 0 Then
   nFactor = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) / Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
End If


For n = 1 To MSFlexGrid1.Rows - 1
    If nFactor <> MSFlexGrid1.TextMatrix(n, 3) / Val(MSFlexGrid1.TextMatrix(n, 2)) Then
       MsgBox "Una de Las Cantidades  no Corresponde al Nro. Calculado para el Armado", vbInformation, "Verifique ....."
       VerificaIngresos = False
       Exit Function
    Else
       Txcant = nFactor
    End If
Next

End Function

