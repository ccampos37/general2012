VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmArmadoKits 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Armado de Kids"
   ClientHeight    =   7125
   ClientLeft      =   2160
   ClientTop       =   3285
   ClientWidth     =   8895
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   8895
   Begin VB.TextBox Txtalma 
      Height          =   285
      Left            =   5040
      TabIndex        =   43
      Text            =   "Text2"
      Top             =   2880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton command5 
      Caption         =   "&Reporte"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5130
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   5985
      Width           =   1065
   End
   Begin VB.Frame Frame5 
      Height          =   1164
      Left            =   2304
      TabIndex        =   37
      Top             =   4035
      Visible         =   0   'False
      Width           =   4512
      Begin MSComctlLib.ProgressBar PBar 
         Height          =   228
         Left            =   324
         TabIndex        =   38
         Top             =   720
         Width           =   3900
         _ExtentX        =   6879
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lbmsg 
         Caption         =   "Procesando...."
         Height          =   264
         Left            =   312
         TabIndex        =   39
         Top             =   360
         Width           =   3900
      End
   End
   Begin VB.TextBox TxCantidad 
      Alignment       =   1  'Right Justify
      Height          =   288
      Left            =   2016
      TabIndex        =   35
      Text            =   "TxCantidad"
      Top             =   3315
      Visible         =   0   'False
      Width           =   768
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlexGrid1 
      Height          =   2490
      Left            =   75
      TabIndex        =   13
      Top             =   3270
      Width           =   8730
      _ExtentX        =   15399
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
   Begin VB.CommandButton Command2 
      Caption         =   "&Agregar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   144
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   5985
      Visible         =   0   'False
      Width           =   1065
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   405
      Top             =   4995
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3960
      Picture         =   "FrmArmadoKits.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5985
      Width           =   1065
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Limpiar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1305
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5985
      Width           =   1065
   End
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
      Height          =   975
      Left            =   2805
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5985
      Width           =   1065
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
      Height          =   2784
      Left            =   90
      TabIndex        =   18
      Top             =   45
      Width           =   8724
      Begin VB.TextBox TxPrecio 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   5040
         MaxLength       =   10
         TabIndex        =   11
         Top             =   2376
         Width           =   1308
      End
      Begin VB.TextBox Txcant 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   1620
         MaxLength       =   10
         TabIndex        =   10
         Top             =   2412
         Width           =   876
      End
      Begin VB.TextBox TxCodKid 
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   1620
         TabIndex        =   9
         Top             =   2052
         Width           =   1956
      End
      Begin VB.TextBox TxCambio 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5052
         TabIndex        =   8
         Top             =   1680
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox CmbMoneda 
         Height          =   288
         ItemData        =   "FrmArmadoKits.frx":0442
         Left            =   1644
         List            =   "FrmArmadoKits.frx":044C
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1680
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox TxTransa 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7530
         MaxLength       =   2
         TabIndex        =   2
         Text            =   "28"
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5052
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1644
         MaxLength       =   2
         TabIndex        =   3
         Top             =   960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1644
         MaxLength       =   20
         TabIndex        =   5
         Top             =   1320
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   5052
         MaxLength       =   3
         TabIndex        =   6
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5052
         MaxLength       =   10
         TabIndex        =   4
         Top             =   960
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   288
         Left            =   1680
         TabIndex        =   0
         Top             =   240
         Width           =   1452
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   108331009
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   36161
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayusalida 
         Height          =   375
         Left            =   1680
         TabIndex        =   45
         Top             =   600
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   661
         XcodMaxLongitud =   0
         xcodwith        =   100
         NomTabla        =   "tabtransa"
         TituloAyuda     =   "Transaciones"
         ListaCampos     =   "tt_codmov(1),tt_descri(1),tt_dr(1),tt_codtrans_auto(1),tt_clie(2),tt_dr(2),intercompanias(1)"
         XcodCampo       =   "tt_codmov"
         XListCampo      =   "tt_descri"
         ListaCamposDescrip=   "Codigo,Descripcion,doc.ref.,trans.auto,Ctrl.Cliente,Doc.ref."
         ListaCamposText =   "tt_codmov,tt_descri,tt_dr,tt_codtrans_auto,tt_clie,tt_dr,intercompanias"
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   36
         Top             =   2445
         Width           =   750
      End
      Begin VB.Label lblnomkits 
         Caption         =   "lblnomkits"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   3816
         TabIndex        =   34
         Top             =   2088
         Width           =   4116
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Codigo de Kit's :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   33
         Top             =   2085
         Width           =   1155
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Cambio :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3810
         TabIndex        =   31
         Top             =   1710
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Moneda :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   1680
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Doc. :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   270
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Transaccion :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   630
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Num. Doc :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   3855
         TabIndex        =   27
         Top             =   270
         Width           =   795
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tip Doc Ref :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   990
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Orden Compra :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   1320
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Autorizacion :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   3855
         TabIndex        =   24
         Top             =   1320
         Width           =   990
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Num. Ref :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   3855
         TabIndex        =   23
         Top             =   990
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label lbltrans 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   22
         Top             =   1110
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lbltipref 
         Caption         =   "lbltipref"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   21
         Top             =   990
         Width           =   2295
      End
      Begin VB.Label lblauto 
         Caption         =   "lblauto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   5652
         TabIndex        =   20
         Top             =   1320
         Width           =   2256
      End
      Begin VB.Label LblCC 
         AutoSize        =   -1  'True
         Caption         =   "Precio :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3855
         TabIndex        =   19
         Top             =   2445
         Width           =   540
      End
   End
   Begin VB.CommandButton Command7 
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
      Height          =   975
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5985
      Width           =   1065
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   2205
      Left            =   195
      TabIndex        =   14
      ToolTipText     =   "Doble Click o F1 para la ayuda"
      Top             =   3315
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   3889
      _Version        =   393216
      RowHeightMin    =   290
      Appearance      =   0
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuALMACEN 
      Height          =   330
      Left            =   1200
      TabIndex        =   42
      Top             =   2880
      Width           =   3780
      _ExtentX        =   6668
      _ExtentY        =   582
      XcodMaxLongitud =   0
      xcodwith        =   100
      NomTabla        =   "tabalm"
      TituloAyuda     =   "Almacenes"
      ListaCampos     =   "TAALMA(1),TADESCRI(1),empresacodigo(1)"
      XcodCampo       =   "TAALMA"
      XListCampo      =   "TADESCRI"
      ListaCamposDescrip=   "Codigo,Descripcion,Empresa"
      ListaCamposText =   "TAALMA,TADESCRI,empresacodigo"
   End
   Begin VB.Label lblalmacen 
      Caption         =   "Label12"
      Height          =   135
      Left            =   5760
      TabIndex        =   44
      Top             =   3000
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Almacen :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   195
      Left            =   315
      TabIndex        =   40
      Top             =   2895
      Width           =   705
   End
End
Attribute VB_Name = "FrmArmadoKits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adodc1 As New ADODB.Recordset
Dim Adodc2 As New ADODB.Recordset
Dim Adodc22 As New ADODB.Recordset
Dim Adodc3 As New ADODB.Recordset
Dim ContaSalida As String
Dim Empresa As String
Dim SQL As String
Dim RSQL As String
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
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

'Eliminar
Private Sub Command1_Click()
LoadReceta (TxCodKid)
Txcant = 0
End Sub

Private Sub Command2_Click()
' MSFlexGrid1_DblClick
End Sub

'Limpiar
Private Sub Command3_Click()
TxCodKid = "": lblnomkits = "": Txcant = "": TxPrecio = ""
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
Dim ndato As String
On Error GoTo GrabErr

Txcant_KeyPress (13)

If Trim(TxTransa) = "" Then
    MsgBox "Debe Ingresar el Movimiento", vbInformation, "Mensaje"
    TxTransa.SetFocus
    Exit Sub
End If

If Trim(Txtalma) = "" Then
    MsgBox "Debe Ingresar el codigo del almacen", vbInformation, "Mensaje"
    Txtalma.SetFocus
    Exit Sub
 Else
    VGAlma = Txtalma.text
End If

If TxCambio.Enabled Then
    If Val(TxCambio) = 0 Then
        MsgBox "Ingrese Tipo de Cambio", vbInformation, "Mensaje"
        TxCambio.SetFocus: Exit Sub
    Else
        VGTipCamb = TxCambio
    End If
End If

If Trim(MSFlexGrid1.TextMatrix(1, 0)) = "" Then
      MsgBox "No se puede grabar,debe adicionar registro", vbInformation, mensaje1
      Exit Sub
End If
'Numeracion
If Trim(Text4) = "" Then muestra

'Verificar I/S
'RMM****************************************
 If VerificaIngresos = False Then Exit Sub
'RMM****************************************


If Not IsNumeric(Text4) Then
     MsgBox "Numero de Documento no consecutivo", vbExclamation, "Aviso"
     Exit Sub
End If

  

Screen.MousePointer = 11

'***
Set cSel1 = Nothing
Frame5.Visible = True
lbmsg.Caption = "Guardando  Datos .......! "
VGCNx.BeginTrans
Set cSel1 = Nothing

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
Set adodc1 = Nothing
criterio = " CANUMDOC = '" & Text4 & "' "
criterio = criterio + " and  CAALMA = '" & VGAlma & "'"
criterio = criterio + " and  CATD = '" & tipo & "'"
RSQL = " select * from movalmcab where " & criterio & ""
adodc1.Open RSQL, VGCNx, adOpenDynamic, adLockOptimistic

If adodc1.RecordCount > 0 Then
   MsgBox "El Nro. del Documento de Ingreso a Generar ya existe......!" & Chr(10) & "Verifique el Correlativo de Salidas del Almacen ", vbInformation, "Verifica"
   adodc1.Close
   Exit Sub
End If

Call grabacabecera("I", "NI", Text4, 1, ndato) 'Entrada
Call grabacabecera("S", "NS", ContaSalida, 0, ndato) 'Salida  ContaSalida
'***
adodc1.Close
FACTOR = 1
contador = 1
'Graba detalle DE INGRESO DEL KIT
NumDoc = Text4
 nIt = 0
criterio = " deNUMDOC = '" & NumDoc & "' "
criterio = criterio + " and  deALMA = '" & VGAlma & "'"
criterio = criterio + " and  deTD ='" & tipo & "'"
cadena = " select * from movalmdet where " & criterio & ""
Adodc2.Open cadena, VGCNx, adOpenDynamic, adLockOptimistic
      
CANTIDAD = MSFlexGrid1.TextMatrix(contador, 2)
If (IIf(VGRegEnt = 1, True, True)) Then      'verificastk
    cadena = TxCodKid  'codigo del Kit
    CANTIDAD = Txcant  'Cantidad del kit

    If Not VGActualizar Then
           Adodc2.AddNew
    Else
           criterio = "DECODIGO = '" & cadena & "'"
           criterio = criterio + " and  DEALMA = '" & VGAlma & "'"
           If Not Adodc2.EOF Then Adodc2.Filter = criterio
    End If
    Adodc2("DEALMA") = VGAlma
    Adodc2("DETD") = "NI"
    Adodc2("DENUMDOC") = Text4.text
    Adodc2("DEITEM") = contador
    Adodc2("DECODIGO") = cadena
    Adodc2("DEDESCRI") = lblnomkits
    Adodc2("DECANTID") = CANTIDAD
    If MSFlexGrid1.TextMatrix(contador, 2) <> "" Then 'Cantidad Ingresada
         Call grabastk(cadena, 1, Val(TxPrecio))
         If Trim(TxPrecio) <> "" Then    'si tiene precio
             Adodc2("DEPRECIO") = TxPrecio '* VGTipCamb '******el precio
             Adodc2("DETIPCAM") = VGTipCamb
         Else
             Adodc2("DEPRECIO") = 0
         End If
         alma = VGAlma
         '****
    End If
    Adodc2.Update
    Adodc2.Filter = ""
  End If
Adodc2.Requery
Adodc2.Close
'*************************************************************************
'Graba Detalle************************************************************
Set Adodc22 = Nothing
PBar.Min = 0
PBar.Max = (MSFlexGrid1.Rows - 1) * 100
tipo = "NS"
nIt = 0
NumDoc = ContaSalida
criterio = " deNUMDOC = '" & NumDoc & "' " 'ContaSalida
criterio = criterio + " and  deALMA = '" & VGAlma & "'"
criterio = criterio + " and  deTD = '" & tipo & "'"
cadena = " select * from movalmdet where " & criterio & ""
Adodc22.Open cadena, VGCNx, adOpenDynamic, adLockOptimistic

For n = 1 To MSFlexGrid1.Rows - 1
        Frame5.Refresh
        PBar.Value = n * 80
        lbmsg.Caption = "Guardando  Datos .......! "
        nIt = nIt + 1
        Call GrabarDetKit(MSFlexGrid1.TextMatrix(n, 0), "NS", nIt, Devolver_Dato(1, MSFlexGrid1.TextMatrix(n, 0), "MaeArt", "Acodigo", False, "Adescri"), MSFlexGrid1.TextMatrix(n, 3), 0)
Next
Adodc22.Close
Frame5.Visible = False
'*************************************************************************
Dim rs As Recordset

If MsgBox("¿Desea Imprimir Guia de Ingreso ?", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    imprimir ("NI")
End If
If MsgBox("¿Desea Imprimir Guia de Salida ?", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    imprimir ("NS")
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
'  Resume
  Screen.MousePointer = 1
  Frame5.Visible = False
   Exit Sub
  Resume
End Sub

Private Sub command5_Click()

Dim cadena As String
Dim cNomRepor  As String

cNomRepor = "al_armadokits.RPT"
If Trim(cNomRepor) <> "" Then
    CrystalReport1.Reset
    CrystalReport1.WindowTitle = "Reporte de Familia de Articulo"
    CrystalReport1.ReportFileName = VGParamSistem.RutaReport + "\" + cNomRepor
 
                        
    CrystalReport1.Connect = VGcadenareport2
    CrystalReport1.StoredProcParam(0) = VGParamSistem.BDEmpresa
    CrystalReport1.StoredProcParam(1) = TxCodKid.text
    CrystalReport1.StoredProcParam(2) = VGAlma
    CrystalReport1.StoredProcParam(3) = Txcant.text
    
    CrystalReport1.formulas(0) = "emp ='" & VGparametros.RucEmpresa & "'"
    CrystalReport1.formulas(1) = "almacen ='" & VGAlma & "'"
    
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.WindowState = crptMaximized
    If CrystalReport1.Status <> 2 Then
        CrystalReport1.Action = 1
    End If
Else
    MsgBox "No existe el nombre del Reporte, verifique en Formatos", vbInformation, "Información"
    Exit Sub
End If

End Sub

Private Sub Command7_Click()
Unload Me
End Sub

Private Sub Ctr_ayuAlmacen_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Txtalma = Ctr_ayuAlmacen.xclave
VGAlma = Ctr_ayuAlmacen.xclave
Empresa = ColecCampos("empresacodigo")
Call muestra
End Sub

Private Sub Ctr_Ayusalida_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
TxTransa = ColecCampos("tt_codmov")
lbltrans = ColecCampos("tt_descri")
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
central Me
Txtalma.text = ""

Set adodc1 = New ADODB.Recordset
Set Adodc2 = New ADODB.Recordset
Set Adodc22 = New ADODB.Recordset
Set Adodc3 = New ADODB.Recordset
     
VGActualizar = False
VGSoles = True
lbltipref = "": lblauto = ""
DTPicker1.MaxDate = VGParamSistem.FechaTrabajo
DTPicker1.Value = UltimoCierreFech(CDate(Format(VGParamSistem.FechaTrabajo, "dd/MM/yyyy")))

VGRegEnt = 1
If VGRegEnt = 1 Then
    Me.Caption = "Registro de Armado de Kits "
    dato = "I"
    tipo = "NI"
    Codigo2 = "NOTA DE INGRESO"
Else
    Me.Caption = "Registro de Desarmado de Kits"
    dato = "S"
    tipo = "NS"
    Codigo2 = "NOTA DE SALIDA"
End If
MSFlexGrid1.Clear
MSFlexGrid1.Rows = 2
inicializar
inicializaFG

SQL = "select * from tabalm where empresacodigo='" & VGparametros.empresacodigo & "' and puntovtacodigo='" & VGparametros.puntovta & "' "
Set adodc1 = VGCNx.Execute(SQL)
If adodc1.RecordCount = 0 Then
   MsgBox (" No existe almacenes para este Punto de Venta y Codigo de empresa ")
   Exit Sub
End If
   
Call Ctr_ayuAlmacen.conexion(VGCNx)
Ctr_ayuAlmacen.filtro = "empresacodigo='" & VGparametros.empresacodigo & "' and puntovtacodigo='" & VGparametros.puntovta & "' "
Call Ctr_Ayusalida.conexion(VGCNx)
Ctr_Ayusalida.filtro = "tt_tipmov='S' and rtrim(tt_codtrans_auto)<>''"

Command2.Picture = MDIPrincipal.ImageList2.ListImages.item("Insertar").Picture
Command3.Picture = MDIPrincipal.ImageList2.ListImages.item("Sacar").Picture
Cmdgrabar.Picture = MDIPrincipal.ImageList2.ListImages.item("Grabar").Picture
'Command1.Picture = MDIPrincipal.ImageList2.ListImages.item("Cancelar").Picture
Command5.Picture = MDIPrincipal.ImageList2.ListImages.item("Imprimir").Picture
Command7.Picture = MDIPrincipal.ImageList2.ListImages.item("Retornar").Picture

End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
       TxCantidad.Visible = False
End Sub

Private Sub MSFlexGrid1_Click()
Alinear
If (MSFlexGrid1.Col = 3) Then
   TxCantidad.FontName = MSFlexGrid1.CellFontName
   TxCantidad.FontSize = MSFlexGrid1.CellFontSize
   TxCantidad.Visible = True
   TxCantidad = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)
   TxCantidad.SelStart = 0
   TxCantidad.SelLength = Len(TxCantidad)
   TxCantidad.SetFocus
Else
     TxCantidad.Visible = False
 End If

End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
Alinear
If (MSFlexGrid1.Col = 3) Then
        If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 46 Then
            TxCantidad.FontName = MSFlexGrid1.CellFontName
            TxCantidad.FontSize = MSFlexGrid1.CellFontSize
            TxCantidad.Visible = True
            TxCantidad = Chr(KeyAscii)
            TxCantidad.SelStart = 1
            TxCantidad.SetFocus
        End If
 Else
     TxCantidad.Visible = False
 End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'************************** NUM REF
If KeyAscii = 13 And Text1.text <> "" Then
       Tabula (KeyAscii)
      MSFlexGrid1.SetFocus
      'Text8.SetFocus
Else
     Tabula (KeyAscii)
End If
End Sub

Private Sub TxCodKid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then TxCodKid_DblClick
End Sub

Private Sub TxPrecio_Change()
   If Not IsNumeric(TxPrecio) And Trim(TxPrecio) <> "" Then
      MsgBox "La Cantidad Tiene que ser Numerica....", vbInformation, "Verifique .."
      TxPrecio.SetFocus
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
     If VerificarSTOCK = True Then
        Txcant = Format(Txcant, "#####0.00")
        Tabula (KeyAscii)
     End If
  End If
End Sub

Private Sub TxCodKid_DblClick()
VGRegEnt = 1: VGForm1 = 4
FormAyuArtKid.Show 1
If TxCodKid <> "" Then Tabula (13)
inicializaFG

End Sub

Private Sub TxCodKid_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Txcant.SetFocus
End Sub

Private Sub TxPrecio_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        If IsNumeric(TxPrecio) Or Trim(TxPrecio) <> "" Then
           TxPrecio = Format(Val(TxPrecio), "###,##0.00")
           Tabula (KeyAscii)
        End If
     End If
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
       MSFlexGrid1.SetFocus
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
    If vGUtil(1) <> "IK" Then
        MsgBox "Debe ser Transferencia por Elaboración de Kids", vbInformation, "Mensaje"
        TxTransa.SetFocus: Exit Sub
    End If
    TxTransa = vGUtil(1)
    lbltrans = Mid(vGUtil(2), 1, 30)
End If
If TxTransa.text <> "" Then TxTransa_KeyPress (13)
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text4 <> "" Then
   Text4 = Format(Text4, "00000000000")
Else
   Tabula (KeyAscii)
End If
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
      Tabula (KeyAscii)
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
              CmbMoneda.SetFocus
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
        Text4.text = Format(Val(numsal) + 1, "00000000000")
        ContaSalida = Format(Val(nument) + 1, "00000000000")
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
cadena = " select * from stkart where " & criterio & ""
Adodc3.Open cadena, VGCNx, adOpenDynamic, adLockOptimistic

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
                   If PreUni <> 0 And (canttemp + CANTIDAD) <> 0 Then   'Valorizacion
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

Private Sub grabacabecera(Dat As String, Tip As String, num As String, RegEs As Integer, Transf As String)
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
      criterio = " CANUMDOC = '" & num & "' AND CATD = '" & Tip & "' "
      criterio = criterio + " and  CAALMA = '" & VGAlma & "'"
      If Not adodc1.EOF Then adodc1.Filter = criterio
    End If
    adodc1("CATIPMOV") = Dat
    adodc1("CATD") = Tip
    adodc1("CAHORA") = Format(Time, "hh:mm:ss")
    adodc1("CAFECDOC") = DTPicker1.Value            ' CDate(Text2.text)
   
    If Trim(Text1.text) <> "" Then
      adodc1("CARFNDOC") = Trim(Text1.text)
    Else
      adodc1("CARFNDOC") = " "
    End If
    adodc1("CACODMOV") = IIf(Dat = "I", "28", "45")
    adodc1("CANUMDOC") = num
    If Trim(Text6) <> "" Then
      adodc1("CARFTDOC") = Mid$(UCase$(Text6.text), 1, 2)
    Else
      adodc1("CARFTDOC") = " "
    End If
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
    adodc1("empresacodigo") = Empresa
    adodc1("CAUSUARI") = UCase(VGUsuario)
    adodc1("CACODMON") = VGCodMon
    adodc1("CASITGUI") = "V"
    adodc1("cafecact") = Now
    adodc1("CAESTIMP") = "V"
    adodc1("catipotransf") = "TR"
    adodc1("canrotransf") = Transf
   
    adodc1.Update
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
Set rs = VGCNx.Execute(RSQL)
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
'lbltrans = "": TxTransa = ""
Text6 = "": Text8 = "": Text4 = "": Text1 = "": Text9 = ""
lbltipref = "": lblauto = "": TxCantidad = "": TxCambio.Enabled = False: TxCambio = "0"
CmbMoneda.ListIndex = 0
End Sub

Private Sub inicializar()
'TxTransa = ""
TxCodKid = "": lblnomkits = "": Txcant = "": TxPrecio = ""
Text6 = "": Text8 = "": Text4 = "": Text1 = "": Text9 = "": TxCantidad = "": TxCambio = "0": TxCambio.Enabled = False
CmbMoneda.ListIndex = 0
lblnomkits = ""

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
  cadena = TxCodKid 'codigo del art
  RSQL = "select SMCANENT,SMCANSAL FROM MoResMes where  SMALMA = '" & almacen & "' and SMMESPRO = '" & mespro & "' AND  SMCODIGO= '" & cadena & "'"  '
  Set adoreg = New ADODB.Recordset
  adoreg.Open RSQL, VGCNx, adOpenDynamic, adLockOptimistic
  If Not adoreg.EOF Then 'existe
      If entrada Then
          Cantent = adoreg(0) + CANTIDAD
          uSql = "Update MoResMes set SMCANENT = " & Cantent & " where SMALMA='" & almacen & "'  and  SMCODIGO ='" & cadena & "' AND SMMESPRO ='" & mespro & "' "
      Else
          Cantsal = adoreg(1) + CANTIDAD
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
''''''''Solo para lote, arreglar
'''''''Private Sub grabalote(alma As String, codigo As String)
'''''''Dim uSql As String
'''''''Dim lote As String
'''''''Dim nuevo_stk As Double
'''''''Dim rSql As String
'''''''Dim rS As Recordset
'''''''Dim fecfab As Date
'''''''Dim fecven As Date
'''''''    If (MSFlexGrid1.TextMatrix(contador, 9)) <> "__/__/____" Then
'''''''      fecfab = MSFlexGrid1.TextMatrix(contador, 9)
'''''''    End If
'''''''    If (MSFlexGrid1.TextMatrix(contador, 8)) <> "__/__/____" Then
'''''''      fecven = MSFlexGrid1.TextMatrix(contador, 8)
'''''''    End If
'''''''    lote = MSFlexGrid1.TextMatrix(contador, 2)
'''''''    rSql = "select STSLKDIS FROM STKLOTE where  STSALMA= '" & alma & "' and  STSCODIGO= '" & codigo & "' and STSLOTE= '" & lote & "'" '
'''''''    Set rS = db.OpenRecordset(rSql, dbOpenSnapshot)
'''''''    If Not rS.EOF Then
'''''''       If Tipo = "NI" Then
'''''''         nuevo_stk = rS(0) + cantidad
'''''''       Else
'''''''         nuevo_stk = rS(0) - cantidad
'''''''       End If
'''''''
'''''''       uSql = "Update STKLOTE set STSLKDIS = " & nuevo_stk & " WHERE  STSALMA='" & alma & "' and STSCODIGO='" & codigo & "'AND STSLOTE='" & lote & "'"
'''''''    Else
'''''''    If (MSFlexGrid1.TextMatrix(contador, 9)) <> "__/__/____" And (MSFlexGrid1.TextMatrix(contador, 8)) = "__/__/____" Then
'''''''        fecfab = MSFlexGrid1.TextMatrix(contador, 9)
'''''''        uSql = "insert into STKLOTE (STSALMA,STSCODIGO,STSLOTE,STSLKDIS,STSFECFAB) VALUES ('" & alma & "','" & codigo & "','" & lote & "'," & cantidad & ",#" & Format(fecfab, "MM/DD/YYYY") & "#) "
'''''''    ElseIf (MSFlexGrid1.TextMatrix(contador, 9)) = "__/__/____" And (MSFlexGrid1.TextMatrix(contador, 8)) <> "__/__/____" Then
'''''''        fecven = MSFlexGrid1.TextMatrix(contador, 8)
'''''''        uSql = "insert into STKLOTE  (STSALMA,STSCODIGO,STSLOTE,STSLKDIS,STSFECVEN)  VALUES ('" & alma & "','" & codigo & "','" & lote & "' ," & cantidad & " ,#" & Format(fecven, "MM/DD/YYYY") & "#) " 'SIN FECFAB
'''''''    ElseIf (MSFlexGrid1.TextMatrix(contador, 9)) <> "__/__/____" And (MSFlexGrid1.TextMatrix(contador, 8)) <> "__/__/____" Then
'''''''        uSql = "insert into STKLOTE  (STSALMA,STSCODIGO,STSLOTE,STSLKDIS,STSFECFAB,STSFECVEN)  VALUES ('" & alma & "','" & codigo & "','" & lote & "' ," & cantidad & " ,#" & Format(fecfab, "MM/DD/YYYY") & "#,#" & Format(fecven, "MM/DD/YYYY") & "#) "
'''''''    Else
'''''''        uSql = "insert into STKLOTE  (STSALMA,STSCODIGO,STSLOTE,STSLKDIS)  VALUES ('" & alma & "','" & codigo & "','" & lote & "' ," & cantidad & ") "
'''''''    End If
'''''''
'''''''    End If
'''''''    db.Execute uSql
'''''''
'''''''End Sub
''''''''Solo para serie arreglar
'''''''Private Sub grabaserie(alma As String, codigo As String)
'''''''Dim uSql As String
'''''''Dim Serie As String
'''''''Dim VALOR As Integer
'''''''Dim rS As Recordset
'''''''Dim rSql As String
'''''''Dim fecfab As Date
'''''''Dim fecven As Date
'''''''    'fecfab = " " '  MSFlexGrid1.TextMatrix(contador, 8)
'''''''    'fecven = " " 'MSFlexGrid1.TextMatrix(contador, 9)
'''''''    Serie = MSFlexGrid1.TextMatrix(contador, 2)
'''''''    rSql = "select STSSKDIS FROM STKSERI where   STSALMA= '" & VGAlma & "' and  STSCODIGO= '" & codigo & "' and STSSERIE= '" & Serie & "'" '
'''''''    Set rS = db.OpenRecordset(rSql, dbOpenSnapshot)
'''''''    If Not rS.EOF Then
'''''''       VALOR = IIf(Tipo = "NI", 1, 0)
'''''''       uSql = "update STKSERI set STSSKDIS = " & VALOR & " WHERE  STSALMA='" & alma & "' and STSCODIGO='" & codigo & "'AND STSSERIE='" & Serie & "'"
'''''''    Else
'''''''       uSql = "insert into STKSERI (STSALMA,STSCODIGO,STSSERIE,STSSKDIS)   VALUES ('" & alma & "','" & codigo & "','" & Serie & "',1) "
'''''''    End If
'''''''    Vgcnx.Execute uSql
'''''''
'''''''End Sub

Private Sub inicializaFG()
MSFlexGrid1.FormatString = "^ Codigo              |<Descripción                                                                  |>Cant.Reg.   |>Cant.Armado  |>Cant.Dispon."
'MSFlexGrid1.FormatString = "^ Codigo          |<Descripción                                                                       |>Cant.Registrada|>Cantidad Armada|>Cant.Disponible"
End Sub
Function existe_numdoc(text As String) As Boolean
Dim criterio  As String
Dim RSQL As String
Dim rs As Recordset
If adodc1.RecordCount > 1 Then
         MsgBox "El Número del documento ya ha sido registrado: " & Format(text, String(11, "0")) & " !", vbExclamation, "Error"
         existe_numdoc = True
  Else
         existe_numdoc = False
  End If
 
End Function

Function validarautorizado(text As TextBox) As String
  Dim RSQL As String
  Dim rs As Recordset
  Dim codayu As String
  codayu = 12
  RSQL = "Select TCLAVE,TDESCRI from TABAYU  where TCOD= '" & codayu & "' and  Tclave ='" & Trim(text) & "'"
  'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
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
   CrystalReport1.formulas(5) = "Dalma = '" & LblCC.Caption & "' "
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
   CrystalReport1.formulas(7) = "Dalmades = '" & LblCC.Caption & "' "
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
Adodc22("DETD") = Tip
Adodc22("DENUMDOC") = NumDoc  ' ContaSalida
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
SQL = " UPDATE STKART SET stskdis = stskdis - " & cant & " where stalma='" & VGAlma & "'"
SQL = SQL & " and stcodigo='" & Art & "'"
VGCNx.Execute (SQL)
Adodc22.Update
Adodc22.Filter = ""
Adodc22.Requery
End Sub

Sub Alinear()
TxCantidad.Width = MSFlexGrid1.CellWidth
TxCantidad.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
TxCantidad.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top - 30
TxCantidad.Height = MSFlexGrid1.CellHeight - 90
End Sub

Private Sub TxCantidad_KeyPress(KeyAscii As Integer)
Dim auxCant As Double
If NumPto(KeyAscii) Then
    Select Case KeyAscii
      Case Is = 13
         If Val(TxCantidad) < MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) Then
            MsgBox "La Cantidad Ingresada No puede Ser Menor al que Registro en el Kit...!", vbInformation, "Verifique el Ingreso"
            TxCantidad.SetFocus
            Exit Sub
         End If
         MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ClsTock.SaldoArti(VGAlma, MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0), VGCNx)
         If Val(TxCantidad) <= MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) Then
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = TxCantidad.text
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4)) - Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3))
            auxCant = IIf(Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)) <> 0, Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)) / Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)), 0)
            TxCantidad.Visible = False
            '*************************************
            Call VerificaIngresos
            If MSFlexGrid1.Row < MSFlexGrid1.Rows - 1 Then
               MSFlexGrid1.Row = MSFlexGrid1.Row + 1
               Alinear
               TxCantidad.Visible = True
               TxCantidad.text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)
               TxCantidad.SetFocus
               TxCantidad.SelStart = 0
               TxCantidad.SelLength = Len(TxCantidad)
            Else
               MSFlexGrid1.SetFocus
            End If
            '*************************************
         Else
            MsgBox "NO Tiene Stock Disponible para ese Componente...!", vbInformation, "Verifique el Ingreso"
            TxCantidad.SetFocus
         End If
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
For n = 1 To MSFlexGrid1.Rows - 1
    MSFlexGrid1.TextMatrix(n, 3) = Round(Val(MSFlexGrid1.TextMatrix(n, 2)) * Val(Txcant), 6)
    MSFlexGrid1.TextMatrix(n, 4) = Val(MSFlexGrid1.TextMatrix(n, 4)) - Val(MSFlexGrid1.TextMatrix(n, 3))
    If Round(Val(MSFlexGrid1.TextMatrix(n, 4)), 4) < 0 Then
       Negativos = False
       Exit For
    End If
Next

If Negativos = False Then
   MsgBox "No Dispone de Componentes para esta Cantidad, fila --> " & Str(n) & "", vbInformation, "Verifique ....."
   LoadReceta (TxCodKid)
   VerificarSTOCK = False
End If
End Function
Sub LoadReceta(ByVal arCodKit As String)
   Dim rs As New ADODB.Recordset
   Dim SQL As String
   'SQL = "SELECT STKART.STALMA, KITS.CODART, MAEART.ADESCRI, KITS.CANART, 0 AS Expr1, STKART.STSKDIS FROM (KITS INNER JOIN STKART ON KITS.CODKIT = STKART.STCODIGO) LEFT JOIN MAEART ON KITS.CODART =MAEART.ACODIGO  where STALMA='" & VGAlma & "' AND  KITS.CODkit='" & arCodKit & "'"
   SQL = "SELECT KITS.CODART, MAEART.ADESCRI, KITS.CANART, 0 AS Expr1, STKART.STSKDIS FROM (KITS INNER JOIN STKART ON KITS.CODart = STKART.STCODIGO) LEFT JOIN MAEART ON KITS.CODART =MAEART.ACODIGO where STALMA='" & Txtalma.text & "' AND  KITS.CODkit='" & arCodKit & "'"
   rs.Open SQL, VGCNx, adOpenForwardOnly, adLockReadOnly
   If Not rs.EOF Then
      'Set MSFlexGrid1.DataSource = rS
      MSFlexGrid1.Rows = 2
      MSFlexGrid1.AddItem rs!codart & Chr(9) & rs!ADESCRI & Chr(9) & rs!CANART & Chr(9) & 0 & Chr(9) & ClsTock.SaldoArti(Txtalma.text, rs!codart, VGCNx), 1
      rs.MoveNext
      Do While Not rs.EOF
         MSFlexGrid1.AddItem rs!codart & Chr(9) & rs!ADESCRI & Chr(9) & rs!CANART & Chr(9) & 0 & Chr(9) & ClsTock.SaldoArti(Txtalma.text, rs!codart, VGCNx), 1
         rs.MoveNext
      Loop
      MSFlexGrid1.Rows = MSFlexGrid1.Rows - 1
   End If
   
   MSFlexGrid1.FormatString = "^ Codigo                 |<Descripción                                                                      |>Cant.Reg.  |>Cant.Armado |>Cant.Dispon."
   rs.Close
End Sub

Function VerificaIngresos() As Boolean
Dim n, fila As Long
Dim nFactor As Double
VerificaIngresos = True

For n = 1 To MSFlexGrid1.Rows - 1
    If Val(MSFlexGrid1.TextMatrix(n, 3)) <= 0 Then
        VerificaIngresos = False
        MsgBox "Las Cantidades  no Corresponde al Nro. Calculado para el Armado", vbInformation, "Verifique ....."
        Exit Function
    End If
Next

If Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)) <> 0 Then
   nFactor = Round(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) / Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)), 4)
End If


For n = 1 To MSFlexGrid1.Rows - 1

    If Round(nFactor, 4) <> Round(MSFlexGrid1.TextMatrix(n, 3) / Val(MSFlexGrid1.TextMatrix(n, 2)), 4) Then
       MsgBox "Una de Las Cantidades  no Corresponde al Nro. Calculado para el Armado", vbInformation, "Verifique ....."
       VerificaIngresos = False
       Exit Function
    Else
       Txcant = nFactor
    End If
Next

End Function
