VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Begin VB.Form FrmRepDocvenciXvence 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Documentos Vencidos x Vencer"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   420
      Left            =   1770
      TabIndex        =   4
      Top             =   3015
      Width           =   1380
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   420
      Left            =   3195
      TabIndex        =   3
      Top             =   3015
      Width           =   1380
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   1125
      TabIndex        =   0
      Top             =   1950
      Width           =   4590
      Begin VB.OptionButton Optres 
         Caption         =   "Vencidos"
         Height          =   195
         Index           =   0
         Left            =   420
         TabIndex        =   2
         Top             =   300
         Width           =   1350
      End
      Begin VB.OptionButton Optres 
         Caption         =   "Por Vencer"
         Height          =   195
         Index           =   1
         Left            =   2565
         TabIndex        =   1
         Top             =   315
         Width           =   1350
      End
   End
   Begin TextFer.TxFer TxtRango 
      Height          =   330
      Left            =   1155
      TabIndex        =   5
      Top             =   1065
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   582
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
      MaxLength       =   100
      Locked          =   -1  'True
      Text            =   ""
      Valor           =   ""
      NoCaracteres    =   "0123456789,"
      NoRangoCadena   =   -1  'True
   End
   Begin MSComCtl2.DTPicker DTPFecha 
      Height          =   330
      Left            =   1185
      TabIndex        =   6
      Top             =   150
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Format          =   59703297
      CurrentDate     =   37697
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda2 
      Height          =   300
      Left            =   1185
      TabIndex        =   7
      Top             =   615
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   529
      XcodMaxLongitud =   0
      xcodwith        =   900
      NomTabla        =   "vt_cliente"
      ListaCampos     =   "clientecodigo(1),clienterazonsocial(1)"
      XcodCampo       =   "clientecodigo"
      XListCampo      =   "clienterazonsocial"
      ListaCamposDescrip=   "Código,Razón_Social"
      ListaCamposText =   "clientecodigo,clienterazonsocial"
      Requerido       =   0   'False
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_Doc 
      Height          =   300
      Left            =   1155
      TabIndex        =   12
      Top             =   1530
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   529
      XcodMaxLongitud =   0
      xcodwith        =   900
      NomTabla        =   "cc_tipodocumento"
      ListaCampos     =   "tdocumentocodigo(1),tdocumentodescripcion(1)"
      XcodCampo       =   "tdocumentocodigo"
      XListCampo      =   "tdocumentodescripcion"
      ListaCamposDescrip=   "Código,Descripción"
      ListaCamposText =   "tdocumentocodigo,tdocumentodescripcion"
      Requerido       =   0   'False
   End
   Begin VB.Label Label3 
      Caption         =   "Documento"
      Height          =   285
      Left            =   165
      TabIndex        =   11
      Top             =   1590
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente :"
      Height          =   255
      Index           =   1
      Left            =   150
      TabIndex        =   10
      Top             =   645
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha :"
      Height          =   285
      Index           =   0
      Left            =   150
      TabIndex        =   9
      Top             =   180
      Width           =   945
   End
   Begin VB.Label Label2 
      Caption         =   "Rango"
      Height          =   285
      Left            =   165
      TabIndex        =   8
      Top             =   1110
      Width           =   945
   End
End
Attribute VB_Name = "FrmRepDocvenciXvence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim op As Integer
Dim criterio As String
Dim Rsaux As ADODB.Recordset
Dim Rango As String
Private Sub cmdAceptar_Click()
'FIXIT: Declare 'Aparam' con un tipo de datos de enlace en tiempo de compilación           FixIT90210ae-R1672-R1B8ZE
Dim Aparam(6) As Variant, Aformu(2) As Variant
Dim vgdll As New dllgeneral.dll_general
'@Base,@Op,@Fecharef,@Cliente,@rango
    Aparam(0) = VGCNx.DefaultDatabase
    Aparam(1) = op
    Aparam(2) = FechS(DTPFecha.Value, Sqlf)
    Aparam(3) = IIf(RTrim$(Ctr_Ayuda2.xclave) = "", "%%", RTrim$(Ctr_Ayuda2.xclave))
    Aparam(4) = RTrim$(TxtRango.Text)
    Aparam(5) = IIf(RTrim$(Ctr_Doc.xclave) = Empty, "%%", RTrim$(Ctr_Doc.xclave))
    Aformu(0) = "crit='" & criterio & "'"
    Aformu(1) = "@documento='" & IIf(RTrim$(Ctr_Doc.xclave) = Empty, "Todos", RTrim$(Ctr_Doc.xclave) & " - " & RTrim$(Ctr_Doc.xnombre)) & "'"
    Call ImpresionRptProc(RutaRepProc & "RepccDocvenciXvence.rpt", Aformu, Aparam)
End Sub
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    Rango = Empty
    Call Ctr_Ayuda2.conexion(VGCNx)
    Call Ctr_Doc.conexion(VGCNx)
    DTPFecha.Value = Date
    Optres(0).Value = True
    Set Rsaux = New ADODB.Recordset
    Rsaux.Open "cc_rangovcto", VGCNx, adOpenKeyset, adLockReadOnly
    Rsaux.MoveFirst
    Do While Not Rsaux.EOF
        Rango = Rango & RTrim$(Rsaux!Cod) & ","
        Rsaux.MoveNext
    Loop
    Rango = Rango & "9999999,"
    TxtRango.Text = Rango
End Sub
Private Sub Optres_Click(Index As Integer)
    op = Index + 1
'FIXIT: Reemplazar la función 'UCase' con la función 'UCase$'.                             FixIT90210ae-R9757-R1B8ZE
    criterio = UCase$(Optres(Index).Caption)
End Sub

