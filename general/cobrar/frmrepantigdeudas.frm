VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Begin VB.Form frmrepantigdeudas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumen de Deudas por Cliente"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4680
      Left            =   105
      TabIndex        =   2
      Top             =   45
      Width           =   6510
      Begin VB.Frame Frame2 
         Caption         =   "Filtro"
         Height          =   1425
         Left            =   540
         TabIndex        =   7
         Top             =   2970
         Width           =   5370
         Begin VB.OptionButton Opt 
            Caption         =   "Vencidos"
            Height          =   360
            Index           =   0
            Left            =   150
            TabIndex        =   11
            Top             =   315
            Width           =   2955
         End
         Begin VB.OptionButton Opt 
            Caption         =   "Por Vencer"
            Height          =   360
            Index           =   1
            Left            =   150
            TabIndex        =   10
            Top             =   795
            Width           =   1305
         End
         Begin VB.ComboBox cmbsimb 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmrepantigdeudas.frx":0000
            Left            =   1590
            List            =   "frmrepantigdeudas.frx":0016
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   810
            Width           =   990
         End
         Begin TextFer.TxFer txdias 
            Height          =   360
            Left            =   3240
            TabIndex        =   8
            Top             =   795
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   635
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
            Enabled         =   0   'False
            ColorIlumina    =   14024183
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            NoCaracteres    =   "0123456789"
            MarcarTextoAlEnfoque=   -1  'True
            NoRangoCadena   =   -1  'True
         End
         Begin VB.Label Label2 
            Caption         =   "D�as"
            Height          =   270
            Left            =   2730
            TabIndex        =   12
            Top             =   870
            Width           =   525
         End
      End
      Begin VB.OptionButton Optres 
         Caption         =   "Resumido"
         Height          =   390
         Index           =   0
         Left            =   1875
         TabIndex        =   5
         Top             =   2430
         Width           =   1185
      End
      Begin VB.OptionButton Optres 
         Caption         =   "Normal"
         Height          =   390
         Index           =   1
         Left            =   3570
         TabIndex        =   4
         Top             =   2445
         Width           =   1185
      End
      Begin VB.ComboBox cmbmon 
         Height          =   315
         ItemData        =   "frmrepantigdeudas.frx":002F
         Left            =   1215
         List            =   "frmrepantigdeudas.frx":0039
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1515
         Width           =   2250
      End
      Begin TextFer.TxFer TxTc 
         Height          =   300
         Left            =   1185
         TabIndex        =   6
         Top             =   1905
         Width           =   1245
         _ExtentX        =   2196
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
         Text            =   "0"
         TipoDato        =   1
         NumeroDecimales =   3
         SignoNegativo   =   0   'False
      End
      Begin MSComCtl2.DTPicker DTPFecha 
         Height          =   330
         Left            =   1215
         TabIndex        =   13
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         Format          =   59703297
         CurrentDate     =   37697
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda2 
         Height          =   300
         Left            =   1215
         TabIndex        =   14
         Top             =   705
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   529
         XcodMaxLongitud =   0
         xcodwith        =   900
         NomTabla        =   "vt_cliente"
         ListaCampos     =   "clientecodigo(1),clienterazonsocial(1)"
         XcodCampo       =   "clientecodigo"
         XListCampo      =   "clienterazonsocial"
         ListaCamposDescrip=   "C�digo,Raz�n_Social"
         ListaCamposText =   "clientecodigo,clienterazonsocial"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Doc 
         Height          =   300
         Left            =   1215
         TabIndex        =   15
         Top             =   1125
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   529
         XcodMaxLongitud =   0
         xcodwith        =   900
         NomTabla        =   "cc_tipodocumento"
         ListaCampos     =   "tdocumentocodigo(1),tdocumentodescripcion(1)"
         XcodCampo       =   "tdocumentocodigo"
         XListCampo      =   "tdocumentodescripcion"
         ListaCamposDescrip=   "C�digo,Descripci�n"
         ListaCamposText =   "tdocumentocodigo,tdocumentodescripcion"
         Requerido       =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha :"
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   20
         Top             =   270
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "Documento :"
         Height          =   255
         Index           =   6
         Left            =   180
         TabIndex        =   19
         Top             =   1185
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente :"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   18
         Top             =   735
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "Mon :"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   17
         Top             =   1575
         Width           =   945
      End
      Begin VB.Label Label3 
         Caption         =   "T/C :"
         Height          =   285
         Left            =   180
         TabIndex        =   16
         Top             =   2025
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   2070
      TabIndex        =   1
      Top             =   4890
      Width           =   1245
   End
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   3330
      TabIndex        =   0
      Top             =   4890
      Width           =   1245
   End
End
Attribute VB_Name = "frmrepantigdeudas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim op As String
Dim NomReporte As String
Dim criterio As String

Private Sub cmdAceptar_Click()
'FIXIT: Declare 'Aparam' con un tipo de datos de enlace en tiempo de compilaci�n           FixIT90210ae-R1672-R1B8ZE
Dim Aparam(8) As Variant, Aformu(3) As Variant
Dim vgdll As New dllgeneral.dll_general
'@Base, @BaseConta, @op, @cliente, @tipdoc, @dias, @simbo
    '@Base,@BaseConta ,@op,@cliente varchar(20),@tipdoc,@dias,@simbo,@fecharef
    Aparam(0) = VGCNx.DefaultDatabase
    Aparam(1) = VGcnxCT.DefaultDatabase
    Aparam(2) = op
    Aparam(3) = IIf(RTrim$(Ctr_Ayuda2.xclave) = "", "%%", RTrim$(Ctr_Ayuda2.xclave))
    Aparam(4) = IIf(RTrim$(Ctr_Doc.xclave) = "", "%%", RTrim$(Ctr_Doc.xclave))
    Aparam(5) = IIf(RTrim$(txdias.Text) = "", " ", RTrim$(txdias.Text))
    Aparam(6) = IIf(RTrim$(cmbsimb.Text) = "", " ", RTrim$(cmbsimb.Text))
    Aparam(7) = FechS(DTPFecha.Value, Sqlf)
    Aformu(0) = "tc=" & IIf(vgdll.ESNULO(TxTc.Text, 0) = 0, 1, TxTc.Text)
'FIXIT: Reemplazar la funci�n 'UCase' con la funci�n 'UCase$'.                             FixIT90210ae-R9757-R1B8ZE
    Aformu(1) = "mon='" & UCase$(RTrim$(Right$(cmbmon.Text, Len(cmbmon.Text) - 4))) & "'"
    Aformu(2) = "crit='" & criterio & "'"
    Call ImpresionRptProc(RutaRepProc & NomReporte, Aformu, Aparam)
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call Ctr_Ayuda2.conexion(VGCNx)
    Call Ctr_Doc.conexion(VGCNx)
    cmbmon.ListIndex = 0
    DTPFecha.Value = Date
    Opt(0).Value = True
    Optres(0).Value = True
End Sub

Private Sub Opt_Click(Index As Integer)
    op = Format(Index + 1, "0")
    Select Case Index
        Case 0
            cmbsimb.Enabled = False: txdias.Enabled = False
            criterio = " Vencidos "
        Case 1
            cmbsimb.Enabled = True: txdias.Enabled = True
            cmbsimb.ListIndex = 0
            criterio = " Por Vencer " & cmbsimb.Text & " " & txdias.Text
    End Select
End Sub

Private Sub Optres_Click(Index As Integer)
    Select Case Index
        Case 0: NomReporte = "RepccAntiguedDeudas.rpt"
        Case 1: NomReporte = "RepccAntiguedDeudasTD.rpt"
    End Select
End Sub
