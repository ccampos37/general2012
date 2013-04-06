VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmRepVtasxForma 
   Caption         =   "Ventas por Forma de Pago"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   6285
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo0 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   4680
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   1305
   End
   Begin Crystal.CrystalReport oCrystalReport 
      Left            =   240
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   4440
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   1560
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   3120
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   720
      Width           =   2895
   End
   Begin MSComCtl2.DTPicker DTHasta 
      Height          =   345
      Left            =   2640
      TabIndex        =   3
      Top             =   1680
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   108986369
      CurrentDate     =   37518
   End
   Begin MSComCtl2.DTPicker DTDesde 
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   108986369
      CurrentDate     =   37518
   End
   Begin VB.Label lbl 
      Caption         =   "Punto de Venta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   720
      TabIndex        =   11
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lbl 
      Caption         =   "Hasta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   1680
      TabIndex        =   8
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label lbl 
      Caption         =   "Desde"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   7
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lbl 
      Caption         =   "Forma de Pago"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   6
      Top             =   720
      Width           =   1935
   End
End
Attribute VB_Name = "FrmRepVtasxForma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adll As New dllgeneral.dll_general
''''''''''''''''''''''''''''''''''''''''''''''''''''
'Agregar:
Dim busca As New dll_apisgen.dll_apis
''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub cmdAceptar_Click(Index As Integer)
Dim aparam(5) As Variant
Dim aform(5) As Variant

 If DTDesde > DTHasta Then
     MsgBox "Fecha Desde debe ser mayor a Fecha Hasta", vbInformation, "AVISO"
     Exit Sub
  End If
        
    aform(0) = "Empresa='" & VGParametros.nomempresa & "'"
    aform(1) = "Desde='" & DTDesde & "'"
    aform(2) = "Hasta='" & DTHasta & "'"
        If Combo0.ListIndex <> -1 Then
            aform(3) = "Puntoventa='" & Combo0.Text & "'"
        Else
            aform(3) = "Puntoventa='TODOS'"
        End If
        If Combo1.ListIndex <> -1 Then
            aform(4) = "Formapago='" & Combo1.Text & "'"
        Else
            aform(4) = "Formapago='TODOS'"
        End If
        aparam(0) = VGCNx.DefaultDatabase
        aparam(1) = IIf(Trim(txt(0)) = "", "%%", Trim(txt(0)))
        aparam(2) = IIf(Trim(txt(1)) = "", "%%", Trim(txt(1)))
        aparam(3) = DTDesde
        aparam(4) = DTHasta
  
  Call ImpresionRpt_SubRpt_Proc("vt_VtasxFormaPago.rpt", aform, aparam, "vt_VtasxFormaPago_sub.rpt", 0, " Forma de Pagos")

End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Combo0_Click()
  If Combo0.ListCount > 0 Then
     txt(0) = adll.ComboDato(Combo0.Text)
  Else
     txt(0) = ""
  End If

End Sub

Private Sub Combo1_Click()
  If Combo1.ListIndex <> -1 Then
    txt(1) = adll.ComboDato(Combo1.Text)
  Else
    txt(1) = ""
  End If
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
       If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub Combo0_KeyDown(KeyCode As Integer, Shift As Integer)
       If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub DtDesde_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub DtHasta_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    MostrarFormVentas Me, "C2"
    Call adll.llenacombo(Combo0, "select puntovtacodigo,puntovtadescripcion from vt_puntoventa", VGCNx)
    Call adll.llenacombo(Combo1, "select formapagocodigo,formapagodescripcion from vt_formapago", VGCNx)
    DTDesde = Date
    DTHasta = Date
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub


