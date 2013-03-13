VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form RptDocumentosxCobrar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Documentos por Cobrar"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   Icon            =   "RptDocumentosxCobrar.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Caption         =   "Seleccionar Rango"
      Height          =   780
      Left            =   180
      TabIndex        =   21
      Top             =   3360
      Width           =   5880
      Begin VB.TextBox txtRangoDias 
         Height          =   315
         Left            =   2265
         TabIndex        =   22
         Text            =   "7*15*30*45*60*"
         Top             =   285
         Width           =   2625
      End
      Begin VB.Label Label5 
         Caption         =   "Cada Nº dias"
         Height          =   270
         Left            =   480
         TabIndex        =   23
         Top             =   315
         Width           =   1605
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3068
      TabIndex        =   17
      Top             =   4350
      Width           =   1260
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1493
      TabIndex        =   16
      Top             =   4350
      Width           =   1260
   End
   Begin VB.Frame Frame4 
      Height          =   3195
      Left            =   180
      TabIndex        =   3
      Top             =   120
      Width           =   5880
      Begin VB.Frame Frame3 
         Height          =   1125
         Left            =   165
         TabIndex        =   19
         Top             =   1035
         Width           =   5490
         Begin VB.ComboBox cboMoneda 
            Height          =   315
            Left            =   795
            TabIndex        =   0
            Text            =   "cboMoneda"
            Top             =   720
            Width           =   1515
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Cliente 
            Height          =   315
            Left            =   795
            TabIndex        =   20
            Top             =   315
            Width           =   4605
            _ExtentX        =   8123
            _ExtentY        =   556
            XcodMaxLongitud =   0
            xcodwith        =   900
            NomTabla        =   "vt_cliente"
            ListaCampos     =   "clientecodigo(1),clienterazonsocial(1)"
            XcodCampo       =   "clientecodigo"
            XListCampo      =   "clienterazonsocial"
            ListaCamposDescrip=   "Código,Razón_Social"
            ListaCamposText =   "clientecodigo,clienterazonsocial"
         End
         Begin VB.Label Label4 
            Caption         =   "Moneda"
            Height          =   225
            Left            =   135
            TabIndex        =   1
            Top             =   780
            Width           =   645
         End
         Begin VB.Label Label3 
            Caption         =   "Cliente"
            Height          =   225
            Left            =   135
            TabIndex        =   2
            Top             =   360
            Width           =   525
         End
      End
      Begin VB.ComboBox cboHojaResumen 
         Height          =   315
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2625
         Width           =   1425
      End
      Begin VB.ComboBox cboResumen 
         Height          =   315
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2235
         Width           =   1425
      End
      Begin VB.Frame Frame1 
         Height          =   705
         Left            =   150
         TabIndex        =   4
         Top             =   180
         Width           =   3435
         Begin MSComCtl2.DTPicker DTP_FechaRef 
            Height          =   345
            Left            =   1350
            TabIndex        =   18
            Top             =   210
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   609
            _Version        =   393216
            Format          =   117768193
            CurrentDate     =   37588
         End
         Begin VB.Label Label1 
            Caption         =   "Hasta Fecha"
            Height          =   255
            Index           =   1
            Left            =   150
            TabIndex        =   5
            Top             =   270
            Width           =   990
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1335
         Left            =   150
         TabIndex        =   6
         Top             =   2955
         Visible         =   0   'False
         Width           =   5145
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda1 
            Height          =   315
            Left            =   2160
            TabIndex        =   10
            Top             =   540
            Width           =   2925
            _ExtentX        =   5159
            _ExtentY        =   556
            Enabled         =   0   'False
            XcodMaxLongitud =   0
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Relacion x Banco"
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   9
            Top             =   870
            Width           =   1935
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Todos Movimientos"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   8
            Top             =   270
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Relacion x Vendedor"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   7
            Top             =   570
            Width           =   1935
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda2 
            Height          =   315
            Left            =   2160
            TabIndex        =   11
            Top             =   900
            Width           =   2925
            _ExtentX        =   5159
            _ExtentY        =   556
            Enabled         =   0   'False
            XcodMaxLongitud =   0
         End
      End
      Begin VB.Label Label2 
         Caption         =   "Hoja Resumen"
         Height          =   165
         Index           =   1
         Left            =   270
         TabIndex        =   13
         Top             =   2685
         Width           =   1125
      End
      Begin VB.Label Label2 
         Caption         =   "Resumen"
         Height          =   165
         Index           =   0
         Left            =   270
         TabIndex        =   12
         Top             =   2295
         Width           =   1125
      End
   End
End
Attribute VB_Name = "RptDocumentosxCobrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adll As New dllgeneral.dll_general
Dim aRango(5) As Integer

Private Sub Form_Load()
   MostrarForm Me, "C2"
   Ctr_Cliente.conexion cn
   Call CargarTipo(cboResumen, 3)
   cboResumen.ListIndex = 1
   
   Call CargarTipo(cboHojaResumen, 3)
   cboMoneda.Clear
   cboMoneda.AddItem g_TipoSol & "-Soles"
   cboMoneda.AddItem g_TipoDolar & "-Dolares"
   'cboMoneda.AddItem "03-Ambos"
   cboMoneda.ListIndex = 0
   DTP_FechaRef.Value = Format(Date, "dd/mm/yyyy")
End Sub

Private Sub cmdAceptar_Click()
   If RangoDias = True Then
     Call Imprimir
   End If
End Sub

Private Sub cmdCancelar_Click()
 Unload Me
End Sub

Sub Imprimir()
'FIXIT: Declare 'arrform' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
Dim arrform(11) As Variant, arrparm(7) As Variant
Dim NombreRep As String, CadOrden As String
Dim NombrePC As String
Dim ValorRango As String
'Dim mon As String
Dim i As Integer
    ValorRango = txtRangoDias.Text
    Randomize   ' Inicializa el generador de números aleatorios.
'FIXIT: Reemplazar la función 'Str' con la función 'Str$'.                                 FixIT90210ae-R9757-R1B8ZE
    NombrePC = RTrim$(Str(CLng(Rnd * 10000000)))
    arrparm(0) = VGCNx.DefaultDatabase
    arrparm(1) = NombrePC
    arrparm(2) = Format(DTP_FechaRef.Value, "dd/mm/yyyy")
    arrparm(3) = ValorRango
    arrparm(4) = ValorRango
    arrparm(5) = IIf(Ctr_Cliente.xclave = Empty, "%", RTrim$(Ctr_Cliente.xclave))
    If cboMoneda.ListIndex = 2 Then
      arrparm(6) = "%"
    Else
      arrparm(6) = Format(cboMoneda.ListIndex + 1, "00")
    End If
    If cboResumen.ListIndex = 0 Then
       NombreRep = "RepccDocumentosPendientesResumen.rpt"
    Else
       NombreRep = "RepccDocumentosPendientes.rpt"
    End If
    CadOrden = ""
    For i = 1 To 5
'FIXIT: Reemplazar la función 'Str' con la función 'Str$'.                                 FixIT90210ae-R9757-R1B8ZE
'FIXIT: Reemplazar la función 'Str' con la función 'Str$'.                                 FixIT90210ae-R9757-R1B8ZE
      arrform(i - 1) = "@Docve" & RTrim$(Str(i)) & "='" & RTrim$(Str(aRango(i - 1))) & "'"
    Next
    For i = 1 To 5
'FIXIT: Reemplazar la función 'Str' con la función 'Str$'.                                 FixIT90210ae-R9757-R1B8ZE
'FIXIT: Reemplazar la función 'Str' con la función 'Str$'.                                 FixIT90210ae-R9757-R1B8ZE
      arrform(4 + i) = "@Docpv" & RTrim$(Str(i)) & "='" & RTrim$(Str(aRango(i - 1))) & "'"
    Next
    arrform(10) = "@Fecha='" & Format(DTP_FechaRef.Value, "dd/mm/yyyy") & "'"
    
    Call ImpresionRptProc(NombreRep, arrform, arrparm, CadOrden, "Documentos Pendientes")
End Sub

Function RangoDias() As Boolean
 Dim pos As Integer
 Dim cadtexto As String
 Dim i As Integer
 If txtRangoDias.Text = Empty Then
    MsgBox "Debe registrar los Valores para el Rango", vbInformation, Caption
    RangoDias = False
    Exit Function
 End If
 
 If Right$(RTrim$(txtRangoDias.Text), 1) <> "*" Then
   MsgBox "El Texto para el Rango debe terminar con un *", vbInformation, Caption
   txtRangoDias.Text = RTrim$(txtRangoDias.Text) & RTrim$("*")
 End If
 
 cadtexto = RTrim$(txtRangoDias.Text)
 For i = 1 To 5
   pos = InStr(1, cadtexto, "*", vbTextCompare)
   If Not IsNumeric(Left$(cadtexto, pos - 1)) Then
      MsgBox "El valor " & Left$(cadtexto, pos - 1) & " No es numérico", vbInformation, Caption
      RangoDias = False
      Exit Function
   End If
   aRango(i - 1) = Left$(cadtexto, pos - 1)
   cadtexto = Right$(cadtexto, Len(cadtexto) - pos)
 Next

RangoDias = True
End Function
