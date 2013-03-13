VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form RepDocuDeta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de Documentox Detallados"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7425
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   7425
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   975
      Left            =   5280
      TabIndex        =   13
      Top             =   1800
      Width           =   1935
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancelar"
         Height          =   645
         Left            =   960
         Picture         =   "RepDocuDeta.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   795
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
         Height          =   645
         Left            =   120
         Picture         =   "RepDocuDeta.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Ordenado Por"
      Height          =   1575
      Left            =   5280
      TabIndex        =   12
      Top             =   120
      Width           =   1935
      Begin VB.OptionButton OptProducto 
         Caption         =   "Producto"
         Height          =   375
         Left            =   360
         TabIndex        =   17
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Optdocumento 
         Caption         =   "Documento"
         Height          =   495
         Left            =   360
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   150
      Top             =   3030
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      Caption         =   "Rango de fechas"
      Height          =   1035
      Left            =   270
      TabIndex        =   5
      Top             =   1770
      Width           =   4845
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   810
         TabIndex        =   6
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   98762753
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   35796
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   3210
         TabIndex        =   7
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   98762753
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   35796
      End
      Begin VB.Label Label5 
         Caption         =   "Desde"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   390
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Left            =   2610
         TabIndex        =   8
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1665
      Left            =   270
      TabIndex        =   0
      Top             =   90
      Width           =   4845
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1050
         Width           =   2385
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   270
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1050
         Width           =   1725
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   270
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   420
         Width           =   4395
      End
      Begin VB.Label Label2 
         Caption         =   "Transaccion"
         Height          =   225
         Index           =   1
         Left            =   2250
         TabIndex        =   10
         Top             =   840
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Movimiento"
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   4
         Top             =   840
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Almacen Origen"
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   3
         Top             =   210
         Width           =   1185
      End
   End
End
Attribute VB_Name = "RepDocuDeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
  Dim rsc As New ADODB.Recordset
  
  Combo4.Clear
  Combo4.AddItem "T-Todos"
  If Left(Combo1.text, 1) Like "[IS]" Then
        Set rsc = VGCNx.Execute("select TT_CODMOV,TT_DESCRI from tabtransa WHERE tt_tipmov='" & Left(Combo1.text, 1) & "'")
  Else
        Set rsc = VGCNx.Execute("select TT_CODMOV,TT_DESCRI from tabtransa ")
  End If
  If rsc.RecordCount > 0 Then
        rsc.MoveFirst
        Do Until rsc.EOF
          Combo4.AddItem rsc.Fields(0) & "-" & IIf(IsNull(rsc.Fields(1)), "****", rsc.Fields(1))
          rsc.MoveNext
        Loop
  End If
  rsc.Close
  
  Set rsc = Nothing
  
  Combo4.ListIndex = 0

End Sub

Private Sub Command1_Click()
    Dim VGDllGeneral As New dllgeneral.dll_general
    Dim aparam(7) As Variant
    Dim aform(6) As Variant
    Dim Reporte As String
    Reporte = "al_documentosdetallado.rpt"
    If Optdocumento.Value = True Then Reporte = "al_documentosdetalladoNumero.rpt"
    aparam(0) = Trim(VGCNx.DefaultDatabase)
    aparam(1) = Left(Combo2.text, 2)
    aparam(2) = Left(Combo1.text, 1)
    aparam(3) = Format(DTPicker1.Value, "dd/mm/yyyy")
    aparam(4) = Format(DTPicker2.Value, "dd/mm/yyyy")
    aparam(5) = "%%"
    aparam(6) = IIf(Trim(VGDllGeneral.ComboDato(Combo4.text)) = "T", "%", VGDllGeneral.ComboDato(Combo4.text))  'transaccion
    
    aform(0) = "almacen ='" & Trim(Combo2) & "  *** Ordenado por : "
    If Optdocumento.Value = True Then
       aform(0) = aform(0) & " Documento  *** '"
     Else
       aform(0) = aform(0) & " Articulo *** '"
    End If
    aform(1) = "fechainicio ='" & DTPicker1 & "'"
    aform(2) = "fechafin ='" & DTPicker2 & "'"
    Select Case Left(Combo1.text, 1)
        Case "I"
            aform(3) = "tipo ='" & "**INGRESOS**" & "'"
        Case "S"
            aform(3) = "tipo ='" & "**SALIDAS**" & "'"
        Case "A"
            aform(3) = "tipo ='" & "**ANULADOS**" & "'"
        Case "T"
            aform(3) = "tipo ='" & "**TODOS**" & "'"
    End Select
    aform(4) = "destino =' Todos'"
    aform(5) = "transa ='" & Trim(Combo4.text) & "'"
    Call ImpresionRptProc(Reporte, aform, aparam, , Reporte + " - Impresion de documentos detallado")
    Screen.MousePointer = 1
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
  Dim rsc As New ADODB.Recordset
  
  Set rs = VGCNx.Execute("select TAALMA,TADESCRI,'','' from tabalm where taalma='*'")
  
  Combo2.Clear
  Set rsc = VGCNx.Execute("select TAALMA,TADESCRI from tabalm ")
  If rsc.RecordCount > 0 Then
      rsc.MoveFirst
      Do Until rsc.EOF
        Combo2.AddItem rsc.Fields(0) & "-" & IIf(IsNull(rsc.Fields(1)), "****", rsc.Fields(1))
        rsc.MoveNext
      Loop
  End If
  rsc.Close
  Set rsc = Nothing
 
 
  Combo4.Clear
  Combo4.AddItem "T-Todos"
  Set rsc = VGCNx.Execute("select TT_CODMOV,TT_DESCRI from tabtransa")
  If rsc.RecordCount > 0 Then
      rsc.MoveFirst
      Do Until rsc.EOF
        Combo4.AddItem rsc.Fields(0) & "-" & IIf(IsNull(rsc.Fields(1)), "****", rsc.Fields(1))
        rsc.MoveNext
      Loop
  End If
  rsc.Close
  Set rsc = Nothing
 
  Combo1.Clear
  Combo1.AddItem "I-Ingreso"
  Combo1.AddItem "S-Salida"
  Combo1.AddItem "A-Anulados"
  Combo1.AddItem "T-Todos"
    
  Combo4.ListIndex = 0
   Combo2.ListIndex = 0
  Combo1.ListIndex = 0
  
  DTPicker1.Value = Date
  DTPicker2.Value = Date
  
End Sub
