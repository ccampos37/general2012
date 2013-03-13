VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form RepKardexValTXDocumento 
   Caption         =   "Reporte De Kardex Valorizado por Documento"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6600
   LinkTopic       =   "Form2"
   ScaleHeight     =   3330
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1665
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   6405
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   270
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   420
         Width           =   4395
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   4950
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   450
         Width           =   1245
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1050
         Width           =   3735
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1050
         Width           =   2145
      End
      Begin VB.Label Label1 
         Caption         =   "Almacen Origen"
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   15
         Top             =   210
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Movimiento"
         Height          =   195
         Index           =   0
         Left            =   4950
         TabIndex        =   14
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label2 
         Caption         =   "Almacen Destino"
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   810
         Width           =   1275
      End
      Begin VB.Label Label2 
         Caption         =   "Transaccion"
         Height          =   225
         Index           =   1
         Left            =   4050
         TabIndex        =   12
         Top             =   840
         Width           =   1275
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   645
      Left            =   4005
      Picture         =   "RepKardexValTXDocumento.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2610
      Width           =   795
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   645
      Left            =   2730
      Picture         =   "RepKardexValTXDocumento.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2610
      Width           =   795
   End
   Begin VB.Frame Frame2 
      Height          =   675
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   6405
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1290
         TabIndex        =   1
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   99483649
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   35796
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   4290
         TabIndex        =   2
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   99483649
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   35796
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Left            =   3090
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Desde"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   270
         Width           =   975
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   0
      Top             =   2940
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "RepKardexValTXDocumento"
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
    Dim arrform(6) As Variant, arrparam(7) As Variant
    Dim reporte As String, titrep As String
  
    titrep = "Inv035 -- Kardex Valorizado Detallado por Documento"
reporte = "inv035.rpt"
   'Ubi_Tab CrystalReport1
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    arrparam(0) = Trim(VGCNx.DefaultDatabase)
    arrparam(1) = Left(Combo2.text, 2)
    arrparam(2) = Left(Combo1.text, 1)
    arrparam(3) = Format(DTPicker1.Value, "dd/mm/yyyy")
    arrparam(4) = Format(DTPicker2.Value, "dd/mm/yyyy")
    arrparam(5) = IIf(Trim(VGDllGeneral.ComboDato(Combo3.text)) = "T", "%", VGDllGeneral.ComboDato(Combo3.text)) 'almacen destino
    arrparam(6) = IIf(Trim(VGDllGeneral.ComboDato(Combo4.text)) = "T", "%", VGDllGeneral.ComboDato(Combo4.text))  'transaccion
    
    arrform(0) = "almacen ='" & Trim(Combo2) & "'"
    arrform(1) = "fechainicio ='" & DTPicker1 & "'"
    arrform(2) = "fechafin ='" & DTPicker2 & "'"
    Select Case Left(Combo1.text, 1)
        Case "I"
            arrform(3) = "tipo ='" & "**INGRESOS**" & "'"
        Case "S"
            arrform(3) = "tipo ='" & "**SALIDAS**" & "'"
        Case "A"
            arrform(3) = "tipo ='" & "**ANULADOS**" & "'"
        Case "T"
            arrform(3) = "tipo ='" & "**TODOS**" & "'"
    End Select
    arrform(4) = "destino ='" & Trim(Combo3.text) & "'"
    arrform(5) = "transa ='" & Trim(Combo4.text) & "'"
      
    Call ImpresionRptProc(reporte, arrform, arrparam, "", titrep)
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
  Dim rsc As New ADODB.Recordset
  
  central Me
  Set Rs = VGCNx.Execute("select TAALMA,TADESCRI,'','' from tabalm where taalma='*'")
  
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
 
  Combo3.Clear
  Combo3.AddItem "T-Todos"
  'Set rsc = VGCNx.Execute("select CENCOST_CODIGO,CENCOST_DESCRIPCION  from CENTRO_COSTOS")
  Set rsc = VGCNx.Execute("select centrocostocodigo,centrocostodescripcion  from ct_centrocosto")
  If rsc.RecordCount > 0 Then
      rsc.MoveFirst
      Do Until rsc.EOF
        Combo3.AddItem rsc.Fields(0) & "-" & IIf(IsNull(rsc.Fields(1)), "****", rsc.Fields(1))
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
  Combo3.ListIndex = 0
  Combo2.ListIndex = 0
  Combo1.ListIndex = 0
  
  DTPicker1.Value = Date
  DTPicker2.Value = Date
  
End Sub


