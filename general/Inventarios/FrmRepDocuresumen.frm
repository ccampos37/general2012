VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRepDocuResumen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de Documentox Detallados"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7020
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   7020
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frameimprimir 
      Caption         =   "Imprimir Por"
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   2400
      TabIndex        =   16
      Top             =   1920
      Width           =   1815
      Begin VB.OptionButton OptUnidades 
         Caption         =   "Unidades"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton OptImportes 
         Caption         =   "Importes"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   4440
      TabIndex        =   13
      Top             =   1920
      Width           =   2175
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancelar"
         Height          =   645
         Left            =   1200
         Picture         =   "FrmRepDocuresumen.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   795
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
         Height          =   645
         Left            =   240
         Picture         =   "FrmRepDocuresumen.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Resumen Por"
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   240
      TabIndex        =   10
      Top             =   1920
      Width           =   1815
      Begin VB.OptionButton OptProducto 
         Caption         =   "Producto"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton OptTransaccion 
         Caption         =   "Transaccion"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   675
      Left            =   270
      TabIndex        =   5
      Top             =   1080
      Width           =   6405
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1290
         TabIndex        =   6
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   57999361
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   35796
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   4290
         TabIndex        =   7
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   57999361
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   35796
      End
      Begin VB.Label Label5 
         Caption         =   "Desde"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Left            =   3090
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   945
      Left            =   270
      TabIndex        =   0
      Top             =   90
      Width           =   6405
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   4950
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   450
         Width           =   1245
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   270
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   420
         Width           =   4395
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Movimiento"
         Height          =   195
         Index           =   0
         Left            =   4950
         TabIndex        =   4
         Top             =   240
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
Attribute VB_Name = "FrmRepDocuResumen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Dim VGDllGeneral As New dllgeneral.dll_general
    Dim aparam(7) As Variant
    Dim aform(5) As Variant
    Dim reporte As String
    Dim titulo As String
    aparam(0) = Trim(VGCNx.DefaultDatabase)
    aparam(1) = Left(Combo2.text, 2)
    aparam(2) = Left(Combo1.text, 1)
    aparam(3) = Format(DTPicker1.Value, "dd/mm/yyyy")
    aparam(4) = Format(DTPicker2.Value, "dd/mm/yyyy")
    aparam(5) = "%%"
    aparam(6) = "%%"
    
    aform(0) = "almacen ='" & Trim(Combo2) & "'"
    aform(1) = "fechainicio ='" & DTPicker1.Value & "'"
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
    aform(4) = "transa =' EXPRESADO EN IMPORTES'"
    If OptProducto = False Then
       titulo = "al_ResumenTransaccion -- Resumen de Transacciones"
       reporte = "al_resumenTransaccion.rpt"
     Else
       titulo = "al_resumenProducto -- Resumen de Productos x Transacciones"
       If OptImportes = True Then
          reporte = "al_resumenProducto.rpt"
        Else
         reporte = "al_resumenProductoUnidades.rpt"
         aform(4) = "transa =' EXPRESADO EN UNIDADES'"
       End If
    End If
    
    Call ImpresionRptProc(reporte, aform, aparam, , titulo)
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
  
 
  Combo1.Clear
  Combo1.AddItem "I-Ingreso"
  Combo1.AddItem "S-Salida"
  Combo1.AddItem "A-Anulados"
  Combo1.AddItem "T-Todos"
    
  Combo2.ListIndex = 0
  Combo1.ListIndex = 0
  
  DTPicker1.Value = Date
  DTPicker2.Value = Date
  
End Sub

Private Sub OptProducto_Click()
Frameimprimir.Enabled = True
End Sub

Private Sub OptTransaccion_Click()
OptProducto.Value = False
OptImportes = True
Frameimprimir.Enabled = False
End Sub
