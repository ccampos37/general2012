VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmStockInicial 
   Caption         =   "Informe de Stock Inicial"
   ClientHeight    =   2760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5265
   LinkTopic       =   "Form2"
   ScaleHeight     =   2760
   ScaleWidth      =   5265
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   120
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   735
      Left            =   1440
      Picture         =   "FrmStockInicial.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   2880
      Picture         =   "FrmStockInicial.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   4695
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   840
         Width           =   3135
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "Transaccion"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Almacen"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmStockInicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    PROCESA
End Sub

Private Sub Command8_Click()
    Unload Me
End Sub

Private Sub Form_Load()
  Dim rsc As New ADODB.Recordset
  
  Combo1.Clear
  Set rsc = VGCNx.Execute("select TAALMA,TADESCRI from tabalm ")
  If rsc.RecordCount > 0 Then
      rsc.MoveFirst
      Do Until rsc.EOF
        Combo1.AddItem rsc.Fields(0) & "-" & IIf(IsNull(rsc.Fields(1)), "****", rsc.Fields(1))
        rsc.MoveNext
      Loop
  End If
  rsc.Close
  Set rsc = Nothing
  
  Combo2.Clear
  Set rsc = VGCNx.Execute("select TT_CODMOV,TT_DESCRI from tabtransa where TT_TIPMOV='I'")
  If rsc.RecordCount > 0 Then
      rsc.MoveFirst
      Do Until rsc.EOF
        Combo2.AddItem rsc.Fields(0) & "-" & IIf(IsNull(rsc.Fields(1)), "****", rsc.Fields(1))
        rsc.MoveNext
      Loop
  End If
  rsc.Close
  Set rsc = Nothing
  
End Sub

Sub PROCESA()
Dim aparam(4) As Variant
Dim aform(1) As Variant
    Dim I1, I2 As Integer
    Dim NCAD1, NCAD2 As String
    
    I1 = InStr(Combo1.text, "-")
    If I1 > 0 Then
       NCAD1 = Left(Combo1.text, I1 - 1)
       If I1 < 1 Then
         Exit Sub
       End If
    Else
     NCAD1 = "%%"
    End If
    I2 = InStr(Combo2.text, "-")
    NCAD2 = Left(Combo2.text, I2 - 1)
        
       
   aparam(0) = VGCNx.DefaultDatabase
   aparam(1) = VGparametros.empresacodigo
   aparam(2) = NCAD1
   aparam(3) = NCAD2
       
   aform(0) = "Transaccion='" & Combo2.text & "'"
  Call ImpresionRptProc("inv038.rpt", aform, aparam, , "Saldos Iniciales")
End Sub
