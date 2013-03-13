VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRepMovFec 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Documento de Almacen"
   ClientHeight    =   3135
   ClientLeft      =   3360
   ClientTop       =   2685
   ClientWidth     =   5895
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   5895
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   1800
      Top             =   2412
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   735
      Left            =   108
      Picture         =   "FrmRepMovFec.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2340
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   936
      Picture         =   "FrmRepMovFec.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2340
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   72
      TabIndex        =   6
      Top             =   72
      Width           =   5745
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   252
         Left            =   1920
         TabIndex        =   1
         Top             =   672
         Width           =   1824
         _ExtentX        =   3228
         _ExtentY        =   450
         _Version        =   393216
         Format          =   47185921
         CurrentDate     =   36928
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   276
         Left            =   1920
         TabIndex        =   0
         Top             =   300
         Width           =   1824
         _ExtentX        =   3228
         _ExtentY        =   476
         _Version        =   393216
         Format          =   47185921
         CurrentDate     =   36928
      End
      Begin VB.ComboBox Combo3 
         Height          =   288
         ItemData        =   "FrmRepMovFec.frx":0884
         Left            =   1920
         List            =   "FrmRepMovFec.frx":0894
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1065
         Width           =   1812
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1908
         MaxLength       =   2
         TabIndex        =   3
         Top             =   1500
         Width           =   432
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   4200
         Picture         =   "FrmRepMovFec.frx":08AB
         Top             =   630
         Width           =   480
      End
      Begin VB.Label Label6 
         Caption         =   "Fec. Final"
         Height          =   330
         Left            =   165
         TabIndex        =   11
         Top             =   705
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Fec. Inicio"
         Height          =   285
         Left            =   135
         TabIndex        =   10
         Top             =   330
         Width           =   1035
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo de Documento"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1095
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Del Movimiento"
         Height          =   255
         Left            =   180
         TabIndex        =   8
         Top             =   1515
         Width           =   1365
      End
      Begin VB.Label lbltrans1 
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
         Left            =   2376
         TabIndex        =   7
         Top             =   1500
         Width           =   3312
      End
   End
End
Attribute VB_Name = "FrmRepMovFec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tipo  As String


Private Sub Command1_Click()
Screen.MousePointer = 11
imprimir
Screen.MousePointer = 1
End Sub

Private Sub Command7_Click()
 Unload Me
End Sub

Private Sub Form_Load()
DTPicker1 = Date
DTPicker2 = Date
Combo3.ListIndex = 0
tipo = "I"
lbtrans1 = ""
End Sub

Private Sub Label5_Click()

End Sub

Private Sub imprimir()
'    cadena = " {MOVALMCAB.CAALMA}='" & VGAlma & "' and ({MOVALMCAB.CAFECDOC} IN DATE (" & Format(DTPicker1, "yyyy") & "," & Format(DTPicker1, "mm") & "," & Format(DTPicker1, "dd") & ") "
'    cadena = cadena & "to DATE (" & Format(DTPicker2, "yyyy") & "," & Format(DTPicker2, "mm") & "," & Format(DTPicker2, "dd") & ")) "
'    If Trim(Text1) <> "" Then
'      cadena = cadena & " and {MOVALMCAB.CACODMOV}='" & Text1 & "'"
'    End If
'    If Mid(Combo3.text, 1, 2) <> "TO" Then
'      cadena = cadena & "and {MOVALMCAB.catd}='" & Mid(Combo3.text, 1, 2) & "' "
'    End If
'    cadena = cadena & " And {MOVALMCAB.CASITGUI} <>'A' "
    CrystalReport1.WindowTitle = "Inv130 -- Control de Inventarios"
    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv130.rpt"
    'Ubi_Tab CrystalReport1
    CrystalReport1.DiscardSavedData = True
    
    
    
    CrystalReport1.Connect = VGcadenareport2
    
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.SelectionFormula = CADENA
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.formulas(0) = "emp = '" & VGparametros.RucEmpresa & "'"
    CrystalReport1.formulas(1) = "Fecha = '" & DTPicker1 & " Al " & DTPicker2 & "' "
    CrystalReport1.formulas(2) = "Almacen = '" & VGNomAlm & "'"
    CrystalReport1.StoredProcParam(0) = VGCNx.DefaultDatabase
    CrystalReport1.StoredProcParam(1) = VGAlma
    CrystalReport1.StoredProcParam(2) = DTPicker1.Value
    CrystalReport1.StoredProcParam(3) = DTPicker2.Value
    CrystalReport1.StoredProcParam(4) = Left(Combo3.text, 2)
    
    If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1

End Sub

Private Sub Text1_DblClick()
Dim Adodc3 As ADODB.Recordset
        If Combo3.ListIndex <> 0 Then
           tipo = "S"
        Else
           tipo = "I"
        End If
        Set Adodc3 = New ADODB.Recordset
        Adodc3.Open "SELECT TT_CODMOV,TT_DESCRI FROM Tabtransa where  TT_tipmov = '" & tipo & "'", VGCNx, adOpenStatic, adLockOptimistic
        frmReferencia.Conectar Adodc3, "SELECT TT_CODMOV,TT_DESCRI FROM Tabtransa where  TT_tipmov = '" & tipo & "'"
        frmReferencia.Label1.Caption = "Transacciones"
        frmReferencia.Show vbModal
        Adodc3.Close
        If vGUtil(1) <> "" Then
                Text1 = vGUtil(1)
                lbltrans1 = Mid(vGUtil(2), 1, 21)
        End If
        If Text1.text <> "" Then Text1_KeyPress (13)
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Text1_DblClick
ElseIf KeyCode = 46 Then
    lbltrans1 = ""
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim Adodc3 As ADODB.Recordset
        If Combo3.ListIndex <> 0 Then
           tipo = "S"
        End If
        If Text1 <> "" And KeyAscii = 13 Then
          Text1 = UCase(Text1)
          Set Adodc3 = New ADODB.Recordset
          Adodc3.Open "SELECT TT_CODMOV,TT_DESCRI FROM Tabtransa where  TT_tipmov = '" & tipo & "' and TT_codmov = '" & Text1 & "' ", VGCNx, adOpenStatic, adLockOptimistic
          If Not Adodc3.EOF Then
            lbtrans1 = Mid(Adodc3(1), 1, 21)
          Else
            MsgBox "No existe el tipo de transacción", vbOKOnly, "Aviso"
          End If
          Adodc3.Close
          SendKeys "{tab}"
        Else
           KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If

End Sub
