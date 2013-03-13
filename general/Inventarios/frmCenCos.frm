VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCenCos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consumo de Centro de Costos"
   ClientHeight    =   3555
   ClientLeft      =   300
   ClientTop       =   1410
   ClientWidth     =   7620
   Icon            =   "frmCenCos.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   7620
   Begin VB.Frame Frame5 
      Height          =   2055
      Left            =   6240
      TabIndex        =   18
      Top             =   720
      Width           =   1215
      Begin VB.CommandButton Command8 
         Caption         =   "&Salir"
         Height          =   708
         Left            =   270
         Picture         =   "frmCenCos.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
         Height          =   708
         Left            =   240
         Picture         =   "frmCenCos.frx":0D0C
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Tipo De Reporte"
      ForeColor       =   &H000000C0&
      Height          =   1335
      Left            =   3240
      TabIndex        =   15
      Top             =   1680
      Width           =   2655
      Begin VB.OptionButton Option2 
         Caption         =   "Listado"
         Height          =   495
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton OptGerencial 
         Caption         =   "Gerencial"
         Height          =   495
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Filtro"
      ForeColor       =   &H000000C0&
      Height          =   1335
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   5655
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1545
         MaxLength       =   6
         TabIndex        =   10
         Top             =   390
         Width           =   1020
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1545
         MaxLength       =   6
         TabIndex        =   9
         Top             =   915
         Width           =   1008
      End
      Begin VB.Label Label1 
         Caption         =   "Del Centro Costo"
         Height          =   360
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1755
      End
      Begin VB.Label Label2 
         Caption         =   "Al Centro de Costo"
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   930
         Width           =   1560
      End
      Begin VB.Label LblCC1 
         Caption         =   "CC1"
         Height          =   240
         Left            =   2670
         TabIndex        =   12
         Top             =   450
         Width           =   2820
      End
      Begin VB.Label LblCC2 
         Caption         =   "CC2"
         Height          =   255
         Left            =   2670
         TabIndex        =   11
         Top             =   975
         Width           =   2805
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Rango de Fechas"
      ForeColor       =   &H000000C0&
      Height          =   1335
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   2775
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   1065
         TabIndex        =   4
         Top             =   360
         Width           =   1410
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         Format          =   57737217
         CurrentDate     =   36755
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   300
         Left            =   1065
         TabIndex        =   5
         Top             =   870
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   529
         _Version        =   393216
         Format          =   57737217
         CurrentDate     =   36755
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   330
         Left            =   120
         TabIndex        =   7
         Top             =   390
         Width           =   840
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta "
         Height          =   360
         Left            =   120
         TabIndex        =   6
         Top             =   855
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccionar"
      Height          =   3360
      Left            =   36
      TabIndex        =   0
      Top             =   0
      Width           =   6090
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   2
         Top             =   2235
         Width           =   1005
      End
      Begin VB.Label Label5 
         Caption         =   "Autorizado"
         Height          =   255
         Left            =   270
         TabIndex        =   1
         Top             =   2235
         Width           =   1500
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   96
      Top             =   3144
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmCenCos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Adodc3 As ADODB.Recordset

Private Sub Command1_Click()
Dim Codigo1 As String, Codigo2 As String
Dim aparam(5) As Variant
Dim aform(5) As Variant
Dim reporte As String
Dim cn As New ADODB.Connection
Screen.MousePointer = 11
''''''''''
Codigo1 = Trim(Text1)
If Text1 = "" Then
        MsgBox "Ingrese el codigo", vbExclamation, "Error"
        Text1.SetFocus
        Exit Sub
End If
Set Adodc3 = New ADODB.Recordset
  Adodc3.Open "SELECT centrocostocodigo,centrocostodescripcion FROM ct_centrocosto Where CENtrocostocodigo='" & Text1 & "'", VGcnxCT, adOpenStatic
  If Adodc3.RecordCount > 0 Then
     Va1 = Adodc3("CENtrocostodescripcion")
 Else
     MsgBox "El codigo:" & Text1 & "   de Centro de Costo No existe", vbExclamation, "Error"
     Screen.MousePointer = 1
    Exit Sub
  End If
  Adodc3.Close
  
  Set Adodc3 = New ADODB.Recordset
  Adodc3.Open "SELECT centrocostocodigo,centrocostodescripcion FROM ct_centrocosto Where CENtrocostocodigo='" & Text2 & "'", VGCNx, adOpenStatic
  If Adodc3.RecordCount > 0 Then
     Va2 = Adodc3("CENtrocostocodigo")
  Else
      MsgBox "El codigo:" & Text2 & "   de Centro de Costo No existe", vbExclamation, "Error"
      Screen.MousePointer = 1
     Exit Sub
  End If
reporte = "al_consumoArticulocentroCosto.rpt"
If OptGerencial.Value = True Then reporte = "al_gastosresumenccostos.rpt"
aparam(0) = VGCNx.DefaultDatabase
aparam(1) = Text1.text
aparam(2) = Text2.text
aparam(3) = DTPicker1
aparam(4) = DTPicker2

aform(0) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
aform(1) = "fecini = '" & Format(DTPicker1, "dd/mm/yyyy") & "'"
aform(2) = "fecfin = '" & Format(DTPicker2, "dd/mm/yyyy") & "'"
aform(3) = "CCini = '" & Text1 & "'"
aform(4) = "CCfin = '" & Text2 & "'"
Call ImpresionRptProc(reporte, aform, aparam, , reporte + " Consumos x centro de costos ")
Screen.MousePointer = 1
End Sub

Private Sub Command8_Click()
 Unload Me
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   SendKeys "{tab}"
End If
End Sub

Private Sub DTPicker2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   SendKeys "{tab}"
End If
End Sub

Private Sub Form_Load()
If VGcc = 1 Then
   frmCenCos.Caption = "Consumo por centro de Costos"
   Text3.Visible = False
   Label5.Visible = False
Else
   frmCenCos.Caption = "Consumo por Persona Autorizado"
   Text3.Visible = True
   Label5.Visible = True
End If
DTPicker1 = DateAdd("m", -1, Date)
DTPicker2 = Date
LblCC1 = ""
LblCC2 = ""
End Sub

Private Sub Option1_Click()

End Sub

Private Sub Text1_DblClick()
 Set Adodc3 = New ADODB.Recordset
 Adodc3.Open "SELECT centrocostocodigo,centrocostodescripcion FROM ct_centrocosto ", VGcnxCT, adOpenStatic, adLockOptimistic 'where  len(cencost_codigo) = '6' "
        frmReferencia.Conectar Adodc3, "SELECT centrocostocodigo,centrocostodescripcion FROM ct_centrocosto " ' where  len(centrocostocodigo) = '6' "
        frmReferencia.Label1.Caption = "Centro de Costos"
        frmReferencia.Show vbModal
        Adodc3.Close
        If vGUtil(1) <> "" Then
                Text1 = vGUtil(1)
                 LblCC1 = vGUtil(2)
        End If
        If Text1.text <> "" Then Text1_KeyPress (13)
  
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
   Text1_DblClick
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    SendKeys "{tab}"
    KeyAscii = 0
 End If
End Sub

Private Sub Text2_DblClick()
 Set Adodc3 = New ADODB.Recordset
 If UCase(Dir$(cRuta4)) = UCase(VGNameCont & ".MDB") Then
        Adodc3.Open "SELECT cencost_codigo,cencost_descripcion FROM centro_costos ", VGcnxCT, adOpenStatic, adLockOptimistic 'where  len(cencost_codigo) = '6'
 Else
        Adodc3.Open "SELECT centrocostocodigo,centrocostodescripcion FROM ct_centrocosto ", VGcnxCT, adOpenStatic, adLockOptimistic  'where  len(cencost_codigo) = '6' "
  End If
        frmReferencia.Conectar Adodc3, "SELECT centrocostocodigo,centrocostodescripcion FROM ct_centrocosto " 'where  len(centrocostocodigo) = '5'
        frmReferencia.Label1.Caption = "Centro de Costos"
        frmReferencia.Show vbModal
        Adodc3.Close
        If vGUtil(1) <> "" Then
                Text2 = vGUtil(1)
                LblCC2 = vGUtil(2)
        End If
  
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
   Text2_DblClick
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

   If Text2 <> "" And KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
   End If
   
End Sub

Private Sub Text3_DblClick()
FormAyuda.Show 1
End Sub
