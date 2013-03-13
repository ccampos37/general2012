VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmExportarExcel 
   Caption         =   "Exportar a Excel"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7680
   LinkTopic       =   "Form2"
   ScaleHeight     =   4890
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2535
      Left            =   1560
      TabIndex        =   3
      Top             =   960
      Width           =   5055
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2040
         MaxLength       =   20
         TabIndex        =   8
         Top             =   960
         Width           =   1575
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "FrmExportarExcel.frx":0000
         Left            =   2040
         List            =   "FrmExportarExcel.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   480
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Todos los Establecimientos"
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Imprimir"
         Height          =   615
         Left            =   840
         TabIndex        =   5
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Salir"
         Height          =   615
         Left            =   2760
         TabIndex        =   4
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Articulo Inicial"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   1215
      End
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   3975
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   7011
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Cruce de Guias de Remision vs Ventas"
      TabPicture(0)   =   "FrmExportarExcel.frx":0004
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Cruce Ventas vs Guias de Remision"
      TabPicture(1)   =   "FrmExportarExcel.frx":0020
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6015
      Left            =   12240
      TabIndex        =   0
      Top             =   1920
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   10610
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "FrmExportarExcel.frx":003C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Command1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "FrmExportarExcel.frx":0058
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Left            =   1080
         TabIndex        =   1
         Top             =   4920
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmExportarExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs1 As New ADODB.Recordset
Private Sub Check1_Click()
If Check1.Value = 1 Then
   Combo2.Enabled = False
Else
   Combo2.Enabled = True
End If
End Sub

Private Sub Command2_Click()
Dim aparam(5) As Variant
Dim aform(3) As Variant
aparam(0) = VGParamSistem.BDEmpresa
aparam(1) = VGparametros.empresacodigo
aparam(2) = Format(VGParamSistem.FechaTrabajo, "yyyy") + Format(VGParamSistem.FechaTrabajo, "mm")
aparam(3) = IIf(Check1.Value = 1, "%%", Left(Combo2.text, 2))
aparam(4) = IIf(Text1.text = "", "%%", Text1.text)
aform(0) = "Titulo='GR VS Ventas'"
aform(1) = "periodo='PERIODO " & Format(VGParamSistem.FechaTrabajo, "yyyy") + " - " + Format(VGParamSistem.FechaTrabajo, "mm") & "'"
aform(2) = "Fecha='" & Date & "'"
Call ImpresionRptProc("al_guias_ventas.rpt", aform, aparam, , " Guias Remision VS Ventas")
End Sub

Private Sub Command3_Click()
Call exportarExcel(rs1, "xxxx")
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call Carga_puntovta
Check1.Value = 1
End Sub
Private Sub Carga_puntovta()
Dim RSQL As String
Dim I As Integer
RSQL = "Select puntovtacodigo,puntovtadescripcion  FROM vt_puntoventa "
Set puntovta = New ADODB.Recordset
puntovta.Open RSQL, VGCNx, adOpenStatic
If puntovta.RecordCount = 0 Then Exit Sub
Do While Not puntovta.EOF
     Combo2.AddItem (Trim(puntovta(0)) & " - " & Trim(puntovta(1)))
     puntovta.MoveNext
     If puntovta.EOF Then Exit Do
Loop
puntovta.MoveFirst
For I = 0 To puntovta.RecordCount - 1
  If puntovta(0) = VGparametros.puntovta Then
    Combo2.ListIndex = I
    Exit For
  Else
    puntovta.MoveNext
  End If
Next
End Sub
Private Sub Text1_DblClick()
    FormAyuArt1.Show 1
    If Text1 <> "" And Text2 <> "" And Text1 > Text2 Then
         MsgBox "Ingrese un codigo menor al fin ", vbOKOnly, "Error"
         Exit Sub
    End If
    If Text1 <> "" Then
         Text2.Enabled = True
         Text2.SetFocus
    End If
End Sub
