VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmArtVen 
   Caption         =   "Artículos Vencidos"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5925
   LinkTopic       =   "Form2"
   ScaleHeight     =   4440
   ScaleWidth      =   5925
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrameRep 
      Caption         =   "Por Lotes  "
      Height          =   2070
      Left            =   240
      TabIndex        =   2
      Top             =   1215
      Width           =   2940
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1215
         MaxLength       =   20
         TabIndex        =   5
         Top             =   1080
         Width           =   1455
      End
      Begin VB.OptionButton OpTodos 
         Caption         =   "Todos los Artículos"
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   240
         Width           =   1905
      End
      Begin VB.OptionButton OpRango 
         Caption         =   "Rango"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Fin"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Inicio "
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   1080
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2055
      Left            =   3330
      TabIndex        =   7
      Top             =   1230
      Width           =   2340
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   540
         TabIndex        =   8
         Top             =   585
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   47448065
         CurrentDate     =   36752
         MaxDate         =   401768
         MinDate         =   36161
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   540
         TabIndex        =   9
         Top             =   1410
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   47448065
         CurrentDate     =   36752
         MaxDate         =   401768
         MinDate         =   36161
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha de Vencimiento"
         Height          =   255
         Left            =   150
         TabIndex        =   16
         Top             =   1065
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha de Inicio"
         Height          =   255
         Left            =   165
         TabIndex        =   13
         Top             =   285
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   735
      Left            =   1815
      Picture         =   "FrmArtVen.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3510
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   735
      Left            =   3360
      Picture         =   "FrmArtVen.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3495
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Consulta"
      Height          =   960
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   3180
      End
      Begin VB.Label Label1 
         Caption         =   "Por Almacen"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   675
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "FrmArtVen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim db As Database
Dim almacen As String

Function Existe_cod_art(text As TextBox) As String
Dim rs As Recordset
Dim rsql As String
rsql = "select  ACODIGO FROM maeart where ACODIGO = '" & text & "'" '
'Set db = Workspaces(0).OpenDatabase(cRuta2)
'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)

Set rs = VGCNx.Execute(rsql)
If Not rs.EOF Then
    Existe_cod_art = rs(0)
Else
    MsgBox "El tipo de codigo no existe !", vbOKOnly, "Error"
    Existe_cod_art = ""
End If
rs.Close
End Function

Private Sub limpiar_t1_t2()
Text1 = ""
Text2 = ""
End Sub

Private Sub imprimir()
Dim Codigo1 As String
Dim Codigo2 As String
Dim CADENA As String
Dim Sq As String
Dim tex1 As String, tex2 As String
Dim Va1 As String, Va2 As String

Dim Adodc3 As New ADODB.Recordset
Codigo1 = Trim(Text1)
If OpTodos.Value Then

    Sq = "Select ACodigo,Adescri from "
    Sq = Sq & "MAEART A Inner Join STKART B on A.ACodigo=B.STCodigo "
    Sq = Sq & "Where Stalma='" & almacen & "' Order by Acodigo"
    
    Adodc3.Open Sq, VGCNx, adOpenDynamic, adLockOptimistic
    If Adodc3.RecordCount > 0 Then
      If Text1 = "" And Text2 = "" Then
        Adodc3.MoveFirst
        tex1 = IIf(IsNull(Adodc3("ACodigo")), "", Adodc3("ACodigo"))
        Va1 = IIf(IsNull(Adodc3("ADescri")), "", Adodc3("Adescri"))
        Adodc3.MoveLast
        tex2 = IIf(IsNull(Adodc3("ACodigo")), "", Adodc3("ACodigo"))
        Va2 = IIf(IsNull(Adodc3("ADescri")), "", Adodc3("ADescri"))
      End If
    End If
    Adodc3.Close
    CrystalReport1.WindowTitle = "Inv096 -- Control de Inventarios"
    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv096.rpt"
    CADENA = "{STKLOTE.STSALMA}='" & almacen & "' and "
    CADENA = CADENA & "({STKLOTE.STSFECVEN} in DATE (" & Format(DTPicker1.Value, "yyyy") & "," & Format(DTPicker1.Value, "mm") & "," & Format(DTPicker1.Value, "dd") & ") "
    CADENA = CADENA & "to  DATE (" & Format(DTPicker2.Value, "yyyy") & "," & Format(DTPicker2.Value, "mm") & "," & Format(DTPicker2.Value, "dd") & "))"

    
    Ubi_Tab CrystalReport1
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.SelectionFormula = CADENA
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.formulas(0) = "emp = '" & VGparametros.RucEmpresa & "'"
    CrystalReport1.formulas(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
    CrystalReport1.formulas(2) = "campoini = '" & tex1 & "'"
    CrystalReport1.formulas(3) = "campofin = '" & tex2 & "'"
    CrystalReport1.formulas(4) = "detaini = '" & Va1 & "'"
    CrystalReport1.formulas(5) = "detafin = '" & Va2 & "'"

    
    CrystalReport1.WindowTitle = "Reporte de Stock por Lote de  Articulos"
    If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
    Exit Sub
End If

If Text1 = "" Then
        MsgBox "Ingrese el código", vbExclamation, "Error"
        Text1.SetFocus
        Exit Sub
End If
Codigo2 = Trim(Text2)
  Sq = "Select ADescri from MAEART Where ACodigo='" & Codigo1 & "'"
  Adodc3.Open Sq, VGCNx, adOpenDynamic, adLockOptimistic
  If Adodc3.RecordCount = 1 Then
    Va1 = Adodc3("Adescri")
  End If
  Adodc3.Close
  
  Sq = "Select ADescri from MAEART Where ACodigo='" & Codigo2 & "'"
  Adodc3.Open Sq, VGCNx, adOpenDynamic, adLockOptimistic
  If Adodc3.RecordCount = 1 Then
    Va2 = Adodc3("Adescri")
  End If
  Adodc3.Close
CrystalReport1.WindowTitle = "Inv096 -- Control de Inventarios"
 CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv096.rpt"
      If Text2 <> "" Then
              Codigo2 = Text2
              CADENA = "{STKLOTE.STSALMA}='" & almacen & "' and ({STKLOTE.STSCODIGO} in '" & Codigo1 & "' to '" & Codigo2 & "') and "
              CADENA = CADENA & "({STKLOTE.STSFECVEN} in DATE (" & Format(DTPicker1.Value, "yyyy") & "," & Format(DTPicker1.Value, "mm") & "," & Format(DTPicker1.Value, "dd") & ") "
              CADENA = CADENA & "to  DATE (" & Format(DTPicker2.Value, "yyyy") & "," & Format(DTPicker2.Value, "mm") & "," & Format(DTPicker2.Value, "dd") & "))"
      Else
              Codigo2 = Codigo1: Va2 = Va1
              CADENA = "{STKLOTE.STSALMA}='" & almacen & "' and {STKLOTE.STSCODIGO} = '" & Codigo1 & "' and "
              CADENA = CADENA & "({STKLOTE.STSFECVEN} in DATE (" & Format(DTPicker1.Value, "yyyy") & "," & Format(DTPicker1.Value, "mm") & "," & Format(DTPicker1.Value, "dd") & ") "
              CADENA = CADENA & "to  DATE (" & Format(DTPicker2.Value, "yyyy") & "," & Format(DTPicker2.Value, "mm") & "," & Format(DTPicker2.Value, "dd") & "))"
      End If
  Ubi_Tab CrystalReport1
  CrystalReport1.DiscardSavedData = True
  CrystalReport1.Destination = crptToWindow
  CrystalReport1.SelectionFormula = CADENA
  CrystalReport1.WindowShowPrintBtn = True
  CrystalReport1.WindowShowRefreshBtn = True
  CrystalReport1.WindowShowSearchBtn = True
  CrystalReport1.WindowShowPrintSetupBtn = True
  CrystalReport1.formulas(0) = "emp = '" & VGparametros.RucEmpresa & "'"
  CrystalReport1.formulas(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
'  CrystalReport1.Formulas(2) = "alm = '" & almacen & "'"
  CrystalReport1.formulas(2) = "campoini = '" & Codigo1 & "'"
  CrystalReport1.formulas(3) = "campofin = '" & Codigo2 & "'"
  CrystalReport1.formulas(4) = "detaini = '" & Va1 & "'"
  CrystalReport1.formulas(5) = "detafin = '" & Va2 & "'"

  If CrystalReport1.Status <> 2 Then
      CrystalReport1.Action = 1
  End If

End Sub


Private Sub Carga_Almacen()
Dim rsql As String
Dim rs As Recordset
 
rsql = "select TAALMA,TADESCRI FROM TabAlm "
'Set db = Workspaces(0).OpenDatabase(cRuta2)
'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
Set rs = VGCNx.Execute(rsql)
While Not rs.EOF
    Combo1.AddItem (rs(0)) & " " & (rs(1))
    rs.MoveNext
Wend
rs.Close
End Sub

Private Sub Combo1_Click()
almacen = Format(Combo1.ListIndex + 1, "00")
Text1 = "": Text2 = ""
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  OpTodos.SetFocus
End If
End Sub

Private Sub Command1_Click()
If DTPicker2 < DTPicker1 Then
     MsgBox "La fecha de Inicio no puede ser mayor que la fecha fin", vbExclamation, "Error"
     Exit Sub
End If
imprimir
End Sub

Private Sub Command7_Click()
  MousePointer = vbHourglass
  
  If Frame2.Visible And Frame3.Visible Then
        MousePointer = vbDefault
        Unload Me
  Else
        Frame2.Visible = True
        Frame3.Visible = True
        FrameRep.Visible = False
        MousePointer = vbDefault
  End If
End Sub

Private Sub DTPicker3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{tab}"
End If
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then DTPicker2.SetFocus
End Sub

Private Sub DTPicker2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Command1.SetFocus
End Sub

Private Sub Form_Load()
Frame2.Visible = True
Frame3.Visible = True
central Me
Carga_Almacen
If Combo1.ListCount > 0 Then Combo1.ListIndex = 0
'central FormStkAlm
'OpArt.Value = True
OpTodos.Value = True
FrameRep.Caption = " Por Articulos"
VGForm1 = 3
DTPicker1.Value = DateAdd("m", -1, Date)
DTPicker2.Value = Date
End Sub

Private Sub OpRango_Click()
If OpRango.Value Then
    Text1.Enabled = True
    Text2.Enabled = True
    Text1.SetFocus
End If
End Sub

Private Sub OpTodos_Click()
Text1.Enabled = False
Text2.Enabled = False
limpiar_t1_t2
End Sub


Private Sub OpTodos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then DTPicker1.SetFocus
End Sub

Private Sub Text1_DblClick()
Dim Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset
         Adodc2.Open "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD from MaeArt p, StkArt n Where n.STSKDIS<> 0 and  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & almacen & "'  ", VGCNx, adOpenStatic, adLockOptimistic
         frmReferencia.Conectar Adodc2, "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD from MaeArt p, StkArt n Where n.STSKDIS<> 0 and  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & almacen & "'  "
         frmReferencia.Label1.Caption = "Artículos"
         frmReferencia.Show vbModal
         Adodc2.Close
         If vGUtil(1) <> "" Then
                 Text1 = (vGUtil(1))
         End If
         If Text1 <> "" And Text2 <> "" And Text1 > Text2 Then
                 MsgBox "Ingrese un código menor al fin ", vbOKOnly, "Error"
                 Exit Sub
        End If
        If Text1 <> "" Then
                 Text2.Enabled = True
                 Text2.SetFocus
        End If
End Sub


Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then Text1_DblClick
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Not OpRango.Value Then
   OpRango = True
End If
If KeyAscii = 13 And Text1 <> "" Then
       If Existe_cod_art(Text1) <> "" Then
               Text2.Enabled = True
               Text2.SetFocus
       End If
End If
End Sub

Private Sub Text2_DblClick()
Dim Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset

    Adodc2.Open "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD from MaeArt p, StkArt n Where n.STSKDIS<> 0 and  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & almacen & "'  ", VGCNx, adOpenStatic, adLockOptimistic
    frmReferencia.Conectar Adodc2, "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD from MaeArt p, StkArt n Where n.STSKDIS<> 0 and  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & almacen & "'  "
    frmReferencia.Label1.Caption = "Artículos"
    frmReferencia.Show vbModal
    Adodc2.Close
    If vGUtil(1) <> "" Then
        Text2 = (vGUtil(1))
    End If
   If Text2 <> "" Then
        DTPicker1.SetFocus
   End If
End Sub


Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then Text2_DblClick
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 And Text2 <> "" Then
        If Existe_cod_art(Text2) <> "" Then
           If Text1 > Text2 Then
                  MsgBox "El codigo fin debe ser mayor que el inicio", vbInformation, mensaje1
                  Exit Sub
           End If
           DTPicker1.SetFocus
        End If
  End If
  If KeyAscii = 13 And Text2 = "" Then
      DTPicker1.SetFocus
  End If
End Sub

