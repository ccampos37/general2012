VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmStkAlm 
   Caption         =   "Stock de Articulos"
   ClientHeight    =   4845
   ClientLeft      =   5325
   ClientTop       =   1215
   ClientWidth     =   5895
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   5895
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   3840
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   3570
      Width           =   1440
   End
   Begin VB.CheckBox ChkSerie 
      Caption         =   "Con Serie o Lote"
      Enabled         =   0   'False
      Height          =   255
      Left            =   315
      TabIndex        =   24
      Top             =   3570
      Width           =   1695
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   300
      Top             =   4140
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Consulta"
      Height          =   1215
      Left            =   240
      TabIndex        =   17
      Top             =   120
      Width           =   5415
      Begin VB.CheckBox Chkprecio 
         Caption         =   "Con Costo"
         Enabled         =   0   'False
         Height          =   195
         Left            =   3000
         TabIndex        =   2
         Top             =   825
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   3180
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Consolidado"
         Enabled         =   0   'False
         Height          =   225
         Left            =   1200
         TabIndex        =   1
         Top             =   795
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Por Almacen"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parámetros"
      Height          =   2055
      Left            =   240
      TabIndex        =   16
      Top             =   1440
      Width           =   1815
      Begin VB.OptionButton Option1 
         Caption         =   "Cod. Ubicacion"
         Enabled         =   0   'False
         Height          =   195
         Left            =   270
         TabIndex        =   23
         Top             =   1665
         Width           =   1425
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Grupos"
         Enabled         =   0   'False
         Height          =   255
         Left            =   255
         TabIndex        =   6
         Top             =   1320
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Líneas"
         Enabled         =   0   'False
         Height          =   255
         Left            =   255
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Familias"
         Enabled         =   0   'False
         Height          =   255
         Left            =   255
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton OpArt 
         Caption         =   "Artículos"
         Height          =   255
         Left            =   255
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame FrameRep 
      Enabled         =   0   'False
      Height          =   2055
      Left            =   2400
      TabIndex        =   15
      Top             =   1440
      Width           =   3255
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   8
         Top             =   510
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   7
         Top             =   210
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.OptionButton OpRango 
         Caption         =   "Rango"
         Enabled         =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   1035
         Width           =   1455
      End
      Begin VB.OptionButton OpTodos 
         Caption         =   "Todos los Artículos"
         Height          =   270
         Left            =   360
         TabIndex        =   9
         Top             =   765
         Width           =   1785
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   11
         Top             =   1320
         Width           =   1470
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   12
         Top             =   1635
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Línea"
         Height          =   225
         Left            =   330
         TabIndex        =   22
         Top             =   525
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Familia"
         Height          =   225
         Left            =   300
         TabIndex        =   21
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Inicio "
         Enabled         =   0   'False
         Height          =   255
         Left            =   765
         TabIndex        =   20
         Top             =   1335
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Fin"
         Enabled         =   0   'False
         Height          =   255
         Left            =   765
         TabIndex        =   19
         Top             =   1635
         Width           =   735
      End
   End
   Begin VB.CommandButton Command7 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   624
      Left            =   3390
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4176
      Width           =   810
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   624
      Left            =   2220
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4176
      Width           =   810
   End
   Begin VB.CheckBox Chkstokcero 
      Caption         =   "Suprimir Stock   0"
      Height          =   255
      Left            =   288
      TabIndex        =   27
      Top             =   3924
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Ordenado por "
      Height          =   255
      Left            =   2505
      TabIndex        =   25
      Top             =   3585
      Width           =   990
   End
End
Attribute VB_Name = "FrmStkAlm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim db As Database
Dim almacen As String
Dim Conexion As String
Dim Adodc3 As ADODB.Recordset

Private Sub Combo1_Click()
'almacen = Format(Combo1.ListIndex + 1, "00")
almacen = Mid(Combo1, 1, 2)
End Sub

Private Sub Command7_Click()
  MousePointer = vbHourglass
  If Frame1.Visible And Frame2.Visible Then
        MousePointer = vbDefault
        Unload Me
  Else
        Frame1.Visible = True
        Frame2.Visible = True
        FrameRep.Visible = False
        MousePointer = vbDefault
  End If
End Sub

Private Sub Command1_Click()
   
    If Text1 <> "" And Text2 <> "" And Text1 > Text2 Then
        MsgBox "Ingrese un código menor al fin ", vbOKOnly, "Error"
        Exit Sub
    End If
    Screen.MousePointer = 11
    If OpArt.Value Then
         imprimir
    ElseIf Option2.Value Then
        Imprimir2
    ElseIf Option3.Value Then
        Imprimir3
    ElseIf Option4.Value Then
        Imprimir4
    ElseIf Option1.Value Then
        Imprimir5
    End If
    Screen.MousePointer = 1
End Sub

Private Sub Form_Load()
Frame1.Visible = True
Frame2.Visible = True
Carga_Almacen
central FrmStkAlm
OpArt.Value = True
OpTodos.Value = True
FrameRep.Caption = " Por Articulos"
VGForm1 = 3
Chkstokcero.Value = 1
End Sub

Private Sub OpRango_Click()
If OpRango.Value Then
    Text1.Enabled = True
    Text2.Enabled = True
    Text1.SetFocus
End If
End Sub

Private Sub OpArt_Click()
OpArt.Value = True
FrameRep.Caption = " Por Articulos "
OpTodos.Caption = "Todos los Articulos"
limpiar_t1_t2
OpTodos.Top = 300: OpRango.Top = 650
Text1.Top = 1100: Label2.Top = 1100
Text2.Top = 1500: Label3.Top = 1500
End Sub

Private Sub Option2_Click()
Option2.Value = True
FrameRep.Caption = " Por Familias "
OpTodos.Caption = "Todos las Familias"
limpiar_t1_t2
OpTodos.Top = 300: OpRango.Top = 650
Text1.Top = 1100: Label2.Top = 1100
Text2.Top = 1500: Label3.Top = 1500
End Sub

Private Sub Option3_Click()
Option3.Value = True
FrameRep.Caption = " Por Lineas "
OpTodos.Caption = "Todos las Líneas "
limpiar_t1_t2
Label4.Visible = True
Text3.Visible = True
OpTodos.Top = 550: OpRango.Top = 900
Text1.Top = 1200: Label2.Top = 1200
Text2.Top = 1600: Label3.Top = 1600
End Sub

Private Sub Option4_Click()
Option4.Value = True
FrameRep.Caption = " Por Grupos "
OpTodos.Caption = "Todos los Grupos"
limpiar_t1_t2
Label4.Visible = True: Text3.Visible = True
Label5.Visible = True: Text4.Visible = True
OpTodos.Top = 850: OpRango.Top = 1100
Text1.Top = 1400: Label2.Top = 1400
Text2.Top = 1700: Label3.Top = 1700
End Sub

Private Sub Carga_Almacen()
Dim rsql As String
Dim rs As Recordset
Dim I As Integer
 
rsql = "select TAALMA,TADESCRI FROM TabAlm "
Set rs = VGCNx.Execute(rsql)
While Not rs.EOF
  Combo1.AddItem (rs(0)) & "  " & (rs(1))
  rs.MoveNext
Wend

rs.MoveFirst
For I = 0 To rs.RecordCount - 1
  If rs(0) = VGAlma Then
    Combo1.ListIndex = I
    Exit For
  Else
    rs.MoveNext
  End If
Next
 rs.Close
 Combo2.AddItem ("codigo")
 Combo2.AddItem ("Descripcion")
 Combo2.ListIndex = 1
End Sub

Private Sub imprimir()
Dim Codigo1 As String
Dim Codigo2 As String
Dim CADENA As String
Dim rsql As String
Dim where As String
Dim tex1 As String, tex2 As String
Dim Va1 As String, Va2 As String
Dim aparam(4) As Variant
Dim aform(1) As Variant

Codigo1 = UCase(Trim(Text1))
Set Adodc3 = New ADODB.Recordset

aparam(0) = VGCNx.DefaultDatabase
aparam(1) = "" & Left(Combo1.text, 2) & ""
aparam(2) = Chkstokcero.Value
aparam(3) = Combo2.ListIndex
    
aform(0) = "almacen='" & Combo1.text & "'"

If OpTodos.Value Then
    where = " "
    If Check1.Value = 0 Then
            If ChkSerie = 1 Then
                    CrystalReport1.WindowTitle = "Inv136 -- Control de Inventarios"
                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & IIf(Combo2.ListIndex = 1, "inv142.rpt", "inv136.rpt")
            ElseIf Chkprecio = 0 Then
                    Call ImpresionRptProc("INV078.rpt", aform, aparam, , "Inv078 -- Control de Inventarios")

            Else
                    CrystalReport1.WindowTitle = "Inv066-- Control de Inventarios"
                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv066.rpt"
            End If

    Else                                     'Consolidado
            If ChkSerie = 1 Then
                    CrystalReport1.WindowTitle = "Inv137 -- Control de Inventarios"
                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & IIf(Combo2.ListIndex = 1, "inv143.rpt", "inv137.rpt")
            ElseIf Chkprecio = 0 Then
                    CrystalReport1.WindowTitle = "Inv079 -- Control de Inventarios"
                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv079.rpt"
            Else
                    CrystalReport1.WindowTitle = "Inv081 -- Control de Inventarios"
                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv081.rpt"
            End If
    End If
    Exit Sub ''
End If

If Text1 = "" Then
        MsgBox "Ingrese el codigo", vbExclamation, "Error"
'        Text1.SetFocus
        Exit Sub
End If
  Codigo2 = UCase(Trim(Text2))
  rsql = "Select ADescri from MAEART Where ACodigo='" & Codigo1 & "'"
  Adodc3.Open rsql, VGCNx, adOpenDynamic, adLockOptimistic
  If Adodc3.RecordCount = 1 Then
    Va1 = Adodc3("Adescri")
  End If
  Adodc3.Close
  
  rsql = "Select ADescri from MAEART Where ACodigo='" & Codigo2 & "'"
  Adodc3.Open rsql, VGCNx, adOpenDynamic, adLockOptimistic
  If Adodc3.RecordCount = 1 Then
    Va2 = Adodc3("Adescri")
  End If
  Adodc3.Close

If OpArt.Value Then           'Un select
    If Check1.Value = 1 Then
            If ChkSerie = 1 Then
                    CrystalReport1.WindowTitle = "Inv137 -- Control de Inventarios"
                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv137.rpt"
            ElseIf Chkprecio = 0 Then
                    CrystalReport1.WindowTitle = "Inv079 -- Control de Inventarios"
                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv079.rpt"
            Else
                    CrystalReport1.WindowTitle = "Inv081 -- Control de Inventarios"
                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv081.rpt"
            End If
            If Text2 <> "" Then
                    Codigo2 = Text2
                    CADENA = "({STKART.STCODIGO} in '" & Codigo1 & "' to '" & Codigo2 & "')"
            Else
                    Codigo2 = Codigo1: Va2 = Va1
                    CADENA = "{STKART.STCODIGO} = '" & Codigo1 & "' "
            End If
    Else
            If ChkSerie = 1 Then
                    CrystalReport1.WindowTitle = "Inv136 -- Control de Inventarios"
                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv136.rpt"
            ElseIf Chkprecio = 0 Then
                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv078.rpt"
                    CrystalReport1.WindowTitle = "Inv078 -- Control de Inventarios"
            Else
                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv066.rpt"
                    CrystalReport1.WindowTitle = "Inv066 -- Control de Inventarios"
            End If
            If Text2 <> "" Then
                    Codigo2 = Text2         '  "23134671"
                    CADENA = " {STKART.STALMA}='" & almacen & "' and ({STKART.STCODIGO} in '" & Codigo1 & "' to '" & Codigo2 & "')"
            Else
                    Codigo2 = Codigo1: Va2 = Va1
                    CADENA = "{STKART.STALMA}='" & almacen & "' and {STKART.STCODIGO} = '" & Codigo1 & "' "
            End If
    End If
    Ubi_Tab CrystalReport1
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.SelectionFormula = CADENA
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    If Combo2.ListIndex = 1 Then
      CrystalReport1.SortFields(0) = "+{MAEART.ADESCRI}"
    Else
      CrystalReport1.SortFields(0) = "+{STKART.STCODIGO}"
    End If
    CrystalReport1.formulas(0) = "emp = '" & VGparametros.RucEmpresa & "'"
    CrystalReport1.formulas(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
    CrystalReport1.formulas(2) = "campoini = '" & Codigo1 & "'"
    CrystalReport1.formulas(3) = "campofin = '" & Codigo2 & "'"
    CrystalReport1.formulas(4) = "detaini = '" & Va1 & "'"
    CrystalReport1.formulas(5) = "detafin = '" & Va2 & "'"
    If CrystalReport1.Status <> 2 Then
        CrystalReport1.Action = 1
    End If
End If
End Sub

Private Sub OpTodos_Click()
Text1.Enabled = False
Text2.Enabled = False
limpiar_t1_t2
If Option3.Value Then
  Label4.Visible = True
  Text3.Visible = True
ElseIf Option4.Value Then
  Label4.Visible = True: Label5.Visible = True
  Text3.Visible = True: Text4.Visible = True
End If
End Sub

Private Sub OpTodos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command1.SetFocus
End Sub

Private Sub Text1_DblClick()
Dim Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset
If OpArt.Value Or Option1 Then
'         Adodc2.Open "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD,p.ACODIGO2 from MaeArt p, StkArt n Where n.STSKDIS<> 0 and  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & almacen & "'  ", Vgcnx, adOpenStatic, adLockOptimistic
'         frmReferencia.conectar Adodc2, "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD,p.ACODIGO2 from MaeArt p, StkArt n Where n.STSKDIS<> 0 and  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & almacen & "'  "
'         frmReferencia.Label1.Caption = "Artículos"
'         frmReferencia.show vbmodal
'         Adodc2.Close
'         If vGUtil(1) <> "" Then
'                 Text1 = (vGUtil(1))
'         End If
'         If Text1 <> "" And Text2 <> "" And Text1 > Text2 Then
'                 MsgBox "Ingrese un codigo menor al fin ", vbOKOnly, "Error"
'                 Exit Sub
'        End If
'        If Text1 <> "" Then
'                 Text2.Enabled = True
'                 Text2.SetFocus
'        End If
         VGForm1 = 3
         FormAyuArt1.Show 1
         If Text1 <> "" And Text2 <> "" And Text1 > Text2 Then
              MsgBox "Ingrese un codigo menor al fin ", vbOKOnly, "Error"
              Exit Sub
         End If
         If Text1 <> "" Then
              Text2.Enabled = True
              Text2.SetFocus
         End If
ElseIf Option2.Value Then
        Adodc2.Open "SELECT FAM_CODIGO,FAM_NOMBRE FROM FAMILIA", VGCNx, adOpenStatic, adLockOptimistic
        frmReferencia.Conectar Adodc2, "SELECT FAM_CODIGO,FAM_NOMBRE FROM FAMILIA"
        frmReferencia.Label1.Caption = "Familias de Artículos"
        frmReferencia.Show vbModal
        Adodc2.Close
        If vGUtil(1) <> "" Then
                Text1 = (vGUtil(1))
        End If
ElseIf Option3.Value Then
        Adodc2.Open "SELECT LIN_CODIGO,LIN_NOMBRE FROM LINEAS Where Fam_Codigo='" & Trim(Text3) & "'", VGCNx, adOpenStatic, adLockOptimistic
        frmReferencia.Conectar Adodc2, "SELECT LIN_CODIGO,LIN_NOMBRE FROM LINEAS Where Fam_Codigo='" & Trim(Text3) & "'"
        frmReferencia.Label1.Caption = "Líneas de Artículos"
        frmReferencia.Show vbModal
        Adodc2.Close
        If vGUtil(1) <> "" Then
                Text1 = (vGUtil(1))
        End If
        If Text1 <> "" Then
                 Text2.Enabled = True
                 Text2.SetFocus
        End If
ElseIf Option4.Value Then
        Adodc2.Open "SELECT GRU_CODIGO,GRU_NOMBRE FROM GRUPO Where Fam_Codigo='" & Trim(Text3) & "' and Lin_Codigo='" & Trim(Text4) & "'", VGCNx, adOpenStatic, adLockOptimistic
        frmReferencia.Conectar Adodc2, "SELECT GRU_CODIGO,GRU_NOMBRE FROM GRUPO Where Fam_Codigo='" & Trim(Text3) & "' and Lin_Codigo='" & Trim(Text4) & "'"
        frmReferencia.Label1.Caption = "Grupos de Artículos"
        frmReferencia.Show vbModal
        Adodc2.Close
        If vGUtil(1) <> "" Then
                Text1 = (vGUtil(1))
        End If
        If Text1 <> "" Then
                 Text2.Enabled = True
                 Text2.SetFocus
        End If
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
    If OpArt.Value Then
       If Existe_cod_art(Text1) <> "" Then
               Text2.Enabled = True
               Text2.SetFocus
       End If
   ElseIf Option2.Value Then
        If Existe(1, Text1, "FAMILIA", "FAM_CODIGO", False) = False Then
                MsgBox "El código de Familia no existe", vbInformation, mensaje1
                Text1.SetFocus: Exit Sub
         Else
                Text2.Enabled = True
                Text2.SetFocus
         End If
   ElseIf Option3.Value Then
        If Existe(1, Text1, "LINEAS", "LIN_CODIGO", False) = False Then
                MsgBox "El código de Línea no existe", vbInformation, mensaje1
                Text1.SetFocus: Exit Sub
         Else
                Text2.Enabled = True
                Text2.SetFocus
         End If
   ElseIf Option4.Value Then
        If Existe(1, Text1, "GRUPO", "GRU_CODIGO", False) = False Then
                MsgBox "El código de Grupo no existe", vbInformation, mensaje1
                Text1.SetFocus: Exit Sub
         Else
                Text2.Enabled = True
                Text2.SetFocus
         End If
     End If
 End If
End Sub

Private Sub Text2_DblClick()
Dim Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset

If OpArt.Value Or Option1.Value Then
'    Adodc2.Open "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD from MaeArt p, StkArt n Where n.STSKDIS<> 0 and  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & almacen & "'  ", Vgcnx, adOpenStatic, adLockOptimistic
'    frmReferencia.conectar Adodc2, "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD from MaeArt p, StkArt n Where n.STSKDIS<> 0 and  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & almacen & "'  "
'    frmReferencia.Label1.Caption = "Artículos"
'    frmReferencia.show vbmodal
'    Adodc2.Close
'    If vGUtil(1) <> "" Then
'        Text2 = (vGUtil(1))
'    End If
'   If Text2 <> "" Then
'        Command1.SetFocus
'   End If
   VGForm1 = 3
   FormAyuArt1.Show 1
   If Text2 <> "" Then
        Command1.SetFocus
   End If
ElseIf Option2.Value Then
    Adodc2.Open "SELECT FAM_CODIGO,FAM_NOMBRE FROM FAMILIA", VGCNx, adOpenStatic, adLockOptimistic
    frmReferencia.Conectar Adodc2, "SELECT FAM_CODIGO,FAM_NOMBRE FROM FAMILIA"
    frmReferencia.Label1.Caption = "Familias de Artículos"
    frmReferencia.Show vbModal
    Adodc2.Close
    If vGUtil(1) <> "" Then
      Text2 = (vGUtil(1))
    End If
ElseIf Option3.Value Then
        Adodc2.Open "SELECT LIN_CODIGO,LIN_NOMBRE FROM LINEAS Where Fam_Codigo='" & Trim(Text3) & "'", VGCNx, adOpenStatic, adLockOptimistic
        frmReferencia.Conectar Adodc2, "SELECT LIN_CODIGO,LIN_NOMBRE FROM LINEAS Where Fam_Codigo='" & Trim(Text3) & "'"
        frmReferencia.Label1.Caption = "Líneas de Artículos"
        frmReferencia.Show vbModal
        Adodc2.Close
        If vGUtil(1) <> "" Then
            Text2 = (vGUtil(1))
        End If
ElseIf Option4.Value Then
        Adodc2.Open "SELECT GRU_CODIGO,GRU_NOMBRE FROM GRUPO Where Fam_Codigo='" & Trim(Text3) & "' and Lin_Codigo='" & Trim(Text4) & "'", VGCNx, adOpenStatic, adLockOptimistic
        frmReferencia.Conectar Adodc2, "SELECT GRU_CODIGO,GRU_NOMBRE FROM GRUPO Where Fam_Codigo='" & Trim(Text3) & "' and Lin_Codigo='" & Trim(Text4) & "'"
        frmReferencia.Label1.Caption = "Grupos de Artículos"
        frmReferencia.Show vbModal
        Adodc2.Close
        If vGUtil(1) <> "" Then
                Text2 = (vGUtil(1))
        End If
        If Text1 <> "" Then
                 Text2.Enabled = True
                 Text2.SetFocus
        End If
End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then Text2_DblClick
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 And Text2 <> "" Then
     If OpArt.Value Then
        If Existe_cod_art(Text2) <> "" Then
           If Text1 > Text2 Then
                  MsgBox "El codigo fin debe ser mayor que el inicio", vbInformation, mensaje1
                  Exit Sub
           End If
           Command1.SetFocus
        End If
    ElseIf Option2.Value Then
         If Existe(1, Text2, "FAMILIA", "FAM_CODIGO", False) = False Then
             MsgBox "El código de Familia no existe", vbInformation, mensaje1
             Text2.SetFocus: Exit Sub
          Else
            Command1.SetFocus
          End If
    ElseIf Option3.Value Then
        If Existe(1, Text1, "LINEAS", "LIN_CODIGO", False) = False Then
                MsgBox "El código de Línea no existe", vbInformation, mensaje1
                Text1.SetFocus: Exit Sub
         Else
                Command1.SetFocus
         End If
    ElseIf Option4.Value Then
        If Existe(1, Text1, "GRUPO", "GRU_CODIGO", False) = False Then
                MsgBox "El código de Grupo no existe", vbInformation, mensaje1
                Text1.SetFocus: Exit Sub
         Else
                Command1.SetFocus
         End If
     End If
  End If
  If KeyAscii = 13 And Text2 = "" Then
      Command1.SetFocus
  End If
End Sub

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
Label4.Visible = False
Text3.Visible = False
Label5.Visible = False
Text4.Visible = False
End Sub

Private Sub Imprimir2() 'Familia
Dim CADENA As String
Dim Codigo1 As String
Dim Codigo2 As String
Dim rsql As String
Dim tex1 As String, tex2 As String
Dim Va1 As String, Va2 As String

Set Adodc3 = New ADODB.Recordset
If OpTodos.Value Then
    rsql = "Select AFamilia,Fam_Nombre from "
    rsql = rsql & "((MAEART A Inner Join STKART B on A.ACodigo=B.STCodigo) "
    rsql = rsql & "Left Join FAMILIA C on A.AFamilia=C.Fam_Codigo) "
    rsql = rsql & "Where Stalma='" & almacen & "' Order by AFamilia"
    
    Adodc3.Open rsql, VGCNx, adOpenDynamic, adLockOptimistic
    If Adodc3.RecordCount > 0 Then
      If Text1 = "" And Text2 = "" Then
        Adodc3.MoveFirst
        tex1 = IIf(IsNull(Adodc3("AFamilia")), "", Adodc3("AFamilia"))
        Va1 = IIf(IsNull(Adodc3("Fam_Nombre")), "", Adodc3("Fam_Nombre"))
        Adodc3.MoveLast
        tex2 = IIf(IsNull(Adodc3("AFamilia")), "", Adodc3("AFamilia"))
        Va2 = IIf(IsNull(Adodc3("Fam_Nombre")), "", Adodc3("Fam_Nombre"))
      End If
    End If
    Adodc3.Close

    If Check1.Value = 0 Then
        If Chkprecio = 0 Then
                CrystalReport1.WindowTitle = "Inv080 -- Control de Inventarios"
                CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv080.rpt"
        Else
                CrystalReport1.WindowTitle = "Inv070 -- Control de Inventarios"
                CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv070.rpt"
        End If
        CADENA = "{STKART.STALMA}='" & almacen & "' "
        
    Else                                     'Consolidado
        If Chkprecio = 0 Then
                CrystalReport1.WindowTitle = "Inv089 -- Control de Inventarios"
                CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv089.rpt"
        Else
                CrystalReport1.WindowTitle = "Inv067 -- Control de Inventarios"
                CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv067.rpt"
        End If
     End If
     Ubi_Tab CrystalReport1
     CrystalReport1.DiscardSavedData = True
     CrystalReport1.Destination = crptToWindow
     CrystalReport1.SelectionFormula = CADENA
     CrystalReport1.WindowShowPrintBtn = True
     CrystalReport1.WindowShowRefreshBtn = True
     CrystalReport1.WindowShowSearchBtn = True
     CrystalReport1.WindowShowPrintSetupBtn = True
     If Combo2.ListIndex = 1 Then
         CrystalReport1.SortFields(0) = "+{MAEART.ADESCRI}"
     Else
         CrystalReport1.SortFields(0) = "+{STKART.STCODIGO}"
     End If
     CrystalReport1.formulas(0) = "emp = '" & VGparametros.RucEmpresa & "'"
     CrystalReport1.formulas(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
     CrystalReport1.formulas(2) = "campoini = '" & tex1 & "'"
     CrystalReport1.formulas(3) = "campofin = '" & tex2 & "'"
     CrystalReport1.formulas(4) = "detaini = '" & Va1 & "'"
     CrystalReport1.formulas(5) = "detafin = '" & Va2 & "'"
     If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
     Exit Sub
End If
If Trim(Text1) = "" Then
      MsgBox "Ingrese el codigo de la familia", vbExclamation, "Error"
      OpRango = True
      Text1.SetFocus
      Exit Sub
End If
Codigo2 = UCase(Trim(Text2))
Codigo1 = UCase(Trim(Text1))

  rsql = "Select Fam_Nombre from FAMILIA Where Fam_Codigo='" & Codigo1 & "'"
  Adodc3.Open rsql, VGCNx, adOpenDynamic, adLockOptimistic
  If Adodc3.RecordCount = 1 Then
    Va1 = Adodc3("Fam_Nombre")
  End If
  Adodc3.Close
  
  rsql = "Select Fam_Nombre from FAMILIA Where Fam_Codigo='" & Codigo2 & "'"
  Adodc3.Open rsql, VGCNx, adOpenDynamic, adLockOptimistic
  If Adodc3.RecordCount = 1 Then
    Va2 = Adodc3("Fam_Nombre")
  End If
  Adodc3.Close

If Check1.Value = 1 Then
        'CrystalReport1.ReportFileName =  VGParamSistem.RutaReport & "stkxfcon.rpt"
        If Chkprecio = 0 Then
                CrystalReport1.WindowTitle = "Inv089 -- Control de Inventarios"
                CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv089.rpt"
        Else
                CrystalReport1.WindowTitle = "Inv067 -- Control de Inventarios"
                CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv067.rpt"
        End If
        Ubi_Tab CrystalReport1
        If Text2 <> "" Then
                CADENA = " ({MAEART.AFAMILIA} in '" & Codigo1 & "' to '" & Codigo2 & "')  and {STKART.STCODIGO}={MAEART.ACODIGO}"
        Else
                Codigo2 = Codigo1: Va2 = Va1
                CADENA = "({MAEART.AFAMILIA} in '" & Codigo1 & "' to '" & Codigo2 & "')  and {STKART.STCODIGO}={MAEART.ACODIGO} "
        End If
Else
        '.ReportFileName =  VGParamSistem.RutaReport & "stkxfam.rpt"
        If Chkprecio = 0 Then
                CrystalReport1.WindowTitle = "Inv080 -- Control de Inventarios"
                CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv080.rpt"
        Else
                CrystalReport1.WindowTitle = "Inv070 -- Control de Inventarios"
                CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv070.rpt"
        End If
        Ubi_Tab CrystalReport1
        If Text2 <> "" Then
                CADENA = "{STKART.STALMA}='" & almacen & "'  and ({MAEART.AFAMILIA} in '" & Codigo1 & "' to '" & Codigo2 & "')  and {STKART.STCODIGO} = {MAEART.ACODIGO} "
        Else
                Codigo2 = Codigo1: Va2 = Va1
                CADENA = "{STKART.STALMA}='" & almacen & "'  and ({MAEART.AFAMILIA} in '" & Codigo1 & "' to '" & Codigo2 & "')  and {STKART.STCODIGO} = {MAEART.ACODIGO} "
        End If
End If
CrystalReport1.DiscardSavedData = True
CrystalReport1.Destination = crptToWindow
CrystalReport1.SelectionFormula = CADENA
CrystalReport1.WindowShowPrintBtn = True
CrystalReport1.WindowShowRefreshBtn = True
CrystalReport1.WindowShowSearchBtn = True
CrystalReport1.WindowShowPrintSetupBtn = True
If Combo2.ListIndex = 1 Then
      CrystalReport1.SortFields(0) = "+{MAEART.ADESCRI}"
Else
      CrystalReport1.SortFields(0) = "+{STKART.STCODIGO}"
End If
CrystalReport1.formulas(0) = "emp = '" & VGparametros.RucEmpresa & "'"
CrystalReport1.formulas(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
CrystalReport1.formulas(2) = "campoini = '" & Codigo1 & "'"
CrystalReport1.formulas(3) = "campofin = '" & Codigo2 & "'"
CrystalReport1.formulas(4) = "detaini = '" & Va1 & "'"
CrystalReport1.formulas(5) = "detafin = '" & Va2 & "'"
If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
End Sub

Private Sub Imprimir3()
Dim CADENA As String
Dim Codigo1 As String
Dim Codigo2 As String
Dim rsql As String
Dim tex1 As String, tex2 As String
Dim Va1 As String, Va2 As String

If Trim(Text3) = "" Then
      MsgBox "Ingrese el código de la familia", vbExclamation, "Error"
      Text3.SetFocus
      Exit Sub
End If

Set Adodc3 = New ADODB.Recordset  'Para sacar la descripcion del rango elegido

If OpTodos.Value Then
    rsql = "Select AModelo,Lin_Nombre from "
    rsql = rsql & "((MAEART A Inner Join STKART B on A.ACodigo=B.STCodigo) "
    rsql = rsql & "Left Join LINEAS C on A.AFamilia=C.Fam_Codigo and A.Amodelo=C.Lin_Codigo) "
    rsql = rsql & "Where AFamilia='" & Text3.text & "' and Stalma='" & almacen & "' Order By Amodelo"
    
    Adodc3.Open rsql, VGCNx, adOpenDynamic, adLockOptimistic
    If Adodc3.RecordCount > 0 Then
      If Text1 = "" And Text2 = "" Then
        Adodc3.MoveFirst
        tex1 = IIf(IsNull(Adodc3("AModelo")), "", Adodc3("AModelo"))
        Va1 = IIf(IsNull(Adodc3("Lin_Nombre")), "", Adodc3("Lin_Nombre"))
        Adodc3.MoveLast
        tex2 = IIf(IsNull(Adodc3("AModelo")), "", Adodc3("AModelo"))
        Va2 = IIf(IsNull(Adodc3("Lin_Nombre")), "", Adodc3("Lin_Nombre"))
      End If
    End If
    Adodc3.Close
    
    If Check1.Value = 0 Then
        If Chkprecio = 0 Then
                CrystalReport1.WindowTitle = "Inv085 -- Control de Inventarios"
                CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv085.rpt"
        Else
                CrystalReport1.WindowTitle = "Inv072 -- Control de Inventarios"
                CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv072.rpt"
        End If
        CADENA = "{STKART.STALMA}='" & almacen & "' and {MAEART.AFAMILIA}='" & Text3.text & "'"
        
    Else                                     'Consolidado
        If Chkprecio = 0 Then
                CrystalReport1.WindowTitle = "Inv084 -- Control de Inventarios"
                CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv084.rpt"
        Else
                CrystalReport1.WindowTitle = "Inv069 -- Control de Inventarios"
                CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv069.rpt"
        End If
       
        CADENA = "{MAEART.AFAMILIA}='" & Text3.text & "'"
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
     CrystalReport1.formulas(2) = "campoini = '" & tex1 & "'"
     CrystalReport1.formulas(3) = "campofin = '" & tex2 & "'"
     CrystalReport1.formulas(4) = "detaini = '" & Va1 & "'"
     CrystalReport1.formulas(5) = "detafin = '" & Va2 & "'"
     If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
     Exit Sub
End If
If Trim(Text1) = "" Then
      MsgBox "Ingrese el codigo de la Línea", vbExclamation, "Error"
      OpRango = True
      Text1.SetFocus
      Exit Sub
End If
Codigo2 = UCase(Trim(Text2))
Codigo1 = UCase(Trim(Text1))
  
  rsql = "Select Lin_Nombre from LINEAS Where Lin_Codigo='" & Codigo1 & "'"
  Adodc3.Open rsql, VGCNx, adOpenDynamic, adLockOptimistic
  If Adodc3.RecordCount = 1 Then
    Va1 = Adodc3("Lin_Nombre")
  End If
  Adodc3.Close
  
  rsql = "Select Lin_Nombre from LINEAS Where Lin_Codigo='" & Codigo2 & "'"
  Adodc3.Open rsql, VGCNx, adOpenDynamic, adLockOptimistic
  If Adodc3.RecordCount = 1 Then
    Va2 = Adodc3("Lin_Nombre")
  End If
  Adodc3.Close

If Check1.Value = 1 Then
        If Chkprecio = 0 Then
              CrystalReport1.WindowTitle = "Inv084 -- Control de Inventarios"
              CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv084.rpt"
        Else
              CrystalReport1.WindowTitle = "Inv069 -- Control de Inventarios"
              CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv069.rpt"
        End If
        Ubi_Tab CrystalReport1
        If Text2 <> "" Then
                CADENA = "({MAEART.AMODELO} in '" & Codigo1 & "' to '" & Codigo2 & "')  and {STKART.STCODIGO}={MAEART.ACODIGO} and {MAEART.AFAMILIA} = '" & Text3.text & "'"
        Else
                Codigo2 = Codigo1: Va2 = Va1
                CADENA = "({MAEART.AMODELO} in '" & Codigo1 & "' to '" & Codigo2 & "')  and {STKART.STCODIGO}={MAEART.ACODIGO} and {MAEART.AFAMILIA} = '" & Text3.text & "'"
        End If
Else
        If Chkprecio = 0 Then
                CrystalReport1.WindowTitle = "Inv085 -- Control de Inventarios"
                CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv085.rpt"
        Else
                CrystalReport1.WindowTitle = "Inv072 -- Control de Inventarios"
                CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv072.rpt"
        End If
        Ubi_Tab CrystalReport1
        If Text2 <> "" Then
                CADENA = "{STKART.STALMA}='" & almacen & "'  and ({MAEART.AMODELO} in '" & Codigo1 & "' to '" & Codigo2 & "')  and {MAEART.AFAMILIA} = '" & Text3.text & "'"
        Else
                Codigo2 = Codigo1: Va2 = Va1
                CADENA = "{STKART.STALMA}='" & almacen & "'  and ({MAEART.AMODELO} in '" & Codigo1 & "' to '" & Codigo2 & "')  and {MAEART.AFAMILIA} = '" & Text3.text & "'"
        End If
End If
CrystalReport1.DiscardSavedData = True
CrystalReport1.Destination = crptToWindow
CrystalReport1.SelectionFormula = CADENA
CrystalReport1.WindowShowPrintBtn = True
CrystalReport1.WindowShowRefreshBtn = True
CrystalReport1.WindowShowSearchBtn = True
CrystalReport1.WindowShowPrintSetupBtn = True
If Combo2.ListIndex = 1 Then CrystalReport1.SortFields(0) = "+{LINEA.NOMBRERE}"
CrystalReport1.formulas(0) = "emp = '" & VGparametros.RucEmpresa & "'"
CrystalReport1.formulas(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
CrystalReport1.formulas(2) = "campoini = '" & Codigo1 & "'"
CrystalReport1.formulas(3) = "campofin = '" & Codigo2 & "'"
CrystalReport1.formulas(4) = "detaini = '" & Va1 & "'"
CrystalReport1.formulas(5) = "detafin = '" & Va2 & "'"
If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
End Sub

Private Sub Imprimir4()
Dim CADENA As String
Dim Codigo1 As String
Dim Codigo2 As String
Dim rsql As String
Dim tex1 As String, tex2 As String
Dim Va1 As String, Va2 As String

If Trim(Text3) = "" Then
      MsgBox "Ingrese el código de la familia", vbExclamation, "Error"
      Text3.SetFocus
      Exit Sub
ElseIf Trim(Text4) = "" Then
      MsgBox "Ingrese el código de la Línea", vbExclamation, "Error"
      Text4.SetFocus
      Exit Sub
End If
Set Adodc3 = New ADODB.Recordset  'Para sacar la descripcion del rango elegido

If OpTodos.Value Then
    rsql = "Select AGrupo,Acodigo,Gru_Nombre from"
    rsql = rsql & "((MAEART A Inner Join STKART B on A.ACodigo=B.STCodigo)"
    rsql = rsql & "Left Join GRUPO C on A.AFamilia=C.Fam_Codigo and A.Amodelo=C.Lin_Codigo and A.Agrupo=C.Gru_Codigo) "
    rsql = rsql & "Where AFamilia='" & Text3.text & "' and Amodelo='" & Text4.text & "' and Stalma='" & almacen & "' Order by Agrupo"

    Adodc3.Open rsql, VGCNx, adOpenDynamic, adLockOptimistic
    If Adodc3.RecordCount > 0 Then
      If Text1 = "" And Text2 = "" Then
        Adodc3.MoveFirst
        tex1 = IIf(IsNull(Adodc3("AGrupo")), "", Adodc3("AGrupo"))
        Va1 = IIf(IsNull(Adodc3("Gru_Nombre")), "", Adodc3("Gru_Nombre"))
        Adodc3.MoveLast
        tex2 = IIf(IsNull(Adodc3("AGrupo")), "", Adodc3("AGrupo"))
        Va2 = IIf(IsNull(Adodc3("Gru_Nombre")), "", Adodc3("Gru_Nombre"))
      End If
    End If
    Adodc3.Close
    
    If Check1.Value = 0 Then
        If Chkprecio = 0 Then
                CrystalReport1.WindowTitle = "Inv083 -- Control de Inventarios"
                CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv083.rpt"
        Else
                CrystalReport1.WindowTitle = "Inv071 -- Control de Inventarios"
                CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv071.rpt"
        End If
        CADENA = "{STKART.STALMA}='" & almacen & "' and {MAEART.AFAMILIA}='" & Text3.text & "' and {MAEART.AMODELO}='" & Text4.text & "'"
        
    Else                                     'Consolidado
        If Chkprecio = 0 Then
                CrystalReport1.WindowTitle = "Inv082 -- Control de Inventarios"
                CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv082.rpt"
        Else
                CrystalReport1.WindowTitle = "Inv068 -- Control de Inventarios"
                CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv068.rpt"
        End If
        CADENA = "{MAEART.AFAMILIA}='" & Text3.text & "' and {MAEART.AMODELO}='" & Text4.text & "'"
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
     CrystalReport1.formulas(2) = "campoini = '" & tex1 & "'"
     CrystalReport1.formulas(3) = "campofin = '" & tex2 & "'"
     CrystalReport1.formulas(4) = "detaini = '" & Va1 & "'"
     CrystalReport1.formulas(5) = "detafin = '" & Va2 & "'"
     If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
     Exit Sub
End If
If Trim(Text1) = "" Then
      MsgBox "Ingrese el código del artículo", vbExclamation, "Error"
      OpRango = True
      Text1.SetFocus
      Exit Sub
End If
Codigo2 = UCase(Trim(Text2))
Codigo1 = UCase(Trim(Text1))

  rsql = "Select Gru_Nombre from GRUPO Where Gru_Codigo='" & Codigo1 & "'"
  Adodc3.Open rsql, VGCNx, adOpenDynamic, adLockOptimistic
  If Adodc3.RecordCount = 1 Then
    Va1 = Adodc3("Gru_Nombre")
  End If
  Adodc3.Close
  
  rsql = "Select Gru_Nombre from Grupo Where Gru_Codigo='" & Codigo2 & "'"
  Adodc3.Open rsql, VGCNx, adOpenDynamic, adLockOptimistic
  If Adodc3.RecordCount = 1 Then
    Va2 = Adodc3("Gru_Nombre")
  End If
  Adodc3.Close

If Check1.Value = 1 Then
        If Chkprecio = 0 Then
              CrystalReport1.WindowTitle = "Inv082 -- Control de Inventarios"
              CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv082.rpt"
        Else
              CrystalReport1.WindowTitle = "Inv068 -- Control de Inventarios"
              CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv068.rpt"
        End If
        Ubi_Tab CrystalReport1
        If Text2 <> "" Then
                CADENA = " ({MAEART.AGRUPO} in '" & Codigo1 & "' to '" & Codigo2 & "')  and {STKART.STCODIGO}={MAEART.ACODIGO}"
                CADENA = CADENA & " and {MAEART.AFAMILIA}='" & Text3.text & "' and {MAEART.AMODELO}='" & Text4.text & "'"
        Else
                Codigo2 = Codigo1: Va2 = Va1
                CADENA = "({MAEART.AGRUPO} in '" & Codigo1 & "' to '" & Codigo2 & "')  and {STKART.STCODIGO}={MAEART.ACODIGO}"
                CADENA = CADENA & " and {MAEART.AFAMILIA}='" & Text3.text & "' and {MAEART.AMODELO}='" & Text4.text & "'"
        End If
Else
        If Chkprecio = 0 Then
                CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv083.rpt"
        Else
                CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv071.rpt"
        End If
        Ubi_Tab CrystalReport1
        If Text2 <> "" Then
                CADENA = "{STKART.STALMA}='" & almacen & "'  and ({MAEART.AGRUPO} in '" & Codigo1 & "' to '" & Codigo2 & "')"
                CADENA = CADENA & " and {MAEART.AFAMILIA}='" & Text3.text & "' and {MAEART.AMODELO}='" & Text4.text & "'"
        Else
                Codigo2 = Codigo1: Va2 = Va1
                CADENA = "{STKART.STALMA}='" & almacen & "'  and ({MAEART.AGRUPO} in '" & Codigo1 & "' to '" & Codigo2 & "')"
                CADENA = CADENA & " and {MAEART.AFAMILIA}='" & Text3.text & "' and {MAEART.AMODELO}='" & Text4.text & "'"
        End If
End If
CrystalReport1.DiscardSavedData = True
CrystalReport1.Destination = crptToWindow
CrystalReport1.SelectionFormula = CADENA
CrystalReport1.WindowShowPrintBtn = True
CrystalReport1.WindowShowRefreshBtn = True
CrystalReport1.WindowShowSearchBtn = True
CrystalReport1.WindowShowPrintSetupBtn = True
CrystalReport1.formulas(0) = "emp = '" & VGparametros.RucEmpresa & "'"
CrystalReport1.formulas(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
CrystalReport1.formulas(2) = "campoini = '" & Codigo1 & "'"
CrystalReport1.formulas(3) = "campofin = '" & Codigo2 & "'"
CrystalReport1.formulas(4) = "detaini = '" & Va1 & "'"
CrystalReport1.formulas(5) = "detafin = '" & Va2 & "'"
If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
End Sub

Private Sub Text3_DblClick()
Dim Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset
If Option3.Value Or Option4.Value Then
         Adodc2.Open "Select FAM_CODIGO,FAM_NOMBRE,FAM_CTA from FAMILIA ", VGCNx, adOpenStatic, adLockOptimistic
         frmReferencia.Conectar Adodc2, "Select FAM_CODIGO,FAM_NOMBRE,FAM_CTA from FAMILIA "
         frmReferencia.Label1.Caption = "Familias"
         frmReferencia.Show vbModal
         Adodc2.Close
         If vGUtil(1) <> "" Then
                 Text3 = (vGUtil(1))
         End If
End If
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then Text3_DblClick
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 And Text3 <> "" Then
     If Text4.Visible = True Then Text4.SetFocus Else OpTodos.SetFocus
  End If
End Sub

Private Sub Text4_DblClick()
Dim Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset
If Option4.Value Then
         Adodc2.Open "Select LIN_CODIGO,LIN_NOMBRE from LINEAS WHERE FAM_CODIGO ='" & Text3 & "' ", VGCNx, adOpenStatic, adLockOptimistic
         frmReferencia.Conectar Adodc2, "Select LIN_CODIGO,LIN_NOMBRE from LINEAS WHERE FAM_CODIGO ='" & Text3 & "' "
         frmReferencia.Label1.Caption = "Líneas"
         frmReferencia.Show vbModal
         Adodc2.Close
         If vGUtil(1) <> "" Then
            Text4 = (vGUtil(1))
         End If
End If
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then Text4_DblClick
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 And Text4 <> "" Then
     OpTodos.SetFocus
  End If
End Sub

Private Sub Imprimir5()
Dim Codigo1 As String
Dim Codigo2 As String
Dim CADENA As String
Dim rsql As String
Dim tex1 As String, tex2 As String
Dim Va1 As String, Va2 As String

Codigo1 = UCase(Trim(Text1))
Set Adodc3 = New ADODB.Recordset

If OpTodos.Value Then
    rsql = "Select ACodigo,Adescri from "
    rsql = rsql & "MAEART A Inner Join STKART B on A.ACodigo=B.STCodigo "
    rsql = rsql & "Where Stalma='" & almacen & "' Order by Acodigo"
    
    Adodc3.Open rsql, VGCNx, adOpenDynamic, adLockOptimistic
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

    If Check1.Value = 0 Then
            If ChkSerie = 1 Then
                    CrystalReport1.WindowTitle = "Inv136 -- Control de Inventarios"
                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv136.rpt"
            ElseIf Chkprecio = 0 Then
                    CrystalReport1.WindowTitle = "Inv138 -- Control de Inventarios"
                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv138.rpt"
            Else
                    CrystalReport1.WindowTitle = "Inv140-- Control de Inventarios"
                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv140.rpt"
            End If
            CADENA = "{STKART.STALMA}='" & almacen & "' "
    Else                                     'Consolidado
            If ChkSerie = 1 Then
                    CrystalReport1.WindowTitle = "Inv137 -- Control de Inventarios"
                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv137.rpt"
            ElseIf Chkprecio = 0 Then
                    CrystalReport1.WindowTitle = "Inv139 -- Control de Inventarios"
                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv139.rpt"
            Else
                    CrystalReport1.WindowTitle = "Inv140 -- Control de Inventarios"
                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv141.rpt"
            End If
    End If
    Ubi_Tab CrystalReport1
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.SelectionFormula = CADENA
    If Combo2.ListIndex = 1 Then
      CrystalReport1.SortFields(0) = "+{MAEART.ADESCRI}"
    Else
      CrystalReport1.SortFields(0) = "+{STKART.STCODIGO}"
    End If
    CrystalReport1.formulas(0) = "emp = '" & VGparametros.RucEmpresa & "'"
    CrystalReport1.formulas(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
    CrystalReport1.formulas(2) = "campoini = '" & tex1 & "'"
    CrystalReport1.formulas(3) = "campofin = '" & tex2 & "'"
    CrystalReport1.formulas(4) = "detaini = '" & Va1 & "'"
    CrystalReport1.formulas(5) = "detafin = '" & Va2 & "'"
    If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
    Exit Sub
End If

If Text1 = "" Then
        MsgBox "Ingrese el codigo", vbExclamation, "Error"
        Text1.SetFocus
        Exit Sub
End If
  Codigo2 = UCase(Trim(Text2))
  rsql = "Select ADescri from MAEART Where ACodigo='" & Codigo1 & "'"
  Adodc3.Open rsql, VGCNx, adOpenDynamic, adLockOptimistic
  If Adodc3.RecordCount = 1 Then
    Va1 = Adodc3("Adescri")
  End If
  Adodc3.Close
  
  rsql = "Select ADescri from MAEART Where ACodigo='" & Codigo2 & "'"
  Adodc3.Open rsql, VGCNx, adOpenDynamic, adLockOptimistic
  If Adodc3.RecordCount = 1 Then
    Va2 = Adodc3("Adescri")
  End If
  Adodc3.Close

If Option1.Value Then           'Un select
    If Check1.Value = 1 Then
            If ChkSerie = 1 Then
                    CrystalReport1.WindowTitle = "Inv137 -- Control de Inventarios"
                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv137.rpt"
            ElseIf Chkprecio = 0 Then
                    CrystalReport1.WindowTitle = "Inv139 -- Control de Inventarios"
                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv139.rpt"
            Else
                    CrystalReport1.WindowTitle = "Inv141 -- Control de Inventarios"
                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv141.rpt"
            End If
            If Text2 <> "" Then
                    Codigo2 = Text2
                    CADENA = "({STKART.STCODIGO} in '" & Codigo1 & "' to '" & Codigo2 & "')"
            Else
                    Codigo2 = Codigo1: Va2 = Va1
                    CADENA = "{STKART.STCODIGO} = '" & Codigo1 & "' "
            End If
    Else
            If ChkSerie = 1 Then
                    CrystalReport1.WindowTitle = "Inv136 -- Control de Inventarios"
                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv136.rpt"
            ElseIf Chkprecio = 0 Then
                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv138.rpt"
                    CrystalReport1.WindowTitle = "Inv138 -- Control de Inventarios"
            Else
                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv040.rpt"
                    CrystalReport1.WindowTitle = "Inv140 -- Control de Inventarios"
            End If
            If Text2 <> "" Then
                    Codigo2 = Text2         '  "23134671"
                    CADENA = " {STKART.STALMA}='" & almacen & "' and ({STKART.STCODIGO} in '" & Codigo1 & "' to '" & Codigo2 & "')"
            Else
                    Codigo2 = Codigo1: Va2 = Va1
                    CADENA = "{STKART.STALMA}='" & almacen & "' and {STKART.STCODIGO} = '" & Codigo1 & "' "
            End If
    End If
    Ubi_Tab CrystalReport1
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.SelectionFormula = CADENA
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    If Combo2.ListIndex = 1 Then
      CrystalReport1.SortFields(0) = "+{MAEART.ADESCRI}"
    Else
      CrystalReport1.SortFields(0) = "+{STKART.STCODIGO}"
    End If
    CrystalReport1.formulas(0) = "emp = '" & VGparametros.RucEmpresa & "'"
    CrystalReport1.formulas(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
    CrystalReport1.formulas(2) = "campoini = '" & Codigo1 & "'"
    CrystalReport1.formulas(3) = "campofin = '" & Codigo2 & "'"
    CrystalReport1.formulas(4) = "detaini = '" & Va1 & "'"
    CrystalReport1.formulas(5) = "detafin = '" & Va2 & "'"
    If CrystalReport1.Status <> 2 Then
        CrystalReport1.Action = 1
    End If
End If
End Sub

