VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRentabilidad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rentabildad"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   Icon            =   "frmRentabilidad.frx":0000
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrameRep 
      Height          =   2055
      Left            =   285
      TabIndex        =   22
      Top             =   990
      Width           =   2985
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1410
         TabIndex        =   12
         Top             =   1635
         Width           =   1275
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1410
         MaxLength       =   20
         TabIndex        =   11
         Top             =   1320
         Width           =   1275
      End
      Begin VB.OptionButton OpTodos 
         Caption         =   "Todos los Artículos"
         Height          =   240
         Left            =   150
         TabIndex        =   9
         Top             =   795
         Value           =   -1  'True
         Width           =   1785
      End
      Begin VB.OptionButton OpRango 
         Caption         =   "Rango"
         Height          =   255
         Left            =   150
         TabIndex        =   10
         Top             =   1035
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1125
         MaxLength       =   20
         TabIndex        =   7
         Top             =   210
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1125
         MaxLength       =   20
         TabIndex        =   8
         Top             =   510
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.Label Label6 
         Caption         =   "Fin"
         Height          =   255
         Left            =   555
         TabIndex        =   26
         Top             =   1635
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Inicio "
         Height          =   255
         Left            =   555
         TabIndex        =   25
         Top             =   1335
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Familia"
         Height          =   225
         Left            =   150
         TabIndex        =   24
         Top             =   225
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Línea"
         Height          =   225
         Left            =   150
         TabIndex        =   23
         Top             =   525
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   735
      Left            =   2400
      Picture         =   "frmRentabilidad.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3330
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   735
      Left            =   3738
      Picture         =   "frmRentabilidad.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3330
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccionar"
      Height          =   705
      Left            =   285
      TabIndex        =   0
      Top             =   135
      Width           =   6270
      Begin VB.OptionButton Option1 
         Caption         =   "Vendedor"
         Height          =   255
         Index           =   5
         Left            =   5145
         TabIndex        =   6
         Top             =   270
         Width           =   1020
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Grupo"
         Height          =   255
         Index           =   4
         Left            =   1065
         TabIndex        =   2
         Top             =   270
         Width           =   1020
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Línea"
         Height          =   255
         Index           =   3
         Left            =   2085
         TabIndex        =   3
         Top             =   270
         Width           =   1020
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Familia"
         Height          =   255
         Index           =   2
         Left            =   3105
         TabIndex        =   4
         Top             =   270
         Width           =   1020
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Clientes"
         Height          =   255
         Index           =   1
         Left            =   4125
         TabIndex        =   5
         Top             =   270
         Width           =   1020
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Articulo"
         Height          =   255
         Index           =   0
         Left            =   45
         TabIndex        =   1
         Top             =   270
         Width           =   1020
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   285
      Top             =   3270
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      Height          =   2085
      Left            =   3555
      TabIndex        =   19
      Top             =   975
      Width           =   2985
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1455
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   705
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   570
         Width           =   1980
      End
      Begin VB.Label Label2 
         Caption         =   "Unidad de Medida"
         Height          =   270
         Left            =   165
         TabIndex        =   21
         Top             =   1125
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "Lista de Precios"
         Height          =   240
         Left            =   180
         TabIndex        =   20
         Top             =   255
         Width           =   1740
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Fecha"
      Height          =   2085
      Left            =   3555
      TabIndex        =   27
      Top             =   975
      Visible         =   0   'False
      Width           =   2985
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   1185
         TabIndex        =   16
         Top             =   1425
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   24576001
         CurrentDate     =   36760
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   1125
         TabIndex        =   15
         Top             =   600
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   529
         _Version        =   393216
         Format          =   24576001
         CurrentDate     =   36760
      End
      Begin VB.Label Label8 
         Caption         =   "Hasta"
         Height          =   255
         Left            =   390
         TabIndex        =   29
         Top             =   1140
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Desde"
         Height          =   225
         Left            =   375
         TabIndex        =   28
         Top             =   330
         Width           =   660
      End
   End
End
Attribute VB_Name = "frmRentabilidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim List As String

Private Sub Combo1_Click()
List = Mid(Combo1, 1, 2)
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Combo2.SetFocus
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command1.SetFocus
End Sub

Private Sub Command1_Click()
If Text1 <> "" And Text2 <> "" And Text1 > Text2 Then
  MsgBox "Ingrese un código menor al fin ", vbOKOnly, "Error"
  Exit Sub
End If
If Option1(1).Value Or Option1(5).Value Then
  If DTPicker1.Value > DTPicker2.Value Then
    MsgBox "Ingrese una fecha menor al fin ", vbOKOnly, "Error"
    Exit Sub
  End If
End If

If Option1(0).Value Then
    Imprimir  'Articulos
ElseIf Option1(1).Value Then
    Imprimir1  'clientes
ElseIf Option1(2).Value Then
    Imprimir2   'Familia
ElseIf Option1(3).Value Then
    Imprimir3   'Linea
ElseIf Option1(4).Value Then
    Imprimir4   'Grupo
ElseIf Option1(5).Value Then
    Imprimir5   'Vendedor
End If
End Sub

Private Sub Command8_Click()
  Unload Me
End Sub

Private Sub Imprimir()
Dim cadena As String
Dim Sql As String, Sql1 As String
Dim Codigo1 As String, Codigo2 As String
Dim tex1 As String, tex2 As String
Dim Va1 As String, Va2 As String
Screen.MousePointer = 11

Codigo1 = UCase(Trim(Text1))
Codigo2 = UCase(Trim(Text2))
Set Adodc1 = New ADODB.Recordset
Set Adodc3 = New ADODB.Recordset

cadena = "{LISTA_PRECIOS.Uni_LisPre}='" & Mid(Combo2.text, 1, 4) & "'"
cadena = cadena & " and {LISTA_PRECIOS.Cod_LisPre}='" & List & "' and {STKART.STALMA}='" & VGAlma & "'"

Sql = "Select ACodigo,Adescri from "
Sql = Sql & "((MAEART A Inner Join STKART B on A.ACodigo=B.STCodigo) "
Sql = Sql & "Left Join LISTA_PRECIOS C on A.ACODIGO=C.COD_ARTI) "
Sql = Sql & "Where Stalma='" & VGAlma & "' AND UNI_LISPRE='" & Mid(Combo2.text, 1, 4) & "' And COD_LISPRE='" & List & "'"

If OpTodos.Value Then
    Sql = Sql & " Order by Acodigo"
    
    Adodc3.Open Sql, cConexCom, adOpenDynamic, adLockOptimistic
    If Adodc3.RecordCount > 0 Then
      If Text1 = "" And Text2 = "" Then
        Adodc3.MoveFirst
        tex1 = IIf(IsNull(Adodc3("ACodigo")), "", Adodc3("ACodigo"))
        Va1 = IIf(IsNull(Adodc3("ADescri")), "", Adodc3("Adescri"))
        Adodc3.MoveLast
        tex2 = IIf(IsNull(Adodc3("ACodigo")), "", Adodc3("ACodigo"))
        Va2 = IIf(IsNull(Adodc3("ADescri")), "", Adodc3("ADescri"))
      End If
    Else
      MsgBox "No existen artículos para esta unidad de medida", vbOKOnly, "Aviso"
      Adodc3.Close
      Screen.MousePointer = 1
      Exit Sub
    End If
    Adodc3.Close
    
    CrystalReport1.ReportFileName = RUTA & "reporteInv\rentabil.rpt"
    Ubi_Tab CrystalReport1
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.SelectionFormula = cadena
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.Formulas(0) = "emp = '" & VGNemp & "'"
    CrystalReport1.Formulas(1) = "Hora = '" & Format(Time, "hh:mm:ss") & "'"
    CrystalReport1.Formulas(2) = "campoini = '" & tex1 & "'"
    CrystalReport1.Formulas(3) = "campofin = '" & tex2 & "'"
    CrystalReport1.Formulas(4) = "detaini = '" & Va1 & "'"
    CrystalReport1.Formulas(5) = "detafin = '" & Va2 & "'"
   CrystalReport1.Destination = crptToWindow
   CrystalReport1.WindowTitle = mensaje1
   If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
   Screen.MousePointer = 1
   Exit Sub
End If
''''''''''''''''''''''''''
If Text1 = "" Then
    MsgBox "Ingrese el codigo", vbExclamation, "Error"
    Text1.SetFocus
    Exit Sub
End If

If Option1(0).Value Then
  If Text2 <> "" Then
      Codigo2 = Text2         '  "23134671"
      cadena = cadena & " and ({STKART.STCODIGO} in '" & Codigo1 & "' to '" & Codigo2 & "')"
      
      Sql = Sql & " and STCODIGO between '" & Codigo1 & "' and '" & Codigo2 & "'"
  Else
      Codigo2 = Codigo1: Va2 = Va1
      cadena = cadena & " and {STKART.STCODIGO} = '" & Codigo1 & "' "
      
      Sql = Sql & " and STCODIGO ='" & Codigo1 & "'"
  End If
    
  Adodc1.Open Sql, cConexCom, adOpenDynamic, adLockOptimistic
  If Adodc1.RecordCount = 0 Then
    MsgBox "     No existen artículos en este rango    ", vbOKOnly, "Aviso"
    Adodc1.Close
    Screen.MousePointer = 1
    Exit Sub
  End If
    
  Adodc3.Open "Select ADescri from MAEART Where ACodigo='" & Codigo1 & "'", cConexCom, adOpenDynamic, adLockOptimistic
  If Adodc3.RecordCount = 1 Then Va1 = Adodc3("Adescri")
  Adodc3.Close
  
  Adodc3.Open "Select ADescri from MAEART Where ACodigo='" & Codigo2 & "'", cConexCom, adOpenDynamic, adLockOptimistic
  If Adodc3.RecordCount = 1 Then Va2 = Adodc3("Adescri")
  Adodc3.Close
    
    CrystalReport1.ReportFileName = RUTA & "reporteInv\rentabil.rpt"
    Ubi_Tab CrystalReport1
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.SelectionFormula = cadena
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.Formulas(0) = "emp = '" & VGNemp & "'"
    CrystalReport1.Formulas(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
    CrystalReport1.Formulas(2) = "campoini = '" & Codigo1 & "'"
    CrystalReport1.Formulas(3) = "campofin = '" & Codigo2 & "'"
    CrystalReport1.Formulas(4) = "detaini = '" & Va1 & "'"
    CrystalReport1.Formulas(5) = "detafin = '" & Va2 & "'"
    If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
End If
Screen.MousePointer = 1
End Sub

Private Sub Imprimir1()
Dim cadena As String
Dim Sql As String, Sql1 As String
Dim Codigo1 As String, Codigo2 As String
Dim tex1 As String, tex2 As String
Dim Va1 As String, Va2 As String

Codigo1 = UCase(Trim(Text1))
Codigo2 = UCase(Trim(Text2))
Set Adodc1 = New ADODB.Recordset
Set Adodc3 = New ADODB.Recordset

Sql1 = "Select DFCODIGO from FACDET A Left Join FACCAB B on A.DFTD=B.CFTD And A.DFNUMSER=B.CFNUMSER And A.DFNUMDOC=B.CFNUMDOC "
Sql1 = Sql1 & "Where (CFFECDOC between #" & Format(DTPicker1, "mm/dd/yyyy") & "# And #" & Format(DTPicker2, "mm/dd/yyyy") & "#) "

cadena = "{FACCAB.CFFECDOC} in date(" & Format(DTPicker1, "yyyy") & "," & Format(DTPicker1, "mm") & "," & Format(DTPicker1, "dd") & " ) "
cadena = cadena & "to date(" & Format(DTPicker2, "yyyy") & "," & Format(DTPicker2, "mm") & "," & Format(DTPicker2, "dd") & " )"

If OpTodos.Value Then
    Adodc1.Open Sql1, cConexCom, adOpenDynamic, adLockOptimistic
    If Adodc1.RecordCount = 0 Then
      MsgBox "No existen Clientes para este rango de fecha", vbOKOnly, "Aviso"
      Adodc1.Close
      Screen.MousePointer = 1
      Exit Sub
    End If

    Sql = "Select CCODCLI,CNOMCLI from "
    Sql = Sql & "MAECLI A Inner Join FACCAB B on A.CCODCLI=B.CFCODCLI "
    Sql = Sql & "Order by CCODCLI"
    
    Adodc3.Open Sql, cConexCom, adOpenDynamic, adLockOptimistic
    If Adodc3.RecordCount > 0 Then
      If Text1 = "" And Text2 = "" Then
        Adodc3.MoveFirst
        tex1 = IIf(IsNull(Adodc3("CCODCLI")), "", Adodc3("CCODCLI"))
        Va1 = IIf(IsNull(Adodc3("CNOMCLI")), "", Adodc3("CNOMCLI"))
        Adodc3.MoveLast
        tex2 = IIf(IsNull(Adodc3("CCODCLI")), "", Adodc3("CCODCLI"))
        Va2 = IIf(IsNull(Adodc3("CNOMCLI")), "", Adodc3("CNOMCLI"))
      End If
    End If
    Adodc3.Close
    
    CrystalReport1.ReportFileName = RUTA & "reporteInv\rentaC.rpt"
    Ubi_Tab CrystalReport1
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.SelectionFormula = cadena
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.Formulas(0) = "emp = '" & VGNemp & "'"
    CrystalReport1.Formulas(1) = "Hora = '" & Format(Time, "hh:mm:ss") & "'"
    CrystalReport1.Formulas(2) = "campoini = '" & tex1 & "'"
    CrystalReport1.Formulas(3) = "campofin = '" & tex2 & "'"
    CrystalReport1.Formulas(4) = "detaini = '" & Va1 & "'"
    CrystalReport1.Formulas(5) = "detafin = '" & Va2 & "'"
    CrystalReport1.Formulas(6) = "FECini = '" & DTPicker1.Value & "'"
    CrystalReport1.Formulas(7) = "FECfin = '" & DTPicker2.Value & "'"
   CrystalReport1.Destination = crptToWindow
   CrystalReport1.WindowTitle = mensaje1
   If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
   Exit Sub
End If

If Text1 = "" Then
        MsgBox "Ingrese el codigo", vbExclamation, "Error"
        Text1.SetFocus
        Exit Sub
End If

If Option1(1).Value Then
  If Text2 <> "" Then
      Codigo2 = Text2         '  "23134671"
      cadena = cadena & " and ({FACCAB.CFCODCLI} in '" & Codigo1 & "' to '" & Codigo2 & "')"
      
      Sql1 = Sql1 & " AND CFCODCLI between '" & Codigo1 & "' and '" & Codigo2 & "'"
  Else
      Codigo2 = Codigo1: Va2 = Va1
      cadena = cadena & " and {FACCAB.CFCODCLI} = '" & Codigo1 & "' "
      
      Sql1 = Sql1 & " AND CFCODCLI='" & Codigo1 & "'"
  End If
  
  Adodc1.Open Sql1, cConexCom, adOpenDynamic, adLockOptimistic
  If Adodc1.RecordCount = 0 Then
    MsgBox "     No existen artículos en este rango    ", vbOKOnly, "Aviso"
    Adodc1.Close
    Screen.MousePointer = 1
    Exit Sub
  End If
    
  Sql = "Select CNOMCLI from MAECLI Where CCODCLI='" & Codigo1 & "'"
  Adodc3.Open Sql, cConexCom, adOpenDynamic, adLockOptimistic
  If Adodc3.RecordCount = 1 Then Va1 = Adodc3("CNOMCLI")
  Adodc3.Close
  
  Sql = "Select CNOMCLI from MAECLI Where CCODCLI='" & Codigo2 & "'"
  Adodc3.Open Sql, cConexCom, adOpenDynamic, adLockOptimistic
  If Adodc3.RecordCount = 1 Then Va2 = Adodc3("CNOMCLI")
  Adodc3.Close
    
    CrystalReport1.ReportFileName = RUTA & "reporteInv\rentaC.rpt"
    Ubi_Tab CrystalReport1
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.SelectionFormula = cadena
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.Formulas(0) = "emp = '" & VGNemp & "'"
    CrystalReport1.Formulas(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
    CrystalReport1.Formulas(2) = "campoini = '" & Codigo1 & "'"
    CrystalReport1.Formulas(3) = "campofin = '" & Codigo2 & "'"
    CrystalReport1.Formulas(4) = "detaini = '" & Va1 & "'"
    CrystalReport1.Formulas(5) = "detafin = '" & Va2 & "'"
    CrystalReport1.Formulas(6) = "FECini = '" & DTPicker1.Value & "'"
    CrystalReport1.Formulas(7) = "FECfin = '" & DTPicker2.Value & "'"

    If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
End If
End Sub

Private Sub Imprimir2()
Dim cadena As String
Dim Sql As String, Sql1 As String
Dim Codigo1 As String, Codigo2 As String
Dim tex1 As String, tex2 As String
Dim Va1 As String, Va2 As String

Codigo1 = UCase(Trim(Text1))
Codigo2 = UCase(Trim(Text2))
Set Adodc1 = New ADODB.Recordset
Set Adodc3 = New ADODB.Recordset

cadena = "{LISTA_PRECIOS.Uni_LisPre}='" & Mid(Combo2.text, 1, 4) & "'"
cadena = cadena & " and {LISTA_PRECIOS.Cod_LisPre}='" & List & "' and {STKART.STALMA}='" & VGAlma & "'"

Sql = "Select AFamilia,Fam_Nombre from "
Sql = Sql & "(((MAEART A Inner Join STKART B on A.ACodigo=B.STCodigo) "
Sql = Sql & "Left Join FAMILIA C on A.AFamilia=C.Fam_Codigo) "
Sql = Sql & "Left Join LISTA_PRECIOS D on B.STCODIGO=D.COD_ARTI) "
Sql = Sql & "Where Stalma='" & VGAlma & "' AND UNI_LISPRE='" & Mid(Combo2.text, 1, 4) & "' "
Sql = Sql & "And COD_LISPRE='" & List & "'"

If OpTodos.Value Then
    Sql = Sql & " Order by AFAMILIA"
    
    Adodc3.Open Sql, cConexCom, adOpenDynamic, adLockOptimistic
    If Adodc3.RecordCount > 0 Then
      If Text1 = "" And Text2 = "" Then
        Adodc3.MoveFirst
        tex1 = IIf(IsNull(Adodc3("AFamilia")), "", Adodc3("AFamilia"))
        Va1 = IIf(IsNull(Adodc3("Fam_Nombre")), "", Adodc3("Fam_Nombre"))
        Adodc3.MoveLast
        tex2 = IIf(IsNull(Adodc3("AFamilia")), "", Adodc3("AFamilia"))
        Va2 = IIf(IsNull(Adodc3("Fam_Nombre")), "", Adodc3("Fam_Nombre"))
      End If
    Else
      MsgBox "No existen artículos para esta unidad de medida", vbOKOnly, "Aviso"
      Adodc3.Close
      Screen.MousePointer = 1
      Exit Sub
    End If
    Adodc3.Close
    
    CrystalReport1.ReportFileName = RUTA & "reporteInv\RentaF.rpt"
    Ubi_Tab CrystalReport1
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.SelectionFormula = cadena
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.Formulas(0) = "emp = '" & VGNemp & "'"
    CrystalReport1.Formulas(1) = "Hora = '" & Format(Time, "hh:mm:ss") & "'"
    CrystalReport1.Formulas(2) = "campoini = '" & tex1 & "'"
    CrystalReport1.Formulas(3) = "campofin = '" & tex2 & "'"
    CrystalReport1.Formulas(4) = "detaini = '" & Va1 & "'"
    CrystalReport1.Formulas(5) = "detafin = '" & Va2 & "'"
   CrystalReport1.Destination = crptToWindow
   CrystalReport1.WindowTitle = mensaje1
   If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
   Exit Sub
End If

If Text1 = "" Then
        MsgBox "Ingrese el codigo", vbExclamation, "Error"
        Text1.SetFocus
        Exit Sub
End If

If Option1(2).Value Then
  If Text2 <> "" Then
      Codigo2 = Text2         '  "23134671"
      cadena = cadena & " and ({MAEART.AFAMILIA} in '" & Codigo1 & "' to '" & Codigo2 & "')"
      
      Sql = Sql & " and AFAMILIA between '" & Codigo1 & "' and '" & Codigo2 & "'"
  Else
      Codigo2 = Codigo1: Va2 = Va1
      cadena = cadena & " and {MAEART.AFAMILIA} = '" & Codigo1 & "' "
          
      Sql = Sql & " and AFAMILIA ='" & Codigo1 & "'"
  End If
    
  Adodc1.Open Sql, cConexCom, adOpenDynamic, adLockOptimistic
  If Adodc1.RecordCount = 0 Then
    MsgBox "     No existen artículos en este rango    ", vbOKOnly, "Aviso"
    Adodc1.Close
    Screen.MousePointer = 1
    Exit Sub
  End If
    
  Adodc3.Open "Select Fam_Nombre from FAMILIA Where Fam_Codigo='" & Codigo1 & "'", cConexCom, adOpenDynamic, adLockOptimistic
  If Adodc3.RecordCount = 1 Then Va1 = Adodc3("Fam_Nombre")
  Adodc3.Close
  
  Adodc3.Open "Select Fam_Nombre from FAMILIA Where Fam_Codigo='" & Codigo2 & "'", cConexCom, adOpenDynamic, adLockOptimistic
  If Adodc3.RecordCount = 1 Then Va2 = Adodc3("Fam_Nombre")
  Adodc3.Close
    
    CrystalReport1.ReportFileName = RUTA & "reporteInv\RentaF.rpt"
    Ubi_Tab CrystalReport1
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.SelectionFormula = cadena
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.Formulas(0) = "emp = '" & VGNemp & "'"
    CrystalReport1.Formulas(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
    CrystalReport1.Formulas(2) = "campoini = '" & Codigo1 & "'"
    CrystalReport1.Formulas(3) = "campofin = '" & Codigo2 & "'"
    CrystalReport1.Formulas(4) = "detaini = '" & Va1 & "'"
    CrystalReport1.Formulas(5) = "detafin = '" & Va2 & "'"
    If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
End If
End Sub

Private Sub Imprimir3()
Dim cadena As String
Dim Sql As String
Dim Codigo1 As String, Codigo2 As String
Dim tex1 As String, tex2 As String
Dim Va1 As String, Va2 As String

Codigo1 = UCase(Trim(Text1))
Codigo2 = UCase(Trim(Text2))
Set Adodc1 = New ADODB.Recordset
Set Adodc3 = New ADODB.Recordset

cadena = "{LISTA_PRECIOS.Uni_LisPre}='" & Mid(Combo2.text, 1, 4) & "' and {STKART.STALMA}='" & VGAlma & "'"
cadena = cadena & " and {LISTA_PRECIOS.Cod_LisPre}='" & List & "' and {MAEART.AFAMILIA}='" & Text3 & "'"

Sql = "Select AModelo,Lin_Nombre from "
Sql = Sql & "(((MAEART A Inner Join STKART B on A.ACodigo=B.STCodigo) "
Sql = Sql & "Left Join LINEAS C on A.AFamilia=C.Fam_Codigo and A.Amodelo=C.Lin_Codigo) "
Sql = Sql & "Left Join LISTA_PRECIOS D on B.STCODIGO=D.COD_ARTI) "
Sql = Sql & "Where AFamilia='" & Text3.text & "' and Stalma='" & VGAlma & "' "
Sql = Sql & "AND UNI_LISPRE='" & Mid(Combo2.text, 1, 4) & "' "
Sql = Sql & "And COD_LISPRE='" & List & "'"

If OpTodos.Value Then
    Sql = Sql & " Order By Amodelo"
   
    Adodc3.Open Sql, cConexCom, adOpenDynamic, adLockOptimistic
    If Adodc3.RecordCount > 0 Then
      If Text1 = "" And Text2 = "" Then
        Adodc3.MoveFirst
        tex1 = IIf(IsNull(Adodc3("AModelo")), "", Adodc3("AModelo"))
        Va1 = IIf(IsNull(Adodc3("Lin_Nombre")), "", Adodc3("Lin_Nombre"))
        Adodc3.MoveLast
        tex2 = IIf(IsNull(Adodc3("AModelo")), "", Adodc3("AModelo"))
        Va2 = IIf(IsNull(Adodc3("Lin_Nombre")), "", Adodc3("Lin_Nombre"))
      End If
    Else
      MsgBox "No existen artículos para esta unidad de medida", vbOKOnly, "Aviso"
      Adodc3.Close
      Screen.MousePointer = 1
      Exit Sub
    End If
    Adodc3.Close
    
    CrystalReport1.ReportFileName = RUTA & "reporteInv\RentaL.rpt"
    Ubi_Tab CrystalReport1
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.SelectionFormula = cadena
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.Formulas(0) = "emp = '" & VGNemp & "'"
    CrystalReport1.Formulas(1) = "Hora = '" & Format(Time, "hh:mm:ss") & "'"
    CrystalReport1.Formulas(2) = "campoini = '" & tex1 & "'"
    CrystalReport1.Formulas(3) = "campofin = '" & tex2 & "'"
    CrystalReport1.Formulas(4) = "detaini = '" & Va1 & "'"
    CrystalReport1.Formulas(5) = "detafin = '" & Va2 & "'"
   CrystalReport1.Destination = crptToWindow
   CrystalReport1.WindowTitle = mensaje1
   If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
   Exit Sub
End If

If Text1 = "" Then
    MsgBox "Ingrese el codigo", vbExclamation, "Error"
    Text1.SetFocus
    Exit Sub
End If

If Option1(3).Value Then
  If Text2 <> "" Then
      Codigo2 = Text2         '  "23134671"
      cadena = cadena & " and ({MAEART.AMODELO} in '" & Codigo1 & "' to '" & Codigo2 & "')"
      
      Sql = Sql & " and AMODELO between '" & Codigo1 & "' and '" & Codigo2 & "'"
  Else
      Codigo2 = Codigo1: Va2 = Va1
      cadena = cadena & " and {MAEART.AMODELO} = '" & Codigo1 & "' "
      
      Sql = Sql & " and AMODELO ='" & Codigo1 & "'"
  End If
    
  Adodc1.Open Sql, cConexCom, adOpenDynamic, adLockOptimistic
  If Adodc1.RecordCount = 0 Then
    MsgBox "     No existen artículos en este rango    ", vbOKOnly, "Aviso"
    Adodc1.Close
    Screen.MousePointer = 1
    Exit Sub
  End If

  Adodc3.Open "Select Lin_Nombre from LINEAS Where Lin_Codigo='" & Codigo1 & "'", cConexCom, adOpenDynamic, adLockOptimistic
  If Adodc3.RecordCount = 1 Then Va1 = Adodc3("Lin_Nombre")
  Adodc3.Close
  
  Adodc3.Open "Select Lin_Nombre from LINEAS Where Lin_Codigo='" & Codigo2 & "'", cConexCom, adOpenDynamic, adLockOptimistic
  If Adodc3.RecordCount = 1 Then Va2 = Adodc3("Lin_Nombre")
  Adodc3.Close
    
    CrystalReport1.ReportFileName = RUTA & "reporteInv\RentaL.rpt"
    Ubi_Tab CrystalReport1
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.SelectionFormula = cadena
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.Formulas(0) = "emp = '" & VGNemp & "'"
    CrystalReport1.Formulas(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
    CrystalReport1.Formulas(2) = "campoini = '" & Codigo1 & "'"
    CrystalReport1.Formulas(3) = "campofin = '" & Codigo2 & "'"
    CrystalReport1.Formulas(4) = "detaini = '" & Va1 & "'"
    CrystalReport1.Formulas(5) = "detafin = '" & Va2 & "'"
    If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
End If
End Sub

Private Sub Imprimir4()
Dim cadena As String
Dim Sql As String
Dim Codigo1 As String, Codigo2 As String
Dim tex1 As String, tex2 As String
Dim Va1 As String, Va2 As String

Codigo1 = UCase(Trim(Text1))
Codigo2 = UCase(Trim(Text2))
Set Adodc1 = New ADODB.Recordset
Set Adodc3 = New ADODB.Recordset

cadena = "{LISTA_PRECIOS.Uni_LisPre}='" & Mid(Combo2.text, 1, 4) & "'"
cadena = cadena & " and {LISTA_PRECIOS.Cod_LisPre}='" & List & "' "
cadena = cadena & "and {MAEART.AFAMILIA}='" & Text3 & "'"
cadena = cadena & "and {MAEART.AMODELO}='" & Text4 & "' and {STKART.STALMA}='" & VGAlma & "'"

Sql = "Select AGrupo,Acodigo,Gru_Nombre from "
Sql = Sql & "(((MAEART A Inner Join STKART B on A.ACodigo=B.STCodigo) "
Sql = Sql & "Left Join GRUPO C on A.AFamilia=C.Fam_Codigo and A.Amodelo=C.Lin_Codigo and A.Agrupo=C.Gru_Codigo) "
Sql = Sql & "Left Join LISTA_PRECIOS D on B.STCODIGO=D.COD_ARTI) "
Sql = Sql & "Where AFamilia='" & Text3.text & "' and Amodelo='" & Text4.text & "' and Stalma='" & VGAlma & "' "
Sql = Sql & "And Uni_LisPre='" & Mid(Combo2.text, 1, 4) & "' "
Sql = Sql & "And Cod_LisPre='" & List & "' "

If OpTodos.Value Then
    Sql = Sql & " Order by Agrupo"

    Adodc3.Open Sql, cConexCom, adOpenDynamic, adLockOptimistic
    If Adodc3.RecordCount > 0 Then
      If Text1 = "" And Text2 = "" Then
        Adodc3.MoveFirst
        tex1 = IIf(IsNull(Adodc3("AGrupo")), "", Adodc3("AGrupo"))
        Va1 = IIf(IsNull(Adodc3("Gru_Nombre")), "", Adodc3("Gru_Nombre"))
        Adodc3.MoveLast
        tex2 = IIf(IsNull(Adodc3("AGrupo")), "", Adodc3("AGrupo"))
        Va2 = IIf(IsNull(Adodc3("Gru_Nombre")), "", Adodc3("Gru_Nombre"))
      End If
    Else
      MsgBox "No existen artículos para esta unidad de medida", vbOKOnly, "Aviso"
      Adodc3.Close
      Screen.MousePointer = 1
      Exit Sub
    End If
    Adodc3.Close
    
    CrystalReport1.ReportFileName = RUTA & "reporteInv\RentaG.rpt"
    Ubi_Tab CrystalReport1
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.SelectionFormula = cadena
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.Formulas(0) = "emp = '" & VGNemp & "'"
    CrystalReport1.Formulas(1) = "Hora = '" & Format(Time, "hh:mm:ss") & "'"
    CrystalReport1.Formulas(2) = "campoini = '" & tex1 & "'"
    CrystalReport1.Formulas(3) = "campofin = '" & tex2 & "'"
    CrystalReport1.Formulas(4) = "detaini = '" & Va1 & "'"
    CrystalReport1.Formulas(5) = "detafin = '" & Va2 & "'"
   CrystalReport1.Destination = crptToWindow
   CrystalReport1.WindowTitle = mensaje1
   If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
   Exit Sub
End If

If Text1 = "" Then
        MsgBox "Ingrese el codigo", vbExclamation, "Error"
        Text1.SetFocus
        Exit Sub
End If

If Option1(4).Value Then
  If Text2 <> "" Then
      Codigo2 = Text2         '  "23134671"
      cadena = cadena & " and ({MAEART.AGRUPO} in '" & Codigo1 & "' to '" & Codigo2 & "')"
      
      Sql = Sql & " and AGRUPO between '" & Codigo1 & "' and '" & Codigo2 & "'"
  Else
      Codigo2 = Codigo1: Va2 = Va1
      cadena = cadena & " and {MAEART.AGRUPO} = '" & Codigo1 & "' "
      
      Sql = Sql & " and AGRUPO ='" & Codigo1 & "'"
  End If
    
  Adodc1.Open Sql, cConexCom, adOpenDynamic, adLockOptimistic
  If Adodc1.RecordCount = 0 Then
    MsgBox "     No existen artículos en este rango    ", vbOKOnly, "Aviso"
    Adodc1.Close
    Screen.MousePointer = 1
    Exit Sub
  End If
    
  Adodc3.Open "Select Gru_Nombre from GRUPO Where Gru_Codigo='" & Codigo1 & "'", cConexCom, adOpenDynamic, adLockOptimistic
  If Adodc3.RecordCount = 1 Then Va1 = Adodc3("Gru_Nombre")
  Adodc3.Close
  
  Adodc3.Open "Select Gru_Nombre from Grupo Where Gru_Codigo='" & Codigo2 & "'", cConexCom, adOpenDynamic, adLockOptimistic
  If Adodc3.RecordCount = 1 Then Va2 = Adodc3("Gru_Nombre")
  Adodc3.Close

  CrystalReport1.ReportFileName = RUTA & "reporteInv\RentaG.rpt"
  Ubi_Tab CrystalReport1
  CrystalReport1.DiscardSavedData = True
  CrystalReport1.Destination = crptToWindow
  CrystalReport1.SelectionFormula = cadena
  CrystalReport1.WindowShowPrintBtn = True
  CrystalReport1.WindowShowRefreshBtn = True
  CrystalReport1.WindowShowSearchBtn = True
  CrystalReport1.WindowShowPrintSetupBtn = True
  CrystalReport1.Formulas(0) = "emp = '" & VGNemp & "'"
  CrystalReport1.Formulas(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
  CrystalReport1.Formulas(2) = "campoini = '" & Codigo1 & "'"
  CrystalReport1.Formulas(3) = "campofin = '" & Codigo2 & "'"
  CrystalReport1.Formulas(4) = "detaini = '" & Va1 & "'"
  CrystalReport1.Formulas(5) = "detafin = '" & Va2 & "'"
  If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
End If
End Sub

Private Sub Imprimir5()
Dim cadena As String
Dim Sql As String
Dim Codigo1 As String, Codigo2 As String
Dim tex1 As String, tex2 As String
Dim Va1 As String, Va2 As String

Codigo1 = UCase(Trim(Text1))
Codigo2 = UCase(Trim(Text2))
Set Adodc1 = New ADODB.Recordset
Set Adodc3 = New ADODB.Recordset

cadena = "{FACCAB.CFFECDOC} in date(" & Format(DTPicker1, "yyyy") & "," & Format(DTPicker1, "mm") & "," & Format(DTPicker1, "dd") & " ) "
cadena = cadena & "to date(" & Format(DTPicker2, "yyyy") & "," & Format(DTPicker2, "mm") & "," & Format(DTPicker2, "dd") & " )"

Sql = "Select COD_VEN,DES_VEN from "
Sql = Sql & "VENDEDOR A Inner Join FACCAB B on A.COD_VEN=B.CFVENDE "
Sql = Sql & " Where (CFFECDOC between #" & Format(DTPicker1, "yyyy/mm/dd") & "# and "
Sql = Sql & " #" & Format(DTPicker2, "yyyy/mm/dd") & "#)"


If OpTodos.Value Then
    Sql = Sql & " Order by COD_VEN"
    
    Adodc3.Open Sql, cConexCom, adOpenDynamic, adLockOptimistic
    If Adodc3.RecordCount > 0 Then
      If Text1 = "" And Text2 = "" Then
        Adodc3.MoveFirst
        tex1 = IIf(IsNull(Adodc3("COD_VEN")), "", Adodc3("COD_VEN"))
        Va1 = IIf(IsNull(Adodc3("DES_VEN")), "", Adodc3("DES_VEN"))
        Adodc3.MoveLast
        tex2 = IIf(IsNull(Adodc3("COD_VEN")), "", Adodc3("COD_VEN"))
        Va2 = IIf(IsNull(Adodc3("DES_VEN")), "", Adodc3("DES_VEN"))
      End If
    Else
      MsgBox "No existen Vendedores para este rango de fechas", vbOKOnly, "Aviso"
      Adodc3.Close
      Screen.MousePointer = 1
      Exit Sub
    End If
    Adodc3.Close
    
    CrystalReport1.ReportFileName = RUTA & "reporteInv\rentaV.rpt"
    Ubi_Tab CrystalReport1
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.SelectionFormula = cadena
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.Formulas(0) = "emp = '" & VGNemp & "'"
    CrystalReport1.Formulas(1) = "Hora = '" & Format(Time, "hh:mm:ss") & "'"
    CrystalReport1.Formulas(2) = "campoini = '" & tex1 & "'"
    CrystalReport1.Formulas(3) = "campofin = '" & tex2 & "'"
    CrystalReport1.Formulas(4) = "detaini = '" & Va1 & "'"
    CrystalReport1.Formulas(5) = "detafin = '" & Va2 & "'"
    CrystalReport1.Formulas(6) = "FECini = '" & DTPicker1.Value & "'"
    CrystalReport1.Formulas(7) = "FECfin = '" & DTPicker2.Value & "'"
   CrystalReport1.Destination = crptToWindow
   CrystalReport1.WindowTitle = mensaje1
   If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
   Exit Sub
End If

If Text1 = "" Then
    MsgBox "Ingrese el código", vbExclamation, "Error"
    Text1.SetFocus
    Exit Sub
End If

If Option1(5).Value Then
  If Text2 <> "" Then
      Codigo2 = Text2         '  "23134671"
      cadena = cadena & " and ({FACCAB.CFVENDE} in '" & Codigo1 & "' to '" & Codigo2 & "')"
      
      Sql = Sql & " and CFVENDE between '" & Codigo1 & "' and '" & Codigo2 & "'"
  Else
      Codigo2 = Codigo1: Va2 = Va1
      cadena = cadena & " and {FACCAB.CFVENDE} = '" & Codigo1 & "' "
      
      Sql = Sql & " and CFVENDE ='" & Codigo1 & "'"
  End If
    
  Adodc1.Open Sql, cConexCom, adOpenDynamic, adLockOptimistic
  If Adodc1.RecordCount = 0 Then
    MsgBox "  No existen Vendedores en este rango de fecha ", vbOKOnly, "Aviso"
    Adodc1.Close
    Screen.MousePointer = 1
    Exit Sub
  End If
    
  Adodc3.Open "Select DES_VEN from VENDEDOR Where COD_VEN='" & Codigo1 & "'", cConexCom, adOpenDynamic, adLockOptimistic
  If Adodc3.RecordCount = 1 Then Va1 = Adodc3("DES_VEN")
  Adodc3.Close
  
  Adodc3.Open "Select DES_VEN from VENDEDOR Where COD_VEN='" & Codigo2 & "'", cConexCom, adOpenDynamic, adLockOptimistic
  If Adodc3.RecordCount = 1 Then Va2 = Adodc3("DES_VEN")
  Adodc3.Close
    
    CrystalReport1.ReportFileName = RUTA & "reporteInv\rentaV.rpt"
    Ubi_Tab CrystalReport1
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.SelectionFormula = cadena
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.Formulas(0) = "emp = '" & VGNemp & "'"
    CrystalReport1.Formulas(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
    CrystalReport1.Formulas(2) = "campoini = '" & Codigo1 & "'"
    CrystalReport1.Formulas(3) = "campofin = '" & Codigo2 & "'"
    CrystalReport1.Formulas(4) = "detaini = '" & Va1 & "'"
    CrystalReport1.Formulas(5) = "detafin = '" & Va2 & "'"
    CrystalReport1.Formulas(6) = "FECini = '" & DTPicker1.Value & "'"
    CrystalReport1.Formulas(7) = "FECfin = '" & DTPicker2.Value & "'"
    
    If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
End If
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then DTPicker2.SetFocus
End Sub

Private Sub DTPicker2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Command1.SetFocus
End Sub

Private Sub Form_Load()
  central Me
  Carga_Lista
  Carga_Uni
  DTPicker1 = DateAdd("m", -1, Date)
  DTPicker2.Value = Date
End Sub

Private Sub renta()
   Dim ds As String

End Sub

Private Sub limpiar_t1_t2()
Text1 = "": Text2 = "": Text3 = "": Text4 = ""
Label4.Visible = False
Text3.Visible = False
Label5.Visible = False
Text4.Visible = False
Frame3.Visible = False
Frame2.Visible = True
End Sub

Private Sub OpRango_Click()
If OpRango.Value Then
    Text1.Enabled = True
    Text2.Enabled = True
    Text1.SetFocus
End If
End Sub

Private Sub Option1_Click(Index As Integer)
limpiar_t1_t2
If Index = 0 Then
  FrameRep.Caption = " Por Articulos "
  OpTodos.Caption = "Todos los Articulos"
  OpTodos.Top = 300: OpRango.Top = 650
  Text1.Top = 1100: Label3.Top = 1100
  Text2.Top = 1500: Label6.Top = 1500
ElseIf Index = 1 Then   'Por Cliente
  FrameRep.Caption = " Por Clientes "
  OpTodos.Caption = "Todos los Clientes"
  OpTodos.Top = 300: OpRango.Top = 650
  Text1.Top = 1100: Label3.Top = 1100
  Text2.Top = 1500: Label6.Top = 1500
  Frame3.Visible = True: Frame2.Visible = False
ElseIf Index = 2 Then
  FrameRep.Caption = " Por Familias "
  OpTodos.Caption = "Todos las Familias"
  OpTodos.Top = 300: OpRango.Top = 650
  Text1.Top = 1100: Label3.Top = 1100
  Text2.Top = 1500: Label6.Top = 1500
ElseIf Index = 3 Then
  FrameRep.Caption = " Por Lineas "
  OpTodos.Caption = "Todos las Líneas "
  Label4.Visible = True
  Text3.Visible = True
  OpTodos.Top = 550: OpRango.Top = 900
  Text1.Top = 1200: Label3.Top = 1200
  Text2.Top = 1600: Label6.Top = 1600
ElseIf Index = 4 Then
  FrameRep.Caption = " Por Grupos "
  OpTodos.Caption = "Todos los Grupos"
  Label4.Visible = True: Text3.Visible = True
  Label5.Visible = True: Text4.Visible = True
  OpTodos.Top = 850: OpRango.Top = 1100
  Text1.Top = 1400: Label3.Top = 1400
  Text2.Top = 1700: Label6.Top = 1700
ElseIf Index = 5 Then   'Por Vendedor
  FrameRep.Caption = " Por Vendedor "
  OpTodos.Caption = "Todos los Vendedor"
  OpTodos.Top = 300: OpRango.Top = 650
  Text1.Top = 1100: Label3.Top = 1100
  Text2.Top = 1500: Label6.Top = 1500
  Frame3.Visible = True: Frame2.Visible = False
End If
End Sub

Private Sub Carga_Lista()
Dim Adodc1 As New ADODB.Recordset
Dim Sql As String

Sql = "Select Cod_LisPre,Des_LisPre from TIPO_PRECIO"
Adodc1.Open Sql, cConexCom, adOpenDynamic, adLockOptimistic
If Adodc1.RecordCount > 0 Then
  Do While Not Adodc1.EOF
    Combo1.AddItem Adodc1(0) & "  " & Adodc1(1)
    Adodc1.MoveNext
  Loop
  Combo1.ListIndex = 0
End If
Adodc1.Close
End Sub

Private Sub Carga_Uni()
Dim Adodc1 As New ADODB.Recordset
Dim Sql As String
Dim Ab As String

Sql = "Select UM_ABREV,UM_NOMBRE from TABUNIMED"
Adodc1.Open Sql, cConexCom, adOpenDynamic, adLockOptimistic
If Adodc1.RecordCount > 0 Then
  Do While Not Adodc1.EOF
    Ab = Space(6)
    Ab = Adodc1("UM_ABREV")
    Combo2.AddItem Ab & "  " & Adodc1("UM_NOMBRE")
    Adodc1.MoveNext
  Loop
  Combo2.ListIndex = 0
End If
Adodc1.Close
End Sub

Private Sub Option1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
  If Index = 4 Or Index = 3 Then
    Text3.SetFocus
  Else
    OpTodos.SetFocus
  End If
End If
End Sub

Private Sub OpTodos_Click()
Text1.Enabled = False: Text1 = ""
Text2.Enabled = False: Text2 = ""
If Option1(3).Value Then
  Label4.Visible = True
  Text3.Visible = True
ElseIf Option1(4).Value Then
  Label4.Visible = True: Label5.Visible = True
  Text3.Visible = True: Text4.Visible = True
End If
End Sub

Private Sub OpTodos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If Frame2.Visible = True Then Combo1.SetFocus
  If Frame3.Visible = True Then DTPicker1.SetFocus
End If
End Sub

Private Sub Text1_DblClick()
Dim Adodc2 As ADODB.Recordset
Screen.MousePointer = 11
Set Adodc2 = New ADODB.Recordset
If Option1(0).Value Then
         Adodc2.Open "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD from MaeArt p, StkArt n Where n.STSKDIS<> 0 and  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & VGAlma & "'  ", cConexCom, adOpenStatic, adLockOptimistic
         frmReferencia.Conectar Adodc2, "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD from MaeArt p, StkArt n Where n.STSKDIS<> 0 and  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & VGAlma & "'  "
         frmReferencia.Label1.Caption = "Artículos"
         frmReferencia.Show vbModal
         Adodc2.Close
         If vGUtil(1) <> "" Then
                 Text1 = (vGUtil(1))
         End If
ElseIf Option1(1).Value Then
         Adodc2.Open "Select CCODCLI,CNOMCLI from MAECLI", cConexCom, adOpenStatic, adLockOptimistic
         frmReferencia.Conectar Adodc2, "Select CCODCLI,CNOMCLI from MAECLI"
         frmReferencia.Label1.Caption = "Clientes"
         frmReferencia.Show vbModal
         Adodc2.Close
         If vGUtil(1) <> "" Then
                 Text1 = (vGUtil(1))
         End If
ElseIf Option1(2).Value Then
        Adodc2.Open "SELECT FAM_CODIGO,FAM_NOMBRE FROM FAMILIA", cConexCom, adOpenStatic, adLockOptimistic
        frmReferencia.Conectar Adodc2, "SELECT FAM_CODIGO,FAM_NOMBRE FROM FAMILIA"
        frmReferencia.Label1.Caption = "Familias de Artículos"
        frmReferencia.Show vbModal
        Adodc2.Close
        If vGUtil(1) <> "" Then
                Text1 = (vGUtil(1))
        End If
ElseIf Option1(3).Value Then
        Adodc2.Open "SELECT LIN_CODIGO,LIN_NOMBRE FROM LINEAS Where Fam_Codigo='" & Trim(Text3) & "'", cConexCom, adOpenStatic, adLockOptimistic
        frmReferencia.Conectar Adodc2, "SELECT LIN_CODIGO,LIN_NOMBRE FROM LINEAS Where Fam_Codigo='" & Trim(Text3) & "'"
        frmReferencia.Label1.Caption = "Líneas de Artículos"
        frmReferencia.Show vbModal
        Adodc2.Close
        If vGUtil(1) <> "" Then
                Text1 = (vGUtil(1))
        End If
ElseIf Option1(4).Value Then
        Adodc2.Open "SELECT GRU_CODIGO,GRU_NOMBRE FROM GRUPO Where Fam_Codigo='" & Trim(Text3) & "' and Lin_Codigo='" & Trim(Text4) & "'", cConexCom, adOpenStatic, adLockOptimistic
        frmReferencia.Conectar Adodc2, "SELECT GRU_CODIGO,GRU_NOMBRE FROM GRUPO Where Fam_Codigo='" & Trim(Text3) & "' and Lin_Codigo='" & Trim(Text4) & "'"
        frmReferencia.Label1.Caption = "Grupos de Artículos"
        frmReferencia.Show vbModal
        Adodc2.Close
        If vGUtil(1) <> "" Then
                Text1 = (vGUtil(1))
        End If
ElseIf Option1(5).Value Then
         Adodc2.Open "Select COD_VEN,DES_VEN from VENDEDOR", cConexCom, adOpenStatic, adLockOptimistic
         frmReferencia.Conectar Adodc2, "Select COD_VEN,DES_VEN from VENDEDOR"
         frmReferencia.Label1.Caption = "Vendedores"
         frmReferencia.Show vbModal
         Adodc2.Close
         If vGUtil(1) <> "" Then
                 Text1 = (vGUtil(1))
         End If
End If
If Text1 <> "" And Text2 <> "" And Text1 > Text2 Then
    MsgBox "Ingrese un codigo menor al fin ", vbOKOnly, "Error"
    Screen.MousePointer = 1
    Exit Sub
End If
If Text1 <> "" Then
    Text2.Enabled = True
    Text2.SetFocus
End If
Screen.MousePointer = 1
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then Text1_DblClick
End Sub

Function Existe_cod_art(text As TextBox) As String
Dim rS As Recordset
Dim rSql As String
rSql = "select  ACODIGO From MAEART where ACODIGO = '" & text & "'" '
Set db = Workspaces(0).OpenDatabase(cRuta2)
Set rS = db.OpenRecordset(rSql, dbOpenSnapshot)
If Not rS.EOF Then
    Existe_cod_art = rS(0)
Else
    MsgBox "El tipo de codigo no existe !", vbOKOnly, "Error"
    Existe_cod_art = ""
End If
rS.Close
End Function

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Not OpRango.Value Then
   OpRango = True
End If
If KeyAscii = 13 And Text1 <> "" Then
    If Option1(0).Value Then
       If Existe_cod_art(Text1) <> "" Then
               Text2.Enabled = True
               Text2.SetFocus
       End If
   ElseIf Option1(1).Value Then
        If Existe(1, Text1, "MAECLI", "CCODCLI", False) = False Then
                MsgBox "El código de Cliente no existe", vbInformation, mensaje1
                Text1.SetFocus: Exit Sub
         Else
                Text2.Enabled = True
                Text2.SetFocus
         End If
   ElseIf Option1(2).Value Then
        If Existe(1, Text1, "FAMILIA", "FAM_CODIGO", False) = False Then
                MsgBox "El código de Familia no existe", vbInformation, mensaje1
                Text1.SetFocus: Exit Sub
         Else
                Text2.Enabled = True
                Text2.SetFocus
         End If
   ElseIf Option1(3).Value Then
        If Existe(1, Text1, "LINEAS", "LIN_CODIGO", False) = False Then
                MsgBox "El código de Línea no existe", vbInformation, mensaje1
                Text1.SetFocus: Exit Sub
         Else
                Text2.Enabled = True
                Text2.SetFocus
         End If
   ElseIf Option1(4).Value Then
        If Existe(1, Text1, "GRUPO", "GRU_CODIGO", False) = False Then
                MsgBox "El código de Grupo no existe", vbInformation, mensaje1
                Text1.SetFocus: Exit Sub
         Else
                Text2.Enabled = True
                Text2.SetFocus
         End If
   ElseIf Option1(5).Value Then
        If Existe(1, Text1, "VENDEDOR", "COD_VEN", False) = False Then
                MsgBox "El código de Vendedor no existe", vbInformation, mensaje1
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
Screen.MousePointer = 11
If Option1(0).Value Then
    Adodc2.Open "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD from MaeArt p, StkArt n Where n.STSKDIS<> 0 and  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & VGAlma & "'  ", cConexCom, adOpenStatic, adLockOptimistic
    frmReferencia.Conectar Adodc2, "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD from MaeArt p, StkArt n Where n.STSKDIS<> 0 and  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & VGAlma & "'  "
    frmReferencia.Label1.Caption = "Artículos"
    frmReferencia.Show vbModal
    Adodc2.Close
    If vGUtil(1) <> "" Then
        Text2 = (vGUtil(1))
    End If
ElseIf Option1(1).Value Then
    Adodc2.Open "Select CCODCLI,CNOMCLI from MAECLI", cConexCom, adOpenStatic, adLockOptimistic
    frmReferencia.Conectar Adodc2, "Select CCODCLI,CNOMCLI from MAECLI"
    frmReferencia.Label1.Caption = "Clientes"
    frmReferencia.Show vbModal
    Adodc2.Close
    If vGUtil(1) <> "" Then
         Text2 = (vGUtil(1))
    End If
ElseIf Option1(2).Value Then
    Adodc2.Open "SELECT FAM_CODIGO,FAM_NOMBRE FROM FAMILIA", cConexCom, adOpenStatic, adLockOptimistic
    frmReferencia.Conectar Adodc2, "SELECT FAM_CODIGO,FAM_NOMBRE FROM FAMILIA"
    frmReferencia.Label1.Caption = "Familias de Artículos"
    frmReferencia.Show vbModal
    Adodc2.Close
    If vGUtil(1) <> "" Then
      Text2 = (vGUtil(1))
    End If
ElseIf Option1(3).Value Then
    Adodc2.Open "SELECT LIN_CODIGO,LIN_NOMBRE FROM LINEAS Where Fam_Codigo='" & Trim(Text3) & "'", cConexCom, adOpenStatic, adLockOptimistic
    frmReferencia.Conectar Adodc2, "SELECT LIN_CODIGO,LIN_NOMBRE FROM LINEAS Where Fam_Codigo='" & Trim(Text3) & "'"
    frmReferencia.Label1.Caption = "Líneas de Artículos"
    frmReferencia.Show vbModal
    Adodc2.Close
    If vGUtil(1) <> "" Then
        Text2 = (vGUtil(1))
    End If
ElseIf Option1(4).Value Then
    Adodc2.Open "SELECT GRU_CODIGO,GRU_NOMBRE FROM GRUPO Where Fam_Codigo='" & Trim(Text3) & "' and Lin_Codigo='" & Trim(Text4) & "'", cConexCom, adOpenStatic, adLockOptimistic
    frmReferencia.Conectar Adodc2, "SELECT GRU_CODIGO,GRU_NOMBRE FROM GRUPO Where Fam_Codigo='" & Trim(Text3) & "' and Lin_Codigo='" & Trim(Text4) & "'"
    frmReferencia.Label1.Caption = "Grupos de Artículos"
    frmReferencia.Show vbModal
    Adodc2.Close
    If vGUtil(1) <> "" Then
        Text2 = (vGUtil(1))
    End If
ElseIf Option1(5).Value Then
    Adodc2.Open "Select COD_VEN,DES_VEN from VENDEDOR", cConexCom, adOpenStatic, adLockOptimistic
    frmReferencia.Conectar Adodc2, "Select COD_VEN,DES_VEN from VENDEDOR"
    frmReferencia.Label1.Caption = "Vendedores"
    frmReferencia.Show vbModal
    Adodc2.Close
    If vGUtil(1) <> "" Then
        Text2 = (vGUtil(1))
    End If
End If
If Text1 <> "" And Text2 <> "" And Text1 > Text2 Then
    MsgBox "El código no puede ser menor al fin", vbOKOnly, "Error"
    Text2.SetFocus: Screen.MousePointer = 1: Exit Sub
End If
If Option1(1).Value Or Option1(5).Value Then DTPicker1.SetFocus Else Combo1.SetFocus
Screen.MousePointer = 1
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then Text2_DblClick
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If Not OpRango.Value Then
   OpRango = True
End If
If KeyAscii = 13 And Text1 <> "" Then
    If Option1(0).Value Then
       If Existe_cod_art(Text2) <> "" Then
            Text2.Enabled = True
            Combo1.SetFocus
       End If
   ElseIf Option1(1).Value Then
        If Existe(1, Text2, "MAECLI", "CCODCLI", False) = False Then
                MsgBox "El código de Cliente no existe", vbInformation, mensaje1
                Text2.SetFocus: Exit Sub
         Else
                Text2.Enabled = True
                DTPicker1.SetFocus
         End If
   ElseIf Option1(2).Value Then
        If Existe(1, Text2, "FAMILIA", "FAM_CODIGO", False) = False Then
                MsgBox "El código de Familia no existe", vbInformation, mensaje1
                Text2.SetFocus: Exit Sub
         Else
                Text2.Enabled = True
                Combo1.SetFocus
         End If
   ElseIf Option1(3).Value Then
        If Existe(1, Text2, "LINEAS", "LIN_CODIGO", False) = False Then
                MsgBox "El código de Línea no existe", vbInformation, mensaje1
                Text2.SetFocus: Exit Sub
         Else
                Text2.Enabled = True
                Combo1.SetFocus
         End If
   ElseIf Option1(4).Value Then
        If Existe(1, Text2, "GRUPO", "GRU_CODIGO", False) = False Then
                MsgBox "El código de Grupo no existe", vbInformation, mensaje1
                Text2.SetFocus: Exit Sub
         Else
                Text2.Enabled = True
                Combo1.SetFocus
         End If
   ElseIf Option1(5).Value Then
        If Existe(1, Text2, "VENDEDOR", "COD_VEN", False) = False Then
                MsgBox "El código de Vendedor no existe", vbInformation, mensaje1
                Text2.SetFocus: Exit Sub
         Else
                Text2.Enabled = True
                DTPicker1.SetFocus
         End If
     End If
 End If
End Sub

Private Sub Text3_DblClick()
Dim Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset
If Option1(3).Value Or Option1(4).Value Then
    Adodc2.Open "Select FAM_CODIGO,FAM_NOMBRE,FAM_CTA from FAMILIA ", cConexCom, adOpenStatic, adLockOptimistic
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
If KeyAscii = 13 Then
   If Option1(3).Value Or Option1(4).Value Then
      If Existe(1, Text3, "FAMILIA", "FAM_CODIGO", False) = False Then
          MsgBox "El código de Familia no existe", vbInformation, mensaje1
          Text3.SetFocus: Exit Sub
       Else
          Text4.Enabled = True
          If Option1(3).Value Then OpTodos.SetFocus Else Text4.SetFocus
       End If
   End If
  
  If Text3 <> "" Then
     If Text4.Visible = True Then Text4.SetFocus Else OpTodos.SetFocus
  End If
End If
End Sub

Private Sub Text4_DblClick()
Dim Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset
If Option1(4).Value Then
    Adodc2.Open "Select LIN_CODIGO,LIN_NOMBRE from LINEAS Where Fam_Codigo='" & Text3 & "'", cConexCom, adOpenStatic, adLockOptimistic
    frmReferencia.Conectar Adodc2, "Select LIN_CODIGO,LIN_NOMBRE from LINEAS Where Fam_Codigo='" & Text3 & "'"
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
If KeyAscii = 13 Then
   If Option1(4).Value Then
      If Existe(1, Text4, "LINEAS", "LIN_CODIGO", False) = False Then
          MsgBox "El código de Línea no existe", vbInformation, mensaje1
          Text4.SetFocus: Exit Sub
       End If
  End If
  If Text4 <> "" Then
     OpTodos.SetFocus
  End If
End If
End Sub
