VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmReposicion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reposicion de Stock"
   ClientHeight    =   4200
   ClientLeft      =   1755
   ClientTop       =   1260
   ClientWidth     =   6165
   Icon            =   "frmReposicion.frx":0000
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   6165
   Begin VB.Frame FrameRep 
      Height          =   2055
      Left            =   2580
      TabIndex        =   17
      Top             =   1035
      Width           =   3255
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1620
         TabIndex        =   12
         Top             =   1635
         Width           =   1275
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1620
         MaxLength       =   20
         TabIndex        =   11
         Top             =   1305
         Width           =   1275
      End
      Begin VB.OptionButton OpTodos 
         Caption         =   "Todos los Artículos"
         Height          =   240
         Left            =   360
         TabIndex        =   9
         Top             =   795
         Width           =   1785
      End
      Begin VB.OptionButton OpRango 
         Caption         =   "Rango"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   1035
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   7
         Top             =   210
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   8
         Top             =   510
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.Label Label3 
         Caption         =   "Fin"
         Height          =   255
         Left            =   765
         TabIndex        =   21
         Top             =   1635
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Inicio "
         Height          =   255
         Left            =   765
         TabIndex        =   20
         Top             =   1335
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Familia"
         Height          =   225
         Left            =   360
         TabIndex        =   19
         Top             =   225
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Línea"
         Height          =   225
         Left            =   360
         TabIndex        =   18
         Top             =   525
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Parámetros"
      Height          =   2055
      Left            =   360
      TabIndex        =   16
      Top             =   1035
      Width           =   1815
      Begin VB.OptionButton OpArt 
         Caption         =   "Artículos"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   405
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Familias"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   765
         Width           =   1095
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Líneas"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1125
         Width           =   975
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Grupos"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1485
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccionar"
      Height          =   795
      Left            =   360
      TabIndex        =   14
      Top             =   135
      Width           =   5475
      Begin VB.OptionButton Option1 
         Caption         =   "Stock Mínimo"
         Height          =   255
         Index           =   0
         Left            =   225
         TabIndex        =   0
         Top             =   315
         Width           =   1365
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Stock Máximo"
         Height          =   255
         Index           =   1
         Left            =   1815
         TabIndex        =   1
         Top             =   330
         Width           =   1410
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Reposicion de stock"
         Height          =   255
         Index           =   2
         Left            =   3465
         TabIndex        =   2
         Top             =   330
         Width           =   1830
      End
   End
   Begin VB.CommandButton Command8 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   735
      Left            =   3270
      Picture         =   "frmReposicion.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3300
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   735
      Left            =   2175
      Picture         =   "frmReposicion.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3300
      Width           =   735
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   225
      Top             =   3300
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmReposicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim db As Database
Dim almacen As String
Dim Conexion As String
Dim Adodc3 As ADODB.Recordset
Dim Adodc1 As ADODB.Recordset

Private Sub Command1_Click()
If Text1 <> "" And Text2 <> "" And Text1 > Text2 Then
    MsgBox "Ingrese un código menor al fin ", vbOKOnly, "Error"
    Exit Sub
End If
If OpArt.Value Then
     imprimir
ElseIf Option2.Value Then
    Imprimir2
ElseIf Option3.Value Then
    Imprimir3
ElseIf Option4.Value Then
    Imprimir4
End If
End Sub

Private Sub Command8_Click()
Unload Me
End Sub

Private Sub Form_Load()
 central Me
 OpTodos.Value = True
 Option1(0).Value = True
 OpArt.Value = True
End Sub

Private Sub imprimir()
Dim Codigo1 As String
Dim Codigo2 As String
Dim cadena As String
Dim SQL As String
Dim tex1 As String, tex2 As String
Dim Va1 As String, Va2 As String
Screen.MousePointer = 11

Codigo1 = UCase(Trim(Text1))
Codigo2 = UCase(Trim(Text2))
Set Adodc1 = New ADODB.Recordset
Set Adodc3 = New ADODB.Recordset

cadena = "{STKART.STALMA}='" & VGAlma & "'"

SQL = "Select ACodigo,Adescri from "
SQL = SQL & "MAEART A Inner Join STKART B on A.ACodigo=B.STCodigo "
SQL = SQL & "Where Stalma='" & VGAlma & "' "

If OpTodos.Value Then
    SQL = SQL & " Order by Acodigo"
    
    Adodc3.Open SQL, cConexCom, adOpenDynamic, adLockOptimistic
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
      MsgBox "    No existen artículos      ", vbOKOnly, "Aviso"
      Adodc3.Close
      Screen.MousePointer = 1
      Exit Sub
    End If
    Adodc3.Close
    
    If Option1(0).Value Then
      CrystalReport1.WindowTitle = "Inv061 -- Control de Inventarios"
      CrystalReport1.ReportFileName = cRutP & "inv061.rpt"
    ElseIf Option1(1).Value Then
      CrystalReport1.ReportFileName = cRutP & "inv056.rpt"
      CrystalReport1.WindowTitle = "Inv056 -- Control de Inventarios"
    ElseIf Option1(2).Value Then
      CrystalReport1.ReportFileName = cRutP & "inv073.rpt"
      CrystalReport1.WindowTitle = "Inv073 -- Control de Inventarios"
    End If
    
    Ubi_Tab CrystalReport1
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.SelectionFormula = cadena
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.Formulas(0) = "emp = '" & VGNemp & "'"
    CrystalReport1.Formulas(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
    CrystalReport1.Formulas(2) = "alm = '" & UCase(VGNomAlm) & "'"
    CrystalReport1.Formulas(3) = "campoini = '" & tex1 & "'"
    CrystalReport1.Formulas(4) = "campofin = '" & tex2 & "'"
    CrystalReport1.Formulas(5) = "detaini = '" & Va1 & "'"
    CrystalReport1.Formulas(6) = "detafin = '" & Va2 & "'"
    'CrystalReport1.Destination = crptToWindow
    If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
    Screen.MousePointer = 1
    Exit Sub
End If
    
If Text1 = "" Then
    MsgBox "Ingrese el codigo", vbExclamation, "Error"
    Text1.SetFocus
    Screen.MousePointer = 1
    Exit Sub
End If

If OpArt.Value Then           'Un select
    If Text2 <> "" Then
        Codigo2 = Text2         '  "23134671"
        cadena = cadena & " and ({STKART.STCODIGO} in '" & Codigo1 & "' to '" & Codigo2 & "')"
        
        SQL = SQL & " and STCODIGO between '" & Codigo1 & "' and '" & Codigo2 & "'"
    Else
        Codigo2 = Codigo1: Va2 = Va1
        cadena = cadena & " and {STKART.STCODIGO} = '" & Codigo1 & "' "
        
        SQL = SQL & " and STCODIGO='" & Codigo1 & "'"
    End If
    
    Adodc1.Open SQL, cConexCom, adOpenDynamic, adLockOptimistic
    If Adodc1.RecordCount = 0 Then
      MsgBox "No existen artículos para este rango", vbOKOnly, "Aviso"
      Screen.MousePointer = 1
      Adodc1.Close
      Exit Sub
    End If
    Adodc1.Close
    
    If Option1(0).Value Then
      CrystalReport1.WindowTitle = "Inv040 -- Control de Inventarios"
      CrystalReport1.ReportFileName = cRutP & "inv061.rpt"
    ElseIf Option1(1).Value Then
      CrystalReport1.WindowTitle = "Inv0456 -- Control de Inventarios"
      CrystalReport1.ReportFileName = cRutP & "inv056.rpt"
    ElseIf Option1(2).Value Then
      CrystalReport1.WindowTitle = "Inv073 -- Control de Inventarios"
      CrystalReport1.ReportFileName = cRutP & "inv073.rpt"
    End If
    
    Adodc3.Open "Select ADescri from MAEART Where ACodigo='" & Codigo1 & "'", cConexCom, adOpenDynamic, adLockOptimistic
    If Adodc3.RecordCount = 1 Then Va1 = Adodc3("Adescri")
    Adodc3.Close
    
    Adodc3.Open "Select ADescri from MAEART Where ACodigo='" & Codigo2 & "'", cConexCom, adOpenDynamic, adLockOptimistic
    If Adodc3.RecordCount = 1 Then Va2 = Adodc3("Adescri")
    Adodc3.Close
    
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
    CrystalReport1.Formulas(2) = "alm = '" & UCase(VGNomAlm) & "'"
    CrystalReport1.Formulas(3) = "campoini = '" & Codigo1 & "'"
    CrystalReport1.Formulas(4) = "campofin = '" & Codigo2 & "'"
    CrystalReport1.Formulas(5) = "detaini = '" & Va1 & "'"
    CrystalReport1.Formulas(6) = "detafin = '" & Va2 & "'"
    If CrystalReport1.Status <> 2 Then
        CrystalReport1.Action = 1
    End If
End If
Screen.MousePointer = 1
End Sub

Private Sub Imprimir2()
Dim Codigo1 As String
Dim Codigo2 As String
Dim cadena As String
Dim SQL As String
Dim tex1 As String, tex2 As String
Dim Va1 As String, Va2 As String
Screen.MousePointer = 11

Codigo1 = UCase(Trim(Text1))
Codigo2 = UCase(Trim(Text2))
Set Adodc1 = New ADODB.Recordset
Set Adodc3 = New ADODB.Recordset

cadena = "{STKART.STALMA}='" & VGAlma & "'"
    
SQL = "Select AFamilia,Fam_Nombre from "
SQL = SQL & "((MAEART A Inner Join STKART B on A.ACodigo=B.STCodigo) "
SQL = SQL & "Left Join FAMILIA C on A.AFamilia=C.Fam_Codigo) "
SQL = SQL & "Where Stalma='" & VGAlma & "' "

If OpTodos.Value Then
    SQL = SQL & " Order by AFamilia"
    
    Adodc3.Open SQL, cConexCom, adOpenDynamic, adLockOptimistic
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
      MsgBox "    No existen artículos      ", vbOKOnly, "Aviso"
      Adodc3.Close
      Screen.MousePointer = 1
      Exit Sub
    End If
    Adodc3.Close
    
    If Option1(0).Value Then
      CrystalReport1.WindowTitle = "Inv062 -- Control de Inventarios"
      CrystalReport1.ReportFileName = cRutP & "inv062.rpt"
    ElseIf Option1(1).Value Then
      CrystalReport1.WindowTitle = "Inv057 -- Control de Inventarios"
      CrystalReport1.ReportFileName = cRutP & "inv057.rpt"
    ElseIf Option1(2).Value Then
      CrystalReport1.WindowTitle = "Inv074 -- Control de Inventarios"
      CrystalReport1.ReportFileName = cRutP & "inv074.rpt"
    End If
    
    Ubi_Tab CrystalReport1
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.SelectionFormula = cadena
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.Formulas(0) = "emp = '" & VGNemp & "'"
    CrystalReport1.Formulas(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
    CrystalReport1.Formulas(2) = "alm = '" & UCase(VGNomAlm) & "'"
    CrystalReport1.Formulas(3) = "campoini = '" & tex1 & "'"
    CrystalReport1.Formulas(4) = "campofin = '" & tex2 & "'"
    CrystalReport1.Formulas(5) = "detaini = '" & Va1 & "'"
    CrystalReport1.Formulas(6) = "detafin = '" & Va2 & "'"
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.WindowTitle = mensaje1
    If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
    Screen.MousePointer = 1
    Exit Sub
End If
    
If Text1 = "" Then
    MsgBox "Ingrese el codigo", vbExclamation, "Error"
    Text1.SetFocus
    Screen.MousePointer = 1
    Exit Sub
End If

If Option2.Value Then           'Un select
    If Text2 <> "" Then
        Codigo2 = Text2         '  "23134671"
        cadena = cadena & " and ({MAEART.AFAMILIA} in '" & Codigo1 & "' to '" & Codigo2 & "')"
        
        SQL = SQL & " and AFAMILIA between '" & Codigo1 & "' and '" & Codigo2 & "'"
    Else
        Codigo2 = Codigo1: Va2 = Va1
        cadena = cadena & " and {MAEART.AFAMILIA} = '" & Codigo1 & "' "
        
        SQL = SQL & " and AFAMILIA='" & Codigo1 & "'"
    End If
    
    Adodc1.Open SQL, cConexCom, adOpenDynamic, adLockOptimistic
    If Adodc1.RecordCount = 0 Then
      MsgBox "No existen artículos para este rango", vbOKOnly, "Aviso"
      Screen.MousePointer = 1
      Adodc1.Close
      Exit Sub
    End If
    Adodc1.Close
    
    If Option1(0).Value Then
      CrystalReport1.WindowTitle = "Inv062 -- Control de Inventarios"
      CrystalReport1.ReportFileName = cRutP & "inv062.rpt"
    ElseIf Option1(1).Value Then
      CrystalReport1.WindowTitle = "Inv057 -- Control de Inventarios"
      CrystalReport1.ReportFileName = cRutP & "inv057.rpt"
    ElseIf Option1(2).Value Then
      CrystalReport1.WindowTitle = "Inv074 -- Control de Inventarios"
      CrystalReport1.ReportFileName = cRutP & "inv074.rpt"
    End If
    
    Adodc3.Open "Select Fam_Nombre from FAMILIA Where Fam_Codigo='" & Codigo1 & "'", cConexCom, adOpenDynamic, adLockOptimistic
    If Adodc3.RecordCount = 1 Then Va1 = Adodc3("Fam_Nombre")
    Adodc3.Close
    
    Adodc3.Open "Select Fam_Nombre from FAMILIA Where Fam_Codigo='" & Codigo2 & "'", cConexCom, adOpenDynamic, adLockOptimistic
    If Adodc3.RecordCount = 1 Then Va2 = Adodc3("Fam_Nombre")
    Adodc3.Close
    
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
    CrystalReport1.Formulas(2) = "alm = '" & UCase(VGNomAlm) & "'"
    CrystalReport1.Formulas(3) = "campoini = '" & Codigo1 & "'"
    CrystalReport1.Formulas(4) = "campofin = '" & Codigo2 & "'"
    CrystalReport1.Formulas(5) = "detaini = '" & Va1 & "'"
    CrystalReport1.Formulas(6) = "detafin = '" & Va2 & "'"
    If CrystalReport1.Status <> 2 Then
        CrystalReport1.Action = 1
    End If
End If
Screen.MousePointer = 1
End Sub

Private Sub Imprimir3()
Dim Codigo1 As String
Dim Codigo2 As String
Dim cadena As String
Dim SQL As String
Dim tex1 As String, tex2 As String
Dim Va1 As String, Va2 As String
Screen.MousePointer = 11

If Trim(Text3) = "" Then
    MsgBox "Ingrese el código de la familia", vbExclamation, "Error"
    Screen.MousePointer = 1: Text3.SetFocus
    Exit Sub
End If

Codigo1 = UCase(Trim(Text1))
Codigo2 = UCase(Trim(Text2))
Set Adodc1 = New ADODB.Recordset
Set Adodc3 = New ADODB.Recordset

cadena = "{STKART.STALMA}='" & VGAlma & "' and {MAEART.AFAMILIA}='" & Text3.text & "'"
    
SQL = "Select AModelo,Lin_Nombre from "
SQL = SQL & "((MAEART A Inner Join STKART B on A.ACodigo=B.STCodigo) "
SQL = SQL & "Left Join LINEAS C on A.AFamilia=C.Fam_Codigo and A.Amodelo=C.Lin_Codigo) "
SQL = SQL & "Where AFamilia='" & Text3.text & "' and Stalma='" & VGAlma & "'"

If OpTodos.Value Then
    SQL = SQL & " Order by AModelo"
    
    Adodc3.Open SQL, cConexCom, adOpenDynamic, adLockOptimistic
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
      MsgBox "    No existen artículos      ", vbOKOnly, "Aviso"
      Adodc3.Close
      Screen.MousePointer = 1
      Exit Sub
    End If
    Adodc3.Close
    
    If Option1(0).Value Then
      CrystalReport1.WindowTitle = "Inv064 -- Control de Inventarios"
      CrystalReport1.ReportFileName = cRutP & "inv064.rpt"
    ElseIf Option1(1).Value Then
      CrystalReport1.WindowTitle = "Inv060 -- Control de Inventarios"
      CrystalReport1.ReportFileName = cRutP & "inv060.rpt"
    ElseIf Option1(2).Value Then
      CrystalReport1.WindowTitle = "Inv076 -- Control de Inventarios"
      CrystalReport1.ReportFileName = cRutP & "inv076.rpt"
    End If

    Ubi_Tab CrystalReport1
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.SelectionFormula = cadena
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.Formulas(0) = "emp = '" & VGNemp & "'"
    CrystalReport1.Formulas(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
    CrystalReport1.Formulas(2) = "alm = '" & UCase(VGNomAlm) & "'"
    CrystalReport1.Formulas(3) = "campoini = '" & tex1 & "'"
    CrystalReport1.Formulas(4) = "campofin = '" & tex2 & "'"
    CrystalReport1.Formulas(5) = "detaini = '" & Va1 & "'"
    CrystalReport1.Formulas(6) = "detafin = '" & Va2 & "'"
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.WindowTitle = mensaje1
    If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
    Screen.MousePointer = 1
    Exit Sub
End If
    
If Text1 = "" Then
    MsgBox "Ingrese el codigo", vbExclamation, "Error"
    Text1.SetFocus
    Screen.MousePointer = 1
    Exit Sub
End If

If Option3.Value Then           'Un select
    If Text2 <> "" Then
        Codigo2 = Text2         '  "23134671"
        cadena = cadena & " and ({MAEART.AMODELO} in '" & Codigo1 & "' to '" & Codigo2 & "')"
        
        SQL = SQL & " and AMODELO between '" & Codigo1 & "' and '" & Codigo2 & "'"
    Else
        Codigo2 = Codigo1: Va2 = Va1
        cadena = cadena & " and {MAEART.AMODELO} = '" & Codigo1 & "' "
        
        SQL = SQL & " and AMODELO='" & Codigo1 & "'"
    End If
    
    Adodc1.Open SQL, cConexCom, adOpenDynamic, adLockOptimistic
    If Adodc1.RecordCount = 0 Then
      MsgBox "No existen artículos para este rango", vbOKOnly, "Aviso"
      Screen.MousePointer = 1
      Adodc1.Close
      Exit Sub
    End If
    Adodc1.Close
    
    If Option1(0).Value Then
      CrystalReport1.WindowTitle = "Inv064 -- Control de Inventarios"
      CrystalReport1.ReportFileName = cRutP & "inv064.rpt"
    ElseIf Option1(1).Value Then
      CrystalReport1.WindowTitle = "Inv060 -- Control de Inventarios"
      CrystalReport1.ReportFileName = cRutP & "inv060.rpt"
    ElseIf Option1(2).Value Then
      CrystalReport1.WindowTitle = "Inv076 -- Control de Inventarios"
      CrystalReport1.ReportFileName = cRutP & "inv076.rpt"
    End If
    
    Adodc3.Open "Select Lin_Nombre from LINEAS Where Lin_Codigo='" & Codigo1 & "'", cConexCom, adOpenDynamic, adLockOptimistic
    If Adodc3.RecordCount = 1 Then Va1 = Adodc3("Lin_Nombre")
    Adodc3.Close
    
    Adodc3.Open "Select Lin_Nombre from LINEAS Where Lin_Codigo='" & Codigo2 & "'", cConexCom, adOpenDynamic, adLockOptimistic
    If Adodc3.RecordCount = 1 Then Va2 = Adodc3("Lin_Nombre")
    Adodc3.Close
    
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
    CrystalReport1.Formulas(2) = "alm = '" & UCase(VGNomAlm) & "'"
    CrystalReport1.Formulas(3) = "campoini = '" & Codigo1 & "'"
    CrystalReport1.Formulas(4) = "campofin = '" & Codigo2 & "'"
    CrystalReport1.Formulas(5) = "detaini = '" & Va1 & "'"
    CrystalReport1.Formulas(6) = "detafin = '" & Va2 & "'"
    If CrystalReport1.Status <> 2 Then
        CrystalReport1.Action = 1
    End If
End If
Screen.MousePointer = 1
End Sub

Private Sub Imprimir4()
Dim Codigo1 As String
Dim Codigo2 As String
Dim cadena As String
Dim SQL As String
Dim tex1 As String, tex2 As String
Dim Va1 As String, Va2 As String
Screen.MousePointer = 11

If Trim(Text3) = "" Then
      MsgBox "Ingrese el código de la familia", vbExclamation, "Error"
      Screen.MousePointer = 1: Text3.SetFocus
      Exit Sub
ElseIf Trim(Text4) = "" Then
      MsgBox "Ingrese el código de la Línea", vbExclamation, "Error"
      Screen.MousePointer = 1: Text4.SetFocus
      Exit Sub
End If

Codigo1 = UCase(Trim(Text1))
Codigo2 = UCase(Trim(Text2))
Set Adodc1 = New ADODB.Recordset
Set Adodc3 = New ADODB.Recordset

cadena = "{STKART.STALMA}='" & VGAlma & "' and {MAEART.AFAMILIA}='" & Text3.text & "' and {MAEART.AMODELO}='" & Text4.text & "'"
    
SQL = "Select AGrupo,Acodigo,Gru_Nombre from "
SQL = SQL & "((MAEART A Inner Join STKART B on A.ACodigo=B.STCodigo) "
SQL = SQL & "Left Join GRUPO C on A.AFamilia=C.Fam_Codigo and A.Amodelo=C.Lin_Codigo and A.Agrupo=C.Gru_Codigo) "
SQL = SQL & "Where AFamilia='" & Text3.text & "' and Amodelo='" & Text4.text & "' and Stalma='" & VGAlma & "' "

If OpTodos.Value Then
    SQL = SQL & " Order by AGrupo"
    
    Adodc3.Open SQL, cConexCom, adOpenDynamic, adLockOptimistic
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
      MsgBox "    No existen artículos      ", vbOKOnly, "Aviso"
      Adodc3.Close
      Screen.MousePointer = 1
      Exit Sub
    End If
    Adodc3.Close
    
    If Option1(0).Value Then
      CrystalReport1.WindowTitle = "Inv063 -- Control de Inventarios"
      CrystalReport1.ReportFileName = cRutP & "inv063.rpt"
    ElseIf Option1(1).Value Then
       CrystalReport1.WindowTitle = "Inv058 -- Control de Inventarios"
      CrystalReport1.ReportFileName = cRutP & "inv058.rpt"
    ElseIf Option1(2).Value Then
       CrystalReport1.WindowTitle = "Inv075 -- Control de Inventarios"
      CrystalReport1.ReportFileName = cRutP & "inv075.rpt"
    End If
    
    Ubi_Tab CrystalReport1
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.SelectionFormula = cadena
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.Formulas(0) = "emp = '" & VGNemp & "'"
    CrystalReport1.Formulas(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
    CrystalReport1.Formulas(2) = "alm = '" & UCase(VGNomAlm) & "'"
    CrystalReport1.Formulas(3) = "campoini = '" & tex1 & "'"
    CrystalReport1.Formulas(4) = "campofin = '" & tex2 & "'"
    CrystalReport1.Formulas(5) = "detaini = '" & Va1 & "'"
    CrystalReport1.Formulas(6) = "detafin = '" & Va2 & "'"
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.WindowTitle = mensaje1
    If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
    Screen.MousePointer = 1
    Exit Sub
End If
    
If Text1 = "" Then
    MsgBox "Ingrese el codigo", vbExclamation, "Error"
    Text1.SetFocus
    Screen.MousePointer = 1
    Exit Sub
End If

If Option4.Value Then           'Un select
    If Text2 <> "" Then
        Codigo2 = Text2         '  "23134671"
        cadena = cadena & " and ({MAEART.AGRUPO} in '" & Codigo1 & "' to '" & Codigo2 & "')"
        
        SQL = SQL & " and AGRUPO between '" & Codigo1 & "' and '" & Codigo2 & "'"
    Else
        Codigo2 = Codigo1: Va2 = Va1
        cadena = cadena & " and {MAEART.AGRUPO} = '" & Codigo1 & "' "
        
        SQL = SQL & " and AGRUPO='" & Codigo1 & "'"
    End If
    
    Adodc1.Open SQL, cConexCom, adOpenDynamic, adLockOptimistic
    If Adodc1.RecordCount = 0 Then
      MsgBox "No existen artículos para este rango", vbOKOnly, "Aviso"
      Screen.MousePointer = 1
      Adodc1.Close
      Exit Sub
    End If
    Adodc1.Close
    
    If Option1(0).Value Then
       CrystalReport1.WindowTitle = "Inv063 -- Control de Inventarios"
      CrystalReport1.ReportFileName = cRutP & "inv063.rpt"
    ElseIf Option1(1).Value Then
       CrystalReport1.WindowTitle = "Inv058-- Control de Inventarios"
       CrystalReport1.ReportFileName = cRutP & "inv058.rpt"
    ElseIf Option1(2).Value Then
        CrystalReport1.WindowTitle = "Inv075 -- Control de Inventarios"
       CrystalReport1.ReportFileName = cRutP & "inv075.rpt"
    End If
    
    Adodc3.Open "Select Gru_Nombre from Grupo Where Gru_Codigo='" & Codigo1 & "'", cConexCom, adOpenDynamic, adLockOptimistic
    If Adodc3.RecordCount = 1 Then Va1 = Adodc3("Gru_Nombre")
    Adodc3.Close
    
    Adodc3.Open "Select Gru_Nombre from Grupo Where Gru_Codigo='" & Codigo2 & "'", cConexCom, adOpenDynamic, adLockOptimistic
    If Adodc3.RecordCount = 1 Then Va2 = Adodc3("Gru_Nombre")
    Adodc3.Close
    
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
    CrystalReport1.Formulas(2) = "alm = '" & UCase(VGNomAlm) & "'"
    CrystalReport1.Formulas(3) = "campoini = '" & Codigo1 & "'"
    CrystalReport1.Formulas(4) = "campofin = '" & Codigo2 & "'"
    CrystalReport1.Formulas(5) = "detaini = '" & Va1 & "'"
    CrystalReport1.Formulas(6) = "detafin = '" & Va2 & "'"
    If CrystalReport1.Status <> 2 Then
        CrystalReport1.Action = 1
    End If
End If
Screen.MousePointer = 1
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

Private Sub OpArt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then OpTodos.SetFocus
End Sub

Private Sub OpRango_Click()
If OpRango.Value Then
    Text1.Enabled = True
    Text2.Enabled = True
    Text1.SetFocus
End If
End Sub

Private Sub Option1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then OpArt.SetFocus
End Sub

Private Sub Option2_Click()
Option2.Value = True
FrameRep.Caption = " Por Familias "
OpTodos.Caption = "Todos las Familias"
limpiar_t1_t2
Text3 = "": Text4 = ""
OpTodos.Top = 300: OpRango.Top = 650
Text1.Top = 1100: Label2.Top = 1100
Text2.Top = 1500: Label3.Top = 1500
End Sub

Private Sub Option2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then OpTodos.SetFocus
End Sub

Private Sub Option3_Click()
Option3.Value = True
FrameRep.Caption = " Por Lineas "
OpTodos.Caption = "Todos las Líneas "
limpiar_t1_t2
Text3 = "": Text4 = ""
Label4.Visible = True
Text3.Visible = True
OpTodos.Top = 550: OpRango.Top = 900
Text1.Top = 1200: Label2.Top = 1200
Text2.Top = 1600: Label3.Top = 1600
End Sub

Private Sub Option3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then OpTodos.SetFocus
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

Private Sub Option4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then OpTodos.SetFocus
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
If OpArt.Value Then
         Adodc2.Open "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD from MaeArt p, StkArt n Where n.STSKDIS<> 0 and  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & VGAlma & "'  ", cConexCom, adOpenStatic, adLockOptimistic
         frmReferencia.Conectar Adodc2, "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD from MaeArt p, StkArt n Where n.STSKDIS<> 0 and  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & VGAlma & "'  "
         frmReferencia.Label1.Caption = "Artículos"
         frmReferencia.Show vbModal
         Adodc2.Close
         If vGUtil(1) <> "" Then
                 Text1 = (vGUtil(1))
         End If
         If Text1 <> "" And Text2 <> "" And Text1 > Text2 Then
                 MsgBox "Ingrese un codigo menor al fin ", vbOKOnly, "Error"
                 Exit Sub
        End If
        If Text1 <> "" Then
                 Text2.Enabled = True
                 Text2.SetFocus
        End If
ElseIf Option2.Value Then
        Adodc2.Open "SELECT FAM_CODIGO,FAM_NOMBRE FROM FAMILIA", cConexCom, adOpenStatic, adLockOptimistic
        frmReferencia.Conectar Adodc2, "SELECT FAM_CODIGO,FAM_NOMBRE FROM FAMILIA"
        frmReferencia.Label1.Caption = "Familias de Artículos"
        frmReferencia.Show vbModal
        Adodc2.Close
        If vGUtil(1) <> "" Then
                Text1 = (vGUtil(1))
        End If
ElseIf Option3.Value Then
        Adodc2.Open "SELECT LIN_CODIGO,LIN_NOMBRE FROM LINEAS Where Fam_Codigo='" & Trim(Text3) & "'", cConexCom, adOpenStatic, adLockOptimistic
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
        Adodc2.Open "SELECT GRU_CODIGO,GRU_NOMBRE FROM GRUPO Where Fam_Codigo='" & Trim(Text3) & "' and Lin_Codigo='" & Trim(Text4) & "'", cConexCom, adOpenStatic, adLockOptimistic
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

If OpArt.Value Then
    Adodc2.Open "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD from MaeArt p, StkArt n Where n.STSKDIS<> 0 and  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & VGAlma & "'  ", cConexCom, adOpenStatic, adLockOptimistic
    frmReferencia.Conectar Adodc2, "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD from MaeArt p, StkArt n Where n.STSKDIS<> 0 and  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & VGAlma & "'  "
    frmReferencia.Label1.Caption = "Artículos"
    frmReferencia.Show vbModal
    Adodc2.Close
    If vGUtil(1) <> "" Then
        Text2 = (vGUtil(1))
    End If
   If Text2 <> "" Then
        Command1.SetFocus
   End If
ElseIf Option2.Value Then
    Adodc2.Open "SELECT FAM_CODIGO,FAM_NOMBRE FROM FAMILIA", cConexCom, adOpenStatic, adLockOptimistic
    frmReferencia.Conectar Adodc2, "SELECT FAM_CODIGO,FAM_NOMBRE FROM FAMILIA"
    frmReferencia.Label1.Caption = "Familias de Artículos"
    frmReferencia.Show vbModal
    Adodc2.Close
    If vGUtil(1) <> "" Then
      Text2 = (vGUtil(1))
    End If
ElseIf Option3.Value Then
        Adodc2.Open "SELECT LIN_CODIGO,LIN_NOMBRE FROM LINEAS Where Fam_Codigo='" & Trim(Text3) & "'", cConexCom, adOpenStatic, adLockOptimistic
        frmReferencia.Conectar Adodc2, "SELECT LIN_CODIGO,LIN_NOMBRE FROM LINEAS Where Fam_Codigo='" & Trim(Text3) & "'"
        frmReferencia.Label1.Caption = "Líneas de Artículos"
        frmReferencia.Show vbModal
        Adodc2.Close
        If vGUtil(1) <> "" Then
            Text2 = (vGUtil(1))
        End If
ElseIf Option4.Value Then
        Adodc2.Open "SELECT GRU_CODIGO,GRU_NOMBRE FROM GRUPO Where Fam_Codigo='" & Trim(Text3) & "' and Lin_Codigo='" & Trim(Text4) & "'", cConexCom, adOpenStatic, adLockOptimistic
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
Dim RSQL As String
RSQL = "select  ACODIGO FROM maeart where ACODIGO = '" & text & "'" '
'Set db = Workspaces(0).OpenDatabase(cRuta2)
'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)

Set rs = cConexCom.Execute(RSQL)
If Not rs.EOF Then
    Existe_cod_art = rs(0)
Else
    MsgBox "El tipo de codigo no existe !", vbOKOnly, "Error"
    Existe_cod_art = ""
End If
rs.Close
End Function

Private Sub limpiar_t1_t2()
Text1 = "": Text2 = ""
Label4.Visible = False
Text3.Visible = False
Label5.Visible = False
Text4.Visible = False
End Sub

Private Sub Text3_DblClick()
Dim Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset
If Option3.Value Or Option4.Value Then
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
  If KeyAscii = 13 And Text3 <> "" Then
     If Text4.Visible = True Then Text4.SetFocus Else OpTodos.SetFocus
  End If
End Sub

Private Sub Text4_DblClick()
Dim Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset
If Option4.Value Then
         Adodc2.Open "Select LIN_CODIGO,LIN_NOMBRE from LINEAS Where FAM_CODIGO='" & Text3 & "'", cConexCom, adOpenStatic, adLockOptimistic
         frmReferencia.Conectar Adodc2, "Select LIN_CODIGO,LIN_NOMBRE from LINEAS Where FAM_CODIGO='" & Text3 & "'"
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

