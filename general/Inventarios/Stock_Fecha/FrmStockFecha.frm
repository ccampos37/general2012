VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmStockFecha 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock de Articulos"
   ClientHeight    =   3045
   ClientLeft      =   1980
   ClientTop       =   1845
   ClientWidth     =   6540
   Icon            =   "FrmStockFecha.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   6540
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   6015
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   210
         Width           =   3855
      End
      Begin VB.Label Label4 
         Caption         =   "Almacen"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   630
      Top             =   2550
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   3720
      Picture         =   "FrmStockFecha.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2175
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   735
      Left            =   2055
      Picture         =   "FrmStockFecha.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2175
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fecha :"
      ForeColor       =   &H80000006&
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   6015
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   24510465
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   35796
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   4320
         TabIndex        =   6
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   24510465
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   35796
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Left            =   3360
         TabIndex        =   4
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Desde"
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   510
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmStockFecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim Kar As New ADODB.Recordset

Private Sub Command8_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Dim rsc As New ADODB.Recordset
    Dim cConexCom As ADODB.Connection   'BdComun
    'Set cConexCom = New ADODB.Connection  'BD. Común
    'Public cConexCom As ADODB.Connection   'BdComun
    
    Call Configurar_Conexiones
    'cConexCom.Open
    Set rsc = cn.Execute("Select  TAALMA,TADESCRI  from  tabalm")
    If rsc.RecordCount > 0 Then
        Combo1.Clear
        rsc.MoveFirst
        Do Until rsc.EOF
            Combo1.AddItem rsc.Fields(0) & "-" & IIf(IsNull(rsc.Fields(1)), "****", rsc.Fields(1))
            rsc.MoveNext
        Loop
    End If
    rsc.Close
     cn.Close
    Set rsc = Nothing
    Set cn = Nothing
    'VGForm1 = 5
    'central Me
    DTPicker1 = DateAdd("m", -2, Date)
    DTPicker2.Value = Date
    
End Sub


Private Sub Command1_Click()
Dim Codigo2 As String
Dim cEmp As String
Dim puntero As Integer
Dim descri As String
      '*********
'         Dim busca As New dll_apisgen.dll_apis
'   Dim strbase As String
'   Dim struser As String
'   Dim strpass As String
'   Dim strserver As String
'
'   Dim strconecta As String
'
'   'RutaRep = busca.LeerIni(App.Path & "\Camtex.ini", "Reporte", "repo", "")
'   'RutaRepProc = busca.LeerIni(App.Path & "\Camtex.ini", "Reporte", "opera", "")
'
'
'   strbase = busca.LeerIni(App.Path & "\Camtex.ini", "Bdatos", "dbase", "")
'   struser = busca.LeerIni(App.Path & "\Camtex.ini", "Bdatos", "duser", "")
'   strpass = busca.LeerIni(App.Path & "\Camtex.ini", "Bdatos", "dpass", "")
'   strserver = busca.LeerIni(App.Path & "\Camtex.ini", "Bdatos", "dserver", "")
'
'
'
      Call Configurar_Conexiones
      
        If Trim(Combo1.Text) = "" Then
            Combo1.SetFocus
            Exit Sub
        
        Else
            cEmp = Left(Combo1.Text, 2)
            puntero = InStr(Combo1.Text, "-")
            descri = Right(Combo1.Text, Len(Combo1.Text) - puntero)
            
        End If
      
      '********
      
      
      
      
        Screen.MousePointer = 11
        If DTPicker1.Value > DTPicker2.Value Then
            MsgBox "Ingrese la Fecha correcta", vbInformation, "Mensaje"
            DTPicker1.SetFocus
            Screen.MousePointer = 1
            Exit Sub
        End If
      ' Codigo2 = Devolver_Dato(1, VGAlma, "TABALM", "TAALMA", False, "TADESCRI")


     '  kardex
       'If Option1.Value Then
           CrystalReport1.Reset
           CrystalReport1.WindowTitle = "Stock de Aticulos"
           CrystalReport1.ReportFileName = App.Path & "\inv036.rpt"
'       Else
'           CrystalReport1.WindowTitle = "Inv033 -- Control de Inventarios"
'           CrystalReport1.ReportFileName = cRutP & "inv033.rpt"
'       End If
       'Ubi_Tab CrystalReport1
       CrystalReport1.DiscardSavedData = True
       
    CrystalReport1.LogOnServer "pdssql.dll", _
                                strserver, _
                                strbase, _
                                struser, _
                                ""
    CrystalReport1.Connect = _
        "DSN=" & strserver & ";" & _
        "DSQ=" & strbase & ";" & _
        "UID=" & struser & ";" & _
        "PWD=''"
       
       CrystalReport1.Destination = crptToWindow
       CrystalReport1.WindowState = crptMaximized
       CrystalReport1.WindowShowPrintBtn = True
       CrystalReport1.WindowShowRefreshBtn = True
       CrystalReport1.WindowShowSearchBtn = True
       CrystalReport1.WindowShowPrintSetupBtn = True
       
       CrystalReport1.StoredProcParam(0) = strbase
       CrystalReport1.StoredProcParam(1) = cEmp
       CrystalReport1.StoredProcParam(2) = DTPicker1
       CrystalReport1.StoredProcParam(3) = DTPicker2
       
       CrystalReport1.Formulas(0) = "almacen ='" & descri & "'"
       CrystalReport1.Formulas(1) = "ini ='" & DTPicker1 & "'"
       CrystalReport1.Formulas(2) = "fin ='" & DTPicker2 & "'"
       CrystalReport1.Formulas(3) = "emp ='" & cEmp & "'"

      
      
      
      
       If CrystalReport1.Status <> 2 Then
          
          CrystalReport1.Action = 1
       End If
       Screen.MousePointer = 1
End Sub


'Function SI_HAY_STOCK(txt As String) As Boolean
' Dim rs As New ADODB.Recordset
' Dim RSQL As String
'  RSQL = "select  n.STSKDIS from  StkArt n where  n.STCODIGO ='" & txt & "'and n.STSKDIS<>0 and  n.STALMA = '" & VGAlma & "' "
'  'Set db = Workspaces(0).OpenDatabase(cRuta2)
'  'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
'
'  Set rs = cConexCom.Execute(RSQL)
'  If Not rs.EOF Then
'      SI_HAY_STOCK = True
'  Else
'      SI_HAY_STOCK = False
'  End If
'   rs.Close
'End Function

'Function cantidadmes(Codigo As String, annomes As String) As Double
Function cantidadmes(Codigo As String, FINI As Date, alma As String) As Double
 Dim RSQL As String
 Dim RSB As Recordset
 Dim ingre, sale As Double
 'RSQL = "select sum(SMCANENT) as col1,sum(SMCANSAL) as col2 from MoResMes where  SMALMA = '" & VGAlma & "'AND SMCODIGO= '" & Codigo & "' AND SMMESPRO <= '" & annomes & "'"  '
 'Set db = Workspaces(0).OpenDatabase(cRuta2)
 'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
 
 RSQL = "select sum(case catipmov when 'I' then decantid else 0 end) as ingreso,sum(case catipmov when 'S' then decantid else 0 end) as salida from movalmdet a inner join movalmcab b " & _
        " on dealma=caalma and detd=catd and denumdoc=canumdoc " & _
        " where decodigo='" & Trim(Codigo) & "' and dealma='" & Trim(alma) & "' and cafecdoc<'" & FINI & "' " & _
        " and catipmov<>'A'"
        
 Set RSB = cConexCom.Execute(RSQL)

 If Not RSB.EOF Then
    cantidadmes = IIf(IsNull(RSB(0)), 0, RSB(0)) - IIf(IsNull(RSB(1)), 0, RSB(1))
 Else
    cantidadmes = 0
 End If
 RSB.Close
End Function

