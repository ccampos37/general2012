VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmKardexLote 
   Caption         =   "Kardex por Lote"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6270
   LinkTopic       =   "Form2"
   ScaleHeight     =   3330
   ScaleWidth      =   6270
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Artículos"
      ForeColor       =   &H80000006&
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   6015
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   4320
         MaxLength       =   20
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   94961665
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   35796
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   4320
         TabIndex        =   9
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   94961665
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   35796
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   375
         Left            =   3120
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Desde"
         Height          =   375
         Left            =   270
         TabIndex        =   12
         Top             =   1110
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Left            =   3120
         TabIndex        =   11
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha :"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   735
      Left            =   1935
      Picture         =   "frmKardexLote.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2535
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   3615
      Picture         =   "frmKardexLote.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2535
      Width           =   735
   End
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   210
         Width           =   3855
      End
      Begin VB.Label Label4 
         Caption         =   "Almacen"
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   510
      Top             =   2550
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo"
      ForeColor       =   &H80000006&
      Height          =   1215
      Left            =   450
      TabIndex        =   19
      Top             =   870
      Width           =   2445
      Begin VB.OptionButton Option2 
         Caption         =   "Resumen"
         Height          =   255
         Left            =   435
         TabIndex        =   20
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Detalle"
         Height          =   255
         Left            =   435
         TabIndex        =   21
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Clase"
      ForeColor       =   &H80000006&
      Height          =   1215
      Left            =   2250
      TabIndex        =   15
      Top             =   960
      Width           =   3255
      Begin VB.OptionButton Option3 
         Caption         =   "Todos los Artículos"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   720
         Width           =   2655
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Solo Movimientos"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   360
         Width           =   2415
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Movimiento y Stock"
         Height          =   255
         Left            =   1635
         TabIndex        =   16
         Top             =   180
         Visible         =   0   'False
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmKardexLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As New ADODB.Connection
Dim Rs As New ADODB.Recordset
Dim Kar As New ADODB.Recordset

'El kardex de articulo utiliza
'la tabla KardexAux para almacena los resultado
'Cuando es Solo movimiento indica los articulos que tuvieron movimiento en ese mes
'Todos los articulos indica que muestra aquellos articulos que no tuvieron movimiento

Public Sub kardex2()
       
    Dim rsql As String
    Dim empxalm As String
    Dim mes As String * 2
    Dim anno As String * 4
    Dim annomes As String * 6
    Dim annomes1 As String * 6
    Dim annomes2 As String * 6
    Dim fecreg As String * 6
    Dim sigue As Boolean
    Dim xAcu As Double
    Dim xGrup As String, xSum As Double, xcantfecha As Double
    Dim balma As String
    Dim puntero As Integer
    
    puntero = InStr(Combo1.text, "-")
    If puntero > 1 Then
       balma = Left(Combo1.text, puntero - 1)
       empxalm = balma
    Else
        Exit Sub
    End If
    

    VGCNx.Execute "DELETE FROM  kardexaux"
    
    rsql = " SELECT b.DECODIGO, a.CATD, a.CAFECDOC, a.CACODMOV, a.CANUMDOC, case a.CATIPMOV when 'S' then b.DECANTID else 0 end AS Salida, case a.CATIPMOV when 'I' then b.DECANTID else 0 end  AS Ingresos, " & _
           " '101' AS stockfin, a.CAALMA + a.CATD + a.CANUMDOC, a.CAHORA, b.DELOTE, b.DESERIE,a.CASITGUI,CARFTDOC,CARFNDOC,catipotransf,canrotransf,stscodprov " & _
           " FROM MovAlmCab a INNER JOIN MovAlmDet b " & _
           " ON a.CAALMA = b.DEALMA  AND a.CATD = b.DETD AND a.CANUMDOC = b.DENUMDOC " & _
           " INNER JOIN MaeArt  maeart ON b.DECODIGO = MaeArt.ACODIGO " & _
           " INNER JOIN StkLote StkLote ON b.delote = StkLote.StsLote and b.decodigo = stkLote.stscodigo " & _
           " WHERE b.decodigo >= '" & Text1 & "' and a.casitgui<>'A' and b.decodigo<= '" & Text2 & "' and b.dealma = '" & balma & "' and " & _
           "  a.CAFECDOC>='" & DTPicker1.Value & "' and a.CAFECDOC<='" & DTPicker2.Value & "' " & _
           " ORDER BY b.DECODIGO,a.CAFECDOC,a.CATIPMOV,a.catd,a.canumdoc"
           
    Kar.Open "kardexaux", VGCNx, adOpenDynamic, adLockOptimistic
    Set Rs = VGCNx.Execute(rsql)
    
    If Rs.RecordCount > 0 Then
    
        Rs.MoveLast
        Rs.MoveFirst
        xGrup = Trim(Rs!decodigo)
        xSum = cantidadmes(xGrup, DTPicker1, empxalm)
        xAcu = xSum
        xcantfecha = xSum
        


        Do Until Rs.EOF
            If Trim(Rs!decodigo) <> Trim(xGrup) Then
                xGrup = Trim(Rs!decodigo)
                xSum = cantidadmes(xGrup, DTPicker1, empxalm)
                xcantfecha = xSum
                xAcu = xSum

            End If
            xSum = xSum + Rs!ingresos - Rs!Salida
            Kar.AddNew
            Kar!c1 = Trim(Rs(0))
            Kar!c2 = Rs(1)
            Kar!c3 = Rs(2)
            Kar!c4 = Rs(3)
            Kar!c5 = Format(Rs(4), "0000000000")
            Kar!c6 = Rs(5)
            Kar!c7 = Rs(6)
            Kar!c8 = xSum
            Kar!c9 = xAcu
            Kar!C10 = Rs!cahora
            Kar!c11 = IIf(IsNull(Rs!DESERIE), (Rs!DELOTE), (Rs!DESERIE))
            Kar!tipdocrf = "" & IIf(RTrim(Rs!CARFTDOC) = "", Rs!catipotransf, Rs!CARFTDOC)
            Kar!numdocrf = "" & IIf(RTrim(Rs!CARFTDOC) = "", Rs!caNROtransf, Rs!CARFNDOC)
            Kar!NOMREFE = "" & Rs!stscodprov
            
            Kar.Update
            Rs.MoveNext
        Loop
    End If
    Kar.Close
    
    Set Kar = Nothing
    
    Rs.Close
    Set Rs = Nothing
    
    
End Sub







Private Sub Command8_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Dim rsc As New ADODB.Recordset
    
    Option1.Value = True
    Option4.Value = True
    Set rsc = VGCNx.Execute("Select  TAALMA,TADESCRI  from  tabalm")
    If rsc.RecordCount > 0 Then
        Combo1.Clear
        rsc.MoveFirst
        Do Until rsc.EOF
            Combo1.AddItem rsc.Fields(0) & "-" & IIf(IsNull(rsc.Fields(1)), "****", rsc.Fields(1))
            rsc.MoveNext
        Loop
    End If
    rsc.Close
     
    Set rsc = Nothing
    
    VGForm1 = 23
    central Me
    DTPicker1 = DateAdd("m", -2, Date)
    DTPicker2.Value = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Rs.Close
    'Db.Close
End Sub

Private Sub Command1_Click()
Dim Codigo2 As String

      
        Screen.MousePointer = 11
        If DTPicker1.Value > DTPicker2.Value Then
            MsgBox "Ingrese la Fecha correcta", vbInformation, mensaje1
            DTPicker1.SetFocus
            Screen.MousePointer = 1
            Exit Sub
        End If
       Codigo2 = Devolver_Dato(1, VGAlma, "TABALM", "TAALMA", False, "TADESCRI")
       If Trim(Text1) = "" Or Trim(Text2) = "" Then
            MsgBox "Ingrese el codigo del Articulo", vbInformation, mensaje1
            Screen.MousePointer = 1
            Exit Sub
       End If
       If Existe_cod_art(Text1) = "" Then
            Text1.SetFocus
            Screen.MousePointer = 1
            Exit Sub
       End If
       If Existe_cod_art(Text2) = "" Then
            Text2.SetFocus
            Screen.MousePointer = 1
            Exit Sub
       End If
       Call kardex2
     '  kardex
       'If Option1.Value Then
           CrystalReport1.Reset
           CrystalReport1.WindowTitle = "RptKardexLote -- Control de Inventarios"
           CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "RptKardexLote.rpt"
'       Else
'           CrystalReport1.WindowTitle = "Inv033 -- Control de Inventarios"
'           CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv033.rpt"
'       End If
       Ubi_Tab CrystalReport1
       CrystalReport1.DiscardSavedData = True
       
                        
    'CrystalReport1.Connect = VGcadenareport2
       CrystalReport1.Connect = "Provider=SQLOLEDB;Password=" & VGParamSistem.PwdGEN & ";Persist Security Info=True;User ID=" & VGParamSistem.UsuarioGEN & ";Initial Catalog=" & VGParamSistem.BDEmpresaGEN & ";SERVER=" & VGParamSistem.ServidorGEN
       
       CrystalReport1.Destination = crptToWindow
       CrystalReport1.WindowState = crptMaximized
       CrystalReport1.WindowShowPrintBtn = True
       CrystalReport1.WindowShowRefreshBtn = True
       CrystalReport1.WindowShowSearchBtn = True
       CrystalReport1.WindowShowPrintSetupBtn = True
       CrystalReport1.StoredProcParam(0) = Trim(VGCNx.DefaultDatabase)
       CrystalReport1.formulas(0) = "almacen ='" & Trim(Combo1) & "'"
       CrystalReport1.formulas(1) = "artinicio='" & Text1 & "'"
       CrystalReport1.formulas(2) = "artfin ='" & Text2 & "'"
       CrystalReport1.formulas(3) = "fechainicio ='" & DTPicker1 & "'"
       CrystalReport1.formulas(4) = "fechafin ='" & DTPicker2 & "'"
       CrystalReport1.formulas(5) = "emp ='" & VGparametros.RucEmpresa & "'"
       CrystalReport1.formulas(6) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
      
       If CrystalReport1.Status <> 2 Then
          
          CrystalReport1.Action = 1
       End If
       Screen.MousePointer = 1
End Sub

Private Sub Option5_Click()
  'SI HA TENIDO MOVIMIENTO HASTA LA FECHA
End Sub

Private Sub Text1_DblClick()
  VGForm1 = 23
  VGAlma = Left(Combo1.text, 2)
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

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then Text1_DblClick
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 And Text1 <> "" Then
      Text1 = UCase(Text1)
      If Existe_cod_art(Text1) <> "" Then
              Text2.Enabled = True
              Text2.SetFocus
      End If
  Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub Text2_DblClick()
   FormAyuArt1.Show 1
   If Text2 <> "" Then
        Command1.SetFocus
   End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then Text2_DblClick
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 And Trim(Text2) <> "" Then
      Text2 = UCase(Text2)
      If Existe_cod_art(Text2) <> "" Then
            If Text1 > Text2 Then
                       MsgBox "El codigo fin debe ser mayor que el inicio", vbExclamation, "Aviso"
                        Exit Sub
           End If
           Command1.SetFocus
      End If
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Function Existe_cod_art(text As TextBox) As String
 Dim Rs As New ADODB.Recordset
 Dim rsql As String
  rsql = "select  ACODIGO FROM maeart where ACODIGO = '" & text & "'" '
  'Set db = Workspaces(0).OpenDatabase(cRuta2)
  'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
  
  Set Rs = VGCNx.Execute(rsql)
  If Not Rs.EOF Then
       Existe_cod_art = Rs(0)
  Else
       MsgBox "El tipo de codigo no existe !", vbOKOnly, "Error"
       Existe_cod_art = ""
  End If
  Rs.Close
End Function

Function SI_HAY_STOCK(txt As String) As Boolean
 Dim Rs As New ADODB.Recordset
 Dim rsql As String
  rsql = "select  n.STSKDIS from  StkArt n where  n.STCODIGO ='" & txt & "'and n.STSKDIS<>0 and  n.STALMA = '" & VGAlma & "' "
  'Set db = Workspaces(0).OpenDatabase(cRuta2)
  'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
  
  Set Rs = VGCNx.Execute(rsql)
  If Not Rs.EOF Then
      SI_HAY_STOCK = True
  Else
      SI_HAY_STOCK = False
  End If
   Rs.Close
End Function

'Function cantidadmes(Codigo As String, annomes As String) As Double
Function cantidadmes(codigo As String, FINI As Date, alma As String) As Double
 Dim rsql As String
 Dim RSB As Recordset
 Dim ingre, sale As Double
 'RSQL = "select sum(SMCANENT) as col1,sum(SMCANSAL) as col2 from MoResMes where  SMALMA = '" & VGAlma & "'AND SMCODIGO= '" & Codigo & "' AND SMMESPRO <= '" & annomes & "'"  '
 'Set db = Workspaces(0).OpenDatabase(cRuta2)
 'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
 
 rsql = "select sum(case catipmov when 'I' then decantid else 0 end) as ingreso,sum(case catipmov when 'S' then decantid else 0 end) as salida from movalmdet a inner join movalmcab b " & _
        " on dealma=caalma and detd=catd and denumdoc=canumdoc " & _
        " where decodigo='" & codigo & "' and dealma='" & alma & "' and cafecdoc<'" & CStr(FINI) & "' " & _
        " and casitgui<>'A'"
        
 Set RSB = VGCNx.Execute(rsql)

 If Not RSB.EOF Then
    cantidadmes = IIf(IsNull(RSB(0)), 0, RSB(0)) - IIf(IsNull(RSB(1)), 0, RSB(1))
 Else
    cantidadmes = 0
 End If
 RSB.Close
End Function



