VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FormKardex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kardex de Articulos"
   ClientHeight    =   3420
   ClientLeft      =   1980
   ClientTop       =   1845
   ClientWidth     =   6540
   Icon            =   "Formkardex.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3420
   ScaleWidth      =   6540
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   240
      TabIndex        =   19
      Top             =   0
      Width           =   6015
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   210
         Width           =   3855
      End
      Begin VB.Label Label4 
         Caption         =   "Almacen"
         Height          =   255
         Left            =   480
         TabIndex        =   20
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
      Left            =   3735
      Picture         =   "Formkardex.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2535
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   735
      Left            =   2055
      Picture         =   "Formkardex.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2535
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Artículos"
      ForeColor       =   &H80000006&
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   6015
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   4320
         MaxLength       =   20
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1320
         TabIndex        =   17
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   41025537
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   35796
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   4320
         TabIndex        =   18
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   41025537
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   35796
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha :"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Left            =   3120
         TabIndex        =   15
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Desde"
         Height          =   375
         Left            =   270
         TabIndex        =   14
         Top             =   1110
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   375
         Left            =   3120
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Clase"
      ForeColor       =   &H80000006&
      Height          =   1215
      Left            =   2370
      TabIndex        =   8
      Top             =   960
      Width           =   3255
      Begin VB.OptionButton Option5 
         Caption         =   "Movimiento y Stock"
         Height          =   255
         Left            =   1635
         TabIndex        =   11
         Top             =   180
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Solo Movimientos"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   360
         Width           =   2415
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Todos los Artículos"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   720
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo"
      ForeColor       =   &H80000006&
      Height          =   1215
      Left            =   570
      TabIndex        =   5
      Top             =   870
      Width           =   2445
      Begin VB.OptionButton Option2 
         Caption         =   "Resumen"
         Height          =   255
         Left            =   435
         TabIndex        =   7
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Detalle"
         Height          =   255
         Left            =   435
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FormKardex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim Kar As New ADODB.Recordset

'El kardex de articulo utiliza
'la tabla KardexAux para almacena los resultado
'Cuando es Solo movimiento indica los articulos que tuvieron movimiento en ese mes
'Todos los articulos indica que muestra aquellos articulos que no tuvieron movimiento

Public Sub kardex2()
    Dim cn As New ADODB.Connection
    
    Dim RSQL As String
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
    
    RSQL = " SELECT b.DECODIGO, a.CATD, a.CAFECDOC, a.CACODMOV, a.CANUMDOC, case a.CATIPMOV when 'S' then b.DECANTID else 0 end AS Salida, case a.CATIPMOV when 'I' then b.DECANTID else 0 end  AS Ingresos, '101' AS stockfin, a.CAALMA + a.CATD + a.CANUMDOC, a.CAHORA, b.DELOTE, b.DESERIE,a.CASITGUI,CARFTDOC,CARFNDOC, " & _
           " canumdoc,caalma FROM MovAlmCab a INNER JOIN MovAlmDet b " & _
           " ON a.CAALMA = b.DEALMA  AND a.CATD = b.DETD AND a.CANUMDOC = b.DENUMDOC " & _
           " INNER JOIN MaeArt  maeart ON b.DECODIGO = MaeArt.ACODIGO " & _
           " WHERE b.decodigo >= '" & Text1 & "' and isnull(a.casitgui,'')<>'A' and b.decodigo<= '" & Text2 & "' and b.dealma = '" & balma & "' and " & _
           "  a.CAFECDOC>='" & DTPicker1.Value & "' and a.CAFECDOC<='" & DTPicker2.Value & "' " & _
           " ORDER BY b.DECODIGO,a.CAFECDOC,a.CATIPMOV,a.CARFTDOC,a.CARFNDOC,a.catd,a.canumdoc"
           
    Kar.Open "kardexaux", VGCNx, adOpenDynamic, adLockOptimistic
    Set rs = VGCNx.Execute(RSQL)
    If rs.RecordCount > 0 Then
    
        rs.MoveFirst
        xGrup = Trim(rs!decodigo)
        xSum = cantidadmes(xGrup, DTPicker1, empxalm)
        xAcu = xSum
        xcantfecha = xSum
        
        Kar.AddNew
        Kar!c1 = Trim(rs(0))
        Kar!c2 = "**"
        Kar!c3 = DTPicker1 - 1
        Kar!c4 = "**"
        Kar!c5 = "Saldo Inicial"
        Kar!c6 = 0
        Kar!c7 = 0
        Kar!c8 = xSum
        Kar!c9 = xAcu
        Kar!C10 = ""
        Kar!c11 = ""
        Kar!tipdocrf = ""
        Kar!numdocrf = ""
        
        Kar.Update

        Do Until rs.EOF
            If Trim(rs!decodigo) <> Trim(xGrup) Then
                xGrup = Trim(rs!decodigo)
                xSum = cantidadmes(xGrup, DTPicker1, empxalm)
                xcantfecha = xSum
                xAcu = xSum
                Kar.AddNew
                Kar!c1 = Trim(xGrup)
                Kar!c2 = "**"
                Kar!c3 = DTPicker1 - 1
                Kar!c4 = "**"
                Kar!c5 = "Saldo Inicial"
                Kar!c6 = 0
                Kar!c7 = 0
                Kar!c8 = xSum
                Kar!c9 = xAcu
                Kar!C10 = ""
                Kar!c11 = ""
                Kar!tipdocrf = ""
                Kar!numdocrf = ""
                Kar.Update
            End If
            xSum = xSum + ESNULO(rs!ingresos, 0) - ESNULO(rs!Salida, 0)
            Kar.AddNew
            Kar!c1 = Trim(rs(0))
            Kar!c2 = rs(1)
            Kar!c3 = rs(2)
            Kar!c4 = rs(3)
            Kar!c5 = Format(rs(4), "0000000000")
            Kar!c6 = rs(5)
            Kar!c7 = rs(6)
            Kar!c8 = xSum
            Kar!c9 = xAcu
            Kar!C10 = rs!cahora
            Kar!c11 = IIf(IsNull(rs!DESERIE), (rs!DELOTE), (rs!DESERIE))
            Kar!c12 = rs!CANUMDOC
            Kar!tipdocrf = rs!CARFTDOC
            Kar!numdocrf = ESNULO(rs!CARFNDOC, "")
            Kar!alma = rs!CAALMA
            Kar.Update
            rs.MoveNext
        Loop
    End If
    Kar.Close
    
    Set Kar = Nothing
    
    rs.Close
    Set rs = Nothing
        
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
    
    VGForm1 = 5
    central Me
    DTPicker1 = DateAdd("m", -2, Date)
    DTPicker2.Value = Date
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
          CrystalReport1.Reset
           CrystalReport1.WindowTitle = "Inv034 -- Control de Inventarios"
           CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv034.rpt"
       CrystalReport1.DiscardSavedData = True
       
 
        If VGsql = 1 Then
           CrystalReport1.Connect = "Provider=SQLOLEDB;Password=" & VGParamSistem.PwdGEN & ";Persist Security Info=True;User ID=" & VGParamSistem.UsuarioGEN & ";Initial Catalog=" & VGParamSistem.BDEmpresaGEN & ";SERVER=" & VGParamSistem.ServidorGEN
          Else
           CrystalReport1.Connect = VGcadenareport2

        End If

       
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
       CrystalReport1.formulas(5) = "emp ='" & VGparametros.NomEmpresa & "'"
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
  VGForm1 = 5
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
 Dim rs As New ADODB.Recordset
 Dim RSQL As String
  RSQL = "select  ACODIGO FROM maeart where ACODIGO = '" & text & "'" '
  'Set db = Workspaces(0).OpenDatabase(cRuta2)
  'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
  
  Set rs = VGCNx.Execute(RSQL)
  If Not rs.EOF Then
       Existe_cod_art = rs(0)
  Else
       MsgBox "El tipo de codigo no existe !", vbOKOnly, "Error"
       Existe_cod_art = ""
  End If
  rs.Close
End Function

Function SI_HAY_STOCK(txt As String) As Boolean
 Dim rs As New ADODB.Recordset
 Dim RSQL As String
  RSQL = "select  n.STSKDIS from  StkArt n where  n.STCODIGO ='" & txt & "'and n.STSKDIS<>0 and  n.STALMA = '" & VGAlma & "' "
  'Set db = Workspaces(0).OpenDatabase(cRuta2)
  'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
  
  Set rs = VGCNx.Execute(RSQL)
  If Not rs.EOF Then
      SI_HAY_STOCK = True
  Else
      SI_HAY_STOCK = False
  End If
   rs.Close
End Function

'Function cantidadmes(Codigo As String, annomes As String) As Double
Function cantidadmes(codigo As String, FINI As Date, alma As String) As Double
 Dim RSQL As String
 Dim RSB As Recordset
 Dim ingre, sale As Double
 'RSQL = "select sum(SMCANENT) as col1,sum(SMCANSAL) as col2 from MoResMes where  SMALMA = '" & VGAlma & "'AND SMCODIGO= '" & Codigo & "' AND SMMESPRO <= '" & annomes & "'"  '
 'Set db = Workspaces(0).OpenDatabase(cRuta2)
 'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
 
 RSQL = "select sum(case catipmov when 'I' then decantid else 0 end) as ingreso,sum(case catipmov when 'S' then decantid else 0 end) as salida from movalmdet a inner join movalmcab b " & _
        " on dealma=caalma and detd=catd and denumdoc=canumdoc " & _
        " where decodigo='" & codigo & "' and dealma='" & alma & "' and cafecdoc<'" & CStr(FINI) & "' " & _
        " and isnull(casitgui,'')<>'A'"
        
 Set RSB = VGCNx.Execute(RSQL)

 If Not RSB.EOF Then
    cantidadmes = IIf(IsNull(RSB(0)), 0, RSB(0)) - IIf(IsNull(RSB(1)), 0, RSB(1))
 Else
    cantidadmes = 0
 End If
 RSB.Close
End Function

