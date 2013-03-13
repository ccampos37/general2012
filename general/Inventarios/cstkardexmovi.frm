VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form cstkardexmovi 
   Caption         =   "Consulta de Movimientos por Articulo"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9975
   LinkTopic       =   "Form2"
   ScaleHeight     =   6330
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   160
      TabIndex        =   20
      Top             =   0
      Width           =   5175
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   210
         Width           =   3975
      End
      Begin VB.Label Label4 
         Caption         =   "Almacen"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MGrid1 
      Height          =   3855
      Left            =   150
      TabIndex        =   19
      Top             =   2370
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   6800
      _Version        =   393216
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "Artículos"
      ForeColor       =   &H80000006&
      Height          =   1575
      Left            =   180
      TabIndex        =   9
      Top             =   690
      Width           =   5325
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   11
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3630
         MaxLength       =   20
         TabIndex        =   10
         Top             =   360
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   50135041
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   35796
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   3630
         TabIndex        =   13
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   50135041
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   35796
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   375
         Left            =   2970
         TabIndex        =   17
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Desde"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Left            =   2970
         TabIndex        =   15
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha :"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo"
      ForeColor       =   &H80000006&
      Height          =   975
      Left            =   5550
      TabIndex        =   6
      Top             =   90
      Width           =   3255
      Begin VB.OptionButton Option1 
         Caption         =   "Detalle"
         Height          =   255
         Left            =   450
         TabIndex        =   8
         Top             =   210
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Resumen"
         Height          =   255
         Left            =   435
         TabIndex        =   7
         Top             =   570
         Width           =   1605
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Clase"
      ForeColor       =   &H80000006&
      Height          =   1065
      Left            =   5550
      TabIndex        =   2
      Top             =   1080
      Width           =   3255
      Begin VB.OptionButton Option3 
         Caption         =   "Todos los Artículos"
         Height          =   255
         Left            =   390
         TabIndex        =   5
         Top             =   720
         Width           =   2655
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Solo Movimientos"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   2415
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Movimiento y Stock"
         Height          =   315
         Left            =   1605
         TabIndex        =   3
         Top             =   210
         Visible         =   0   'False
         Width           =   1485
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   825
      Left            =   8880
      Picture         =   "cstkardexmovi.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   915
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&Salir"
      Height          =   795
      Left            =   8880
      Picture         =   "cstkardexmovi.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1260
      Width           =   945
   End
End
Attribute VB_Name = "cstkardexmovi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim Kar As New ADODB.Recordset
Dim rsbusca As New ADODB.Recordset

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
    Dim xAcu, k As Double
    Dim xGrupo As String, xSum As Double, xcantfecha As Double
    Dim ningresa, nsalidas As Double
    Dim balma As String
    Dim puntero As Integer
    
    puntero = InStr(Combo1.text, "-")
    If puntero > 1 Then
       balma = Left(Combo1.text, puntero - 1)
       empxalm = balma
    Else
        Exit Sub
    End If
    
    VGcnx.Execute "DELETE FROM  kardexaux"
    Set rs = VGcnx.Execute("select * from sysobjects where name like 'i1%'")
    If rs.RecordCount > 0 Then
       VGcnx.Execute "drop table [dbo].[i1]"
    End If
    rs.Close
    Set rs = Nothing

    rsql = " SELECT MovAlmDet.DECODIGO, MovAlmCab.CATD, MovAlmCab.CAFECDOC, MovAlmCab.CACODMOV, MovAlmCab.CANUMDOC, case MovAlmCab.CATIPMOV when 'S' then MovAlmDet.DECANTID else 0 end AS Salida, case MovAlmCab.CATIPMOV when 'I' then MovAlmDet.DECANTID else 0 end  AS Ingresos, '101' AS stockfin, MovAlmCab.CAALMA + MovAlmCab.CATD + MovAlmCab.CANUMDOC, MovAlmCab.CAHORA, MovAlmDet.DELOTE, MovAlmDet.DESERIE,MovAlmCab.CASITGUI,CARFTDOC,CARFNDOC " & _
           " FROM MovAlmCab movalmcab INNER JOIN MovAlmDet movalmdet " & _
           " ON MovAlmCab.CAALMA = MovAlmDet.DEALMA  AND MovAlmCab.CATD = MovAlmDet.DETD AND MovAlmCab.CANUMDOC = MovAlmDet.DENUMDOC " & _
           " INNER JOIN MaeArt  maeart ON MovAlmDet.DECODIGO = MaeArt.ACODIGO " & _
           " WHERE movalmdet.decodigo >= '" & Text1 & "' and movalmdet.decodigo<= '" & Text2 & "' and MovAlmCab.casitgui<>'A' and movalmdet.dealma = '" & balma & "' and " & _
           "  MovAlmCab.CAFECDOC>='" & DTPicker1.Value & "' and MovAlmCab.CAFECDOC<='" & DTPicker2.Value & "' " & _
           " ORDER BY MovAlmDet.DECODIGO,MovAlmCab.CAFECDOC,MovAlmCab.CATIPMOV,MovAlmCab.catd,MovAlmCab.canumdoc"
           
  
    Set rs = VGcnx.Execute(rsql)
    
    Kar.Open "kardexaux", VGcnx, adOpenDynamic, adLockOptimistic
    
    If rs.RecordCount > 0 Then
    
   
        rs.MoveLast
        rs.MoveFirst
                
        xGrupo = Trim(rs!decodigo)
        xSum = cantidadmes(xGrupo, DTPicker1, empxalm)
        xAcu = xSum
        xcantfecha = xSum
        
        Kar.AddNew

        Kar!c1 = xGrupo
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
            If Trim(rs!decodigo) <> Trim(xGrupo) Then
                xGrupo = Trim(rs!decodigo)
                xSum = cantidadmes(xGrupo, DTPicker1, empxalm)
                xcantfecha = xSum
                xAcu = xSum
                Kar.AddNew

                Kar!c1 = Trim(xGrupo)
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
            xSum = xSum + rs!ingresos - rs!Salida
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
            Kar!tipdocrf = "" & rs!CARFTDOC
            Kar!numdocrf = "" & rs!CARFNDOC
            
            Kar.Update
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
    Kar.Close
    Set Kar = Nothing
    
    MGrid1.Clear
    Set rsbusca = Nothing
    Set rsbusca = VGcnx.Execute("select c1 as Articulo,c3 as Fecha,c2 as TP,c5 as Documento,c4 as TMovi,c7 as Ingreso,c6 as Salida,c8 as Saldo from kardexaux order by c1,c3,c5,c2")
    If rsbusca.RecordCount > 0 Then
        Call Configura_Grilla
        rsbusca.MoveLast
        rsbusca.MoveFirst
        rsql = rsbusca.Fields(0): ningresa = 0: nsalidas = 0
        Do Until rsbusca.EOF
            MGrid1.Rows = MGrid1.Rows + 1
            MGrid1.Row = MGrid1.Rows - 1
            For k = 0 To 7
              If k = 0 And (rsql <> rsbusca.Fields(0) Or MGrid1.Row = 1) Then
                  MGrid1.TextMatrix(MGrid1.Row, k) = rsbusca.Fields(k)
                  rsql = rsbusca.Fields(0)
              ElseIf k <> 0 Then
                 If k Like "[567]" Then
                    MGrid1.TextMatrix(MGrid1.Row, k) = "" & Format(rsbusca.Fields(k), "#,###,##0.00")
                 Else
                    MGrid1.TextMatrix(MGrid1.Row, k) = "" & rsbusca.Fields(k)
                 End If
              End If
              If k = 5 Then
                ningresa = ningresa + rsbusca.Fields(5)
              ElseIf k = 6 Then
                nsalidas = nsalidas + rsbusca.Fields(6)
              End If
            Next k
            rsbusca.MoveNext
            If Not rsbusca.EOF Then
                If rsql <> rsbusca.Fields(0) Then
                    MGrid1.Rows = MGrid1.Rows + 1
                    MGrid1.Row = MGrid1.Rows - 1
                    MGrid1.TextMatrix(MGrid1.Row, 4) = "TOTAL  ==>"
                    MGrid1.TextMatrix(MGrid1.Row, 5) = Format(ningresa, "#,###,##0.00")
                    MGrid1.Row = MGrid1.Row: MGrid1.Col = 5: MGrid1.CellBackColor = RGB(0, 120, 120)
                    MGrid1.TextMatrix(MGrid1.Row, 6) = Format(nsalidas, "#,###,##0.00")
                    MGrid1.Row = MGrid1.Row: MGrid1.Col = 6: MGrid1.CellBackColor = RGB(0, 120, 120)
                    'MGrid1.TextMatrix(MGrid1.Row, 7) = Format(MGrid1.TextMatrix(MGrid1.Rows - 2, 7), "#,###,##0.00")
                    MGrid1.Row = MGrid1.Row: MGrid1.Col = 7: MGrid1.CellBackColor = RGB(0, 120, 120)
                    MGrid1.Rows = MGrid1.Rows + 1
                    MGrid1.Row = MGrid1.Rows - 1
                    ningresa = 0: nsalidas = 0
                                        
                End If
                
            End If
        Loop
'        If RSQL <> rsbusca.Fields(0) Then
             MGrid1.Rows = MGrid1.Rows + 1
             MGrid1.Row = MGrid1.Rows - 1
             MGrid1.TextMatrix(MGrid1.Row, 4) = "TOTAL  ==>"
             MGrid1.TextMatrix(MGrid1.Row, 5) = Format(ningresa, "#,###,##0.00")
             MGrid1.Row = MGrid1.Row: MGrid1.Col = 5: MGrid1.CellBackColor = RGB(0, 120, 120)
             MGrid1.TextMatrix(MGrid1.Row, 6) = Format(nsalidas, "#,###,##0.00")
             MGrid1.Row = MGrid1.Row: MGrid1.Col = 6: MGrid1.CellBackColor = RGB(0, 120, 120)
             'MGrid1.TextMatrix(MGrid1.Row, 7) = Format(MGrid1.TextMatrix(MGrid1.Rows - 2, 7), "#,###,##0.00")
             MGrid1.Row = MGrid1.Row: MGrid1.Col = 7: MGrid1.CellBackColor = RGB(0, 120, 120)
             
             MGrid1.Rows = MGrid1.Rows + 1
             MGrid1.Row = MGrid1.Rows - 1
             ningresa = 0: nsalidas = 0
 '        End If
        
    End If

End Sub

Public Sub Configura_Grilla()
        MGrid1.Rows = 1: MGrid1.Cols = 8
        MGrid1.TextMatrix(0, 0) = "Articulo"
        MGrid1.TextMatrix(0, 1) = "Fecha"
        MGrid1.Row = 0: MGrid1.Col = 1: MGrid1.ColWidth(3) = 1500
        MGrid1.TextMatrix(0, 2) = "TP"
        MGrid1.TextMatrix(0, 3) = "Documento"
        MGrid1.Row = 0: MGrid1.Col = 3: MGrid1.ColWidth(3) = 1900
        MGrid1.TextMatrix(0, 4) = "TMov."
        MGrid1.TextMatrix(0, 5) = "Ingresos"
        MGrid1.TextMatrix(0, 6) = "Salidas"
        MGrid1.TextMatrix(0, 7) = "Saldos"
End Sub





Private Sub Command8_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    
    Dim rsc As New ADODB.Recordset
    
    Set rsc = VGcnx.Execute("Select  TAALMA,TADESCRI  from  tabalm")
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
        
    Call Configura_Grilla
    Option1.Value = True
    Option4.Value = True
    VGForm1 = 5
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

       Screen.MousePointer = 1
End Sub

Private Sub Option5_Click()
  'SI HA TENIDO MOVIMIENTO HASTA LA FECHA
End Sub

Private Sub Text1_DblClick()
  VGForm1 = 8
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
 Dim rsql As String
  rsql = "select  ACODIGO FROM maeart where ACODIGO = '" & text & "'" '
  'Set db = Workspaces(0).OpenDatabase(cRuta2)
  'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
  
  Set rs = VGcnx.Execute(rsql)
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
 Dim rsql As String
  rsql = "select  n.STSKDIS from  StkArt n where  n.STCODIGO ='" & txt & "'and n.STSKDIS<>0 and  n.STALMA = '" & VGAlma & "' "
  'Set db = Workspaces(0).OpenDatabase(cRuta2)
  'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
  
  Set rs = VGcnx.Execute(rsql)
  If Not rs.EOF Then
      SI_HAY_STOCK = True
  Else
      SI_HAY_STOCK = False
  End If
   rs.Close
End Function

'Function cantidadmes(Codigo As String,annomes As String) As Double
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
        
 Set RSB = VGcnx.Execute(rsql)

 If Not RSB.EOF Then
    cantidadmes = IIf(IsNull(RSB(0)), 0, RSB(0)) - IIf(IsNull(RSB(1)), 0, RSB(1))
 Else
    cantidadmes = 0
 End If
 RSB.Close
End Function


