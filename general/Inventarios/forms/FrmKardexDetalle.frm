VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmKardexDetalle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kardex de Articulos con Referencias"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "FrmKardexDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   288
      Top             =   2052
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Caption         =   "Artículos"
      ForeColor       =   &H80000006&
      Height          =   3405
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   6015
      Begin VB.CheckBox Check1 
         Caption         =   "Todos los Almacenes"
         Height          =   195
         Left            =   1320
         TabIndex        =   14
         Top             =   720
         Width           =   3255
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   360
         Width           =   3180
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
         Height          =   732
         Left            =   2280
         Picture         =   "FrmKardexDetalle.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2520
         Width           =   735
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Salir"
         Height          =   732
         Left            =   3225
         Picture         =   "FrmKardexDetalle.frx":0D0C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2535
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   4320
         MaxLength       =   20
         TabIndex        =   4
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   2
         Top             =   1200
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Top             =   1920
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   52428801
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   36161
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   4320
         TabIndex        =   9
         Top             =   1920
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   52428801
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   36161
      End
      Begin VB.Label Label1 
         Caption         =   "Almacen :"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha :"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Left            =   3120
         TabIndex        =   6
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Desde"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   375
         Left            =   3120
         TabIndex        =   3
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   1200
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmKardexDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset, Kar  As ADODB.Recordset
Dim TempoTab As String

Private Sub kardex()
On Error GoTo Mensaje
    Dim rsql        As String
    Dim empxalm     As String * 4
    Dim mes         As String * 2
    Dim anno        As String * 4
    Dim annomes     As String * 6
    Dim annomes1    As String * 6
    Dim annomes2    As String * 6
    Dim fecreg      As String * 6
    Dim sigue       As Boolean
    Dim xAcu        As Double
    
    
    rsql = " SELECT MovAlmDet.DECODIGO, MovAlmCab.CATD, MovAlmCab.CAFECDOC, MovAlmCab.CACODMOV, MovAlmCab.CANUMDOC, " & _
    "(CASE MovAlmCab.CATIPMOV WHEN 'S' THEN MovAlmDet.DECANTID ELSE 0 END) AS Salida, " & _
    "(CASE MovAlmCab.CATIPMOV WHEN 'I' THEN MovAlmDet.DECANTID ELSE 0 END) AS Ingresos, 101 AS stockfin, " & _
    "MovAlmCab.CAALMA + MovAlmCab.CATD + MovAlmCab.CANUMDOC AS Expr1, MovAlmCab.CAHORA, MovAlmDet.DELOTE, " & _
    " MovAlmDet.DESERIE,MovAlmCab.CARFTDOC,MovAlmCab.CARFNDOC,CANOMPRO,CANOMCLI,CASITGUI,CACODCLI"
    rsql = rsql & ",CARFALMA, movalmdet.dealma"
    rsql = rsql & " FROM (MovAlmCab INNER JOIN MovAlmDet ON "
    rsql = rsql & " (MovAlmCab.CAALMA = MovAlmDet.DEALMA) AND (MovAlmCab.CATD = MovAlmDet.DETD) AND (MovAlmCab.CANUMDOC = MovAlmDet.DENUMDOC)) "
    rsql = rsql & " INNER JOIN MaeArt ON MovAlmDet.DECODIGO = MaeArt.ACODIGO "
    
    rsql = rsql & " WHERE (((MovAlmCab.CASITGUI)<>'A')) and movalmdet.decodigo >= '" & Text1 & "' and movalmdet.decodigo<= '" & Text2 & "'"
    If Me.Check1.Value = 0 Then
        rsql = rsql & " AND movalmdet.dealma = '" & Trim(Left(Combo1.text, 4)) & "'"
    End If
    rsql = rsql & " ORDER BY MovAlmDet.DECODIGO, MovAlmCab.CAFECDOC, MovAlmCab.CATIPMOV, MovAlmCab.CATD, MovAlmCab.CANUMDOC"
    
    Rem mvv RSQL = RSQL & " WHERE (((MovAlmCab.CASITGUI)<>'A')) and movalmdet.decodigo >= '" & Text1 & "' and movalmdet.decodigo<= '" & Text2 & "' and movalmdet.dealma = '" & VGAlma & "'"
    Set rs = New ADODB.Recordset
    rs.Open rsql, VGCNx, adOpenStatic, adLockOptimistic
    
    
    Dim xGrup As String, xSum As Double, xcantfecha As Double
    
    VGCNx.Execute "DELETE FROM  kardexaux"
    Set Kar = New ADODB.Recordset
    Kar.Open "select * from kardexaux", VGCNx, adOpenDynamic, adLockOptimistic
     
    If Month(DTPicker1) = 1 Then
      mes = "12"
      anno = Format(Year(DTPicker1) - 1, "0000")
    Else
      mes = Format(Month(DTPicker1) - 1, "00")
      anno = Format(Year(DTPicker1), "0000")
    End If
    annomes = anno & mes
    'empxalm = VGAlma
    If Not rs.EOF Then
        xGrup = rs!decodigo
        xSum = cantidadmes(xGrup, annomes)
        xAcu = xSum
        xcantfecha = xSum
        Do While Not rs.EOF
            'If Left(rs(8), 2) <> Trim(empxalm) Then
                'rs.MoveNext
            'Else
                If Not (cNull(rs!CATD) = "GS" And cNull(rs!cacodmov) = "GF" And cNull(rs!CASITGUI) = "F") Then  'rS!catd <> "GS" And rS!cacodmov <> "GF"
                    fecreg = Format(Year(rs!CAFECDOC), "0000") & Format(Month(rs!CAFECDOC), "00")
                    annomes1 = Year(DTPicker1) & Format(Month(DTPicker1), "00")
                    annomes2 = Year(DTPicker2) & Format(Month(DTPicker2), "00")
                    If (rs!decodigo >= Text1) And (rs!decodigo <= Text2) And (fecreg >= annomes1) And (fecreg <= annomes2) Then  'coment  (fecreg >= annomes1) And (fecreg <= annomes2)
                        sigue = True
                        If sigue Then
                            If rs!decodigo <> xGrup Then
                                xGrup = rs!decodigo
                                xSum = cantidadmes(xGrup, annomes)
                                xcantfecha = xSum
                                xAcu = xSum
                            End If
                            If (rs!CAFECDOC < DTPicker1) Then                   '(diadeinicio >= Rs!cafecdoc) And (Rs!cafecdoc < DTPicker1) Then
                                xcantfecha = xcantfecha + rs!ingresos - rs!Salida
                                xSum = xcantfecha
                                xAcu = xcantfecha
                            Else
                                xSum = xSum + rs!ingresos - rs!Salida
                                If (rs(2) >= DTPicker1) And (rs(2) <= DTPicker2) Then
                                    Kar.AddNew
                                    Kar!c1 = rs(0)
                                    Kar!c2 = rs(1)
                                    Kar!c3 = rs(2)
                                    Kar!c4 = rs(3)
                                    Kar!c5 = Format(rs(4), "0000000000")
                                    Kar!c6 = rs(5)
                                    Kar!c7 = rs(6)
                                    Kar!c8 = xSum
                                    Kar!c9 = xAcu
                                    Kar!C10 = rs!cahora
                                    Kar!c11 = IIf(IsNull(rs!DESERIE), (rs!DELOTE), IIf(Trim(rs!DESERIE) = "", (rs!DELOTE), (rs!DESERIE)))
                                    Kar!tipdocrf = cNull(rs!CARFTDOC)
                                    Kar!numdocrf = cNull(rs!CARFNDOC)
                                    If Trim(cNull(rs!CANOMpro)) = "" Then
                                        If Trim(cNull(rs!CACODCLI)) <> "" Then
                                            Kar!NOMREFE = rs!CACODCLI
                                        ElseIf Kar!c4 = "TD" Then
   '                                         Kar!NOMREFE = funcESNULO(Devolver_DatoInv(1, rs!CARFALMA, "TABALM", "TAALMA", False, "TADESCRI"), "")
                                        Else
                                            Kar!NOMREFE = ""
                                        End If
                                    Else
                                        Kar!NOMREFE = cNull(rs!CANOMpro)
                                    End If
                                    Kar!alma = rs!DEALMA
                                    Kar.Update
                                End If
                            End If    ' if de sigue
                        End If  'If del anno
                    End If
                    rs.MoveNext
                Else
                    rs.MoveNext
                End If
            'End If
        Loop
    End If
    rs.Close
Exit Sub
Mensaje:
    Captura_error
End Sub

Sub TemporalKardex()
On Error GoTo Mensaje
Dim NTRAN As Byte
    If ExisteElem(0, VGCNx, TempoTab) Then VGCNx.Execute "DROP TABLE " & TempoTab
    SQL = "select * INTO " & TempoTab & " from KARDEXAUX INNER JOIN MAEART ON  MAEART.ACODIGO = KARDEXAUX.c1 "
    NTRAN = 1
    VGCNx.BeginTrans
    VGCNx.Execute SQL
    VGCNx.CommitTrans
    NTRAN = 0
Exit Sub
Mensaje:
    Captura_error
    If NTRAN = 1 Then VGCNx.RollbackTrans
End Sub

Private Sub Check1_Click()
If Me.Check1.Value = 1 Then
    Me.Combo1.Enabled = False
Else
    Me.Combo1.Enabled = True
End If
End Sub

Private Sub Command8_Click()
  Unload Me
End Sub
Sub Carga_Almacen()
Dim rsql      As String
Dim rs        As ADODB.Recordset
Dim I         As Integer
 
rsql = "select TAALMA,TADESCRI FROM TabAlm "
Set rs = New ADODB.Recordset
rs.Open rsql, VGCNx, adOpenStatic, adLockOptimistic

While Not rs.EOF
  Combo1.AddItem (rs(0)) & Space(4) & (rs(1))
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
End Sub
Private Sub Form_Load()
On Error GoTo Mensaje
    TempoTab = "##" + ComputerName
    VGForm1 = 23
    central Me
    Carga_Almacen
    DTPicker1 = DateAdd("m", -2, VG_FecTrab)
    DTPicker2.Value = VG_FecTrab
Exit Sub
Mensaje:
    Captura_error
End Sub

Private Sub Command1_Click()
Dim Codigo2 As String
Dim arrform(1) As Variant, arrparm(2) As Variant
Dim NombreRep As String, CadOrden As String
Screen.MousePointer = 11
If DTPicker1.Value > DTPicker2.Value Then
   MsgBox "Ingrese la Fecha correcta", vbInformation, mensaje1
   DTPicker1.SetFocus
   Screen.MousePointer = 1
   Exit Sub
End If
'       Codigo2 = Devolver_DatoInv(1, VGAlma, "TABALM", "TAALMA", False, "TADESCRI")
If Trim(Text1) = "" Or Trim(Text2) = "" Then
   MsgBox "Ingrese el codigo del Articulo", vbInformation, mensaje1
   Screen.MousePointer = 1
   Exit Sub
End If
'*****************************
       kardex
       TemporalKardex
'*****************************
CrystalReport1.Reset
NombreRep = "inv505.rpt"
If Me.Check1.Value = 0 Then
   CrystalReport1.formulas(0) = "almacen ='" & Trim(Combo1) & "'"
 Else
  CrystalReport1.formulas(0) = "@almacen='TODOS LOS ALMACENES'"
End If

Call PropCrystal(CrystalReport1)
CrystalReport1.StoredProcParam(0) = VGCNx.DefaultDatabase
CrystalReport1.StoredProcParam(1) = "kardexaux"

CrystalReport1.formulas(1) = "artinicio='" & Text1 & "'"
CrystalReport1.formulas(2) = "artfin ='" & Text2 & "'"
CrystalReport1.formulas(3) = "fechainicio ='" & DTPicker1 & "'"
CrystalReport1.formulas(4) = "fechafin ='" & DTPicker2 & "'"

CrystalReport1.LogOnServer "pdssql.dll", VGParamSistem.ServidorGEN, VGParamSistem.BDEmpresaGEN, VGParamSistem.UsuarioGEN, ""
CrystalReport1.Connect = VGcadenareport2

CrystalReport1.ReportFileName = VGParamSistem.RutaReport & NombreRep
If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
'Call ImpresionRptProc(NombreRep, arrform, arrparm, CadOrden, "Inv505 ")
Screen.MousePointer = 1
Exit Sub
ImprimeRegCompras:
    MsgBox Err.Description
End Sub

Private Sub Text1_DblClick()
    FormAyuArt1.Show 1
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
   End If
End Sub

Private Sub Text2_DblClick()
   FormAyuArt1.Show 1
    If vGUtil(1) <> "" Then
        Text2 = (vGUtil(1))
    End If

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
   End If
   
End Sub

Function Existe_cod_art(text As TextBox) As String
 
 Dim rs As ADODB.Recordset
 Dim rsql As String
  
  rsql = "select  ACODIGO FROM maeart where ACODIGO = '" & text & "'" '
  Set rs = New ADODB.Recordset
  rs.Open rsql, cConexCom, adOpenStatic, adLockOptimistic
  
  If Not rs.EOF Then
       Existe_cod_art = rs(0)
  Else
       MsgBox "El tipo de codigo no existe !", vbOKOnly, "Error"
       Existe_cod_art = ""
  End If
  rs.Close
  
End Function

Function SI_HAY_STOCK(txt As String) As Boolean
 Dim rs As ADODB.Recordset
 Dim rsql As String
  
  rsql = "select  n.STSKDIS from  StkArt n where  n.STCODIGO ='" & txt & "'and n.STSKDIS<>0 and  n.STALMA = '" & VGAlma & "' "
  Set rs = New ADODB.Recordset
  rs.Open rsql, cConexCom, adOpenStatic, adLockOptimistic
    
  If Not rs.EOF Then
      SI_HAY_STOCK = True
  Else
      SI_HAY_STOCK = False
  End If
   rs.Close
End Function

Function cantidadmes(codigo As String, annomes As String) As Double
 Dim rsql As String
 Dim rs As ADODB.Recordset
 
 'RSQL = "select sum(SMCANENT) as col1,sum(SMCANSAL) as col2 from MoResMes where  SMALMA = '" & VGAlma & "'AND SMCODIGO= '" & codigo & "' AND SMMESPRO <= '" & annomes & "'"  '
 'RSQL = "select sum(SMCANENT) as col1,sum(SMCANSAL) as col2 from MoResMes where  SMCODIGO= '" & codigo & "' AND SMMESPRO <= '" & annomes & "'"  '
    If Me.Check1.Value = 0 Then
        rsql = "select sum(SMCANENT) as col1,sum(SMCANSAL) as col2 from MoResMes where  SMCODIGO= '" & codigo & "' AND SMMESPRO <= '" & annomes & "' AND SMALMA = '" & Trim(Left(Me.Combo1.text, 4)) & "'"  '
    Else
        rsql = "select sum(SMCANENT) as col1,sum(SMCANSAL) as col2 from MoResMes where  SMCODIGO= '" & codigo & "' AND SMMESPRO <= '" & annomes & "'"  '
    End If
 Set rs = New ADODB.Recordset
 rs.Open rsql, VGCNx, adOpenStatic, adLockOptimistic
 
 If Not rs.EOF Then
    cantidadmes = IIf(IsNull(rs(0)), 0, rs(0)) - IIf(IsNull(rs(1)), 0, rs(1))
 Else
    cantidadmes = 0
 End If
 rs.Close
End Function
