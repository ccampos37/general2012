VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FormValArtPed 
   Caption         =   "Valorización de Articulo Pendientes"
   ClientHeight    =   5730
   ClientLeft      =   2535
   ClientTop       =   2370
   ClientWidth     =   8040
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   5730
   ScaleWidth      =   8040
   Begin VB.CommandButton Command1 
      Caption         =   "&Validar"
      Height          =   735
      Left            =   2520
      Picture         =   "FormValArtPed.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   4905
      Width           =   840
   End
   Begin VB.Frame Frame1 
      Caption         =   "Valorizado"
      ForeColor       =   &H80000007&
      Height          =   4575
      Left            =   195
      TabIndex        =   1
      Top             =   150
      Width           =   7515
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "FormValArtPed.frx":0442
         Left            =   5400
         List            =   "FormValArtPed.frx":044C
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   2760
         Width           =   1455
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "FormValArtPed.frx":0460
         Left            =   2280
         List            =   "FormValArtPed.frx":046D
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   2280
         TabIndex        =   23
         Top             =   3840
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   5400
         TabIndex        =   21
         Top             =   3528
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2280
         TabIndex        =   20
         Top             =   3480
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   5400
         TabIndex        =   19
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2280
         TabIndex        =   18
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label20 
         Caption         =   "Label20"
         Height          =   255
         Left            =   3000
         TabIndex        =   31
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label Label19 
         Caption         =   "Label1"
         Height          =   255
         Left            =   3000
         TabIndex        =   28
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label18 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   285
         Left            =   2280
         TabIndex        =   27
         Top             =   2400
         Width           =   4575
      End
      Begin VB.Label Label17 
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   480
         TabIndex        =   26
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label16 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   288
         Left            =   2280
         TabIndex        =   25
         Top             =   2052
         Width           =   1452
      End
      Begin VB.Label Label15 
         Caption         =   "Codigo Art."
         Height          =   255
         Left            =   480
         TabIndex        =   24
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Total"
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   22
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Cantidad"
         Height          =   255
         Index           =   2
         Left            =   4080
         TabIndex        =   17
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Costo Unitario"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   16
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Documento"
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Transaccion"
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Proveedor"
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Doc Referencial"
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Serie"
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Moneda"
         Height          =   255
         Left            =   4080
         TabIndex        =   10
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Factura"
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Conversion"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Tipo de Cambio"
         Height          =   255
         Index           =   0
         Left            =   4080
         TabIndex        =   7
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   285
         Left            =   2280
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lbltransa 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   285
         Left            =   2280
         TabIndex        =   5
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label12 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   285
         Left            =   2280
         TabIndex        =   4
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   285
         Left            =   2280
         TabIndex        =   3
         Top             =   960
         Width           =   4215
      End
      Begin VB.Label Label14 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   285
         Left            =   2280
         TabIndex        =   2
         Top             =   1680
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   4575
      Left            =   225
      TabIndex        =   32
      Top             =   120
      Width           =   7470
      Begin VB.TextBox TxtBuscar 
         Height          =   285
         Left            =   1320
         TabIndex        =   34
         Top             =   360
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FormValArtPed.frx":0491
         Left            =   4560
         List            =   "FormValArtPed.frx":049B
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   360
         Width           =   1695
      End
      Begin MSFlexGridLib.MSFlexGrid FG 
         Height          =   3495
         Left            =   135
         TabIndex        =   37
         Top             =   1020
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   6165
         _Version        =   393216
      End
      Begin VB.Label Label22 
         Caption         =   "Filtro"
         Height          =   255
         Left            =   360
         TabIndex        =   36
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label21 
         Caption         =   "Indice"
         Height          =   255
         Left            =   3600
         TabIndex        =   35
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Data Data3 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Inventarios.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "StkArt"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   4080
      Picture         =   "FormValArtPed.frx":04B4
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4920
      Width           =   840
   End
End
Attribute VB_Name = "FormValArtPed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''Dim db As Database
'Dim db1 As Database
Dim RSQL As String
Dim precio As Double
Dim CANTIDAD As Double
Dim tipcam As Double
Dim RS As Recordset
Dim Rs1 As Recordset
Dim mRsql As String
Dim mRsql1 As String
Dim sCodMon As String
Dim Fecha As Date   'Fecha del documento
'***********************************
'**************RMM  07/07/2001
Dim rsSTKART As New adodb.Recordset


Private Sub Combo1_Click()
     FG.Col = Combo1.ListIndex
     FG.Sort = 5
End Sub

Private Sub Combo4_Click()
  If Combo4.ListIndex = 2 Then
     Text2.SetFocus
  End If
End Sub

Private Sub Command1_Click()
  Dim CANT As String
  Dim Lote As String
  Dim Serie As String
  Dim uSql As String
  Dim RSQL As String
  Dim codmon As String * 2
  If Frame1.Visible Then
      If Not IsNumeric(Text3) Then
            MsgBox "Ingrese el Precio unitario !", vbOKOnly, "Error"
            Text3.SetFocus
            Exit Sub
     End If
     If Not IsNumeric(Text4) Then
            MsgBox "Ingrese la cantidad !", vbOKOnly, "Error"
            Text4.SetFocus
            Exit Sub
     End If
     If Combo3.ListIndex = 0 Then
            codmon = "MN"
     Else
            codmon = "ME"
     End If
     If sCodMon <> codmon Then
           If MsgBox("Desea Ud. cambiar el Tipo de moneda declarado inicialmente?", vbYesNo, "Aviso") = vbNo Then
                Exit Sub
           End If
     End If
     If Not IsNumeric(Text2) Then
           MsgBox "Ingrese el tipo de cambio !", vbOKOnly, "Error"
           Text2.SetFocus
           Exit Sub
     Else
           tipcam = Val(Text2)
     End If
     If Val(Text2) = 0 And codmon = "ME" Then
           MsgBox "Ingrese el tipo de cambio !", vbOKOnly, "Error"
           Text2.SetFocus
           Exit Sub
     End If
     If codmon = "MN" Then
        precio = Val(Text3.text) '* tipcam
     Else
        precio = Val(Text3.text) '* Val(Text2)
     End If
     CANTIDAD = Val(Text4.text)
     uSql = "Update MovAlmCab set CACODMON = '" & codmon & "', CATIPCAM = " & Val(Text2) & " where CANUMDOC='" & Trim(FG.TextMatrix(FG.Row, 3)) & "' and CAALMA = '" & VGAlma & "'    AND CATD='" & Trim(FG.TextMatrix(FG.Row, 2)) & "' "
     cConexCom.Execute uSql
     uSql = "Update MovAlmDet set DEPRECIO = " & precio & ",DETIPCAM = " & Val(Text2) & ",DECODMON = '" & codmon & "' where DENUMDOC='" & Trim(FG.TextMatrix(FG.Row, 3)) & "' and DECODIGO ='" & Trim(FG.TextMatrix(FG.Row, 0)) & "'and DEALMA = '" & VGAlma & "'  and  DETD='" & Trim(FG.TextMatrix(FG.Row, 2)) & "' "
     cConexCom.Execute uSql
     '*RMM************************************************************
     grabastk  'valoriza
     '*RMM************************************************************
     Frame1.Visible = False
     Text4 = ""
     Text5 = ""
     limpiaGrid
  Else
     Text2 = "0"
     Text3 = "0"
     If FG.Rows = 1 Then Exit Sub
     Frame1.Visible = True
     Command1.Caption = "&Aceptar"
     RSQL = "select  cacodmon,catipcam,cafecdoc from  MovAlmCab  where   CAALMA ='" & VGAlma & "'  and CATD='" & Trim(FG.TextMatrix(FG.Row, 2)) & "'  and CANUMDOC= '" & Trim(FG.TextMatrix(FG.Row, 3)) & "'" '    "'  n.DENUMDOC "
     'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
     
     Set RS = cConexCom.Execute(RSQL)
     If Not RS.EOF Then
            If RS("CACODMON") = "MN" Then
                Combo3.ListIndex = 0
                sCodMon = "MN"
            Else
                Combo3.ListIndex = 1
                sCodMon = "ME"
            End If
            If RS(1) <> 0 Then
                Text2 = RS(1)
            End If
     Fecha = RS(2)
     End If
     RS.Close
     Label10 = FG.TextMatrix(FG.Row, 3)   'numdoc
     Lbltransa = FG.TextMatrix(FG.Row, 2)
     Label16 = FG.TextMatrix(FG.Row, 0)  ' çod
     Label18 = FG.TextMatrix(FG.Row, 1)
     Label13 = FG.TextMatrix(FG.Row, 4)     ' proveedor
     Label12 = FG.TextMatrix(FG.Row, 5)     'RFTDOC
     'Label14 = FG.TextMatrix(FG.Row, 5)     'RFTDOC
     Text1 = FG.TextMatrix(FG.Row, 6)
     If Label12 <> "" Then
             Label19 = tipref(Label12)
     End If
     If Lbltransa <> "" Then Label20 = Transa(Lbltransa)
     Call cantidad_art(CANT, Serie, Lote)
     Text4.Enabled = True
     Text4 = CANT
     Text4.Enabled = False
     If Lote = "" Then
            Label14 = Serie
     Else
            Label14 = Lote
     End If
     Text3 = UltimoPrecio(Label16, sCodMon) 'precio Sugerido
     Text3.SetFocus
  End If
End Sub

Private Sub Command7_Click()
  If Frame1.Visible Then
        Frame1.Visible = False
        Command1.Caption = "&Validar"
        Text4 = ""
        Text5 = ""
  Else
'        db.Close
        Unload Me
  End If
  
End Sub

Private Sub Form_Activate()
  If Rs1.EOF Then
     MsgBox "No hay articulos por valorizar", vbExclamation, mensaje1
     Rs1.Close
     Unload Me
     Exit Sub
  End If
  'Set RS = db1.OpenRecordset(mRsql, dbOpenSnapshot)
  
  Set RS = cConexCom.Execute(mRsql)
  If RS.RecordCount = 0 Then
      MsgBox "No hay articulos por valorizar que esten Pendientes", vbExclamation, mensaje1
     RS.Close
     Form_Unload (0)
  Else
    FG.Rows = 1
    RS.MoveFirst
    FG.Visible = False
    While Not RS.EOF
          FG.AddItem (RS(0) & vbTab & RS(1) & vbTab & RS(2) & vbTab & RS(3) & vbTab & RS(4) & vbTab & RS(5) & vbTab & RS(6))
          RS.MoveNext
    Wend
    RS.Close
    FG.Visible = True
  End If
End Sub

Private Sub Form_Load()
 '****************************************************RMM 07/07/2001
  Set rsSTKART = New adodb.Recordset
  rsSTKART.Open "Select * from STKART WHERE STALMA='" & VGAlma & "'", cConexCom, adOpenDynamic, adLockOptimistic
 '******************************************************************************************
  
  Data3.DatabaseName = ""
  central FormValArtPed
  FG.FormatString = "Codigo Art.|Descripcion| TD |Num.Doc| |"
  FG.Row = 0
  Label12 = ""
  Label13 = ""
  Label14 = ""
  Label19 = ""
  Text3 = ""
  Combo1.ListIndex = 0
  Combo3.ListIndex = 0
  Combo4.ListIndex = 0
  
  FG.Cols = 7
  FG.ColWidth(0) = 1500
  FG.ColWidth(1) = 3000
  FG.ColWidth(2) = 800
  FG.ColWidth(3) = 1500
  FG.ColWidth(4) = 2
  FG.ColWidth(5) = 2
  FG.ColWidth(6) = 2
  FG.ColAlignment(0) = 1
  Frame1.Visible = False
  'mRsql = "select  n.DECODIGO, ADESCRI, m.CACODMOV ,n.DENUMDOC, m.CANOMPRO, m.CARFTDOC, m.CARFNDOC from MovAlmCab m, MovAlmDet n ,MaeArt  Where  m.CAALMA ='" & VGAlma & _
                     "' AND n.DEALMA = m.CAALMA and CATD='NI'  and  n.DEPRECIO = 0   and ACODIGO  = n.DECODIGO      And   n.DENUMDOC = m.CANUMDOC  and n.DETD= m.CATD and m.CASITGUI<>'A'  ORDER BY m.CANUMDOC"
  'RMM 02/07/2001****************
  mRsql = "select  n.DECODIGO, ADESCRI, N.DETD,n.DENUMDOC, m.CANOMPRO, m.CARFTDOC, m.CARFNDOC from MovAlmCab m, MovAlmDet n ,MaeArt  Where  m.CAALMA ='" & VGAlma & _
                     "' AND n.DEALMA = m.CAALMA and (CATD='NI' OR CATD='NC' )   and  n.DEPRECIO = 0   and ACODIGO  = n.DECODIGO      And   n.DENUMDOC = m.CANUMDOC  and n.DETD= m.CATD and m.CASITGUI<>'A'  ORDER BY m.CANUMDOC"
  '*********************************
  'Set db1 = Workspaces(0).OpenDatabase(cRuta2)
  mRsql1 = "select n.STCODIGO FROM  StkArt n where n.STALMA = '" & VGAlma & "'"
  'Set db = Workspaces(0).OpenDatabase(cRuta2)
  'Set Rs1 = db.OpenRecordset(mRsql1, dbOpenSnapshot)
  Set Rs1 = cConexCom.Execute(mRsql1)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        rsSTKART.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
     Unload Me
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
   SendKeys "{tab}"
 Else
   If ((Chr$(KeyAscii) > "9" Or Chr$(KeyAscii) < "0")) And Chr(KeyAscii) <> "." And KeyAscii <> 8 Then KeyAscii = 0
 End If
End Sub

Private Sub Text5_Change()
  If Text4 <> "" And IsNumeric(Text5) Then
             Text3 = Format(Val(Text5) / Val(Text4), "###0.0000")
  End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 And IsNumeric(Text3) Then
            Command1.SetFocus
            Exit Sub
  End If
  If ((Chr$(KeyAscii) > "9" Or Chr$(KeyAscii) < "0")) And Chr(KeyAscii) <> "." And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If IsNumeric(Text4) And KeyAscii = 13 And IsNumeric(Text3) Then
        If Not IsNumeric(Text3) Then Exit Sub
Text5 = Val(Text3) * Val(Text4)
ElseIf KeyAscii = 13 And IsNumeric(Text5) And IsNumeric(Text4) Then
        Text3 = Format(Val(Text5) / Val(Text4), "##0.0000")
Else
        If Chr$(KeyAscii) = "." Then Exit Sub
        If ((Chr$(KeyAscii) > "9" Or Chr$(KeyAscii) < "0")) And KeyAscii <> 8 Then KeyAscii = 0
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And IsNumeric(Text5) And IsNumeric(Text4) Then
      Text3 = Val(Text5) / Val(Text4)
      Text3 = Format(Text3, "##0.0000")
      Command1.SetFocus
Else
      If Chr$(KeyAscii) = "." Then Exit Sub
      If ((Chr$(KeyAscii) > "9" Or Chr$(KeyAscii) < "0")) And KeyAscii <> 8 Then KeyAscii = 0
End If
End Sub

Public Sub grabastk()
   Dim criterio As String
   Dim cadena As String
   Dim auxdisp As Double
   Dim AUXPRECIO As Double
   cadena = Label16
   '***************************RMM  07/07/2001***************************
   criterio = " STCODIGO ='" & cadena & "' and  STALMA ='" & VGAlma & "'"
   rsSTKART.Filter = criterio
        
   If Combo3.ListIndex = 0 Then
       AUXPRECIO = precio
   Else
       If Val(Text2) <> 0 Then
          AUXPRECIO = precio * Val(Text2)
       End If
   End If
   '***************************RMM  01/08/2001***************************
   
   If Not rsSTKART.EOF Then
     'Data3.*Recordset.Edit
     auxdisp = rsSTKART("STSKDIS")
     If rsSTKART("STKPREPRO") <> 0 And (CANTIDAD + auxdisp) <> 0 Then   'no se registrado algun precio
        'rsSTKART("STKPREPRO") = (Precio * cantidad + auxdisp * rsSTKART("STKPREPRO")) / (cantidad + auxdisp)
         rsSTKART("STKPREPRO") = (AUXPRECIO * CANTIDAD + auxdisp * rsSTKART("STKPREPRO")) / (CANTIDAD + auxdisp)
     Else
        'rsSTKART("STKPREPRO") = precio
        rsSTKART("STKPREPRO") = AUXPRECIO
        '***************************RMM  01/08/2001***************************
     End If
     If IsNull(rsSTKART("STKFECULT")) Or (rsSTKART("STKFECULT") <= Fecha) Then
        rsSTKART("STKFECULT") = Fecha
        rsSTKART("STKPREULT") = AUXPRECIO '*RMM***********  'Precio
     End If
   End If
     
   rsSTKART.Update
   'Data3.*Refresh
End Sub

Private Sub cantidad_art(pcantidad As String, pserie As String, plote As String)
 Dim RS As Recordset
 Dim RSQL As String
 RSQL = "select decantid,delote,deserie from MovAlmdet where DENUMDOC='" & Trim(FG.TextMatrix(FG.Row, 3)) & "' and DECODIGO ='" & Trim(FG.TextMatrix(FG.Row, 0)) & "' and DEALMA = '" & VGAlma & "' AND  DETD='" & Trim(FG.TextMatrix(FG.Row, 2)) & "'"
 'Set RS = db1.OpenRecordset(RSQL, dbOpenSnapshot)
 
 Set RS = cConexCom.Execute(RSQL)
  If RS.EOF Then
            pcantidad = "0"
            pserie = ""
            plote = ""
  Else
            pcantidad = Str(RS(0))
            pserie = IIf(Not IsNull(RS(1)), RS(1), "")
            pserie = IIf(Not IsNull(RS(2)), RS(2), "")
  End If
End Sub

Function tipref(text As Label) As String
 
 Dim RS As Recordset
 Dim RSQL As String
 RSQL = "select  TDO_DESCRI FROM TIPO_DOCU  where TDO_TIPDOC= '" & text & "'" '
 'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
 Set RS = cConexCom.Execute(RSQL)
 tipref = IIf(Not RS.EOF, RS(0), "")
 RS.Close
End Function

Function Transa(text As Label) As String
 Dim RS As Recordset
 Dim RSQL As String
 Dim dato As String
  dato = "I"
  RSQL = "select  TT_DESCRI FROM TabTransa where TT_CODMOV= '" & text & "' and TT_TIPMOV ='" & dato & "'" '
  'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
  Set RS = cConexCom.Execute(RSQL)
  Transa = IIf(Not RS.EOF, RS(0), "")
  RS.Close
End Function

Private Sub limpiaGrid()
Dim i As Integer
 If FG.Rows = 1 Then Exit Sub
 i = FG.RowSel
 If FG.Rows > 2 Then
        FG.RemoveItem i
 Else
        FG.Clear
        FG.Rows = 1
        FG.FormatString = "Cod. Articulo.|Descripcion| Tr| Num.Doc."
        FG.Row = 0
        FG.ColWidth(0) = 1200
        FG.ColWidth(1) = 3000
        FG.ColWidth(2) = 800
        FG.ColWidth(3) = 1500
        FG.ColWidth(4) = 2
        FG.ColWidth(5) = 2
  End If
End Sub

Private Sub Txtbuscar_Change()
Dim i As Integer
Dim n As Integer
n = Combo1.ListIndex
If TxtBuscar <> "" Then
      For i = 1 To FG.Rows - 1
          If UCase(Left(FG.TextMatrix(i, n), Len(TxtBuscar))) = UCase(Trim(TxtBuscar)) Then
             Exit For
          End If
      Next i
      If i >= FG.Rows Then
            FG.HighLight = flexHighlightNever
      Else
            FG.HighLight = flexHighlightAlways
            FG.TopRow = i
            FG.Row = i
            FG.Col = 0
            FG.ColSel = FG.Cols - 1
      End If
End If
End Sub
