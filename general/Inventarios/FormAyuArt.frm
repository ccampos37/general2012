VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FormAyuArt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Articulos"
   ClientHeight    =   5190
   ClientLeft      =   435
   ClientTop       =   1935
   ClientWidth     =   9765
   ControlBox      =   0   'False
   LinkTopic       =   "FormAyuArt"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   9765
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
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
      Height          =   4095
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   9525
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FormAyuArt.frx":0000
         Left            =   7680
         List            =   "FormAyuArt.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   390
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2160
         MaxLength       =   20
         TabIndex        =   0
         Top             =   360
         Width           =   2925
      End
      Begin MSFlexGridLib.MSFlexGrid FG 
         Height          =   3015
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   9300
         _ExtentX        =   16404
         _ExtentY        =   5318
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   360
         Picture         =   "FormAyuArt.frx":0035
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "Indice"
         Height          =   255
         Left            =   6840
         TabIndex        =   7
         Top             =   375
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Filtro"
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   735
      Left            =   3000
      Picture         =   "FormAyuArt.frx":08FF
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4320
      Width           =   835
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   5235
      Picture         =   "FormAyuArt.frx":0D41
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4320
      Width           =   835
   End
End
Attribute VB_Name = "FormAyuArt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
'*************************************************************
'******Modificado RMM 07/07/2001
'*************************************************************
Dim rs As New ADODB.Recordset
Dim varform As Form


Private Sub FG_Click()
  If FG.Row = 0 Then Exit Sub
  If FG.TextMatrix(FG.Row, 0) = ">>" Then
     FG.TextMatrix(FG.Row, 0) = " "
  Else
     FG.TextMatrix(FG.Row, 0) = ">>"
  End If
End Sub


Private Sub Combo1_Click()
  If Combo1.text = "Codigo" Or Combo1.text = "CodFabricante" Then
     FG.Col = Combo1.ListIndex + 1
     FG.Sort = 5
  Else
     FG.Col = Combo1.ListIndex + 1
     FG.Sort = 5
     If FG.Col = 3 Then
       FG.Col = 8:  FG.Sort = 5
     End If
  End If
  Label1.Caption = Combo1.text
  End Sub

Private Sub Command1_Click()
   Dim I As Integer
   Dim J As Integer
   Dim cad As String
   Dim varform As Form
   J = 0
  
   Select Case VGForm1
     Case 1
       Set varform = FormCreacion
        varform.Salida.Rows = 1
     Case 2
       Set varform = FrmCreacionSin
        varform.Salida.Rows = 1
     Case 3
      Set varform = FrmCreacionSal
       varform.Salida.Rows = 1
    Case 4
      Set varform = FrmRegPlantilladeKits
      FrmRegPlantilladeKits.Txtarticulo = FG.TextMatrix(I, 1)
      
    End Select
   
     For I = 0 To FG.Rows - 1

      If FG.TextMatrix(I, 0) = ">>" Then
        If VGForm1 = 4 Then
        Else
            varform.Salida.AddItem (FG.TextMatrix(I, 1) & vbTab & FG.TextMatrix(I, 2) & vbTab & FG.TextMatrix(I, 4) & vbTab & FG.TextMatrix(I, 6) & vbTab & FG.TextMatrix(I, 7))
        End If
        J = J + 1
        'FormCreacion.Salida.AddItem (FG.TextMatrix(i, 1) & vbTab & FG.TextMatrix(i, 2) & vbTab & FG.TextMatrix(i, 3))
      End If
     Next I

'     If rS.RecordCount > 0 Then
'
'        If VGForm1 = 4 Then
'            varform.Salida.AddItem (" " & vbTab & FG.TextMatrix(i, 1) & vbTab & FG.TextMatrix(i, 2) & vbTab & "1")
'            varform.Salida.AddItem (" " & vbTab & DBGrid1.Columns(1).text & vbTab & DBGrid1.Columns(2).text & vbTab & "1")
'
'        Else
'            varform.Salida.AddItem (DBGrid1.Columns(1).text & vbTab & DBGrid1.Columns(2).text & vbTab & DBGrid1.Columns(3).text & vbTab & DBGrid1.Columns(5).text & vbTab & DBGrid1.Columns(6).text)
'        End If
'     Next i


  Unload Me
End Sub

Private Sub Command8_Click()
  Unload Me
End Sub

Private Sub FG_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then
        FG_Click
    ElseIf KeyAscii = 13 Then
        FG_Click
        Command1_Click
   End If
End Sub

Public Sub Activa_ayuda()
Dim cCod As String
Dim nStock As Double
Dim rsa As New ADODB.Recordset

If rs.RecordCount = 0 Then
'    MsgBox "No hay articulos disponibles en el almacen", vbInformation, "Aviso"
    'Form_Unload (0)
    Exit Sub
  End If
  rs.MoveFirst
  FG.Visible = False
  'Recorre la tabla y recoge el stock del almacen que pertenece
  cCod = rs("acodigo"): nStock = 0
  

  Do While Not rs.EOF

     If cCod <> rs("acodigo") Then
            rs.MovePrevious
            FG.AddItem (" " & vbTab & rs(0) & vbTab & rs(1) & vbTab & rs("ACODIGO2") & vbTab & rs(3) & vbTab & Format(0, "##0.000") & vbTab & rs(5) & vbTab & rs(6) & vbTab & rs(7))
            rs.MoveNext
            nStock = 0
            cCod = rs("acodigo")
     ElseIf rs("alma1") = VGAlma And cCod = rs("acodigo") Then

            nStock = IIf(IsNull(rs(4)), 0, rs(4))
            FG.AddItem (" " & vbTab & rs(0) & vbTab & rs(1) & vbTab & rs("ACODIGO2") & vbTab & rs(3) & vbTab & IIf(IsNull(rs(4)), 0, Format(rs(4), "##0.000")) & vbTab & rs(5) & vbTab & rs(6) & vbTab & rs(7))
           Do While rs("acodigo") = cCod
                           rs.MoveNext
                           If rs.EOF Then Exit Do
           Loop
            If rs.EOF Then Exit Do
           cCod = rs("acodigo")
           nStock = 0
     Else
          rs.MoveNext
     End If
     If nStock = 0 And rs.EOF Then
            rs.MovePrevious
            FG.AddItem (" " & vbTab & rs(0) & vbTab & rs(1) & vbTab & rs("ACODIGO2") & vbTab & rs(3) & vbTab & Format("0", "##0.000") & vbTab & rs(5) & vbTab & rs(6) & vbTab & rs(7))
            Exit Do
     End If
     If rs.EOF Then Exit Do

 Loop
  FG.Visible = True

End Sub

Private Sub Form_Load()
  Dim real As Double
  Dim Cod As String
  Dim RSQL As String
Combo1.ListIndex = 0
   Cod = ""
   Select Case VGForm1
     Case 1
          Set varform = FormCreacion
     Case 2
          Set varform = FrmCreacionSin
     Case 3
          Set varform = FrmCreacionSal
     Case 4
          Set varform = FrmRegPlantilladeKits
      
  End Select
  'RMM***********************************************************
  If VGForm1 <> 4 Then
    varform.Salida.Clear
  End If
  'RMM***********************************************************
  
  If VGRegEnt = 1 Then
     If Val(varform.Txtarticulo) = 1 Then
        RSQL = "select  p.ACODIGO, p.ADESCRI,p.ACODIGO2 ,p.AUNIDAD,n.STSKDIS,p.AFSERIE, p.AFLOTE,p.AFAMILIA ,isnull(n.stalma,'xx') as Alma1 " & _
               " from MaeArt p LEFT JOIN StkArt n  ON  p.ACODIGO = n.STCODIGO WHERE p.ADESCRI LIKE '" & varform.Txtarticulo & "%'  order by p.ADESCRI "
     ElseIf Val(varform.Txtarticulo) = 2 Then
        RSQL = "select  p.ACODIGO, p.ADESCRI,p.ACODIGO2 ,p.AUNIDAD,n.STSKDIS,p.AFSERIE, p.AFLOTE,p.AFAMILIA ,isnull(n.stalma,'xx') as Alma1 " & _
               " from MaeArt p LEFT JOIN StkArt n  ON  p.ACODIGO = n.STCODIGO WHERE p.ADESCRI2 LIKE '" & varform.Txtarticulo & "%'  order by p.ADESCRI2 "
        Else
           RSQL = "select  p.ACODIGO,p.ADESCRI,p.ACODIGO2 ,p.AUNIDAD,n.STSKDIS,p.AFSERIE, p.AFLOTE,p.AFAMILIA ,isnull(n.stalma,'xx') as Alma1" & _
               " from MaeArt p LEFT JOIN StkArt n  ON  p.ACODIGO = n.STCODIGO   WHERE p.ACODIGO LIKE '" & varform.Txtarticulo & "%'  order by p.acodigo "
               
     End If
  Else   'If VGForm1 = 3 Then
         If Val(varform.Txtarticulo) = 0 Then
             RSQL = "select  p.ACODIGO, p.ADESCRI,ACODIGO2,p.AUNIDAD, n.STSKDIS, p.AFSERIE, p.AFLOTE  ,p.AFAMILIA,n.STALMA as Alma1  from MaeArt p, StkArt n where p.ACODIGO =  n.STCODIGO and n.STALMA = '" & VGAlma & "'  and n.STSKDIS >0 and p.ADESCRI LIKE '" & varform.Txtarticulo & "%' ORDER BY p.ADESCRI "
         Else
             RSQL = "select  p.ACODIGO, p.ADESCRI ,ACODIGO2 ,p.AUNIDAD, n.STSKDIS, p.AFSERIE, p.AFLOTE  ,p.AFAMILIA,n.STALMA as Alma1 from MaeArt p, StkArt n where p.ACODIGO =  n.STCODIGO and n.STALMA = '" & VGAlma & "'   and n.STSKDIS >0 and p.ACODIGO LIKE '" & varform.Txtarticulo & "%' ORDER BY ACODIGO "
         End If
  End If

  Set rs = New ADODB.Recordset
  
  Set rs = VGCNx.Execute(RSQL)
  '********************************************************
  FG.Clear
  FG.FormatString = "^Seleccion|  Codigo|   Descripcion | Cod. Fabr.|  Unidad | Stock   |Se | Lt | Familia "
  FG.Row = 0
  FG.ColWidth(0) = 800
  FG.ColWidth(1) = 1500
  FG.ColWidth(2) = 3250
  FG.ColWidth(3) = 1200
  FG.ColWidth(4) = 800
  FG.ColWidth(5) = 1000
  FG.ColWidth(6) = 300
  FG.ColWidth(7) = 300
  FG.ColWidth(8) = 800


  FG.ColAlignment(1) = 1
  If VGForm1 <> 4 Then
    varform.Salida.FormatString = "Codigo|Descripcion|Unidad|sr|lt"
    varform.Salida.Row = 0
    varform.Salida.ColWidth(0) = 1500
    varform.Salida.ColWidth(1) = 1500
    varform.Salida.ColWidth(2) = 500
    varform.Salida.ColWidth(3) = 200
    varform.Salida.ColWidth(4) = 200
  End If
  FG.Rows = 1
'  If Val(varform.TxtArticulo) = 0 Then
'      Combo1.ListIndex = 1
'      Text1 = IIf(IsNull(varform.TxtArticulo), "", varform.TxtArticulo)
'      Text1.SelStart = Len(Text1)
'  Else
     Combo1.ListIndex = 1
     text1 = IIf(IsNull(varform.Txtarticulo), "", varform.Txtarticulo)
     text1.SelStart = Len(text1)
'  End If
  
 
  Label1.Caption = Combo1.text
  AlinearAyuda Me
  DoEvents
  
  Call Activa_ayuda
  
End Sub


Sub LISTAR()
   Dim RSQL As String
   Select Case VGForm1
     Case 1
          Set varform = FormCreacion
     Case 2
          Set varform = FrmCreacionSin
     Case 3
          Set varform = FrmCreacionSal
     Case 4
          Set varform = FrmRegKit
  End Select
  'RMM***********************************************************
  If VGForm1 <> 4 Then
    varform.Salida.Clear
  End If
  'RMM***********************************************************
  
  If VGRegEnt = 1 Then
     'If Val(CStr(StrReverse(Text1))) = 0 Then
     If Combo1.ListIndex = 1 Then
        RSQL = "select  p.ACODIGO, p.ADESCRI,ACODIGO2,p.AUNIDAD,n.STSKDIS,p.AFSERIE, p.AFLOTE,p.AFAMILIA ,isnull(n.stalma,'xx') as Alma1  " & _
               "from MaeArt p LEFT JOIN StkArt n  ON  p.ACODIGO = n.STCODIGO WHERE p.ADESCRI LIKE '%" & text1 & "%'  order by p.ADESCRI "
     ElseIf Combo1.ListIndex = 2 Then
            RSQL = "select  p.ACODIGO, p.ADESCRI,ACODIGO2,p.AUNIDAD,n.STSKDIS,p.AFSERIE, p.AFLOTE,p.AFAMILIA ,isnull(n.stalma,'xx') as Alma1 " & _
               "from MaeArt p LEFT JOIN StkArt n  ON  p.ACODIGO = n.STCODIGO WHERE p.acodigo2 LIKE '" & text1 & "%'  order by p.acodigo2 "
        Else
            RSQL = "select  p.ACODIGO, p.ADESCRI,ACODIGO2,p.AUNIDAD,n.STSKDIS,p.AFSERIE, p.AFLOTE,p.AFAMILIA ,isnull(n.stalma,'xx') as Alma1  " & _
               "from MaeArt p LEFT JOIN StkArt n  ON  p.ACODIGO = n.STCODIGO   WHERE p.ACODIGO LIKE '" & text1 & "%'  order by p.acodigo " 'group by p.ACODIGO ,p.ADESCRI,p.AUNIDAD, p.AFSERIE, p.AFLOTE ,n.stskdis,p.AFAMILIA  " '  " ', and n.STALMA = '" & VGAlma & "'ORDER BY ACODIGO   ' WHERE  n.STALMA = '" & VGAlma & "'
     End If
  Else   'If VGForm1 = 3 Then
         'If Val(CStr(StrReverse(Text1))) = 0 Then
         If Combo1.ListIndex = 1 Then
             RSQL = "select  p.ACODIGO, p.ADESCRI,p.acodigo2,p.AUNIDAD, n.STSKDIS, p.AFSERIE, p.AFLOTE  ,p.AFAMILIA,n.STALMA as Alma1 ,ACODIGO2 from MaeArt p, StkArt n where p.ACODIGO =  n.STCODIGO and n.STALMA = '" & VGAlma & "'  and n.STSKDIS >0 and p.ADESCRI LIKE '%" & text1 & "%' ORDER BY p.ADESCRI "
         ElseIf Combo1.ListIndex = 2 Then
                  RSQL = "select  p.ACODIGO, p.ADESCRI,p.acodigo2,p.AUNIDAD, n.STSKDIS, p.AFSERIE, p.AFLOTE  ,p.AFAMILIA,n.STALMA as Alma1 ,ACODIGO2 from MaeArt p, StkArt n where p.ACODIGO =  n.STCODIGO and n.STALMA = '" & VGAlma & "'   and n.STSKDIS >0 and p.ACODIGO2 LIKE '%" & text1 & "%' ORDER BY ACODIGO2 "
               Else
                 RSQL = "select  p.ACODIGO, p.ADESCRI,p.acodigo2,p.AUNIDAD, n.STSKDIS, p.AFSERIE, p.AFLOTE  ,p.AFAMILIA,n.STALMA as Alma1 ,ACODIGO2 from MaeArt p, StkArt n where p.ACODIGO =  n.STCODIGO and n.STALMA = '" & VGAlma & "'   and n.STSKDIS >0 and p.ACODIGO LIKE '%" & text1 & "%' ORDER BY ACODIGO "
         End If
  End If

  FG.Clear
  DoEvents
  Set rs = New ADODB.Recordset
  Set rs = VGCNx.Execute(RSQL)
  '********************************************************
  FG.FormatString = "^Seleccion|  Codigo|   Descripcion| Cod. Fabr. |  Unidad | Stock   |Se | Lt | Familia "
  FG.Row = 0
  FG.ColWidth(0) = 800
  FG.ColWidth(1) = 1500
  FG.ColWidth(2) = 3500
  FG.ColWidth(3) = 1500
  FG.ColWidth(4) = 800
  FG.ColWidth(5) = 1000
  FG.ColWidth(6) = 300
  FG.ColWidth(7) = 300
  FG.ColWidth(8) = 800
  

  FG.ColAlignment(1) = 1
  If VGForm1 <> 4 Then
    varform.Salida.FormatString = "Codigo|Descripcion|Unidad|sr|lt"
    varform.Salida.Row = 0
    varform.Salida.ColWidth(0) = 1500
    varform.Salida.ColWidth(1) = 3500
    varform.Salida.ColWidth(2) = 400
    varform.Salida.ColWidth(3) = 200
    varform.Salida.ColWidth(4) = 200
  End If
  FG.Rows = 1
  Call Activa_ayuda
  DoEvents
End Sub



Private Sub Form_Unload(Cancel As Integer)
  Unload Me
End Sub

Private Sub Text1_Change()
  Call LISTAR
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  
   Dim I As Integer
   Dim n As Integer
   n = Combo1.ListIndex + 1
   If n = 3 Then
      n = 8
   End If
   
   If text1 <> "" And KeyAscii = 13 Then

      For I = 1 To FG.Rows - 1
          If UCase(Left(FG.TextMatrix(I, n), Len(text1))) = UCase(Trim(text1)) Then
             Exit For
          End If
      Next I
      If I >= FG.Rows Then
            FG.HighLight = flexHighlightNever
      Else
            FG.HighLight = flexHighlightAlways
            FG.TopRow = I
            FG.Row = I
            FG.Col = 0
            FG.ColSel = FG.Cols - 1
      End If
      FG.SetFocus
   End If
End Sub

