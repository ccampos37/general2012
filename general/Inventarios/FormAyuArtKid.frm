VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FormAyuArtKid 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Articulos"
   ClientHeight    =   5250
   ClientLeft      =   240
   ClientTop       =   990
   ClientWidth     =   8355
   ControlBox      =   0   'False
   LinkTopic       =   "FormAyuArt"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   8355
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
      Width           =   8055
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FormAyuArtKid.frx":0000
         Left            =   4920
         List            =   "FormAyuArtKid.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2160
         MaxLength       =   20
         TabIndex        =   0
         Top             =   360
         Width           =   1695
      End
      Begin MSFlexGridLib.MSFlexGrid FG 
         Height          =   3015
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   5318
         _Version        =   393216
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   360
         Picture         =   "FormAyuArtKid.frx":0023
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "Indice"
         Height          =   255
         Left            =   4080
         TabIndex        =   7
         Top             =   360
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
      Picture         =   "FormAyuArtKid.frx":08ED
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4320
      Width           =   835
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   4680
      Picture         =   "FormAyuArtKid.frx":0D2F
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4320
      Width           =   835
   End
End
Attribute VB_Name = "FormAyuArtKid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As Recordset

Private Sub FG_Click()
  If FG.Row = 0 Then Exit Sub
  If FG.TextMatrix(FG.Row, 0) = ">>" Then
     FG.TextMatrix(FG.Row, 0) = " "
  Else
     FG.TextMatrix(FG.Row, 0) = ">>"
  End If
End Sub

Private Sub Combo1_Click()
  If Combo1.text = "Codigo" Then
     FG.Col = Combo1.ListIndex + 1
     FG.Sort = 5
  Else
     FG.Col = Combo1.ListIndex + 1
     FG.Sort = 5
  End If
  Label1.Caption = Combo1.text
  End Sub

Private Sub Command1_Click()
   Dim I As Integer
   Dim varform As Form
   Select Case VGForm1
   Case 4
      Set varform = FrmArmadoKits
   Case 5
      Set varform = FrmDesKits
   End Select
   If VGForm1 <> 4 And VGForm1 <> 5 Then varform.MSFlexGrid1.Rows = 1
   For I = 0 To FG.Rows - 1
      If FG.TextMatrix(I, 0) = ">>" Then
            Dim rs As New ADODB.Recordset
            Dim SQL As String
            'SQL = "SELECT KITS.CODART, MAEART.ADESCRI, KITS.CANART, 0 AS Expr1, STKART.STSKDIS FROM (KITS INNER JOIN STKART ON KITS.CODART = STKART.STCODIGO) INNER JOIN MAEART ON STKART.STCODIGO = MAEART.ACODIGO where STALMA='" & VGAlma & "' AND  KITS.CODkit='" & FG.TextMatrix(i, 1) & "'"
            Call ClsTock.VerificaKIT(VGAlma, FG.TextMatrix(I, 1), VGCNx)
            SQL = "SELEct KITS.CODART, MAEART.ADESCRI, KITS.CANART, 0 AS Expr1, STKART.STSKDIS FROM (KITS INNER JOIN STKART ON KITS.CODKIT = STKART.STCODIGO) LEFT JOIN MAEART ON KITS.CODART =MAEART.ACODIGO where STALMA='" & VGAlma & "' AND  KITS.CODkit='" & FG.TextMatrix(I, 1) & "'"
            rs.Open SQL, VGCNx, adOpenStatic, adLockReadOnly
            If Not rs.EOF Then
               varform.TxCodKid = FG.TextMatrix(I, 1)
               varform.lblnomkits = FG.TextMatrix(I, 2)
               'Set varform.MSFlexGrid1.DataSource = rS
               varform.MSFlexGrid1.Clear
               varform.MSFlexGrid1.Rows = 2
               'varform.MSFlexGrid1.Row = 1
               varform.MSFlexGrid1.AddItem rs!codart & Chr(9) & rs!ADESCRI & Chr(9) & rs!CANART & Chr(9) & 0 & Chr(9) & ClsTock.SaldoArti(VGAlma, rs!codart, VGCNx), 1
               rs.MoveNext
               Do While Not rs.EOF
                  varform.MSFlexGrid1.AddItem rs!codart & Chr(9) & rs!ADESCRI & Chr(9) & rs!CANART & Chr(9) & 0 & Chr(9) & ClsTock.SaldoArti(VGAlma, rs!codart, VGCNx), 1
                  rs.MoveNext
               Loop
               varform.MSFlexGrid1.Rows = varform.MSFlexGrid1.Rows - 1
            
             If VGForm1 = 4 Then
                varform.MSFlexGrid1.FormatString = "^Codigo             |<Descripción                                                                                    |>Cant.Reg.    |>Cant.Desarm.  |>Cant.Dispon."
             Else
                varform.MSFlexGrid1.FormatString = "^Codigo             |<Descripción                                                                                    |>Cant.Reg.    |>Cant.Armada   |>Cant.Dispon."
             End If
               
            End If
            rs.Close
            'Exit Sub
            'varform.MSFlexGrid1.AddItem (FG.TextMatrix(I, 1) & vbTab & FG.TextMatrix(I, 2) & vbTab & "1" & vbTab & "0")
      End If
     Next I
  Unload Me
End Sub

Private Sub Command8_Click()
  Unload Me
End Sub

Private Sub FG_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then
        FG_Click
   End If
End Sub

Private Sub Form_Activate()
Dim cCod As String
Dim nStock As Double
If rs.RecordCount = 0 Then
    MsgBox "No hay KITS  disponibles ", vbInformation, "Aviso"
    Form_Unload (0)
    Exit Sub
  End If
  rs.MoveFirst
  FG.Visible = False
  'Recorre la tabla y recoge el stock del almacen que pertenece
  cCod = rs("acodigo"): nStock = 0
  Do While Not rs.EOF
     If cCod <> rs("acodigo") Then
            rs.MovePrevious
            FG.AddItem (" " & vbTab & rs(0) & vbTab & rs(1) & vbTab & rs(2) & vbTab & Format(0, "##0.000") & vbTab & rs(4) & vbTab & rs(5) & vbTab & rs(6))
            rs.MoveNext
            nStock = 0
            cCod = rs("acodigo")
     ElseIf rs("alma1") = VGAlma And cCod = rs("acodigo") Then
     
            nStock = IIf(IsNull(rs(3)), 0, rs(3))
            FG.AddItem (" " & vbTab & rs(0) & vbTab & rs(1) & vbTab & rs(2) & vbTab & IIf(IsNull(rs(3)), 0, Format(rs(3), "##0.000")) & vbTab & rs(4) & vbTab & rs(5) & vbTab & rs(6))
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
            FG.AddItem (" " & vbTab & rs(0) & vbTab & rs(1) & vbTab & rs(2) & vbTab & Format("0", "##0.000") & vbTab & rs(4) & vbTab & rs(5) & vbTab & rs(6))
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
  Dim varform As Form
  Cod = ""
  Select Case VGForm1
     Case 4
       'Set varform = FrmArmadoKits
     Case 5
       'Set varform = FrmDesKits
  End Select
  If VGForm1 <> 4 And VGForm1 <> 5 Then varform.MSFlexGrid1.Rows = 1
  'AND STALMA = '" & VGAlma & "'
  If VGRegEnt = 1 Then
        RSQL = "SELECT distinct(ACODIGO),ADESCRI,AUNIDAD,AFSERIE,AFLOTE,AFAMILIA,0 as Alma1 FROM KITS,MaeArt WHERE  ACODIGO = CODKIT ORDER BY ACODIGO"
  Else
        RSQL = "SELECT distinct(ACODIGO),ADESCRI,AUNIDAD,STSKDIS,AFSERIE,AFLOTE,AFAMILIA,STALMA as Alma1 FROM KITS,stkart,MaeArt WHERE STCODIGO = CODKIT  AND ACODIGO = CODKIT  AND STALMA = '" & VGAlma & "'    ORDER BY ACODIGO"
  End If
  'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
  Set rs = VGCNx.Execute(RSQL)
  If VGRegEnt = 1 Then
    FG.FormatString = "^Seleccion|  Codigo|   Descripcion|  Unidad | Se | Lt | Familia "
    FG.Row = 0
    FG.ColWidth(0) = 910
    FG.ColWidth(1) = 1200
    FG.ColWidth(2) = 3250
    FG.ColWidth(3) = 800
    FG.ColWidth(4) = 2
    FG.ColWidth(5) = 2
    FG.ColWidth(6) = 1000
  Else
    FG.FormatString = "^Seleccion|  Codigo|   Descripcion|  Unidad | Stock   |Se | Lt | Familia "
    FG.Row = 0
    FG.ColWidth(0) = 910
    FG.ColWidth(1) = 1200
    FG.ColWidth(2) = 3250
    FG.ColWidth(3) = 800
    FG.ColWidth(4) = 1200
    FG.ColWidth(5) = 2
    FG.ColWidth(6) = 2
    FG.ColWidth(7) = 1000
  End If
  FG.ColAlignment(1) = 1
  If VGForm1 <> 4 And VGForm1 <> 5 Then
    varform.MSFlexGrid1.FormatString = "Codigo|Descripcion|Unidad|sr|lt"
    varform.MSFlexGrid1.Row = 0
    varform.MSFlexGrid1.ColWidth(0) = 1500
    varform.MSFlexGrid1.ColWidth(1) = 1500
    varform.MSFlexGrid1.ColWidth(2) = 500
    varform.MSFlexGrid1.ColWidth(3) = 200
    varform.MSFlexGrid1.ColWidth(4) = 200
  End If
  FG.Rows = 1
  Combo1.ListIndex = 0
  Label1.Caption = Combo1.text
  AlinearAyuda Me
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Unload Me
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  
   Dim I As Integer
   Dim n As Integer
   n = Combo1.ListIndex + 1
   
   
   If Text1 <> "" Then
      For I = 1 To FG.Rows - 1
          If UCase(Left(FG.TextMatrix(I, n), Len(Text1))) = UCase(Trim(Text1)) Then
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
   
   End If

End Sub

