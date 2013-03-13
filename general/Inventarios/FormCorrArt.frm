VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FormCorrArt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Correción de Artículos"
   ClientHeight    =   6570
   ClientLeft      =   915
   ClientTop       =   1710
   ClientWidth     =   8025
   Icon            =   "FormCorrArt.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   8025
   Begin VB.Frame Frame1 
      Caption         =   "Correción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   4785
      Left            =   270
      TabIndex        =   11
      Top             =   900
      Visible         =   0   'False
      Width           =   7530
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1920
         TabIndex        =   18
         Top             =   4200
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   6165
         TabIndex        =   17
         Top             =   1755
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1920
         TabIndex        =   16
         Top             =   3720
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   5280
         TabIndex        =   15
         Top             =   3240
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1920
         TabIndex        =   14
         Top             =   2880
         Width           =   1455
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "FormCorrArt.frx":08CA
         Left            =   5280
         List            =   "FormCorrArt.frx":08D4
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2880
         Width           =   1575
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "FormCorrArt.frx":08E8
         Left            =   1920
         List            =   "FormCorrArt.frx":08F5
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label Label19 
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   41
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label18 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1920
         TabIndex        =   40
         Top             =   2520
         Width           =   4935
      End
      Begin VB.Label Label17 
         Caption         =   "Descripción"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   480
         TabIndex        =   39
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label16 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1920
         TabIndex        =   38
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "Código Artículo"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   495
         TabIndex        =   37
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Total"
         ForeColor       =   &H80000007&
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   36
         Top             =   4200
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Cantidad"
         ForeColor       =   &H80000007&
         Height          =   255
         Index           =   2
         Left            =   3960
         TabIndex        =   35
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Costo Unitario"
         ForeColor       =   &H80000007&
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   34
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Documento"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   480
         TabIndex        =   33
         Top             =   375
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Transacción"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   480
         TabIndex        =   32
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Proveedor"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   480
         TabIndex        =   31
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Doc Referencial"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   480
         TabIndex        =   30
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Serie"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   480
         TabIndex        =   29
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Moneda"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   3960
         TabIndex        =   28
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Factura"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   480
         TabIndex        =   27
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Conversion"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   480
         TabIndex        =   26
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Tipo de Cambio"
         ForeColor       =   &H80000007&
         Height          =   255
         Index           =   0
         Left            =   3960
         TabIndex        =   25
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1920
         TabIndex        =   24
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1920
         TabIndex        =   23
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label12 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1920
         TabIndex        =   22
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1920
         TabIndex        =   21
         Top             =   1080
         Width           =   4695
      End
      Begin VB.Label Label14 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1920
         TabIndex        =   20
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label20 
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   19
         Top             =   720
         Width           =   3975
      End
   End
   Begin VB.Frame Frame4 
      Height          =   825
      Left            =   240
      TabIndex        =   8
      Top             =   0
      Width           =   7485
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuAlmacen 
         Height          =   375
         Left            =   1320
         TabIndex        =   10
         Top             =   240
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   661
         xclave          =   "taalma"
         xnombre         =   "tadescri"
         XcodMaxLongitud =   0
         xcodwith        =   200
         NomTabla        =   "tabalm"
         ListaCampos     =   "taalma(1),tadescri(1)"
         ListaCamposDescrip=   "Almacen, descripcion"
         ListaCamposText =   "taalma,tadescri"
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Almacen"
         Height          =   195
         Left            =   405
         TabIndex        =   9
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Data Data3 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5865
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Corregir"
      Height          =   735
      Left            =   2400
      Picture         =   "FormCorrArt.frx":0920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5715
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   735
      Left            =   4440
      Picture         =   "FormCorrArt.frx":0D62
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5745
      Width           =   855
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
      Height          =   4740
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   7500
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FormCorrArt.frx":11A4
         Left            =   4440
         List            =   "FormCorrArt.frx":11B1
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox TxtBuscar 
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
      Begin MSFlexGridLib.MSFlexGrid FG 
         Height          =   3615
         Left            =   75
         TabIndex        =   7
         Top             =   990
         Width           =   7350
         _ExtentX        =   12965
         _ExtentY        =   6376
         _Version        =   393216
      End
      Begin VB.Label Label22 
         Caption         =   "Indice"
         Height          =   255
         Left            =   3480
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label21 
         Caption         =   "Filtro"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "FormCorrArt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''Dim db As Database
Dim rs As Recordset
Dim RSQL As String
Dim precioant As Double  'costo anterior
Dim sCodMon As String 'El codigo de moneda
Dim Fecha As Date     'Graba la Fecha del Documento
'***********************************
'**************RMM  07/07/2001
Dim rsSTKART As New ADODB.Recordset


Private Sub Combo1_Click()
   If Combo1.text = "Codigo" Then
        FG.Col = Combo1.ListIndex + 1
        FG.Sort = 5
  Else
        FG.Col = Combo1.ListIndex + 1
        FG.Sort = 5
  End If
  'Label1.Caption = Combo1.text
End Sub

Private Sub Command1_Click()
  Dim precio As Double  ' corregir la
  Dim CANTIDAD As Double
  Dim uSql As String
  Dim RSQL As String
  Dim cant As String
  Dim Serie As String
  Dim Lote As String
  Dim codmon As String
  
  If Frame1.Visible Then
        'Text3.Text = "0"
        If Not IsNumeric(Text3) Then
                MsgBox "Ingrese el Precio unitario !", vbOKOnly + vbExclamation, "Error"
                Text3.SetFocus
                Exit Sub
        End If
        If Not IsNumeric(Text4) Then
                MsgBox "Ingrese la cantidad !", vbOKOnly + vbExclamation, "Error"
                Text4.SetFocus
                Exit Sub
        End If
        If Combo3.ListIndex <> 0 Then
                If Val(Text2) = 0 Then
                    MsgBox "Ingrese el tipo de cambio !", vbOKOnly + vbExclamation, "Error"
                    Text2.SetFocus
                    Exit Sub
                End If
        End If
        If Combo3.ListIndex = 0 Then
            codmon = "01"
        Else
            codmon = "02"
        End If
        If sCodMon <> codmon Then
            If MsgBox("Desea Ud. cambiar el Tipo de moneda declarado inicialmente?", vbYesNo, "Aviso") = vbNo Then
                Exit Sub
            End If
        End If
        If codmon = "01" Then
            precio = Val(Text3.text)
        Else
            precio = Val(Text3.text) '* Val(Text2)
        End If
        CANTIDAD = Val(Text4.text)
        uSql = "Update MovAlmCab set CACODMON = '" & codmon & "', CATIPCAM = " & Val(Text2) & " where CANUMDOC='" & Trim(FG.TextMatrix(FG.Row, 3)) & "' and CAALMA = '" & VGAlma & "'    AND CATD='" & Label11.Caption & "' "
        VGCNx.Execute uSql
        uSql = "Update MovAlmDet set DEPRECIO = " & precio & ",DETIPCAM = " & Val(Text2) & " ,DECODMON = '" & codmon & "' where  DEALMA ='" & VGAlma & "'  and DENUMDOC='" & Trim(FG.TextMatrix(FG.Row, 3)) & "' and DECODIGO ='" & Trim(FG.TextMatrix(FG.Row, 0)) & "' and DeTD='" & Label11.Caption & "'"
        VGCNx.Execute uSql
        If Text3 <> precioant Then grabastk
        Frame1.Visible = False
        Text4 = ""
        Text5 = ""
   Else
        Text2 = "0"
        Text3 = "0"
        If FG.Rows = 1 Then
                Command7_Click
                Exit Sub
        End If
        Frame1.Visible = True
        Command1.Caption = "&Aceptar"
        RSQL = "select  cacodmon,cafecdoc from  MovAlmCab  where   CAALMA ='" & VGAlma & "'  and CATD='" & Trim(FG.TextMatrix(FG.Row, 2)) & "' AND CANUMDOC= '" & Trim(FG.TextMatrix(FG.Row, 3)) & "'"
        
        Set rs = VGCNx.Execute(RSQL)
        If Not rs.EOF Then
            If rs("CACODMON") = "01" Then
                Combo3.ListIndex = 0
                sCodMon = "01"
            Else
                Combo3.ListIndex = 1
                sCodMon = "02"
            End If
         Fecha = rs(1)
         End If
         rs.Close
         Set rs = Nothing
         RSQL = "select   deprecio=isnull(n.DEPRECIO,0),detipcam=isnull(n.DETIPCAM,0),decantid from  MovAlmDet n where   n.DEALMA ='" & VGAlma & "' and DETD='" & Trim(FG.TextMatrix(FG.Row, 2)) & "' AND n.DECODIGO='" & Trim(FG.TextMatrix(FG.Row, 0)) & "' and n.DENUMDOC= '" & Trim(FG.TextMatrix(FG.Row, 3)) & "'"
          Set rs = VGCNx.Execute(RSQL)
         If rs.EOF Then
                Exit Sub
         End If
        Text4.text = rs!DECANTID

         If sCodMon = "01" Then
             precioant = rs(0)
         Else
             If IsNull(rs(1)) Then
                MsgBox "Ud no ha ingresado el tipo de cambio", vbInformation, "Aviso"
             End If
             If rs(1) <> 0 Then
                precioant = rs(0)
             Else
                precioant = rs(0)
             End If
         End If
         Text2 = IIf(Not IsNull(rs(1)), rs(1), 0)
         Text3 = precioant
          Label10 = FG.TextMatrix(FG.Row, 3)
          Label11 = FG.TextMatrix(FG.Row, 2)
          Label16 = FG.TextMatrix(FG.Row, 0)
          Label18 = FG.TextMatrix(FG.Row, 1)
          Label13 = FG.TextMatrix(FG.Row, 4)     ' proveedor
          Label12 = FG.TextMatrix(FG.Row, 5)
          'Label14 = FG.TextMatrix(FG.Row, 6)
          Text1 = FG.TextMatrix(FG.Row, 6)
          If Label12 <> "" Then
              Label19 = tipref(Label12)
          End If
          If Label11 <> "" Then Label20 = transa(Label11)
          Call cantidad_art(cant, Serie, Lote)
          Text4.Enabled = True
          Text4 = cant
          If Lote = "" Then
                Label14 = Serie
          Else
                Label14 = Lote
          End If
          Text4.Enabled = False
          Text3.SetFocus
 End If
End Sub

Private Sub Command7_Click()
  If Frame1.Visible Then
        Frame1.Visible = False
        Command1.Caption = "&Corregir"
        Text4 = ""
        Text5 = ""
  Else
        Unload Me
  End If
End Sub



Private Sub FG_Click()
  'Dim db As Database
  Dim rs As Recordset
  Dim RSQL As String

End Sub

Private Sub Form_Load()
Call Ctr_ayuAlmacen.conexion(VGCNx)

End Sub
Private Sub Ctr_ayuAlmacen_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
VGAlma = Ctr_ayuAlmacen.xclave

Frame1.Visible = True
Frame2.Visible = True
cargar
End Sub

Private Sub cargar()

  
  Dim RSQL As String

 '****************************************************RMM 07/07/2001
  Set rsSTKART = New ADODB.Recordset
  rsSTKART.Open "Select * from STKART WHERE STALMA='" & VGAlma & "'", VGCNx, adOpenDynamic, adLockOptimistic
 '******************************************************************************************
  
  FG.FormatString = "Codigo Art.|Descripcion| TD | Num.Doc||"
  FG.Row = 0
  'Text1.SetFocus
  Label14 = ""
  Label19 = ""
  FG.Cols = 7
  FG.ColWidth(0) = 1400
  FG.ColWidth(1) = 3200
  FG.ColWidth(2) = 800
  FG.ColWidth(3) = 1500
  FG.ColWidth(4) = 2
  FG.ColWidth(5) = 2
  FG.ColWidth(6) = 2
  FG.ColAlignment(0) = 1
  Combo3.ListIndex = 0
  Combo4.ListIndex = 0
  Combo1.ListIndex = 0
  RSQL = "Select  n.DECODIGO, ADESCRI, N.DETD,n.DENUMDOC, m.CANOMPRO, m.CARFTDOC,m.CARFNDOC " & _
             "from MovAlmCab m, MovAlmDet n ,MaeArt  WHERE   m.CaALMA ='" & VGAlma & "'  AND " & _
             " isnull(m.CACIERRE,0)=0  AND  ( m.CATD ='NI' or m.CATD ='NC')    AND  n.DENUMDOC =m.CANUMDOC  and ACODIGO  = n.DECODIGO   and n.DEALMA=m.CAALMA  and n.DETD= m.CATD AND CASITGUI<>'A' ORDER BY n.DECODIGO, n.DENUMDOC"
   
  Set rs = VGCNx.Execute(RSQL)
  
  FG.Rows = 1
   Frame1.Visible = False
   
  If rs.EOF Then
     MsgBox "No hay articulo valorizados para corregir", vbInformation, mensaje1
      central Me
     Exit Sub
  End If
  rs.MoveFirst
  FG.Visible = False
  
  While Not rs.EOF
       'FG.AddItem (rS(0) & vbTab & rS(1) & vbTab & rS(2) & vbTab & -rS(3) & vbTab & rS(4) & vbTab & rS(5) & vbTab & rS(6))
        FG.AddItem (rs(0) & vbTab & rs(1) & vbTab & rs(2) & vbTab & rs(3) & vbTab & rs(4) & vbTab & rs(5) & vbTab & rs(6))
        rs.MoveNext
  Wend
  FG.Visible = True
  rs.Close
  central Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        rsSTKART.Close
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
       SendKeys "{tab}"
  End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And IsNumeric(Text2) Then
      SendKeys "{tab}"
    Else
      If Chr$(KeyAscii) = "." Then Exit Sub
      If ((Chr$(KeyAscii) > "9" Or Chr$(KeyAscii) < "0")) And KeyAscii <> 8 Then KeyAscii = 0
    End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
      If KeyAscii = 13 And IsNumeric(Text3) Then
                Command1.SetFocus
                Exit Sub
      End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   If IsNumeric(Text4) And KeyAscii = 13 And IsNumeric(Text3) Then
        If Not IsNumeric(Text3) Then Exit Sub
        Text5 = Val(Text3) * Val(Text4)
   ElseIf KeyAscii = 13 And IsNumeric(Text5) <> 0 And IsNumeric(Text4) <> 0 Then
         Text3 = Format(Val(Text5) / Val(Text4), "##0.0000")
   Else
        If Chr$(KeyAscii) = "." Then Exit Sub
        If ((Chr$(KeyAscii) > "9" Or Chr$(KeyAscii) < "0")) And KeyAscii <> 8 Then KeyAscii = 0
   End If
End Sub

Private Sub Text5_Change()
If Trim(Text4) <> "" Then
   Text3 = Format(Val(Text5) / Val(Text4), "###0.0000")
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

Function tipref(text As Label) As String
 Dim rs As Recordset
 Dim RSQL As String
  RSQL = "select  TDO_DESCRI  FROM TIPO_DOCU  where TDO_TIPDOC = '" & text & "'" '
  'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
  
  Set rs = VGCNx.Execute(RSQL)
  tipref = IIf(Not rs.EOF, rs(0), "")
  rs.Close
End Function

Function transa(text As Label) As String
 Dim rs As Recordset
 Dim RSQL As String
 Dim dato As String
  dato = "I"
  RSQL = "select  TT_DESCRI FROM TabTransa where TT_CODMOV= '" & text & "' and TT_TIPMOV ='" & dato & "'" '
  'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
  Set rs = VGCNx.Execute(RSQL)
  transa = IIf(Not rs.EOF, rs(0), "")
  rs.Close
End Function

Private Sub cantidad_art(pcantidad As String, pserie As String, plote As String)
 Dim Adoreg1 As ADODB.Recordset
 Dim RSQL As String
 RSQL = "select decantid,delote,deserie from MovAlmdet where DENUMDOC='" & Trim(FG.TextMatrix(FG.Row, 3)) & "' and DECODIGO ='" & Trim(FG.TextMatrix(FG.Row, 0)) & "' and DEALMA = '" & VGAlma & "'  AND DETD='" & Trim(FG.TextMatrix(FG.Row, 2)) & "' "
 Set Adoreg1 = New ADODB.Recordset
Adoreg1.Open RSQL, VGCNx, adOpenDynamic, adLockOptimistic
 If Adoreg1.RecordCount = 0 Then
    pcantidad = ""
    pserie = ""
    plote = ""
  Else
    pcantidad = Str(Adoreg1(0))
    pserie = IIf(Not IsNull(Adoreg1(2)), Adoreg1(2), "")
    plote = IIf(Not IsNull(Adoreg1(1)), Adoreg1(1), "")
  End If
End Sub

Public Sub grabastk()
   Dim criterio As String
   Dim cadena As String
   Dim auxdisp As Double
   Dim AUXPRECIO As Double
   Dim AUXPRECIOANT As Double
   cadena = Label16
   '**************RMM  07/07/2001************************
   criterio = " STCODIGO='" & cadena & "' and  STALMA ='" & VGAlma & "'"
   rsSTKART.Filter = criterio
   '*****************************************************
      
  If Combo3.ListIndex = 0 Then
       AUXPRECIO = Text3
       AUXPRECIOANT = precioant
   Else
       If Val(Text2) <> 0 Then
          AUXPRECIO = Val(Text3) * Val(Text2)
          AUXPRECIOANT = precioant * Val(Text2)
       Else
          AUXPRECIO = 0
          AUXPRECIOANT = 0
       End If
   End If

   If Not rsSTKART.EOF Then
     'Data3.*Recordset.Edit
     auxdisp = rsSTKART("STSKDIS")
     If rsSTKART("STKPREPRO") <> 0 And (auxdisp <> 0) Then  'no se registrado algun precio
        rsSTKART("STKPREPRO") = (auxdisp * rsSTKART("STKPREPRO") - Val(Text4) * (AUXPRECIOANT - Val(AUXPRECIO))) / auxdisp
        If IsNull(rsSTKART("stkultfechacompra")) Or (rsSTKART("stkultfechacompra") <= Fecha) Then rsSTKART("stkultfechacompra") = Fecha
     End If
     If IsNull(rsSTKART("STKFECULT")) Or (rsSTKART("STKFECULT") <= Format(Fecha, "DD/MM/YYYY")) Then
       rsSTKART("STKPREPRO") = AUXPRECIO    'IIf(sCodMon = "01", Val(Text3), Val(Text3) * Val(Text2))
       rsSTKART("STKPREULT") = AUXPRECIOANT 'IIf(sCodMon = "01", Val(Text3), Val(Text3) * Val(Text2))
       rsSTKART("STKFECULT") = Format(Fecha, "DD/MM/YYYY")
     End If
     
   Else
     rsSTKART.AddNew
     rsSTKART!stalma = Ctr_ayuAlmacen.xclave
    rsSTKART!stcodigo = cadena
    rsSTKART!STSKDIS = Text4.text
    rsSTKART("STKPREPRO") = Val(AUXPRECIO)
    rsSTKART("STKPREULT") = AUXPRECIOANT
    rsSTKART("STKFECULT") = Format(Fecha, "DD/MM/YYYY")
   End If
   rsSTKART.Update
   'Data3.*Refresh
End Sub

Private Sub Txtbuscar_Change()
Dim I As Integer
   Dim n As Integer
   n = Combo1.ListIndex
   If n = 2 Then n = n + 1
   If TxtBuscar <> "" And n >= 0 Then
      For I = 1 To FG.Rows - 1
          If UCase(Left(FG.TextMatrix(I, n), Len(TxtBuscar))) = UCase(Trim(TxtBuscar)) Then
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
