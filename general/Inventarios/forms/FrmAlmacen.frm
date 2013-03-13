VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmAlmacen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Almacen"
   ClientHeight    =   4455
   ClientLeft      =   2025
   ClientTop       =   435
   ClientWidth     =   7755
   Icon            =   "FrmAlmacen.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   7755
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   975
      Left            =   195
      TabIndex        =   21
      Top             =   3480
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CommandButton Command20 
         Caption         =   "&Grabar"
         Height          =   735
         Left            =   2160
         Picture         =   "FrmAlmacen.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   840
      End
      Begin VB.CommandButton Command21 
         Caption         =   "&Salir"
         Height          =   735
         Left            =   4920
         Picture         =   "FrmAlmacen.frx":0D0C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   840
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   285
      Top             =   3765
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Height          =   3330
      Left            =   210
      TabIndex        =   16
      Top             =   90
      Visible         =   0   'False
      Width           =   7185
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Almacen Valorizado"
         Height          =   255
         Left            =   3960
         TabIndex        =   32
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   29
         Text            =   "Text8"
         Top             =   3030
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   4410
         MaxLength       =   7
         TabIndex        =   9
         Text            =   "Text10"
         Top             =   2745
         Width           =   1335
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   1800
         MaxLength       =   7
         TabIndex        =   8
         Text            =   "Text9"
         Top             =   2700
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "FrmAlmacen.frx":114E
         Left            =   1800
         List            =   "FrmAlmacen.frx":1158
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   7
         Text            =   "Text7"
         Top             =   2370
         Width           =   1695
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   6
         Text            =   "Text6"
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   4695
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1800
         MaxLength       =   25
         TabIndex        =   1
         Top             =   600
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   0
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1800
         MaxLength       =   40
         TabIndex        =   2
         Top             =   960
         Width           =   4215
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   3
         Top             =   1335
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Num. Serie"
         Height          =   255
         Left            =   390
         TabIndex        =   30
         Top             =   3030
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Num. Final"
         Height          =   255
         Left            =   3600
         TabIndex        =   27
         Top             =   2790
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Ult.Guia Imp"
         Height          =   255
         Left            =   390
         TabIndex        =   26
         Top             =   2730
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Numeración"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Nro. P/Salida"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Nro .P/Ingreso"
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Telefono"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3720
         TabIndex        =   22
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Direccion"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Distrito"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Nombre"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Salir"
      Height          =   675
      Left            =   5760
      Picture         =   "FrmAlmacen.frx":1170
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3615
      Width           =   840
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Eliminar"
      Height          =   675
      Left            =   3405
      Picture         =   "FrmAlmacen.frx":15B2
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3615
      Width           =   840
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Modificar"
      Height          =   675
      Left            =   2280
      Picture         =   "FrmAlmacen.frx":19F4
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3615
      Width           =   840
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Adicionar"
      Height          =   675
      Left            =   1170
      Picture         =   "FrmAlmacen.frx":1E36
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3615
      Width           =   855
   End
   Begin VB.CommandButton command4 
      Caption         =   "&Reporte"
      Height          =   675
      Left            =   4620
      Picture         =   "FrmAlmacen.frx":2278
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   3615
      Width           =   825
   End
   Begin MSDataGridLib.DataGrid DBGrid1 
      Height          =   3255
      Left            =   210
      TabIndex        =   31
      Top             =   90
      Width           =   7185
      _ExtentX        =   12674
      _ExtentY        =   5741
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmAlmacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim modomodificar As Boolean
Dim rs As New ADODB.Recordset
Dim VGDllGeneral As New dllgeneral.dll_general


'Adiciconar
Private Sub Command1_Click()
   modomodificar = False
   Text1.Enabled = True
   Frame1.Visible = True
   Frame2.Visible = True
   Text1.SetFocus
End Sub
'Modificar
Private Sub Command2_Click()
Dim rb As New ADODB.Recordset
If rs.RecordCount > 0 Then
       Frame1.Visible = True
       Frame2.Visible = True
       Set rb = VGCNx.Execute("select * from tabalm where taalma='" & DBGrid1.Columns(0).text & "'")
       If rb.RecordCount > 0 Then
           Text1 = rb.Fields("TAALMA")  'guardo el codigo
           If Not IsNull(rb.Fields("TADESCRI")) Then
                Text2 = rb.Fields("TADESCRI")
           End If
           If Not IsNull(rb.Fields("TADIRECC")) Then
                Text3 = rb.Fields("TADIRECC")
           End If
           If Not IsNull(rb.Fields("TADISTRI")) Then
                Text4 = rb.Fields("TADISTRI")
           End If
           If Not IsNull(rb.Fields("TATELEFO")) Then
                Text5 = rb.Fields("TATELEFO")
           End If
           If Not IsNull(rb.Fields("TANUMENT")) Then Text6 = rb.Fields("TANUMENT")
           If Not IsNull(rb.Fields("TANUMSAL")) Then Text7 = rb.Fields("TANUMSAL")
           If Not IsNull(rb.Fields("TANUMFAC")) Then Text9 = rb.Fields("TANUMFAC")
           If Not IsNull(rb.Fields("TANUMND")) Then Text10 = rb.Fields("TANUMND")
           Check1.Value = ESNULO(rb.Fields("almacenvalorizado"), 0)
           Combo1.ListIndex = 0
    '       If Data1.Recordset("TACTLNUM") = "A" Then
    '           Combo1.ListIndex = 0
    '       Else
    '           Combo1.ListIndex = 1
    '       End If
            modomodificar = True
            Text1.Enabled = False
       End If
End If
End Sub

Private Sub Command4_Click()
    Dim cadena As String
    Dim cNomRepor  As String

cNomRepor = "almacenes.RPT"
If Trim(cNomRepor) <> "" Then
    CrystalReport1.Reset
    CrystalReport1.WindowTitle = "Reporte de Almacenes"
    CrystalReport1.ReportFileName = VGParamSistem.RutaReport + "\" + cNomRepor
 
    CrystalReport1.Connect = VGCadenaReport2
    CrystalReport1.StoredProcParam(0) = VGParamSistem.BDEmpresa
    
    CrystalReport1.formulas(0) = "emp ='" & VGParametros.RucEmpresa & "'"
    
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.WindowState = crptMaximized
    If CrystalReport1.Status <> 2 Then
        CrystalReport1.Action = 1
    End If
Else
    MsgBox "No existe el nombre del Reporte, verifique en Formatos", vbInformation, "Información"
    Exit Sub
End If

End Sub
'Grabar
Private Sub Command20_Click()
 Dim criterio As String
 Dim RSG As New ADODB.Recordset
 Dim rsb As New ADODB.Recordset
 
 If Frame1.Visible Then
  If modomodificar Then
     criterio = "TAALMA = " & "'" + Text1.text + "'"
     Set rsb = VGCNx.Execute("SELECT * FROM TABALM")
     If rsb.RecordCount > 0 Then
           VGCNx.Execute "Update Tabalm " & _
                             "Set TADESCRI ='" & LTrim(Text2) & "'," & _
                             "TADIRECC='" & LTrim(Text3) & "'," & _
                             "TADISTRI='" & LTrim(Text4) & "'," & _
                             "TATELEFO='" & LTrim(Text5) & "'," & _
                             "TANUMENT=" & Text6 & "," & _
                             "TANUMSAL=" & Text7 & "," & _
                             "TANUMNC= " & CDbl(0) & "," & _
                             "TANUMFAC = '" & Text9 & "'," & _
                             "TANUMND = '" & Text10 & "'," & _
                             "almacenvalorizado =" & Check1.Value & "," & _
                             "tactlnum='" & IIf(Combo1.ListIndex = 0, "A", "M") & "'" & _
                             " Where TAALMA='" & Text1.text & "'"
     End If
     rsb.Close
     Set rsb = Nothing
     
  ElseIf Len(Trim(Text1)) = 2 Then
         criterio = "TAALMA = " & "'" + Text1.text + "'"
         Set rsb = VGCNx.Execute("SELECT * FROM TABALM WHERE " & criterio)
         If rsb.RecordCount > 0 Then
            MsgBox "El Código del Almacen ya existe "
            Text1.SetFocus
            Exit Sub
         Else
          SQL = "INSERT INTO Tabalm (TAALMA,TADESCRI,TADIRECC,TADISTRI,TATELEFO,TANUMENT,empresacodigo,"
          SQL = SQL & "puntovtacodigo,TANUMSAL,TANUMNC,TANUMFAC,TANUMND,almacenvalorizado,tactlnum)"
          SQL = SQL & " VALUES('" & Text1 & "','" & LTrim(Text2) & "','" & LTrim(Text3) & "',"
          SQL = SQL & "'" & LTrim(Text4) & "','" & LTrim(Text5) & "'," & IIf(Len(Trim(Text6)) = 0, 0, Text6) & ","
          SQL = SQL & "'" & VGParametros.empresacodigo & "','" & VGParametros.puntovta & "',"
          SQL = SQL & "" & IIf(Len(Trim(Text7)) = 0, 0, Text7) & "," & CDbl(0) & ","
          SQL = SQL & "" & IIf(Len(Trim(Text9)) = 0, 0, Text9) & "," & IIf(Len(Trim(Text10)) = 0, 0, Text10) & ","
          SQL = SQL & "" & Check1.Value & ",'" & IIf(Combo1.ListIndex = 0, "A", "M") & "')"
          VGCNx.Execute (SQL)
          criterio = "TAALMA = " & "'" + Text1.text + "' and puntovtacodigo='" & VGParametros.puntovta & "'"
          Call Listado("select * from tabalm where " & criterio & "")
         End If
  Else
      MsgBox "Ingrese el codigo", vbExclamation, "Control de Inventarios'"
      Text1.SetFocus
      Exit Sub
  End If
     limpia
     Text1.Enabled = True
 End If
 modomodificar = True
' Unload Me
  Command21_Click
End Sub
'Salir de una opción
Private Sub Command21_Click()
     limpia
     Frame1.Visible = False
     Frame2.Visible = False
     modomodificar = False
End Sub
'Eliminar
Private Sub Command3_Click()
Dim csql As String
Dim cCodigo1 As String
Dim Cod As String
Dim cSqlA As String, cSelA As ADODB.Recordset

If rs.RecordCount > 0 Then
    cSqlA = "Select * FROM STKART WHERE STALMA = '" & Trim(rs.Fields("taalma")) & "'"
    Set cSelA = New ADODB.Recordset
    cSelA.Open cSqlA, VGCNx, adOpenStatic
    If cSelA.RecordCount > 0 Then
            MsgBox "El Almacen tiene registrado articulos, no se puede eliminar", vbInformation, "Eliminacion de Registro"
            cSelA.Close: Exit Sub
    End If
    cSelA.Close
    Cod = Trim(rs.Fields("taalma"))
    If MsgBox("Seguro de Eliminar ?", vbQuestion + vbOKCancel, mensaje1) = vbOK Then
        csql = "delete  from tabalm where taalma = '" & Cod & " '"
        VGCNx.BeginTrans
        VGCNx.Execute csql
        VGCNx.CommitTrans
        Call Listado("select * from tabalm")
    End If
End If
End Sub


Sub Listado(wcad)
  Set DBGrid1.DataSource = Nothing
  Set rs = Nothing
  
  Set rs = VGCNx.Execute(wcad)
  Set DBGrid1.DataSource = rs
  With DBGrid1
      .Columns(0).Caption = "Codigo"
      .Columns(0).Width = 1000
      .Columns(1).Caption = "Nombre"
      .Columns(1).Width = 1800
      .Columns(2).Caption = "Numeracion"
      .Columns(2).Width = 1000
      .Columns(3).Caption = "Ult. Entrada"
      .Columns(3).Width = 1000
      .Columns(4).Caption = "Ult. Salida"
      .Columns(4).Width = 1000
      .MarqueeStyle = dbgHighlightRow
      .Refresh
  End With

End Sub
'Salir el Formulario
Private Sub Command7_Click()
    Unload Me
End Sub

Private Sub Form_Load()
  central FrmAlmacen
  
  'Data1.DatabaseName = cRuta2
  'Data1.ConnectionString = ""
'  Data1.Refresh
  Call Listado("SELECT * FROM TABALM where empresacodigo='" & VGParametros.empresacodigo & "' and puntovtacodigo='" & VGParametros.puntovta & "'")
  
  'Init_ControlDBGrid DBGrid1
  Combo1.ListIndex = 0
  limpia
End Sub

Private Sub Text1_GotFocus()
Enfoque Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      Dim criterio As String
      Text1 = Format(Text1, "00")
      If Len(Text1.text) = 2 Then
         criterio = "TAALMA = " & "'" + Text1.text + "'"
         If VGDllGeneral.VerificaDatoExistente(VGCNx, "select * from tabalm where " & criterio) = 1 Then
               MsgBox "El Código del Almacen ya existe "
               Text1.SetFocus
         Else
               Text2.SetFocus
         End If
     End If
 End If
End Sub

Private Sub graba()
'    If Text2 <> "" Then
'        Data1.Recordset("TADESCRI") = LTrim(Text2)
'     End If
'     If Text3 <> "" Then
'       Data1.Recordset("TADIRECC") = LTrim(Text3)
'     End If
'     If Text4 <> "" Then
'       Data1.Recordset("TADISTRI") = LTrim(Text4)
'     End If
'     If Text5 <> "" Then
'       Data1.Recordset("TATELEFO") = LTrim(Text5)
'     End If
'     If IsNumeric(Text6) Then Data1.Recordset("TANUMENT") = Text6
'     If IsNumeric(Text7) Then Data1.Recordset("TANUMSAL") = Text7
'     'If IsNumeric(Text8) Then
'     Data1.Recordset("TANUMNC") = 0
'     If IsNumeric(Text9) Then Data1.Recordset("TANUMFAC") = Text9
'     If IsNumeric(Text10) Then Data1.Recordset("TANUMND") = Text10
'     If Combo1.ListIndex = 0 Then
'       Data1.Recordset("TACTLNUM") = "A"
'     Else
'       Data1.Recordset("TACTLNUM") = "M"
'     End If
'     Data1.Recordset.Update
'     Data1.Refresh
End Sub

Private Sub limpia()
     Text1 = ""
     Text2 = ""
     Text3 = ""
     Text4 = ""
     Text5 = ""
     Text6 = ""
     Text7 = ""
'     Text8 = ""
     Text9 = ""
     Text10 = ""
End Sub

Private Sub Text10_GotFocus()
Enfoque Text10
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If NumPto(KeyAscii) Then
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
Else
    KeyAscii = 0
End If
End Sub

Private Sub Text2_GotFocus()
Enfoque Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{tab}"
  KeyAscii = 0
End If
End Sub

Private Sub Text3_GotFocus()
Enfoque Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     SendKeys "{tab}"
    KeyAscii = 0
  End If
End Sub

Private Sub Text4_GotFocus()
Enfoque Text4
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
    SendKeys "{tab}"
    KeyAscii = 0
   End If
End Sub

Private Sub Text5_GotFocus()
Enfoque Text5
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      SendKeys "{tab}"
       KeyAscii = 0
   End If
End Sub

Private Sub Text6_GotFocus()
Enfoque Text6
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If NumPto(KeyAscii) Then
    If KeyAscii = 13 Then
            SendKeys "{tab}"
            KeyAscii = 0
    End If
Else
    KeyAscii = 0
End If
End Sub

Private Sub Text7_GotFocus()
Enfoque Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If NumPto(KeyAscii) Then
    If KeyAscii = 13 Then
            SendKeys "{tab}"
            KeyAscii = 0
    End If
Else
        KeyAscii = 0
End If
End Sub


Private Sub Text9_GotFocus()
Enfoque Text9
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If NumPto(KeyAscii) Then
    If KeyAscii = 13 Then
              SendKeys "{tab}"
              KeyAscii = 0
    End If
Else
    KeyAscii = 0
End If
End Sub
