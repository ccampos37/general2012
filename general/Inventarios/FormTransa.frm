VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FormTransa 
   Caption         =   "Transacciones"
   ClientHeight    =   5070
   ClientLeft      =   2160
   ClientTop       =   1605
   ClientWidth     =   6795
   Icon            =   "FormTransa.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   6795
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   915
      Left            =   120
      TabIndex        =   0
      Top             =   4110
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CommandButton Command20 
         Caption         =   "&Grabar"
         Height          =   735
         Left            =   1500
         Picture         =   "FormTransa.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   105
         Width           =   855
      End
      Begin VB.CommandButton Command21 
         Caption         =   "&Salir"
         Height          =   735
         Left            =   4215
         Picture         =   "FormTransa.frx":0D0C
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   105
         Width           =   855
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   30
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Salir"
      Height          =   690
      Left            =   5265
      Picture         =   "FormTransa.frx":114E
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4155
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Eliminar"
      Height          =   690
      Left            =   2955
      Picture         =   "FormTransa.frx":1590
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4185
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Modificar"
      Height          =   690
      Left            =   1755
      Picture         =   "FormTransa.frx":19D2
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4170
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Adicionar"
      Height          =   690
      Left            =   585
      Picture         =   "FormTransa.frx":1E14
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4170
      Width           =   855
   End
   Begin VB.CommandButton command4 
      Caption         =   "&Reporte"
      Height          =   675
      Left            =   4125
      Picture         =   "FormTransa.frx":211E
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4185
      Width           =   775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detalle"
      Height          =   3975
      Left            =   30
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   6735
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FormTransa.frx":2560
         Left            =   4680
         List            =   "FormTransa.frx":256A
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.Frame Frame3 
         Caption         =   "Configurar datos del Documento"
         Height          =   2655
         Left            =   210
         TabIndex        =   23
         Top             =   1110
         Width           =   6255
         Begin VB.CheckBox Check12 
            Caption         =   "Maquinas/Equipos"
            Height          =   375
            Left            =   3210
            TabIndex        =   27
            Top             =   1800
            Width           =   1755
         End
         Begin VB.CheckBox Check11 
            Caption         =   "Orden de Fabricación"
            Height          =   375
            Left            =   360
            TabIndex        =   26
            Top             =   1800
            Width           =   2355
         End
         Begin VB.CheckBox Check10 
            Caption         =   "Orden de consumo"
            Height          =   375
            Left            =   360
            TabIndex        =   12
            Top             =   2190
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.CheckBox Check9 
            Caption         =   "Cliente"
            Height          =   375
            Left            =   360
            TabIndex        =   11
            Top             =   1470
            Width           =   1695
         End
         Begin VB.CheckBox Check8 
            Caption         =   "Valorizado"
            Height          =   375
            Left            =   3210
            TabIndex        =   17
            Top             =   2190
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CheckBox Check7 
            Caption         =   "Almacen"
            Height          =   375
            Left            =   3210
            TabIndex        =   13
            Top             =   360
            Width           =   1695
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Comentario"
            Height          =   375
            Left            =   3210
            TabIndex        =   16
            Top             =   1440
            Width           =   1695
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Orden de Compra"
            Height          =   375
            Left            =   3210
            TabIndex        =   15
            Top             =   1080
            Width           =   2055
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Centro de Costo"
            Height          =   375
            Left            =   3210
            TabIndex        =   14
            Top             =   720
            Width           =   1695
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Doc. Ref"
            Height          =   375
            Left            =   360
            TabIndex        =   9
            Top             =   720
            Width           =   1695
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Autorizado"
            Height          =   375
            Left            =   360
            TabIndex        =   10
            Top             =   1080
            Width           =   1695
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Proveedor"
            Height          =   375
            Left            =   360
            TabIndex        =   8
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2040
         MaxLength       =   40
         TabIndex        =   7
         Top             =   720
         Width           =   4095
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo de Trans."
         Height          =   255
         Left            =   3120
         TabIndex        =   24
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Transaccion"
         Height          =   375
         Left            =   360
         TabIndex        =   18
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo Trans."
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSDataGridLib.DataGrid DbGrid1 
      Height          =   3855
      Left            =   180
      TabIndex        =   28
      Top             =   210
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   6800
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
Attribute VB_Name = "FormTransa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim modomodificar As Boolean
Dim rs As New ADODB.Recordset
Dim adll As New dllgeneral.dll_general


Private Sub Combo1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
End If
End Sub

Private Sub Command1_Click()
   Frame1.Visible = True
   Frame2.Visible = True
   Combo1.Enabled = True
   Text1.Enabled = True
   Clear_Cheks
End Sub
'Modificar
Private Sub Command2_Click()
       Dim rb As New ADODB.Recordset
       On Local Error GoTo ERRAR
       
       If rs.RecordCount = 0 Then
          Exit Sub
       End If
       Set rb = cConexCom.Execute("SELECT * FROM TABTRANSA WHERE TT_TIPMOV='" & DBGrid1.Columns(0).text & "' AND TT_CODMOV='" & DBGrid1.Columns(1).text & "'")
       If rb.RecordCount > 0 Then
         Clear_Cheks
         Frame1.Visible = True
         Frame2.Visible = True
         Text1 = rb.Fields("TT_CODMOV") 'guardo el codigo
         If Not IsNull(rb.Fields("TT_DESCRI")) Then
              Text2 = rb.Fields("TT_DESCRI")
         End If
         If rb.Fields("TT_TIPMOV") = "I" Then
             Combo1.text = "Ingreso"
         Else
             Combo1.text = "Salida"
         End If
         If rb.Fields("TT_PRV") = "S" Then
             Check1.Value = 1
         End If
         If rb.Fields("TT_DR") = "S" Then
             Check3.Value = 1
         End If
         If rb.Fields("TT_AT") = "S" Then
             Check2.Value = 1
         End If
         If rb.Fields("TT_CC") = "S" Then
             Check4.Value = 1
         End If
         If rb.Fields("TT_OC") = "S" Then
             Check5.Value = 1
         End If
         If rb.Fields("TT_CO") = "S" Then
             Check6.Value = 1
         End If
         If rb.Fields("TT_ALMA") = "S" Then
             Check7.Value = 1
         End If
         If rb.Fields("TT_ORCON") = "S" Then
             Check10.Value = 1
         End If
         If rb.Fields("TT_CLIE") = "S" Then
             Check9.Value = 1
         End If
         If rb.Fields("TT_VAL") = "V" Then
             Check8.Value = 1
         End If
         
         If cNull(rb.Fields("TT_ORDFAB")) = "S" Then
             Check11.Value = 1
         End If
        
         If cNull((rb.Fields("TT_EQUIP"))) = "S" Then
             Check12.Value = 1
         End If
        
         Text1.Enabled = False
         Combo1.Enabled = False
         modomodificar = True
    End If
Exit Sub
ERRAR:
      MsgBox "Error en Estructura,Verifique la Estructura de Base de Datos", vbCritical, "Error en Base de Datos"
End Sub
'Grabar
Private Sub Command20_Click()

On Error GoTo Err
If Frame1.Visible Then
  If modomodificar Then
    If adll.VerificaDatoExistente(cConexCom, "select * from tabtransa Where TT_TIPMOV='" & DBGrid1.Columns(0).text & "' and TT_CODMOV='" & Text1 & "'") = 1 Then
        cConexCom.Execute "Update Tabtransa " & _
                          " Set TT_CODMOV='" & Text1 & "'," & _
                          "TT_DESCRI='" & IIf(IsNull(Text2), "", Text2) & "'," & _
                          "TT_TIPMOV='" & Left(Combo1.text, 1) & "'," & _
                          "TT_PRV='" & IIf(Check1.Value = 1, "S", "N") & "'," & _
                          "TT_DR='" & IIf(Check3.Value = 1, "S", "N") & "'," & _
                          "TT_AT='" & IIf(Check2.Value = 1, "S", "N") & "'," & _
                          "TT_CC='" & IIf(Check4.Value = 1, "S", "N") & "'," & _
                          "TT_OC='" & IIf(Check5.Value = 1, "S", "N") & "', " & _
                          "TT_CO='" & IIf(Check6.Value = 1, "S", "N") & "'," & _
                          "TT_ALMA='" & IIf(Check7.Value = 1, "S", "N") & "'," & _
                          "TT_ORCON='" & IIf(Check10.Value = 1, "S", "N") & "'," & _
                          "TT_CLIE='" & IIf(Check9.Value = 1, "S", "N") & "'," & _
                          "TT_VAL='" & IIf(Check8.Value = 1, "V", "F") & "'," & _
                          "TT_ORDFAB='" & IIf(Check11.Value = 1, "S", "N") & "'," & _
                          "TT_EQUIP='" & IIf(Check12.Value = 1, "S", "N") & "' " & _
                          " Where TT_TIPMOV='" & DBGrid1.Columns(0).text & "' and TT_CODMOV='" & Text1 & "'"
                          
        Call listado("select  TT_TIPMOV,TT_CODMOV,TT_DESCRI from tabtransa")
     End If
  Else
   If Text1 <> "" Then
        If adll.VerificaDatoExistente(cConexCom, "select * from tabtransa Where TT_TIPMOV='" & Left(Combo1.text, 1) & "' and TT_CODMOV='" & Text1 & "'") = 0 Then
            cConexCom.Execute "INSERT INTO Tabtransa " & _
                            "(TT_CODMOV,TT_DESCRI,TT_TIPMOV,TT_PRV," & _
                            "TT_DR,TT_AT,TT_CC,TT_OC,TT_CO,TT_ALMA," & _
                            "TT_ORCON,TT_CLIE,TT_VAL,TT_ORDFAB,TT_EQUIP,TT_CONT)" & _
                            " VALUES (" & _
                            "'" & Text1 & "'," & _
                            "'" & IIf(IsNull(Text2), "", Text2) & "'," & _
                            "'" & Left(Combo1.text, 1) & "'," & _
                            "'" & IIf(Check1.Value = 1, "S", "N") & "'," & _
                            "'" & IIf(Check3.Value = 1, "S", "N") & "'," & _
                            "'" & IIf(Check2.Value = 1, "S", "N") & "'," & _
                            "'" & IIf(Check4.Value = 1, "S", "N") & "'," & _
                            "'" & IIf(Check5.Value = 1, "S", "N") & "', " & _
                            "'" & IIf(Check6.Value = 1, "S", "N") & "'," & _
                            "'" & IIf(Check7.Value = 1, "S", "N") & "'," & _
                            "'" & IIf(Check10.Value = 1, "S", "N") & "'," & _
                            "'" & IIf(Check9.Value = 1, "S", "N") & "'," & _
                            "'" & IIf(Check8.Value = 1, "V", "F") & "'," & _
                            "'" & IIf(Check11.Value = 1, "S", "N") & "'," & _
                            "'" & IIf(Check12.Value = 1, "S", "N") & "',0)"
                              
            Call listado("select  TT_TIPMOV,TT_CODMOV,TT_DESCRI from tabtransa")
         End If
    Else
      MsgBox "Ingrese el Codigo"
    End If
   End If
     Text1 = ""
     Text2 = ""
     Check1.Value = 0
     Check2.Value = 0
     Check3.Value = 0
     Check4.Value = 0
     Check5.Value = 0
     Check6.Value = 0
     Check7.Value = 0
     Check8.Value = 0
     Check9.Value = 0
     Check10.Value = 0
     Text1.Enabled = True
 End If
 modomodificar = False
  Frame1.Visible = False
  Frame2.Visible = False
Exit Sub
Err:
  MsgBox Err.Description

End Sub
'Salir
Private Sub Command21_Click()
     Text1 = ""
     Text2 = ""
     Check1.Value = 0
     Check2.Value = 0
     Check3.Value = 0
     Check4.Value = 0
     Check5.Value = 0
     Check6.Value = 0
     Check7.Value = 0
     Check8.Value = 0
     Check9.Value = 0
     Check10.Value = 0
     modomodificar = False
     Frame1.Visible = False
     Frame2.Visible = False
End Sub
'Eliminar
Private Sub Command3_Click()
Dim csql As String
Dim cCodigo1 As String
Dim cSqlA As String, cSelA As ADODB.Recordset
Dim codmov As String
Dim codtipo As String

If rs.RecordCount = 0 Then Exit Sub

  codmov = Trim(rs.Fields("TT_CODMOV"))
  codtipo = rs.Fields("TT_TIPMOV")
  cSqlA = "Select * FROM MOVALMCAB WHERE CACODMOV = '" & Trim(rs.Fields("TT_CODMOV")) & "' AND CATIPMOV ='" & rs.Fields("TT_TIPMOV") & "' "
  Set cSelA = New ADODB.Recordset
  cSelA.Open cSqlA, cConexCom, adOpenStatic
  If cSelA.RecordCount > 0 Then
       If MsgBox("No se puede eliminarla porque se registro transaciones", vbYesNo, "Eliminacion de Registro") = vbNo Then
              cSelA.Close: Exit Sub
       End If
  End If
  cSelA.Close
  If MsgBox("Seguro de Eliminar", vbQuestion + vbOKCancel, mensaje1) = vbOK Then
    'cCodigo1 = RS.Fields, "TT_codmov")
    csql = "DELETE from tabtransa where tt_codmov = '" & codmov & "' and tt_tipmov = '" & codtipo & "' "
    cConexCom.BeginTrans
    cConexCom.Execute csql
    cConexCom.CommitTrans
    'Data1.Refresh
'    If cCodigo1 <> "" Then
'                RS.Find "tt_codmov='" & cCodigo1 & "'"
'    End If
    Call listado("select  TT_TIPMOV,TT_CODMOV,TT_DESCRI from tabtransa")
    'DBGrid1.Refresh
  End If
 ' cConexCom.Execute "delete from tabtransa where  tt_tipmov = ' " & RS.FIELDS(0) & "'  and  tt_codmov = '" & RS.FIELDS(1) & " '  "
' Data1.Refresh
End Sub

Private Sub Command4_Click()
    Dim cadena As String
    Dim cNomRepor  As String

cNomRepor = "transac.RPT"
If Trim(cNomRepor) <> "" Then
    CrystalReport1.Reset
    CrystalReport1.WindowTitle = "Reporte de Transacciones"
    CrystalReport1.ReportFileName = cRutP + "\" + cNomRepor
    CrystalReport1.LogOnServer "pdssql.dll", _
                VGServer, _
                VGBase3, _
                VGBUsuario, _
                VGPassw
                        
    CrystalReport1.Connect = "DSN=" & VGServer & ";DSQ=" & VGBase3 & ";UID=" & VGUsuario & ";PWD=" & VGPassw
    CrystalReport1.StoredProcParam(0) = VGBase
    
    CrystalReport1.Formulas(0) = "emp ='" & VGNemp & "'"
    
    
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

Private Sub Command7_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  central FormTransa
  'Data1.DatabaseName = cRuta2
  
  Combo1.text = "Ingreso"
  'Init_ControlDBGrid DBGrid1
  
  If VGTip_Alma = "V" Then
     Check11.Visible = False
     Check12.Visible = False
  Else
     Check11.Visible = True
     Check12.Visible = True
  End If
  Call listado("select  TT_TIPMOV,TT_CODMOV,TT_DESCRI from tabtransa")
  
End Sub

Sub listado(wcad)
  Set DBGrid1.DataSource = Nothing
  Set rs = Nothing
  
  Set rs = cConexCom.Execute(wcad)
  Set DBGrid1.DataSource = rs
   With DBGrid1
        .MarqueeStyle = dbgHighlightRow
        .Columns(0).Caption = "TMOVI"
        .Columns(0).Width = 800
        .Columns(1).Caption = "TRANSACCION"
        .Columns(1).Width = 1000
        .Columns(2).Caption = "DESCRIPCION"
        .Columns(2).Width = 4000
        .Refresh
   End With
   
End Sub


Private Sub graba()
Dim contador As Integer
contador = 0
     Text2 = UCase(Text2)
     If Trim(Text2) <> "" Then
       rs.Fields("TT_DESCRI") = Text2
     End If
      If Check1.Value = 1 Then
            rs.Fields("TT_PRV") = "S"
            contador = contador + 1
       Else
             rs.Fields("TT_PRV") = "N"
       End If
        If Check3.Value = 1 Then
            rs.Fields("TT_DR") = "S"
            contador = contador + 1
       Else
             rs.Fields("TT_DR") = "N"
       End If
        If Check2.Value = 1 Then
            rs.Fields("TT_AT") = "S"
            contador = contador + 1
       Else
             rs.Fields("TT_AT") = "N"
       End If
        If Check4.Value = 1 Then
            rs.Fields("TT_CC") = "S"
            contador = contador + 1
       Else
             rs.Fields("TT_CC") = "N"
       End If
      
       If Check5.Value = 1 Then
           rs.Fields("TT_OC") = "S"
          contador = contador + 1
       Else
           rs.Fields("TT_OC") = "N"
       End If
        If Check6.Value = 1 Then
            rs.Fields("TT_CO") = "S"
           
       Else
            rs.Fields("TT_CO") = "N"
       End If
       If Check7.Value = 1 Then
            rs.Fields("TT_ALMA") = "S"
            contador = contador + 1
       Else
             rs.Fields("TT_ALMA") = "N"
       End If
        If Check10.Value = 1 Then
            rs.Fields("TT_ORCON") = "S"
           ' contador = contador + 1
       Else
             rs.Fields("TT_ORCON") = "N"
       End If
       If Check9.Value = 1 Then
            rs.Fields("TT_CLIE") = "S"
            contador = contador + 1
       Else
             rs.Fields("TT_CLIE") = "N"
       End If
       If Check8.Value = 1 Then
            rs.Fields("TT_VAL") = "V"
       Else
             rs.Fields("TT_VAL") = "F"
       End If
       
       If Check11.Value = 1 Then
            rs.Fields("TT_ORDFAB") = "S"
       Else
             rs.Fields("TT_ORDFAB") = "N"
       End If
       
       If Check12.Value = 1 Then
            rs.Fields("TT_EQUIP") = "S"
       Else
             rs.Fields("TT_EQUIP") = "N"
       End If
       
       If Combo1.text = "Ingreso" Then
            rs.Fields("TT_TIPMOV") = "I"
       Else
            rs.Fields("TT_TIPMOV") = "S"
       End If
     rs.Fields("TT_CONT") = contador
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 And Text1 <> "" Then
        Text1 = UCase(Text1)
        SendKeys "{tab}"
        KeyAscii = 0
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 And Text2 <> "" Then
         Text2 = UCase(Text2)
        SendKeys "{tab}"
        KeyAscii = 0
End If
End Sub
Sub Clear_Cheks()
Check1.Value = 0
Check2.Value = 0
Check3.Value = 0
Check4.Value = 0
Check5.Value = 0
Check6.Value = 0
Check7.Value = 0
Check8.Value = 0
Check9.Value = 0
Check10.Value = 0
Check11.Value = 0
Check12.Value = 0
End Sub
