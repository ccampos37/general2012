VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmTransaccion 
   Caption         =   "Transacciones"
   ClientHeight    =   7575
   ClientLeft      =   2160
   ClientTop       =   1605
   ClientWidth     =   7140
   Icon            =   "FrmTransaccion.frx":0000
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   7140
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   915
      Left            =   120
      TabIndex        =   0
      Top             =   6390
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CommandButton Command20 
         Caption         =   "&Grabar"
         Height          =   735
         Left            =   1500
         Picture         =   "FrmTransaccion.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   105
         Width           =   855
      End
      Begin VB.CommandButton Command21 
         Caption         =   "&Salir"
         Height          =   735
         Left            =   4215
         Picture         =   "FrmTransaccion.frx":0D0C
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
      Picture         =   "FrmTransaccion.frx":114E
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6555
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Eliminar"
      Height          =   690
      Left            =   2955
      Picture         =   "FrmTransaccion.frx":1590
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6585
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Modificar"
      Height          =   690
      Left            =   1755
      Picture         =   "FrmTransaccion.frx":19D2
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6570
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Adicionar"
      Height          =   690
      Left            =   585
      Picture         =   "FrmTransaccion.frx":1E14
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6570
      Width           =   855
   End
   Begin VB.CommandButton command4 
      Caption         =   "&Reporte"
      Height          =   675
      Left            =   4125
      Picture         =   "FrmTransaccion.frx":211E
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6585
      Width           =   775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detalle"
      Height          =   6255
      Left            =   30
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   6735
      Begin VB.Frame Frame4 
         Height          =   1455
         Left            =   120
         TabIndex        =   29
         Top             =   4680
         Width           =   6495
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayutransa 
            Height          =   495
            Left            =   1440
            TabIndex        =   30
            Top             =   360
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   873
            XcodMaxLongitud =   0
            xcodwith        =   100
            NomTabla        =   "tabtransa"
            TituloAyuda     =   "Transaciones"
            ListaCampos     =   "tt_codmov(1),tt_descri(1)"
            XcodCampo       =   "tt_codmov"
            XListCampo      =   "tt_descri"
            ListaCamposDescrip=   "Codigo,Descripcion"
            ListaCamposText =   "tt_codmov,tt_descri"
            Requerido       =   0   'False
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayumerma 
            Height          =   495
            Left            =   1440
            TabIndex        =   32
            Top             =   840
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   873
            XcodMaxLongitud =   0
            xcodwith        =   100
            NomTabla        =   "tabtransa"
            TituloAyuda     =   "Transaciones"
            ListaCampos     =   "tt_codmov(1),tt_descri(1)"
            XcodCampo       =   "tt_codmov"
            XListCampo      =   "tt_descri"
            ListaCamposDescrip=   "Codigo,Descripcion"
            ListaCamposText =   "tt_codmov,tt_descri"
            Requerido       =   0   'False
         End
         Begin VB.Label Label4 
            Caption         =   "Transaccion merma"
            Height          =   375
            Left            =   240
            TabIndex        =   33
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Transacion automatica"
            Height          =   375
            Left            =   240
            TabIndex        =   31
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FrmTransaccion.frx":2560
         Left            =   4680
         List            =   "FrmTransaccion.frx":256A
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.Frame Frame3 
         Caption         =   "Configurar datos del Documento"
         Height          =   3495
         Left            =   210
         TabIndex        =   23
         Top             =   1110
         Width           =   6255
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "FrmTransaccion.frx":257F
            Left            =   480
            List            =   "FrmTransaccion.frx":2592
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   2880
            Width           =   1815
         End
         Begin VB.CheckBox Check14 
            Caption         =   "Intercompanias"
            Height          =   375
            Left            =   3240
            TabIndex        =   35
            Top             =   2160
            Width           =   1815
         End
         Begin VB.CheckBox Check7 
            Caption         =   "Almacen"
            Height          =   375
            Left            =   3240
            TabIndex        =   13
            Top             =   360
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
         Begin VB.CheckBox Check13 
            Caption         =   "Ingreso Comprometido"
            Height          =   255
            Left            =   360
            TabIndex        =   34
            Top             =   2160
            Width           =   2295
         End
         Begin VB.CheckBox Check12 
            Caption         =   "Maquinas/Equipos/ Proyectos"
            Height          =   375
            Left            =   3240
            TabIndex        =   27
            Top             =   1800
            Width           =   1815
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
            Left            =   3240
            TabIndex        =   12
            Top             =   2910
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
            Top             =   2550
            Width           =   1695
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Comentario"
            Height          =   375
            Left            =   3240
            TabIndex        =   16
            Top             =   1440
            Width           =   1695
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Orden de Compra"
            Height          =   375
            Left            =   3240
            TabIndex        =   15
            Top             =   1080
            Width           =   2055
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Centro de Costo"
            Height          =   375
            Left            =   3240
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
         Begin VB.Label Label5 
            Caption         =   "Tipo de transacciones"
            Height          =   255
            Left            =   360
            TabIndex        =   37
            Top             =   2520
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
      Height          =   4935
      Left            =   180
      TabIndex        =   28
      Top             =   210
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   8705
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
Attribute VB_Name = "FrmTransaccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim modomodificar As Boolean
Dim rs As New ADODB.Recordset
Dim VGDllGeneral As New dllgeneral.dll_general
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
       Set rb = VGCNx.Execute("SELECT * FROM TABTRANSA WHERE TT_TIPMOV='" & DbGrid1.Columns(0).text & "' AND TT_CODMOV='" & DbGrid1.Columns(1).text & "'")
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
         Combo2.ListIndex = ESNULO(rb!tt_orcon, 1) - 1
         If rb.Fields("TT_CLIE") = "S" Then
             Check9.Value = 1
         End If
         If rb.Fields("estadocosto") = 1 Then
             Check8.Value = 1
         End If
         
         If cNull(rb.Fields("TT_ORDFAB")) = "S" Then
             Check11.Value = 1
         End If
        
         If cNull((rb.Fields("TT_EQUIP"))) = "S" Then
             Check12.Value = 1
         End If
         If cNull((rb.Fields("ingresosfuturos"))) = "S" Then
             Check13.Value = 1
         End If
         If cNull((rb.Fields("intercompanias"))) = "S" Then
             Check14.Value = 1
         End If
         
         If cNull((rb.Fields("TT_CODTRANS_AUTO"))) <> "" Then
             Ctr_Ayutransa.xclave = rb.Fields("TT_CODTRANS_AUTO")
             Ctr_Ayutransa.Ejecutar
         End If
         If cNull((rb.Fields("TT_CODTRANS_merma"))) <> "" Then
             Ctr_Ayumerma.xclave = rb.Fields("TT_CODTRANS_merma")
             Ctr_Ayumerma.Ejecutar
         End If
        
         Text1.Enabled = False
         Combo1.Enabled = False
         modomodificar = True
    End If
Exit Sub
ERRAR:
      MsgBox "Error en Estructura,Verifique la Estructura de Base de Datos", vbCritical, "Error en Base de Datos"
Exit Sub
Resume
End Sub

'Grabar
Private Sub Command20_Click()

On Error GoTo Err
If Frame1.Visible Then
  If modomodificar Then
    If VGDllGeneral.VerificaDatoExistente(VGCNx, "select * from tabtransa Where TT_TIPMOV='" & DbGrid1.Columns(0).text & "' and TT_CODMOV='" & Text1 & "'") = 1 Then
        VGCNx.Execute "Update Tabtransa " & _
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
                          "TT_ORCON='" & Left(Combo2.text, 1) & "'," & _
                          "TT_CLIE='" & IIf(Check9.Value = 1, "S", "N") & "'," & _
                          "estadocosto=" & Check8.Value & "," & _
                          "TT_ORDFAB='" & IIf(Check11.Value = 1, "S", "N") & "'," & _
                          "TT_EQUIP='" & IIf(Check12.Value = 1, "S", "N") & "', " & _
                          "ingresosfuturos='" & IIf(Check13.Value = 1, "S", "N") & "', " & _
                          "intercompanias='" & IIf(Check14.Value = 1, "S", "N") & "', " & _
                          "TT_CODTRANS_AUTO='" & Ctr_Ayutransa.xclave & "'," & _
                          "tt_codtrans_merma='" & Ctr_Ayumerma.xclave & "' " & _
                          " Where TT_TIPMOV='" & DbGrid1.Columns(0).text & "' and TT_CODMOV='" & Text1 & "'"
  End If
 Else
   If Text1 <> "" Then
        If VGDllGeneral.VerificaDatoExistente(VGCNx, "select * from tabtransa Where TT_TIPMOV='" & Left(Combo1.text, 1) & "' and TT_CODMOV='" & Text1 & "'") = 0 Then
            VGCNx.Execute "INSERT INTO Tabtransa " & _
                            "(TT_CODMOV,TT_DESCRI,TT_TIPMOV,TT_PRV," & _
                            "TT_DR,TT_AT,TT_CC,TT_OC,TT_CO,TT_ALMA," & _
                            "TT_ORCON,TT_CLIE,estadocosto,TT_ORDFAB,TT_EQUIP,TT_CONT,TT_CODTRANS_AUTO" & _
                            ",tt_codtrans_merma,ingresosfuturos,intercompanias ) " & _
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
                            "'" & Left(Combo2.text, 1) & "'," & _
                            "'" & IIf(Check9.Value = 1, "S", "N") & "'," & _
                            Check8.Value & ",'" & _
                            "" & IIf(Check11.Value = 1, "S", "N") & "'," & _
                            "'" & IIf(Check12.Value = 1, "S", "N") & "'," & _
                            " 0 , " & _
                            "'" & Ctr_Ayutransa.xclave & "','" & Ctr_Ayumerma.xclave & "'," & _
                            "'" & IIf(Check13.Value = 1, "S", "N") & "'," & _
                            "'" & IIf(Check14.Value = 1, "S", "N") & "')"
                              
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
     Check13.Value = 0
     Check14.Value = 0
     Text1.Enabled = True
     Ctr_Ayutransa.xclave = ""
     Ctr_Ayutransa.xnombre = ""
     Ctr_Ayumerma.xclave = ""
     Ctr_Ayumerma.xnombre = ""
 End If
  Call Listado("select  TT_TIPMOV,TT_CODMOV,TT_DESCRI from tabtransa ORDER BY 1,2")
  modomodificar = False
  Frame1.Visible = False
  Frame2.Visible = False
Exit Sub
Err:
  MsgBox Err.Description
  Exit Sub
  Resume

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
     Check11.Value = 0
     Check12.Value = 0
     Check13.Value = 0
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
  cSelA.Open cSqlA, VGCNx, adOpenStatic
  If cSelA.RecordCount > 0 Then
       If MsgBox("No se puede eliminarla porque se registro transaciones", vbYesNo, "Eliminacion de Registro") = vbNo Then
              cSelA.Close: Exit Sub
       End If
  End If
  cSelA.Close
  If MsgBox("Seguro de Eliminar", vbQuestion + vbOKCancel, mensaje1) = vbOK Then

    csql = "DELETE from tabtransa where tt_codmov = '" & codmov & "' and tt_tipmov = '" & codtipo & "' "
    VGCNx.BeginTrans
    VGCNx.Execute csql
    VGCNx.CommitTrans
    Call Listado("select  TT_TIPMOV,TT_CODMOV,TT_DESCRI from tabtransa")
  End If
End Sub

Private Sub Command4_Click()
    Dim cadena As String
    Dim cNomRepor  As String

cNomRepor = "transac.RPT"
If Trim(cNomRepor) <> "" Then
    CrystalReport1.Reset
    CrystalReport1.WindowTitle = "Reporte de Transacciones"
    CrystalReport1.ReportFileName = VGParamSistem.RutaReport + "\" + cNomRepor

    CrystalReport1.Connect = VGcadenareport2
    CrystalReport1.StoredProcParam(0) = VGParamSistem.BDEmpresa
    
    CrystalReport1.formulas(0) = "emp ='" & VGparametros.RucEmpresa & "'"
    
    
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
  central FrmTransaccion
  'Data1.DatabaseName = cRuta2
  Call Ctr_Ayutransa.conexion(VGCNx): Ctr_Ayutransa.filtro = "tt_tipmov='I'"
  Call Ctr_Ayumerma.conexion(VGCNx): Ctr_Ayumerma.filtro = "tt_tipmov='S'"
  Combo1.text = "Ingreso"

  Call Listado("select  TT_TIPMOV,TT_CODMOV,TT_DESCRI from tabtransa order by 1,2")
  
End Sub

Sub Listado(wcad)
  Set DbGrid1.DataSource = Nothing
  Set rs = Nothing
  
  Set rs = VGCNx.Execute(wcad)
  Set DbGrid1.DataSource = rs
   With DbGrid1
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
            rs.Fields("estadocosto") = 1
       Else
             rs.Fields("estadocosto") = 0
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
       If Check13.Value = 1 Then
            rs.Fields("intercompanias") = "S"
       Else
             rs.Fields("intercompanias") = "N"
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
Ctr_Ayutransa.xclave = ""
End Sub
