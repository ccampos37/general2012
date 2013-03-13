VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmTranspor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualiza Datos Generales del Transportista"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   105
      TabIndex        =   5
      Top             =   4380
      Width           =   7365
      Begin VB.CommandButton cmdreporte 
         Caption         =   "&Reporte"
         Height          =   675
         Left            =   4695
         Picture         =   "FrmTranspor.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   240
         Width           =   825
      End
      Begin VB.CommandButton CmdEli 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   3600
         Picture         =   "FrmTranspor.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5880
         Picture         =   "FrmTranspor.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton CmdModi 
         Caption         =   "&Modificar"
         Height          =   675
         Left            =   2400
         Picture         =   "FrmTranspor.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   255
         Width           =   775
      End
      Begin VB.CommandButton CmdIng 
         Caption         =   "&Ingreso"
         Height          =   675
         Left            =   1200
         Picture         =   "FrmTranspor.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   775
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   0
         Top             =   150
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   10
      Top             =   4320
      Visible         =   0   'False
      Width           =   7335
      Begin VB.CommandButton CmdSalir2 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   675
         Left            =   4320
         Picture         =   "FrmTranspor.frx":154A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   2160
         Picture         =   "FrmTranspor.frx":198C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   775
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4155
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   7395
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   10
         Left            =   5220
         MaxLength       =   15
         TabIndex        =   37
         Text            =   "Text10"
         Top             =   1500
         Width           =   1665
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   9
         Left            =   2160
         MaxLength       =   25
         TabIndex        =   24
         Text            =   "Text10"
         Top             =   3510
         Width           =   2280
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   8
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   22
         Text            =   "Text9"
         Top             =   3135
         Width           =   4740
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   6
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   20
         Text            =   "Text6"
         Top             =   2385
         Width           =   4740
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   16
         Text            =   "Text6"
         Top             =   720
         Width           =   4755
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   19
         Text            =   "Text5"
         Top             =   1800
         Width           =   1860
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   2160
         MaxLength       =   25
         TabIndex        =   18
         Text            =   "Text4"
         Top             =   1440
         Width           =   1860
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   17
         Text            =   "Text3"
         Top             =   1080
         Width           =   4755
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   4965
         MaxLength       =   11
         TabIndex        =   15
         Text            =   "Text2"
         Top             =   390
         Width           =   1965
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   2160
         MaxLength       =   11
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   390
         Width           =   1875
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   7
         Left            =   2160
         MaxLength       =   11
         TabIndex        =   21
         Text            =   "Text7"
         Top             =   2775
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "Brevete"
         Height          =   225
         Left            =   4305
         TabIndex        =   36
         Top             =   1500
         Width           =   690
      End
      Begin VB.Label Label13 
         Caption         =   "Datos de la Empresa  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   34
         Top             =   2145
         Width           =   1935
      End
      Begin VB.Label Label11 
         Caption         =   "Dirección"
         Height          =   255
         Left            =   360
         TabIndex        =   33
         Top             =   3165
         Width           =   1515
      End
      Begin VB.Label Label9 
         Caption         =   "Telefono"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   32
         Top             =   1470
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Teléfono"
         Height          =   195
         Left            =   360
         TabIndex        =   31
         Top             =   3525
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Placa"
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   1830
         Width           =   1635
      End
      Begin VB.Label Label6 
         Caption         =   "Nombre "
         Height          =   195
         Left            =   360
         TabIndex        =   29
         Top             =   810
         Width           =   1515
      End
      Begin VB.Label Label5 
         Caption         =   "Razón Social"
         Height          =   195
         Left            =   360
         TabIndex        =   28
         Top             =   2445
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Dirección "
         Height          =   195
         Left            =   360
         TabIndex        =   27
         Top             =   1170
         Width           =   915
      End
      Begin VB.Label Label2 
         Caption         =   "RUC   "
         Height          =   255
         Left            =   4350
         TabIndex        =   26
         Top             =   420
         Width           =   750
      End
      Begin VB.Label Label1 
         Caption         =   "Código "
         Height          =   195
         Left            =   360
         TabIndex        =   25
         Top             =   480
         Width           =   915
      End
      Begin VB.Label Label4 
         Caption         =   "RUC"
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   2805
         Width           =   855
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FrmTranspor.frx":1DCE
      Height          =   3375
      Left            =   135
      TabIndex        =   35
      Top             =   855
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   5953
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   17
      TabAction       =   2
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "TRACODIGO"
         Caption         =   "              CODIGO"
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
         DataField       =   "TRANOMBRE"
         Caption         =   "                                  RAZON SOCIAL"
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
      BeginProperty Column02 
         DataField       =   "TRARUC"
         Caption         =   "            RUC"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "TRADIR"
         Caption         =   "                                   DIRECCION"
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
      BeginProperty Column04 
         DataField       =   "TRATELEF"
         Caption         =   "             TELEFONO"
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
      BeginProperty Column05 
         DataField       =   "TRARAZEMP"
         Caption         =   "                                  RAZON SOCIAL"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "TRATELEMP"
         Caption         =   "             TELEFONO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         ScrollBars      =   3
         BeginProperty Column00 
            Alignment       =   2
            Locked          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4680
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   4665.26
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            Locked          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   2099.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   4245.166
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2039.811
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   7320
      Begin VB.TextBox TxFiltro 
         Height          =   300
         Left            =   1200
         TabIndex        =   2
         Text            =   "TxFiltro"
         Top             =   240
         Width           =   2655
      End
      Begin VB.ComboBox CmbOrden 
         Height          =   315
         ItemData        =   "FrmTranspor.frx":1DE3
         Left            =   5280
         List            =   "FrmTranspor.frx":1DED
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label32 
         Caption         =   "Buscar   :"
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label33 
         Caption         =   "Orden :"
         Height          =   255
         Left            =   4320
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmTranspor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adodc1 As ADODB.Recordset
Dim cSel1 As ADODB.Recordset
Dim cSql1 As String, CSQL2 As String
Dim cSql3 As String, nT As Integer
Dim cCod As String, cDes As String
Dim nCom As Integer, nExiste As Integer
Dim nTra2 As Integer, nCursor As Integer
Dim nTra As Integer
Private Sub OculObj01(nTipo As Boolean)
Frame5.Visible = nTipo
Frame1.Visible = Not nTipo
Frame2.Visible = nTipo
Frame3.Visible = Not nTipo
DataGrid1.Visible = nTipo
End Sub
Private Sub CmbOrden_Click()            ' Ordenar por
nCom = CmbOrden.ListIndex
Set adodc1 = New ADODB.Recordset
Select Case nCom
Case 0
    adodc1.Open "Select TRACODIGO,TRANOMBRE,TRATELEF FROM al_transporte ORDER BY TRACODIGO", VGCNx, adOpenStatic
Case 1
    adodc1.Open "Select TRACODIGO,TRANOMBRE,TRADIR,TRATELEF FROM al_transporte ORDER BY TRANOMBRE", VGCNx, adOpenStatic
End Select
TxFiltro = ""
Set DataGrid1.DataSource = adodc1
If DataGrid1.Visible Then DataGrid1.SetFocus
End Sub

Private Sub CmdEli_Click()              ' Elimina
Dim nPosi As Integer
On Error GoTo EliErr

If adodc1.RecordCount > 0 Then
    If MsgBox("Desea Eliminar Datos ?", vbQuestion + vbOKCancel, "Mensaje") = vbOK Then
        If Existe(1, adodc1("TRACODIGO"), "MovAlmcab", "Cacodpro", False) Then
'            MsgBox "No se puede eliminar el Transportista, porque tiene documentos Anexados", vbInformation, "Información"
 '           Exit Sub
        Else
            cSql1 = "Delete from al_transporte where TRACODIGO= '" & adodc1("TRACODIGO") & "'"
            nPosi = Pos_Dato(adodc1)
            nTra = 1
            VGCNx.BeginTrans
            VGCNx.Execute cSql1
            VGCNx.CommitTrans
            nTra = 0
            adodc1.Requery
            If nPosi <> 0 Then adodc1.AbsolutePosition = nPosi
        End If
    End If
    If DataGrid1.Visible Then DataGrid1.SetFocus
Else
    MsgBox "No existe registros para Eliminar", vbInformation, "Mensaje"
    Exit Sub
End If
Exit Sub
EliErr:
    MsgBox Err.Description
    If nTra = 1 Then VGCNx.RollbackTrans
End Sub

Private Sub CmdGrabar_Click()           ' Grabar
Dim cMon As String
On Error GoTo GrabErr

If nT = 1 Then
  If Text1(0) <> "" Then
    If Existe(1, Text1(0), "al_transporte", "TRACODIGO", False) Then
        MsgBox "El código de Transportista, ya existe", vbInformation, "Mensaje"
        Text1(0).SetFocus: Exit Sub
    End If
  Else
        MsgBox "Ingrese código de Transportista", vbInformation, "Mensaje"
        Text1(0).SetFocus: Exit Sub
  End If
End If
    
If Trim(Text1(2)) = "" Then
    MsgBox "Ingrese Nombre de Transportista", vbInformation, "Mensaje"
    Text1(2).SetFocus: Exit Sub
End If
If Trim(Text1(3)) = "" Then
    MsgBox "Ingrese Dirección de Transportista", vbInformation, "Mensaje"
    Text1(3).SetFocus: Exit Sub
End If
If Text1(1) <> "" Then
   If Validar_RUC(Text1(1)) = False Then
      Text1(1).SetFocus: Exit Sub
   End If
ElseIf Text1(7) <> "" Then
   If Validar_RUC(Text1(7)) = False Then
      Text1(7).SetFocus: Exit Sub
   End If
End If

If MsgBox("Es correcta la Información ?", vbQuestion + vbOKCancel, "Mensaje") = vbOK Then
    
    If nT = 1 Then      'Ingreso
        CSQL2 = "Insert Into al_transporte (TRACODIGO,TRANOMBRE,TRATELEF,TRAPLACA,"
        CSQL2 = CSQL2 & "TRABREVE) VALUES "
        CSQL2 = CSQL2 & "('" & Text1(0) & "','" & Text1(2) & "','" & (Text1(4)) & "',"
        CSQL2 = CSQL2 & "'" & Text1(5) & "','" & Text1(10) & "')"
        cCod = Text1(0)
        
    ElseIf nT = 2 Then     'Modificar
        CSQL2 = "Update al_transporte Set TRACODIGO='" & Text1(0) & "',TRARUC='" & Text1(1) & "',"
        CSQL2 = CSQL2 & "TRANOMBRE='" & SupCadSQL(Text1(2)) & "',TRADIR='" & SupCadSQL(Text1(3)) & "',"
        CSQL2 = CSQL2 & "TRATELEF='" & Text1(4) & "',TRAPLACA='" & Text1(5) & "',"
        CSQL2 = CSQL2 & "TRARAZEMP='" & Text1(6) & "',TRARUCEMP='" & Text1(7) & "',"
        CSQL2 = CSQL2 & "TRADIREMP='" & Text1(8) & "',TRATELEMP='" & SupCadSQL(Text1(9)) & "',"
        CSQL2 = CSQL2 & "TRABREVE='" & Text1(10) & "'"
        CSQL2 = CSQL2 & "Where TRACODIGO= '" & Trim(Text1(0)) & "'"
        cCod = Text1(0)
    End If
    
    nTra = 1
    VGCNx.BeginTrans
    VGCNx.Execute CSQL2
    VGCNx.CommitTrans
    nTra = 0
    adodc1.Requery
    
    adodc1.Find "TRACODIGO = '" & cCod & "'"
End If


If nT = 1 Then
    Limpiar
    Text1(0).SetFocus
Else
    CmdSalir2_Click
End If
Exit Sub
GrabErr:
    MsgBox Err.Description
    If nTra = 1 Then VGCNx.RollbackTrans
    If nTra2 = 1 Then VGCNx.RollbackTrans
End Sub

Private Sub CmdIng_Click()      'Ingresar
nT = 1
Frame3.Caption = "Ingreso de Datos del Transportista"
OculObj01 (False)
Limpiar
Text1(0).Enabled = True
Text1(0).SetFocus
End Sub

Private Sub CmdModi_Click()     'Modificar
If adodc1.RecordCount > 0 Then
    nT = 2
    Frame3.Caption = "Modificación de Datos de Transportista"
    OculObj01 (False)
    Limpiar
    cCod = adodc1("TRACODIGO")
    Mostrar (cCod)
    Text1(0).Enabled = False
    If Text1(1).Visible Then Text1(1).SetFocus
Else
    MsgBox "No existen registros", vbInformation, "Mensaje"
End If
End Sub

Private Sub cmdreporte_Click()
    Dim cadena As String
    Dim cNomRepor  As String

cNomRepor = "transportista.RPT"
If Trim(cNomRepor) <> "" Then
    CrystalReport1.Reset
    CrystalReport1.WindowTitle = "Reporte de Transportistas"
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

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub CmdSalir2_Click()   'Salida de la segunda pantalla
If Not VGtransp Then
   Unload Me
Else
  OculObj01 (True)
  DataGrid1.SetFocus
End If
End Sub

Private Sub Command4_Click()

End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyBack Then
 If Len(TxFiltro) - 1 > 0 Then
  TxFiltro = Left(TxFiltro, Len(TxFiltro) - 1)
 Else
  TxFiltro = ""
 End If
 KeyAscii = 0
ElseIf KeyAscii <> 13 Then
 TxFiltro = TxFiltro & Chr(KeyAscii)
End If
End Sub

Private Sub Form_Activate()
TxFiltro = ""
CmbOrden.ListIndex = 0
If DataGrid1.Visible Then DataGrid1.SetFocus
If Not VGtransp Then
   CmdIng_Click
    Text1(0) = FrmGuiaSal.TxtTransp.text
End If
End Sub
Private Sub Form_Load()
Dim RUTA As String
Dim NAMEBD As String
central Me          ' Centrar Formulario
Init_ControlDataGrid DataGrid1
Limpiar
OculObj01 (True)
Set adodc1 = New ADODB.Recordset
adodc1.Open "Select TRACODIGO,TRANOMBRE,TRATELEF FROM al_transporte ORDER BY TRACODIGO", VGCNx, adOpenStatic
Set DataGrid1.DataSource = adodc1
DataGrid1.Refresh
End Sub
Private Sub Limpiar()   'Limpia variables
Dim n As Integer
For n = 0 To 10: Text1(n) = "": Next
End Sub

Private Sub Text1_GotFocus(Index As Integer)
Enfoque Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If Index = 0 Then
        If Trim(Text1(0)) <> "" Then
            If Existe(1, Text1(0), "al_transporte", "TRACODIGO", False) Then
                MsgBox "El código de Transportista, ya existe", vbInformation, "Mensaje"
                Text1(0).SetFocus: Exit Sub
            End If
            Text1(1).SetFocus: Exit Sub
        Else
            MsgBox "Ingrese código del Transportista", vbInformation, "Mensaje"
            Text1(0).SetFocus: Exit Sub
        End If
    ElseIf Index = 1 Then
             Text1(1) = Trim(Text1(1))
             If Text1(1) <> "" Then
                   If Validar_RUC(Text1(1)) = False Then
                        Text1(1).SetFocus
                        Exit Sub
                  End If
            End If
            SendKeys "{tab}"
      
    ElseIf Index = 2 Then
        If Trim(Text1(2)) <> "" Then
            Text1(3).SetFocus: Exit Sub
        Else
            MsgBox "Ingrese Nombre de Transportista", vbInformation, "Mensaje"
            Text1(2).SetFocus: Exit Sub
        End If
    ElseIf Index = 3 Then
        If Trim(Text1(3)) <> "" Then
            Text1(4).SetFocus: Exit Sub
        Else
            MsgBox "Ingrese Dirección de Transportista", vbInformation, "Mensaje"
            Text1(3).SetFocus: Exit Sub
        End If
    ElseIf Index = 7 Then
          If Text1(7) <> "" Then
            If Validar_RUC(Text1(7)) = False Then
                Text1(7).SetFocus
                Exit Sub
            End If
          End If
         SendKeys "{tab}"
    
    Else
        If Index <> 9 Then
           If Index = 1 Then
              If Text1(1) <> "" Then
                 If Validar_RUC(Text1(1)) = False Then
                    Text1(1).SetFocus: Exit Sub
                 Else
                    Text1(2).SetFocus
                 End If
              End If
           End If
           SendKeys "{tab}"
        Else
           Cmdgrabar.SetFocus
        End If
    End If
    
Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
    If Index = 1 Then
    If (Chr$(KeyAscii) > "9" Or Chr$(KeyAscii) < "0") And KeyAscii <> 8 Then KeyAscii = 0
    End If
End If
End Sub

Private Sub TxFiltro_Change()
If adodc1.RecordCount > 0 Then
    If Trim(TxFiltro) <> "" And TxFiltro.Visible Then
        nCursor = adodc1.Bookmark
        adodc1.AbsolutePosition = 1
        adodc1.MoveFirst
        
        Select Case CmbOrden.ListIndex
        Case 0
            adodc1.Find "TRACODIGO LIKE '" & Trim(UCase(TxFiltro)) & "*'"
        Case 1
            adodc1.Find "TRANOMBRE LIKE '" & Trim(UCase(TxFiltro)) & "*' "
        End Select
        If adodc1.EOF Then adodc1.AbsolutePosition = nCursor
    End If
End If
End Sub

Private Sub Mostrar(cC1 As String) 'Muestra los datos
Dim cSqlM As String, cSelM As ADODB.Recordset
If Trim(cC1) = "" Then
    MsgBox "No hay registros para mostrar", vbInformation, "Mensaje"
    Exit Sub
End If

cSqlM = "Select * From al_transporte Where TRACODIGO= '" & cC1 & "'"
Set cSelM = New ADODB.Recordset
cSelM.Open cSqlM, VGCNx, adOpenStatic
If cSelM.RecordCount > 0 Then
    Text1(0) = cSelM("TRACODIGO")
    'If Not IsNull(cSelM("TRARUC")) Then Text1(1) = cSelM("TRARUC")
    If Not IsNull(cSelM("TRANOMBRE")) Then Text1(2) = cSelM("TRANOMBRE")
'    If Not IsNull(cSelM("TRADIR")) Then Text1(3) = cSelM("TRADIR")
    If Not IsNull(cSelM("TRATELEF")) Then Text1(4) = cSelM("TRATELEF")
    If Not IsNull(cSelM("TRAPLACA")) Then Text1(5) = cSelM("TRAPLACA")
'    If Not IsNull(cSelM("TRARAZEMP")) Then Text1(6) = cSelM("TRARAZEMP")
'    If Not IsNull(cSelM("TRARUCEMP")) Then Text1(7) = cSelM("TRARUCEMP")
'    If Not IsNull(cSelM("TRADIREMP")) Then Text1(8) = cSelM("TRADIREMP")
'    If Not IsNull(cSelM("TRATELEMP")) Then Text1(9) = cSelM("TRATELEMP")
    If Not IsNull(cSelM("TRABREVE")) Then Text1(10) = cSelM("TRABREVE")

Else

    MsgBox "No existe registro", vbInformation, "Mensaje"
    CmdSalir2_Click
End If
cSelM.Close
End Sub

