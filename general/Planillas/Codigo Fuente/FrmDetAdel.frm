VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmDetAdel 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6450
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8910
   Icon            =   "FrmDetAdel.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   8910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   7230
      TabIndex        =   22
      Top             =   4875
      Width           =   1665
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Grabar "
      Height          =   345
      Left            =   7230
      TabIndex        =   21
      Top             =   4485
      Width           =   1665
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   945
      Left            =   45
      TabIndex        =   13
      Top             =   105
      Width           =   4485
      Begin VB.TextBox xCod 
         Height          =   285
         Left            =   870
         TabIndex        =   14
         Top             =   120
         Width           =   1380
      End
      Begin MSDataListLib.DataCombo DCmcampo 
         Height          =   315
         Left            =   870
         TabIndex        =   15
         Top             =   450
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   "DataCombo1"
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Código "
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   60
         TabIndex        =   17
         Top             =   180
         Width           =   675
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Columna "
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   60
         TabIndex        =   16
         Top             =   510
         Width           =   840
      End
   End
   Begin VB.CommandButton CmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   315
      Left            =   4635
      TabIndex        =   8
      Top             =   705
      Width           =   1200
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "&Agregar"
      Height          =   315
      Left            =   4635
      TabIndex        =   7
      Top             =   360
      Width           =   1200
   End
   Begin VB.CommandButton CmdImpPla 
      Caption         =   "&Imprimir Adelanto"
      Height          =   345
      Left            =   7245
      TabIndex        =   6
      Top             =   4095
      Width           =   1665
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Totales"
      Height          =   345
      Left            =   5550
      TabIndex        =   5
      Top             =   4095
      Width           =   1650
   End
   Begin Crystal.CrystalReport Reporte 
      Left            =   6555
      Top             =   8220
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton CmdCfgVista 
      Caption         =   "&Config de Vista"
      Height          =   345
      Left            =   5535
      TabIndex        =   4
      Top             =   4500
      Width           =   1665
   End
   Begin MSComctlLib.ListView LvColumn 
      Height          =   1140
      Left            =   135
      TabIndex        =   0
      Top             =   5265
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   2011
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Descrip"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Codigo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ColumnaRef"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Height          =   2970
      Left            =   45
      ScaleHeight     =   2910
      ScaleWidth      =   8805
      TabIndex        =   9
      Top             =   1065
      Width           =   8865
      Begin VB.TextBox XFormula 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   750
         TabIndex        =   18
         Top             =   30
         Width           =   7050
      End
      Begin VB.CommandButton Command1 
         Height          =   285
         Left            =   7890
         Picture         =   "FrmDetAdel.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Ejecutar Fórmula"
         Top             =   45
         Width           =   810
      End
      Begin MSDataGridLib.DataGrid DgDet 
         Height          =   2460
         Left            =   75
         TabIndex        =   10
         Top             =   375
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   4339
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
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
         Caption         =   "Detalles del Consolidado "
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
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
               LCID            =   2058
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
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Fórmula "
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   11
         Top             =   90
         Width           =   585
      End
   End
   Begin VB.TextBox xVarFormu 
      Height          =   285
      Left            =   3165
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3435
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.TextBox SqlSel 
      Height          =   315
      Left            =   3165
      TabIndex        =   2
      Top             =   3120
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.TextBox SqlCad 
      Height          =   330
      Left            =   3150
      TabIndex        =   1
      Top             =   2790
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   8205
      Picture         =   "FrmDetAdel.frx":0C0C
      Top             =   495
      Width           =   480
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Detalle de Adelantos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   360
      Left            =   6060
      TabIndex        =   20
      Top             =   90
      Width           =   2595
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Detalle de Adelantos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   6105
      TabIndex        =   19
      Top             =   90
      Width           =   2595
   End
End
Attribute VB_Name = "FrmDetAdel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim POSCOL As Integer
Dim RSDET As New ADODB.Recordset
Dim RsDetClone As New ADODB.Recordset
Dim FLAGCOLUM As Boolean
Dim INTO As String
Public FMES As Date
Public CodigoTrab As String
Dim SQLSTR As String
Private Sub CMDAGREGAR_CLICK()
Dim XLIST As ListItem
Dim SQLBASE As String
Dim RSDETAUX As New ADODB.Recordset
    If VALIDCAMP(DCmcampo.BoundText) Then Exit Sub
    CambiaPanelBD True
    Screen.MousePointer = 11
    
    If Not ExisteCampo("MONTO", " [##TMPDETCOLUM" & VGL_COMPUTER & "] ", DBSYSTEM) Then DBSYSTEM.Execute "ALTER TABLE  [##TMPDETCOLUM" & VGL_COMPUTER & "]  ADD MONTO  Numeric(20,2) "
    
    DBSYSTEM.Execute "ALTER TABLE  [##TMPDETCOLUM" & VGL_COMPUTER & "]  ADD " & Trim(DCmcampo.BoundText) & "  Numeric(20,2) "
     RSDETAUX.Open " [##TMPDETCOLUM" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic
    If RSDETAUX.RecordCount = 0 Then Exit Sub
    RSDETAUX.MoveFirst
    Do While Not RSDETAUX.EOF
        RSDETAUX.Fields(Trim(DCmcampo.BoundText)) = DEVOLVERIMPORTE(RSDETAUX!CODTRAB, Trim(DCmcampo.BoundText))
        RSDETAUX.Update
        RSDETAUX.MoveNext
    Loop
    
    Set XLIST = LvColumn.ListItems.Add(, , Trim(DCmcampo.BoundText))
        XLIST.SubItems(1) = DCmcampo.BoundText
        XLIST.SubItems(2) = DCmcampo.Text
    SqlSel.Text = ""
    LvColumn.Refresh
    'ACTUALIZA EL MONTO TOTAL
    DBSYSTEM.Execute "Update  [##TMPDETCOLUM" & VGL_COMPUTER & "]  Set MONTO=MONTO+" & Trim(DCmcampo.BoundText)
    Set RSDET = Nothing
    RSDET.Open SQLSTR, DBSYSTEM, adOpenKeyset, adLockReadOnly
    Set DgDet.DataSource = RSDET
    DgDet.Refresh
    FORMHEAD
    DBSYSTEM.Execute "Update  [##TMPDETCOLUM" & VGL_COMPUTER & "]  Set MONTO=" & ""
    frCfgDet2.Command2_Click
    frCfgDet2.CARGAR
    Screen.MousePointer = 1
    CambiaPanelBD False
End Sub
Private Function DEVOLVERIMPORTE(CODTRAB As String, CAMPO As String) As Double
    Dim RSAUX As New ADODB.Recordset
    Dim SQL As String
    DEVOLVERIMPORTE = 0
    SQL = "SELECT " & xCod.Text & " FROM [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  WHERE CODTRAB='" & Trim(CODTRAB) & " ' "
    RSAUX.Open SQL, DBSYSTEM, adOpenKeyset, adLockReadOnly
    If RSAUX.RecordCount = 0 Then Exit Function

    DEVOLVERIMPORTE = RSAUX.Fields(0)
End Function

Private Sub CMDCFGVISTA_Click()
    frCfgDet2.Show 1
End Sub

Private Sub CMDIMPPLA_Click()
    Dim REG As Long
    CambiaPanelBD True
    Screen.MousePointer = 11
    Call CREARQUERYTABLA
    DBSYSTEM.Execute "UPDATE ##TMPTRAB SET CODTRAB=CODTRAB", REG
'    With Reporte
'        .Reset
'        .WindowTitle = "PLAN0068.RPT - RESUMEN"
'        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0068.RPT"
'        .Connect = "DSN=" & VGL_SERVER & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL & ""
'        .Destination = crptToWindow
'        .WindowState = crptMaximized
'        .WindowShowPrintBtn = True
'        .WindowShowRefreshBtn = True
'        .WindowShowSearchBtn = True
'        .WindowShowPrintSetupBtn = True
'        .Formulas(0) = "XCABEZA='" & Trim(xTitulo(0).Text) & "'"
'        .Formulas(1) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
'        .Formulas(2) = "XRUC='" & REGSISTEMA.RUC & "'"
'        .Formulas(3) = "XREG='" & Str(REG) & "'"
'        If .Status <> 2 Then .Action = 1
'    End With
    Screen.MousePointer = 1
    CambiaPanelBD False
End Sub

Private Sub CREARQUERYTABLA()
  Dim INSQLCONC As String
  Dim SQL As String
  Dim XITEM As ListItem
  Dim ORDEN As Integer
  
  If ExisteTablaAux(" [##TMPLIST" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPLIST" & VGL_COMPUTER & "] "
  DBSYSTEM.Execute "CREATE TABLE  [##TMPLIST" & VGL_COMPUTER & "]  (CODIGO VARCHAR(6),ORDEN INT)"
  DBSYSTEM.Execute "CREATE Index CODIGO ON  [##TMPLIST" & VGL_COMPUTER & "]  (CODIGO)"
  
  For Each XITEM In frCfgDet2.LColumnas.ListItems
    If XITEM.INDEX > 2 Then
        If XITEM.Checked Then
            ORDEN = ORDEN + 1
            DBSYSTEM.Execute "INSERT INTO  [##TMPLIST" & VGL_COMPUTER & "]  VALUES('" & Trim(XITEM.Text) & "'," & Str(ORDEN) & ")"
        End If
    End If
  Next
  
 DBSTARPLAN.Execute "EXECUTE QUERY_DETALLE"
 
End Sub
Private Sub CmdImprimir_Click()
    frCfgDet2.CARGAR
    If ExisteTablaAux(" [##PRTPLANILLA" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##PRTPLANILLA" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "CREATE TABLE  [##PRTPLANILLA" & VGL_COMPUTER & "]  (FILA VARCHAR(8000))"
    Dim X As Integer, CAD As String, XSTR As String, XNUMCHAR As Integer, xValor As Variant, XN As Byte, NCOUNT As Integer
    Dim RSAUX As ADODB.Recordset, CADTABLA As String
    Dim RSTOT As New ADODB.Recordset
    Set RSAUX = New ADODB.Recordset
    Screen.MousePointer = 11
    

    'PROCESO DEL ENCABEZADO DE LAS PLANILLAS
    '---------------------------------------
    RSAUX.Open " [##TMPDETCOLUM" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic
    RSAUX.MoveFirst
    For X = 0 To DgDet.Columns.Count - 1
        If DgDet.Columns(X).Visible Then
            CAD = RSAUX.Fields(Trim(frCfgDet2.LColumnas.ListItems(X + 1))).Type
            Select Case RSAUX.Fields(Trim(frCfgDet2.LColumnas.ListItems(X + 1))).Type
                    Case adSingle, 3, 5, 6 'SI ES NUMERO SIMPLE
                        XNUMCHAR = 10
                    Case adVarChar
                        XNUMCHAR = RSAUX.Fields(Trim(frCfgDet2.LColumnas.ListItems(X + 1))).DefinedSize + 1
                    Case adUnsignedTinyInt
                        XNUMCHAR = 3
            End Select
            CAD = frCfgDet2.LColumnas.ListItems(X + 1).SubItems(1)
            XSTR = XSTR & Left(CAD & String(XNUMCHAR, " "), XNUMCHAR)
        End If
    Next
    CAD = XSTR
    NCOUNT = 0
    
    'GENERACION DE LA CADENA DE IMPRESION
    '------------------------------------
    RSAUX.MoveFirst
    Do While Not RSAUX.EOF
        XSTR = ""
        For X = 0 To DgDet.Columns.Count - 1
            xValor = RSAUX.Fields(Trim(frCfgDet2.LColumnas.ListItems(X + 1))).Value
            If DgDet.Columns(X).Visible Then
                Select Case RSAUX.Fields(Trim(frCfgDet2.LColumnas.ListItems(X + 1))).Type
                    Case adSingle, 3, 5, 6 'SI ES NUMERO SIMPLE
                        XSTR = XSTR & " " & Right("         " & Format$(xValor, "0.00"), 9)
                    Case adVarChar
                        XN = RSAUX.Fields(Trim(frCfgDet2.LColumnas.ListItems(X + 1))).DefinedSize
                        XSTR = XSTR & " " & Left(xValor & String(XN, " "), XN)
                    Case adUnsignedTinyInt
                        XSTR = XSTR & " " & Format(xValor, "00")
                    Case Else
                        'MSGBOX "TIPO NO ENCONTRADO: " & RSAUX.FIELDS(TRIM(frCfgDet2.LCOLUMNAS.LISTITEMS(X + 1))).TYPE, VBCRITICAL
                End Select
            End If
        Next
        NCOUNT = NCOUNT + 1
        If NCOUNT = 1 Then
            DBSTARPLAN.Execute "INSERT INTO  [##PRTPLANILLA" & VGL_COMPUTER & "]  VALUES ('" & String(Len(XSTR), "_") & "')"
            'INCLUIR AQUI LA CABEZAR
            DBSTARPLAN.Execute "INSERT INTO  [##PRTPLANILLA" & VGL_COMPUTER & "]  VALUES ('" & CAD & "')"
            DBSTARPLAN.Execute "INSERT INTO  [##PRTPLANILLA" & VGL_COMPUTER & "]  VALUES ('" & String(Len(XSTR), "_") & "')"
        End If
        DBSYSTEM.Execute "INSERT INTO  [##PRTPLANILLA" & VGL_COMPUTER & "]  VALUES ('" & XSTR & "')"
        RSAUX.MoveNext
    Loop
    CAD = String(Len(XSTR), "_")
    DBSTARPLAN.Execute "INSERT INTO  [##PRTPLANILLA" & VGL_COMPUTER & "]  VALUES ('" & CAD & "')"
    
    
    'PROCESO DEL CÁLCULO DE TOTALES DE LA PLANILLA
    '---------------------------------------------
    RSAUX.MoveFirst
    XSTR = ""
    For X = 0 To DgDet.Columns.Count - 1
        If DgDet.Columns(X).Visible Then
            Select Case RSAUX.Fields(Trim(frCfgDet2.LColumnas.ListItems(X + 1))).Type
                    Case adDate 'SI ES POR FECHA
                        XNUMCHAR = 9
                    Case adSingle, 3, 5, 6  'SI ES NUMERO SIMPLE
                        XNUMCHAR = 10
                    Case adVarChar
                        XNUMCHAR = RSAUX.Fields(Trim(frCfgDet2.LColumnas.ListItems(X + 1))).DefinedSize + 1
                    Case adUnsignedTinyInt
                        XNUMCHAR = 3
            End Select
            If RSAUX.Fields(Trim(frCfgDet2.LColumnas.ListItems(X + 1))).Type = 5 Then
                RSTOT.Open "SELECT SUM(" & Trim(frCfgDet2.LColumnas.ListItems(X + 1)) & ") AS TOTAL FROM  [##TMPDETCOLUM" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenStatic
                If RSTOT.RecordCount = 0 Or IsNull(RSTOT!TOTAL) Then xValor = 0 Else xValor = RSTOT!TOTAL
                CAD = Right("         " & Format$(xValor, "0.00"), XNUMCHAR)
                RSTOT.Close
            Else
                CAD = String(XNUMCHAR, " ")
            End If
            XSTR = XSTR & CAD
        End If
    Next
    XSTR = "TOTAL GENERAL: " & Right(XSTR, Len(XSTR) - Len("TOTAL GENERAL: "))
    Set RSAUX = Nothing
    DBSYSTEM.Execute "INSERT INTO  [##PRTPLANILLA" & VGL_COMPUTER & "]  VALUES ('" & XSTR & "')"
    CAD = String(Len(XSTR), "_")
    DBSYSTEM.Execute "INSERT INTO  [##PRTPLANILLA" & VGL_COMPUTER & "]  VALUES ('" & CAD & "')"
    Screen.MousePointer = 1
'    With Reporte
'        .WindowTitle = "PLAN0023 - " & DgDet.Caption
'        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0023.RPT"
'        .Connect = "DSN=" & VGL_SERVER & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL & ""
'        '.DataFiles(0) = App.PATH & "\BDAUXCOM.MDB"
'        .StoredProcParam(0) = " [##PRTPLANILLA"  & VGL_COMPUTER & "] "
'        .Destination = crptToWindow
'        .WindowShowPrintBtn = True
'        .WindowState = crptMaximized
'        .WindowShowPrintBtn = True
'        .WindowShowRefreshBtn = True
'        .WindowShowSearchBtn = True
'        .WindowShowPrintSetupBtn = True
'        If xTitulo(0).Text = "" Then
'            .Formulas(0) = "XTITULO=' " & DgDet.Caption & "'"
'        Else
'            .Formulas(0) = "XTITULO=' " & xTitulo(0).Text & "'"
'        End If
'        .Formulas(1) = "XCABEZA1=' " & xTitulo(1).Text & "'"
'        .Formulas(2) = "XCABEZA2=' " & xTitulo(2).Text & "'"
'        .Formulas(3) = "XCABEZA3=' " & xTitulo(3).Text & "'"
'        If .Status <> 2 Then .Action = 1
'    End With
End Sub

Private Sub Command1_Click()
Dim CAD As String
Dim XLIST As ListItem
Dim SQLBASE As String, REMP As String
Dim I As Integer, X As Integer, j As Integer
 On Error GoTo ERRFORMULA
    xVarFormu.Text = ""
    CAD = Getcad(":", 1, xFormula.Text)
    If InStr(Trim(CAD), ",") > 0 Then
        MsgBox "No se admiten caracteres especiales para los nombres del campo", vbInformation
        Exit Sub
    End If
    
    If CAD <> "" Then
        xVarFormu.Text = "(" & Right(xFormula.Text, Len(xFormula) - (Len(CAD) + 1)) & ")  "
     Else: Exit Sub
    End If
    INTO = " INTO [" & App.PATH & "\BDAUXCOM.MDB].TMPDETCOLUM"
    If LvColumn.ListItems.Count = 0 Then
        MsgBox "Por lo menos debe existir una columna numerica", vbExclamation
        xFormula.SetFocus
        Exit Sub
    End If
    If VALIDCAMP(Trim(CAD)) Then Exit Sub
    REMP = ""
    For I = 1 To LvColumn.ListItems.Count
        If InStr(1, xFormula.Text, Trim(LvColumn.ListItems.Item(I))) > 0 Then X = X + 1
        If X > 0 Then Exit For
    Next
    If X = 0 Then
        MsgBox "Algunos de los codigo de campo utilizado en la fórmula no existen", vbExclamation
        xFormula.SelStart = X: xFormula.SelLength = 1
        xFormula.SetFocus
        Exit Sub
    End If
    Screen.MousePointer = 11
    CambiaPanelBD True
    DBSTARPLAN.Execute "ALTER TABLE  [##TMPDETCOLUM" & VGL_COMPUTER & "]  ADD " & Trim(CAD) & "  Numeric(20,2) "
    
    DBSTARPLAN.Execute "UPDATE  [##TMPDETCOLUM" & VGL_COMPUTER & "]  SET " & Trim(CAD) & "=" & Trim(xVarFormu.Text)
    
    Set XLIST = LvColumn.ListItems.Add(, , Trim(CAD))
    XLIST.SubItems(1) = Trim(xVarFormu.Text)
    XLIST.SubItems(2) = Trim(CAD)
    XLIST.SubItems(3) = "X"
    XLIST.SubItems(4) = xCod.Text
    LvColumn.Refresh
    
    'ACTUALIZA EL MONTO TOTAL
    DBSYSTEM.Execute "Update  [##TMPDETCOLUM" & VGL_COMPUTER & "]  Set MONTO=MONTO+" & Trim(DCmcampo.BoundText)

    Set RSDET = New ADODB.Recordset
    RSDET.Open SQLSTR, DBSTARPLAN, adOpenKeyset, adLockOptimistic
    Set DgDet.DataSource = RSDET
    DgDet.Refresh
    FORMHEAD ' FORMATEA LAS COLUMNAS DEL GRID
    
    frCfgDet2.Command2_Click
    frCfgDet2.CARGAR
    Screen.MousePointer = 1
    CambiaPanelBD False
        
    Exit Sub
    
ERRFORMULA:
    Select Case ERR.Number
        Case -2147217887: MsgBox "EL CODIGO DEL CAMPO : """ & Trim(CAD) & """ YA ESTA SIENDO UTILIZADO", vbInformation
        Case 5:
            MsgBox "NO SE ENCUENTRA NINGÚN CAMPO NUMERICO PARA ELABORAR FÓRMULA", vbInformation
        Case -2147217900
            MsgBox "Fórmula incorrecta; revise su sintaxis", vbInformation
            DBSTARPLAN.Execute "ALTER TABLE  [##TMPDETCOLUM" & VGL_COMPUTER & "]  DROP COLUMN " & Trim(CAD)
        Case Else
            MsgBox ERR.Description, vbInformation
            
    End Select
    
    
    Screen.MousePointer = 1
    Resume Next
    xFormula.SetFocus
End Sub
Public Function REPLCADENA(FIND As String, PALABRA As String, FRASE As String, _
                           CARACTERES As String, Optional ByRef POSCAD As Integer) As String
Dim I As Integer, POS As Integer, IND As Integer, j As Integer, FIN As Integer
Dim Aux As String
    FIND = UCase(FIND)
    FRASE = UCase(FRASE)
    IND = 1
    Aux = ""
    POSCAD = 0
    For I = 1 To Len(FRASE)
        POS = InStr(IND, FRASE, FIND)
        If POS = 0 Then Exit For
        If POS = 1 Then POS = 2
        FIN = Len(FIND) + 1
        If (POS - 1) + (FIN - 1) = Len(FRASE) Then
            FIN = 0
        End If
        
        If InStr(CARACTERES, Mid(FRASE, POS - 1, 1)) <> 0 And _
        InStr(CARACTERES, Mid(FRASE, (POS - 1) + FIN, 1)) <> 0 Then
            POSCAD = POS
            For j = 1 To Len(FRASE)
                If Not ((j >= POS) And j <= (POS + Len(FIND) - 1)) Then
                   Aux = Aux + Mid(FRASE, j, 1)
                   Else:
                        If j = POS Then Aux = Aux + "º"
                End If
            Next
            Aux = Replace(Aux, "º", PALABRA)
            FRASE = Aux
            Aux = ""
        End If
            
        IND = ((POS - 1) + I) + Len(PALABRA)
    Next
    REPLCADENA = FRASE

End Function
Public Sub FORMHEAD()
Dim I As Integer
    DgDet.Columns(0).Width = 1200
    DgDet.Columns(0).Caption = "CÓDIGO"
    DgDet.Columns(1).Caption = "APELLIDOS Y NOMBRES"
    DgDet.Columns(1).Width = 3200
    For I = 1 To LvColumn.ListItems.Count
        DgDet.Columns(I + 1).Width = 1200
        DgDet.Columns(I + 1).Caption = LvColumn.ListItems.Item(I).SubItems(2)
        DgDet.Columns(I + 1).Alignment = dbgRight
        DgDet.Columns(I + 1).NumberFormat = "###,###,##0.00 "
        DgDet.Columns(I + 1).WrapText = True
    Next
End Sub
Private Function VALIDCAMP(CAMPO As String) As Boolean
Dim I As Integer
    VALIDCAMP = False
    If LvColumn.ListItems.Count = 0 Then Exit Function
    For I = 1 To LvColumn.ListItems.Count
      If Trim(CAMPO) = Trim(LvColumn.ListItems.Item(I)) Then
         MsgBox "La columna de la fórmula :""" & CAMPO & """  ya esta ingresada ", vbExclamation
         xFormula.SetFocus
         VALIDCAMP = True
         Exit Function
      End If
    Next
End Function

Private Function MOSTRARCOLUM(Codigo As String)
Dim CADENA As String
    CADENA = " ( SELECT   T.SUMADEMONTO FROM TMPDETAGROUP T   " & _
             " WHERE T.CODTRAB=TMPDETAGROUP.CODTRAB AND " & _
             " T.CODCONCEP='" & Trim(Codigo) & "' ) AS " & Codigo
    MOSTRARCOLUM = CADENA
End Function

Private Sub DEVCADSQL()
Dim I As Integer
    If LvColumn.ListItems.Count = 0 Then Exit Sub
    SqlSel.Text = ""
    For I = 1 To LvColumn.ListItems.Count
      If I = LvColumn.ListItems.Count Then
        SqlSel.Text = SqlSel.Text & LvColumn.ListItems.Item(I).SubItems(1)
      Else: SqlSel.Text = SqlSel.Text & LvColumn.ListItems.Item(I).SubItems(1) & ","
      End If
    Next
End Sub

Private Sub CMDELIMINAR_CLICK()
Dim RSDET As New ADODB.Recordset
Dim SQLBASE As String
Dim X As Integer, I As Integer, j As Integer
    
    'VALIDANDO QUE NO ELIMINE UN CAMPO QUE ESTE RELACIONADO CON UNA FORMULA
    If POSCOL < 2 Then Exit Sub
    For I = 1 To LvColumn.ListItems.Count
        If LvColumn.ListItems.Item(I).SubItems(3) = "X" And _
           LvColumn.ListItems.Item(POSCOL - 1) <> _
           LvColumn.ListItems.Item(I) Then
                REPLCADENA Trim(LvColumn.ListItems.Item(POSCOL - 1)), " ", _
                LvColumn.ListItems.Item(I).SubItems(1), "/*-+\() ", X
                If X <> 0 Then
                    MsgBox "EL CAMPO " & LvColumn.ListItems.Item(POSCOL - 1).SubItems(2) & " ESTA RELACIONADO CON UN CAMPO FORMULA"
                    Screen.MousePointer = 1
                    Exit Sub
                End If
        End If
    Next
    Screen.MousePointer = 11
    DBSYSTEM.Execute "ALTER TABLE  [##TMPDETCOLUM" & VGL_COMPUTER & "]  DROP COLUMN " & Trim(LvColumn.ListItems(POSCOL - 1).Text)
    'ELIMINANDO LOS REGISTROS DE FORMULAS
    
    LvColumn.ListItems.Remove (POSCOL - 1)
    LvColumn.Refresh
    Set RSDET = Nothing
    RSDET.Open " [##TMPDETCOLUM" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic
    Set DgDet.DataSource = RSDET
    FORMHEAD
    frCfgDet2.LColumnas.ListItems.Remove (POSCOL + 1)
    frCfgDet2.Command2_Click
    frCfgDet2.CARGAR
    
    Screen.MousePointer = 1
End Sub

Private Sub Command2_Click()
    Dim RSTOT As ADODB.Recordset
    Dim COMA As String
    Dim I As Integer
    Dim RSAUX As New ADODB.Recordset
    Dim XLIST As ListItem
    Dim CAMPOS As String
    CAMPOS = ""
    'GENERACION DE CAMPOS PARA LOS TOTALES
    For I = 1 To LvColumn.ListItems.Count
        If I <> LvColumn.ListItems.Count Then
            COMA = ","
           Else: COMA = ""
        End If
        CAMPOS = CAMPOS & " SUM(" & LvColumn.ListItems(I) & ") AS S" & Trim(LvColumn.ListItems(I)) & COMA
    Next
    If CAMPOS = "" Then
        MsgBox "No existe ningún campo númerico ha calcular", vbExclamation
        Exit Sub
    End If
    
    RSAUX.Open "SELECT " & CAMPOS & " FROM  [##TMPDETCOLUM" & VGL_COMPUTER & "]  ", DBSYSTEM, adOpenKeyset, adLockOptimistic
    
    RSAUX.MoveFirst
    For I = 1 To LvColumn.ListItems.Count
         Set XLIST = FrmTotales.ListView1.ListItems.Add(, , Trim(LvColumn.ListItems(I)), , 1)
         XLIST.SubItems(1) = Trim(LvColumn.ListItems(I).SubItems(2))
         XLIST.SubItems(2) = Format(RSAUX.Fields("S" & Trim(LvColumn.ListItems(I))).Value, "###,###,##0.00")
         If frCfgDet2.LColumnas.ListItems.Item(I + 2).Checked = False Then
            FrmTotales.ListView1.ListItems.Remove (I)
         End If
    Next
    FrmTotales.Show 1
End Sub

Private Sub Command3_Click()
Dim X As Integer
Dim CARGACC  As Boolean
Dim RSAUX As ADODB.Recordset
Set RSAUX = New ADODB.Recordset
If Not ExisteTabla("DETADELANTOS") Then DBSYSTEM.Execute "CREATE TABLE DETADELANTOS (NUMBOL INT, MES DATETIME, CODTRAB VARCHAR(8), DESCRIP VARCHAR(50), MONTO  Numeric(20,2) )"
If LvColumn.ListItems.Count = 0 Then Exit Sub
    If UCase(Dir$(App.PATH & "\ADELCC.DYB")) = "ADELCC.DYB" Then
        If MsgBox("Esta a punto de cargar los Debitos por Cuentas Corrientes de Trabajadores.. desea hacer efectivo los debitos especificados", vbYesNo + vbQuestion) = vbNo Then CARGACC = False Else CARGACC = True
    End If
    If MsgBox("Desea continuar", vbYesNo) = vbYes Then
        If CARGACC Then
            RSAUX.Open App.PATH & "\ADELCC.DYB", , adOpenStatic, adLockReadOnly, adCmdFile
            Do While Not RSAUX.EOF
                DBSYSTEM.Execute "INSERT INTO PAGOSCTA (CODMOV,NUMBOL,CODNOMBOL,TIPOBOLETA,MONTO,DOLAR,CODTRAB,TIPO) VALUES (" & RSAUX!CODMOV & ",0,0,'A'," & Round(RSAUX!DEBITO, 2) & ",0,'" & RSAUX!CODTRAB & "'," & IIf(RSAUX!Tip = "E", 2, 1) & ")"
                DBSYSTEM.Execute "UPDATE MOVICTA SET SALDO=SALDO-" & RSAUX!DEBITO & " WHERE CODMOV=" & RSAUX!CODMOV
                RSAUX.MoveNext
            Loop
        End If
                With RSDET
                    .MoveFirst
                    While Not .EOF
                            'GRABA EL ADELANTOS
                            Do While Not RSDET.EOF
                                DBSYSTEM.Execute "INSERT INTO " & REGSISTEMA.TABLAADEL & " (CODTRAB,MES,FECHAING,MONTO,NUMBOL,NOMBOL, ORIGEN) VALUES ('" & RSDET!CODTRAB & "'," & DateSQL(FMES) & "," & DateSQL(Date) & "," & RSDET!MONTO & ",0,0,1)"
                                RSDET.MoveNext
                            Loop
                            'GRABA EL DETALLE DE ADELANTOS
                            For X = 0 To LvColumn.ListItems.Count
                                If LvColumn.ListItems(X).SubItems(3) = "X" Then
                                                                        
                                End If
                            Next
                        .MoveNext
                    Wend
                End With
            MsgBox "Se han grabado los datos satisfactoriamente", vbInformation
        End If
    Set RSDET = Nothing
    Set RSAUX = Nothing
    

End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub DCMCAMPO_CHANGE()
    xCod.Text = DCmcampo.BoundText
End Sub

Private Sub DGDET_ROWCOLCHANGE(LASTROW As Variant, ByVal LASTCOL As Integer)
    POSCOL = DgDet.COL
End Sub

Private Sub Form_Load()
'DETALLE ADELANTOS
Dim I As Integer
Dim INTO As String
    LvColumn.ListItems.Clear
    If ExisteTablaAux(" [##TMPDETCOLUM" & VGL_COMPUTER & "] ") Then DBSTARPLAN.Execute "Drop Table  [##TMPDETCOLUM" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "SELECT CODTRAB, NOMBRES INTO  [##TMPDETCOLUM" & VGL_COMPUTER & "]  FROM [##CALCINPUT" & Trim(VGL_COMPUTER) & "] "
    If MsgBox("Desea dar Adelanto al Trabajador Seleccionado(Si) o a Todos(No)?", vbYesNo, "Confirmar") = vbYes Then
        If Len(CodigoTrab) > 0 Then
            SQLSTR = "SELECT * from  [##TMPDETCOLUM" & VGL_COMPUTER & "]  where CODTRAB='" & CodigoTrab & "'"
            RSDET.Open "SELECT * from  [##TMPDETCOLUM" & VGL_COMPUTER & "]  where CODTRAB='" & CodigoTrab & "'", DBSTARPLAN, adOpenStatic, adLockOptimistic
        Else
            MsgBox "Escoga al Trabajador ", vbCritical
            Unload Me
            Exit Sub
        End If
    Else
        RSDET.Open " [##TMPDETCOLUM" & VGL_COMPUTER & "] ", DBSTARPLAN, adOpenStatic, adLockOptimistic
        SQLSTR = " [##TMPDETCOLUM" & VGL_COMPUTER & "] "
    End If
    Set DgDet.DataSource = RSDET
    DgDet.Refresh
    FORMHEAD
    'frCfgDet2.Show: frCfgDet2.Visible = False
    PROCSIS.PrCambiarEstilo Me.hWnd, Frame2.hWnd
    PROCSIS.PrCambiarEstilo Me.hWnd, Picture2.hWnd
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSDET = Nothing
End Sub

