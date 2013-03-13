VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmDetalle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Utilitario de Calculo"
   ClientHeight    =   8160
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11835
   Icon            =   "FrmDetalle.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   11835
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   945
      Left            =   45
      TabIndex        =   23
      Top             =   105
      Width           =   4485
      Begin VB.TextBox xCod 
         Height          =   285
         Left            =   870
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   120
         Width           =   1380
      End
      Begin MSDataListLib.DataCombo DCmcampo 
         Height          =   315
         Left            =   870
         TabIndex        =   25
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
         TabIndex        =   27
         Top             =   180
         Width           =   675
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Columna "
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   60
         TabIndex        =   26
         Top             =   510
         Width           =   840
      End
   End
   Begin VB.CommandButton CmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   315
      Left            =   4635
      TabIndex        =   18
      Top             =   705
      Width           =   1200
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "&Agregar"
      Height          =   315
      Left            =   4635
      TabIndex        =   17
      Top             =   375
      Width           =   1200
   End
   Begin VB.CommandButton CmdImpPla 
      Caption         =   "&Imprimir Planilla"
      Height          =   345
      Left            =   9690
      TabIndex        =   16
      Top             =   7425
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Totales"
      Height          =   345
      Left            =   9690
      TabIndex        =   15
      Top             =   7770
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Configuración de la Impresión"
      Height          =   1815
      Left            =   60
      TabIndex        =   6
      Top             =   6300
      Width           =   9285
      Begin AplisetControlText.Aplitext xTitulo 
         Height          =   300
         Index           =   0
         Left            =   1515
         TabIndex        =   7
         Top             =   450
         Width           =   7605
         _ExtentX        =   13414
         _ExtentY        =   529
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xTitulo 
         Height          =   300
         Index           =   1
         Left            =   1515
         TabIndex        =   8
         Top             =   765
         Width           =   7605
         _ExtentX        =   13414
         _ExtentY        =   529
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xTitulo 
         Height          =   300
         Index           =   2
         Left            =   1515
         TabIndex        =   9
         Top             =   1080
         Width           =   7605
         _ExtentX        =   13414
         _ExtentY        =   529
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xTitulo 
         Height          =   300
         Index           =   3
         Left            =   1515
         TabIndex        =   10
         Top             =   1395
         Width           =   7605
         _ExtentX        =   13414
         _ExtentY        =   529
         Text            =   ""
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Titulo del Informe"
         Height          =   195
         Left            =   195
         TabIndex        =   14
         Top             =   510
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Encabezado 01"
         Height          =   195
         Index           =   3
         Left            =   225
         TabIndex        =   13
         Top             =   825
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Encabezado 02"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   12
         Top             =   1140
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Encabezado 03"
         Height          =   195
         Index           =   2
         Left            =   225
         TabIndex        =   11
         Top             =   1455
         Width           =   1125
      End
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
      Caption         =   "&Configuración de Vista"
      Height          =   345
      Left            =   9705
      TabIndex        =   5
      Top             =   6435
      Width           =   2055
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir Text"
      Height          =   345
      Left            =   9690
      TabIndex        =   4
      Top             =   7080
      Width           =   2055
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Height          =   5145
      Left            =   45
      ScaleHeight     =   5085
      ScaleWidth      =   11670
      TabIndex        =   19
      Top             =   1065
      Width           =   11730
      Begin VB.TextBox XFormula 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   750
         TabIndex        =   28
         Top             =   30
         Width           =   9750
      End
      Begin VB.CommandButton Command1 
         Height          =   285
         Left            =   10575
         Picture         =   "FrmDetalle.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Ejecutar Fórmula"
         Top             =   30
         Width           =   810
      End
      Begin MSDataGridLib.DataGrid DgDet 
         Height          =   4515
         Left            =   75
         TabIndex        =   20
         Top             =   360
         Width           =   11340
         _ExtentX        =   20003
         _ExtentY        =   7964
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
         TabIndex        =   21
         Top             =   90
         Width           =   585
      End
   End
   Begin MSComctlLib.ListView LvColumn 
      Height          =   2280
      Left            =   1980
      TabIndex        =   0
      Top             =   3165
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   4022
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
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
   End
   Begin VB.TextBox SqlCad 
      Height          =   330
      Left            =   7815
      TabIndex        =   1
      Top             =   3885
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.TextBox SqlSel 
      Height          =   315
      Left            =   7830
      TabIndex        =   2
      Top             =   4215
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.TextBox xVarFormu 
      Height          =   285
      Left            =   7830
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   4515
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Utilitario de Cálculo"
      ForeColor       =   &H00000040&
      Height          =   360
      Left            =   8790
      TabIndex        =   30
      Top             =   45
      Width           =   2385
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Utilitario de Calculo"
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   8790
      TabIndex        =   29
      Top             =   60
      Width           =   2385
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   11280
      Picture         =   "FrmDetalle.frx":0C0C
      Top             =   60
      Width           =   480
   End
End
Attribute VB_Name = "FrmDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim POSCOL As Integer
Dim RSDET As New ADODB.Recordset
Dim FLAGCOLUM As Boolean
Dim INTO As String

Private Sub CMDAGREGAR_CLICK()
Dim XLIST As ListItem
Dim SQLBASE As String
Dim RSDETAUX As New ADODB.Recordset
    If Trim(DCmcampo.Text) = "" Then
        MsgBox "Debe seleccionar un concepto", vbExclamation
        Exit Sub
    End If
    If VALIDCAMP(DCmcampo.BoundText) Then Exit Sub
    CambiaPanelBD True
    Screen.MousePointer = 11
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
    Set RSDET = Nothing
    RSDET.Open " [##TMPDETCOLUM" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockReadOnly
    Set DgDet.DataSource = RSDET
    FORMHEAD
    
    frCfgDet.Command2_Click
    frCfgDet.CARGAR
    Screen.MousePointer = 1
    CambiaPanelBD False
End Sub
Private Function DEVOLVERIMPORTE(CODTRAB As String, CAMPO As String) As Double
    Dim RSAUX As New ADODB.Recordset
    Dim SQL As String
    DEVOLVERIMPORTE = 0
    SQL = "SELECT * FROM  [##TMPDETAGROUP" & VGL_COMPUTER & "]  WHERE CODTRAB='" & Trim(CODTRAB) & " ' AND CODCONCEP='" & Trim(CAMPO) & "'"
    RSAUX.Open SQL, DBSYSTEM, adOpenKeyset, adLockReadOnly
    If RSAUX.RecordCount = 0 Then Exit Function
    
    DEVOLVERIMPORTE = RSAUX!SUMADEMONTO
End Function

Private Sub CMDCFGVISTA_Click()
    frCfgDet.Show 1
End Sub

Private Sub CMDIMPPLA_Click()
    Dim REG As Long
    CambiaPanelBD True
    Screen.MousePointer = 11
    Call CREARQUERYTABLA
    
    DBSTARPLAN.Execute "EXEC QUERY_DETALLE " & VGL_COMPUTER
    DBSYSTEM.Execute "UPDATE [##TMPTRAB" & VGL_COMPUTER & "] SET CODTRAB=CODTRAB", REG
    With Reporte
        .Reset
        .WindowTitle = "PLAN0068.RPT - RESUMEN"
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0068.RPT"
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=" & VGL_BASE & ""
        .StoredProcParam(0) = VGL_COMPUTER
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .Formulas(0) = "XCABEZA='" & Trim(xTitulo(0).Text) & "'"
        .Formulas(1) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(2) = "XRUC='" & REGSISTEMA.RUC & "'"
        .Formulas(3) = "XREG='" & Str(REG) & "'"
        If .Status <> 2 Then .Action = 1
    End With
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
  
  For Each XITEM In frCfgDet.LColumnas.ListItems
    If XITEM.INDEX > 2 Then
        If XITEM.Checked Then
            ORDEN = ORDEN + 1
            DBSYSTEM.Execute "INSERT INTO  [##TMPLIST" & VGL_COMPUTER & "]  VALUES('" & Trim(XITEM.Text) & "'," & Str(ORDEN) & ")"
        End If
    End If
  Next
End Sub
Private Sub CmdImprimir_Click()
    frCfgDet.CARGAR
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
            CAD = RSAUX.Fields(Trim(frCfgDet.LColumnas.ListItems(X + 1))).Type
            Select Case RSAUX.Fields(Trim(frCfgDet.LColumnas.ListItems(X + 1))).Type
                    Case adSingle, 3, 5, 6 'SI ES NUMERO SIMPLE
                        XNUMCHAR = 10
                    Case adVarChar
                        XNUMCHAR = RSAUX.Fields(Trim(frCfgDet.LColumnas.ListItems(X + 1))).DefinedSize + 1
                    Case adUnsignedTinyInt
                        XNUMCHAR = 3
            End Select
            CAD = frCfgDet.LColumnas.ListItems(X + 1).SubItems(1)
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
            xValor = RSAUX.Fields(Trim(frCfgDet.LColumnas.ListItems(X + 1))).Value
            If DgDet.Columns(X).Visible Then
                Select Case RSAUX.Fields(Trim(frCfgDet.LColumnas.ListItems(X + 1))).Type
                    Case adSingle, 3, 5, 6 'SI ES NUMERO SIMPLE
                        XSTR = XSTR & " " & Right("         " & Format$(xValor, "0.00"), 9)
                    Case adVarChar
                        XN = RSAUX.Fields(Trim(frCfgDet.LColumnas.ListItems(X + 1))).DefinedSize
                        XSTR = XSTR & " " & Left(xValor & String(XN, " "), XN)
                    Case adUnsignedTinyInt
                        XSTR = XSTR & " " & Format(xValor, "00")
                    Case Else
                        'MSGBOX "TIPO NO ENCONTRADO: " & RSAUX.FIELDS(TRIM(FRCFGDET.LCOLUMNAS.LISTITEMS(X + 1))).TYPE, VBCRITICAL
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
            Select Case RSAUX.Fields(Trim(frCfgDet.LColumnas.ListItems(X + 1))).Type
                    Case adDate 'SI ES POR FECHA
                        XNUMCHAR = 9
                    Case adSingle, 3, 5, 6  'SI ES NUMERO SIMPLE
                        XNUMCHAR = 10
                    Case adVarChar
                        XNUMCHAR = RSAUX.Fields(Trim(frCfgDet.LColumnas.ListItems(X + 1))).DefinedSize + 1
                    Case adUnsignedTinyInt
                        XNUMCHAR = 3
            End Select
            If RSAUX.Fields(Trim(frCfgDet.LColumnas.ListItems(X + 1))).Type = 5 Then
                RSTOT.Open "SELECT SUM(" & Trim(frCfgDet.LColumnas.ListItems(X + 1)) & ") AS TOTAL FROM  [##TMPDETCOLUM" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenStatic
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
    With Reporte
        .Reset
        .WindowTitle = "PLAN0023 - " & DgDet.Caption
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0023.RPT"
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=" & VGL_BASE & ""
        '.DataFiles(0) = App.PATH & "\BDAUXCOM.MDB"
        .StoredProcParam(0) = " [##PRTPLANILLA" & VGL_COMPUTER & "] "
        .Destination = crptToWindow
        .WindowShowPrintBtn = True
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        If xTitulo(0).Text = "" Then
            .Formulas(0) = "XTITULO=' " & DgDet.Caption & "'"
        Else
            .Formulas(0) = "XTITULO=' " & xTitulo(0).Text & "'"
        End If
        .Formulas(1) = "XCABEZA1=' " & xTitulo(1).Text & "'"
        .Formulas(2) = "XCABEZA2=' " & xTitulo(2).Text & "'"
        .Formulas(3) = "XCABEZA3=' " & xTitulo(3).Text & "'"
        If .Status <> 2 Then .Action = 1
    End With
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
    'LLENAR LOS REGISTROS DE FORMULA
    Dim RS As New ADODB.Recordset
    RS.Open " [##TMPDETCOLUM" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic
    RS.MoveFirst
    Do While Not RS.EOF
        If RS.Fields(Trim(CAD)).Value <> 0 Then
            DBSYSTEM.Execute "INSERT INTO  [##TMPDETAGROUP" & VGL_COMPUTER & "] (CODTRAB,NOMBRES,CODCONCEP,NOMCONCEP,SUMADEMONTO) VALUES('" & _
                              Trim(RS!CODTRAB) & "','" & Trim(RS!NOMBRES) & "','" & Trim(CAD) & "','" & Trim(CAD) & "'," & Str(RS.Fields(Trim(CAD)).Value) & ")"
        End If
        RS.MoveNext
    Loop
    Set RSDET = Nothing
    
    Set XLIST = LvColumn.ListItems.Add(, , Trim(CAD))
    XLIST.SubItems(1) = Trim(xVarFormu.Text)
    XLIST.SubItems(2) = Trim(CAD)
    XLIST.SubItems(3) = "X"
    LvColumn.Refresh
    
    RSDET.Open " [##TMPDETCOLUM" & VGL_COMPUTER & "] ", DBSTARPLAN, adOpenKeyset, adLockOptimistic
    Set DgDet.DataSource = RSDET
    DgDet.Refresh
    FORMHEAD ' FORMATEA LAS COLUMNAS DEL GRID
    
    frCfgDet.Command2_Click
    frCfgDet.CARGAR
    Screen.MousePointer = 1
    CambiaPanelBD False
        
    Exit Sub
    
ERRFORMULA:
    Resume Next
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
    If Trim(LvColumn.ListItems(POSCOL - 1).SubItems(3)) = "X" Then
        DBSYSTEM.Execute "DELETE FROM  [##TMPDETAGROUP" & VGL_COMPUTER & "]  WHERE CODCONCEP='" & Trim(LvColumn.ListItems(POSCOL - 1).Text) & "'"
    End If
    LvColumn.ListItems.Remove (POSCOL - 1)
    LvColumn.Refresh
    Set RSDET = Nothing
    RSDET.Open " [##TMPDETCOLUM" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic
    Set DgDet.DataSource = RSDET
    FORMHEAD
    frCfgDet.LColumnas.ListItems.Remove (POSCOL + 1)
    frCfgDet.Command2_Click
    frCfgDet.CARGAR
    
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
         If frCfgDet.LColumnas.ListItems.Item(I + 2).Checked = False Then
            FrmTotales.ListView1.ListItems.Remove (I)
         End If
    Next
    FrmTotales.Show 1
End Sub

Private Sub DCMCAMPO_CHANGE()
    xCod = DCmcampo.BoundText
End Sub

Private Sub DGDET_HEADCLICK(ByVal COLINDEX As Integer)
Dim CAMP As String
Dim RS As New ADODB.Recordset
Dim INTO As String
    POSCOL = COLINDEX
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT * FROM  [##TMPDETCOLUM" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic
    Set DgDet.DataSource = RS
    FORMHEAD
    If LvColumn.ListItems.Count > 0 And COLINDEX > 1 Then
        CAMP = LvColumn.ListItems.Item(COLINDEX - 1)
    End If
    Select Case COLINDEX
        Case 0: CAMP = "CODTRAB"
        Case 1: CAMP = "NOMBRES"
    End Select
    
    If Not FLAGCOLUM Then
        FLAGCOLUM = True
        RS.Sort = CAMP & " ASC "
      Else: RS.Sort = CAMP & " DESC "
          FLAGCOLUM = False
    End If
    frCfgDet.Command2_Click
End Sub

Private Sub DGDET_ROWCOLCHANGE(LASTROW As Variant, ByVal LASTCOL As Integer)
    POSCOL = DgDet.COL
End Sub

Private Sub Form_Load()
Dim I As Integer
Dim INTO As String
    DBSTARPLAN.Execute "EXECUTE SP_ARMAR_CONSULTA2 '" & VGL_COMPUTER & "'"
    RSDET.Open " [##TMPDETCOLUM" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic
    Set DgDet.DataSource = RSDET
    FORMHEAD
    frCfgDet.Show: frCfgDet.Visible = False
    PROCSIS.PrCambiarEstilo Me.hWnd, Frame2.hWnd
    PROCSIS.PrCambiarEstilo Me.hWnd, Picture2.hWnd
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSDET = Nothing
End Sub

