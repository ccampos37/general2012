VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrRepNivAr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imprimir Areas de Trabajo x Nivel"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   Icon            =   "frRepNivAr.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport Reporte 
      Left            =   240
      Top             =   2670
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Salir"
      Height          =   360
      Left            =   3585
      TabIndex        =   6
      Top             =   2730
      Width           =   1020
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Imprimir"
      Height          =   360
      Left            =   2460
      TabIndex        =   5
      Top             =   2730
      Width           =   1020
   End
   Begin VB.TextBox SqlCad 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1545
      Visible         =   0   'False
      Width           =   1080
   End
   Begin MSComctlLib.ListView LvNivel 
      Height          =   2100
      Left            =   75
      TabIndex        =   1
      Top             =   540
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   3704
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   4304
      EndProperty
   End
   Begin VB.ComboBox CmbNivel 
      Height          =   315
      Left            =   990
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   150
      Width           =   1755
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3135
      Top             =   1635
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frRepNivAr.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frRepNivAr.frx":0D1E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      X1              =   4125
      X2              =   4230
      Y1              =   405
      Y2              =   405
   End
   Begin VB.Label lMarcaTodos 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Marcar Todos"
      Height          =   195
      Left            =   3585
      TabIndex        =   4
      Top             =   225
      Width           =   990
   End
   Begin VB.Image xMarcaTodos 
      Height          =   240
      Left            =   3105
      Picture         =   "frRepNivAr.frx":1172
      Top             =   195
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Niveles"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   195
      Width           =   540
   End
End
Attribute VB_Name = "FrRepNivAr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim COD As Integer, KEY As Integer
Private Sub CMBNIVEL_CLICK()
    CARGAREGS
End Sub

Public Sub CARGAREGS()
    Dim VRELAT As String
    Dim XLIST As ListItem
    Dim RSTEMP As New ADODB.Recordset
    'CARGAR LOS REGISTROS DEPENDIENDO DEL NIVEL ESCOGIDO
    LvNivel.ListItems.Clear
    SqlCad = "SELECT * FROM  [##TMPAREASTRAB" & VGL_COMPUTER & "]  WHERE NIVEL=" & Str(CmbNivel.ListIndex)
    RSTEMP.Open SqlCad.Text, DBSTARPLAN, adOpenKeyset, adLockReadOnly
    
    If RSTEMP.RecordCount = 0 Then
        MsgBox "NO SE ENCONTRO NIVELES", vbExclamation
        Exit Sub
    End If
    'CARGAR LOS REGISTROS AL LISTVIEW
    RSTEMP.MoveFirst
    Do While Not RSTEMP.EOF
        Set XLIST = LvNivel.ListItems.Add(, , RSTEMP!CODCCOSTO, , _
                            IIf(InStr(RSTEMP!CODCCOSTO, ".") > 0, 1, 2))
        XLIST.SubItems(1) = RSTEMP!NOMBRE
        RSTEMP.MoveNext
    Loop
    Set RSTEMP = Nothing
End Sub

Private Sub Command1_Click()
    Dim XITEM As ListItem
    
    'RECORRE LE LW PARA ARMAR LA CADENA DE CONSULTA DEPENDIENDO DE LOS CHECKS
    SqlCad.Text = ""
    For Each XITEM In LvNivel.ListItems
        If XITEM.Checked Then
            SqlCad.Text = SqlCad.Text & "T.AREA LIKE '" & Trim(XITEM.Text) & "%' OR "
        End If
    Next
    If SqlCad.Text = "" Then
        MsgBox "SELECCIONE POR LO MENOS UNA AREA", vbExclamation
        Exit Sub
    End If
    SqlCad.Text = Left(SqlCad.Text, Len(SqlCad.Text) - 4)
    'LLAMA AL PROCEDIMIENTO IMPRIMIR
    IMPRIMIR SqlCad
End Sub
Private Sub IMPRIMIR(SELE As TextBox)
    Dim X As Long
    Screen.MousePointer = 11
    
    If ExisteTablaAux(" [##TMPREPARTRAB" & VGL_COMPUTER & "] ") Then DBSTARPLAN.Execute "DROP TABLE  [##TMPREPARTRAB" & VGL_COMPUTER & "] "
    'SELECCION LOS DATOS ESCOGIDOS DEL LISTVIEW CREANDO CODAUX UN CAMPO MAS PARA IDENTIFICAR
    'LAS AREA CORRESPONDIENTES AL NIVEL ESCOGIDO
    DBSTARPLAN.Execute "" & _
        "SELECT T.*,(LTRIM(T.APEPAT) + ' ' + LTRIM(T.APEMAT) + ' ' + LTRIM(T.NOMBRE)) AS NOMBRES,  " & _
        "STR(0) AS CODAUX  " & _
        "INTO  [##TMPREPARTRAB" & VGL_COMPUTER & "]  " & _
        "FROM [" & REGSISTEMA.BASESQL & "].[dbo].TRABAJADORES T WHERE " & SELE.Text, X
    If X = 0 Then
      MsgBox "MENSAJE DEL SISTEMA: " & _
      " NO SE ENCONTRARÓN REGISTROS ", vbInformation
      Screen.MousePointer = 1
      Exit Sub
    End If
    
    'SE RECORRE EL TEMPORAL Y EN EL CODAUX DE COLOCA EL CODIGO DEL NIVEL SELECCIONADO
    'PARA LA IMPRESION CON LA FUNCION GETCAD
    Dim RSTEMP As New ADODB.Recordset
    RSTEMP.Open "SELECT CODTRAB,AREA,CODAUX FROM  [##TMPREPARTRAB" & VGL_COMPUTER & "] ", _
    DBSTARPLAN, adOpenDynamic, adLockOptimistic
    RSTEMP.MoveFirst
    Do While Not RSTEMP.EOF
        RSTEMP!CODAUX = Getcad(".", CmbNivel.ListIndex + 1, RSTEMP!AREA)
        RSTEMP.Update
        RSTEMP.MoveNext
    Loop
    Set RSTEMP = Nothing
    DBSTARPLAN.Execute "EXECUTE CREA_TMP_AREA '" & REGSISTEMA.BASESQL & "','" & VGL_COMPUTER & "'"
    
    Screen.MousePointer = 1
    With Reporte
        .Reset
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0002.RPT"
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
        .Destination = crptToWindow
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .StoredProcParam(0) = REGSISTEMA.BASESQL
        .StoredProcParam(1) = VGL_COMPUTER
        .WindowTitle = "PLAN0002 - TRABAJADORES POR NIVELES Y AREAS DE TRABAJO "
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(2) = "XNIVEL='" & CmbNivel.Text & "'"
        .Formulas(3) = "XRUC='" & REGSISTEMA.RUC & "'"
        .Formulas(4) = "XHORA='" & Format(Time, "HH:MM") & "'"
        If .Status <> 2 Then .Action = 1
    End With
    Exit Sub
ERRADO:
    MsgBox "POR FAVOR INTENTELO DE NUEVO"
    Screen.MousePointer = 1
End Sub


Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    MsgBox (Str(COD) & "," & Str(KEY))
End Sub

Private Sub FORM_KEYDOWN(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 84 And Shift = 4 Then XMARCATODOS_CLICK
End Sub

Private Sub TEXT1_KEYDOWN(KEYCODE As Integer, Shift As Integer)
    COD = KEYCODE
    KEY = Shift
End Sub

Private Sub XMARCATODOS_CLICK()
    Dim XITEM As ListItem
    For Each XITEM In LvNivel.ListItems
        If Not XITEM.Checked Then
            XITEM.Checked = True
        Else: XITEM.Checked = False
        End If
    Next
End Sub

