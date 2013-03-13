VERSION 5.00
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos Generales del Trabajador"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4905
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4905
   Begin VB.Frame Frame1 
      Caption         =   "Seleccionar Trabajador"
      Height          =   1065
      Left            =   112
      TabIndex        =   2
      Top             =   120
      Width           =   4680
      Begin AplisetControlText.Aplitext xTrab 
         Height          =   285
         Left            =   1065
         TabIndex        =   3
         Top             =   525
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   503
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trabajador"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   135
         TabIndex        =   4
         Top             =   525
         Width           =   765
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   405
      Left            =   3232
      TabIndex        =   1
      Top             =   2115
      Width           =   1560
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   405
      Left            =   3232
      TabIndex        =   0
      Top             =   1530
      Width           =   1560
   End
   Begin Crystal.CrystalReport Reporte 
      Left            =   150
      Top             =   1725
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS_TABLA As ADODB.Recordset
Dim RS_AUX As ADODB.Recordset
Private Sub Command1_Click()
Dim SQLSTR As String
Set RS_TABLA = New ADODB.Recordset
Set RS_AUX = New ADODB.Recordset

    If Len(Trim(xTrab.Tag)) = 0 Then
        MsgBox "Escoga un Trabjador", vbInformation
        xTrab.SetFocus
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    If Not ExisteTabla("DATATRAB") Then
        MsgBox "ERROR NO SE ENCONTRO EL ARCHIVO O LA TABLA DATA TRABAJADOR", vbCritical, "INFORMACION"
        Exit Sub
    End If
        If Not ExisteTablaSQL(" [##_TMPCNP" & VGL_COMPUTER & "] ", DBSYSTEM) Then DBSTARPLAN.Execute "CREATE TABLE  [##_TMPCNP" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8), CODCNP VARCHAR(30), CONCEPTO VARCHAR(100), VALOR VARCHAR(50))"
        DBSTARPLAN.Execute "DELETE FROM  [##_TMPCNP" & VGL_COMPUTER & "] "
        RS_AUX.Open "SELECT * FROM DATATRAB", DBSYSTEM
        If RS_AUX.RecordCount Then
            While Not RS_AUX.EOF
                SQLSTR = "SELECT CODTRAB, " & RS_AUX.Fields(0) & " FROM TRABAJADORES WHERE CODTRAB='" & xTrab.Tag & "'"
                RS_TABLA.Open SQLSTR, DBSYSTEM
                If Not IsNull(RS_TABLA.Fields(1)) Then
                    SQL = "INSERT INTO  [##_TMPCNP" & VGL_COMPUTER & "]  VALUES('" & RS_TABLA.Fields(0) & "','" & RS_AUX.Fields(0) & "','" & RS_AUX.Fields(1) & "','" & RS_TABLA.Fields(1) & "')"
                    DBSTARPLAN.Execute SQL, E
                End If
                RS_TABLA.Close
                RS_AUX.MoveNext
            Wend
        End If
        RS_AUX.Close
    With Reporte
        .Reset
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL
        .ReportFileName = REGSISTEMA.REPORTES & "PLRH0015.RPT"
        .SelectionFormula = "{TRABAJADORES.CODTRAB}='" & xTrab.Tag & "'"
        .WindowTitle = "PLRH0015.RPT - " & Me.Caption
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .Formulas(0) = "EMPRESA='" & REGSISTEMA.EMPRESA & "'"
        
        .SubreportToChange = "PlRH0001.rpt"
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL & ""

        .SubreportToChange = "PlRH0003.rpt"
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL & ""
        
        .SubreportToChange = "PlRH0002.rpt"
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL & ""
        
        .SubreportToChange = "PlRH0006.rpt"
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL & ""
        
        
        .SubreportToChange = "PlRH0014.rpt"
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL & ""
        
        
        If .Status <> 2 Then .Action = 1
    End With
     Screen.MousePointer = vbNormal
End Sub
Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub XTRAB_DBLCLICK()
    Dim RSTRAB As New ADODB.Recordset
    RSTRAB.Open "SELECT CODTRAB, NOMBRES FROM VWTRABAJ", DBSYSTEM, adOpenKeyset, adLockReadOnly
    If RSTRAB.EOF Or RSTRAB.RecordCount = 0 Then
        MsgBox "NO SE HAN ENCONTRADO REGISTRO DE TRABAJADORES", vbCritical
        Set RSTRAB = Nothing
        Exit Sub
    End If
    frmComun.CONECTAR RSTRAB
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xTrab.Tag = RSTRAB!CODTRAB
        xTrab.Text = RSTRAB!CODTRAB & " : " & RSTRAB!NOMBRES
    End If
    Set RSTRAB = Nothing
End Sub

