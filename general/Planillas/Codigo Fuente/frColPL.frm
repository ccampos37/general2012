VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frColPL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Columnas de Planilla"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   Icon            =   "frColPL.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   5640
   Tag             =   "Panel de Columnas de Planilla"
   Begin VB.CommandButton Command1 
      Caption         =   "&Preparar Tabla"
      Height          =   375
      Left            =   4005
      TabIndex        =   5
      Top             =   2130
      Width           =   1425
   End
   Begin VB.CommandButton cmEditar 
      Caption         =   "&Editar"
      Height          =   375
      Left            =   3990
      TabIndex        =   4
      Top             =   710
      Width           =   1425
   End
   Begin VB.CommandButton cmCerrar 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   3990
      TabIndex        =   3
      Top             =   1650
      Width           =   1425
   End
   Begin VB.CommandButton cmDel 
      Caption         =   "&Borrar"
      Height          =   375
      Left            =   3990
      TabIndex        =   2
      Top             =   1180
      Width           =   1425
   End
   Begin VB.CommandButton cmAdd 
      Caption         =   "&Insertar"
      Height          =   375
      Left            =   3990
      TabIndex        =   1
      Top             =   240
      Width           =   1425
   End
   Begin MSDataGridLib.DataGrid dgColumnas 
      Height          =   4215
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   3585
      _ExtentX        =   6324
      _ExtentY        =   7435
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
         MarqueeStyle    =   3
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   4815
      Picture         =   "frColPL.frx":0442
      Top             =   3720
      Width           =   240
   End
End
Attribute VB_Name = "frColPL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSCOLPL As New ADODB.Recordset
Dim REGACT As REGWIN

Private Sub CMADD_CLICK()
    If RSCOLPL.EOF Then
        VPNUMTMP = 0
    Else
        VPNUMTMP = RSCOLPL!INDICE
    End If
    VPTAREA = "NUEVO"
    frEColPL.Show 1
    RSCOLPL.Requery
End Sub

Private Sub cmCerrar_Click()
    Unload Me
End Sub

Private Sub CMDEL_CLICK()
    On Error GoTo Err1
    If RSCOLPL.EOF Then Exit Sub
    
    If InStr("'TOTING','TOTEGR','NETO'", "'" & RSCOLPL!Codigo & "'") > 0 Then
        MsgBox "columna del sistema,no se puede eliminar", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Seguro de eliminar el registro seleccionado de columnas de planilla", vbInformation + vbYesNo) = vbNo Then Exit Sub
    RSCOLPL.Delete
    RSCOLPL.MoveFirst
    Dim I As Byte
    I = 0
    Do While Not RSCOLPL.EOF
       I = I + 1
       RSCOLPL!INDICE = I
       RSCOLPL.Update
       RSCOLPL.MoveNext
    Loop
    RSCOLPL.MoveFirst
    Exit Sub
Err1:
    Unload Me
End Sub

Private Sub CMEDITAR_CLICK()
    If RSCOLPL.EOF Then Exit Sub
    VPTAREA = RSCOLPL!Codigo
    VPNUMTMP = RSCOLPL!INDICE
    frEColPL.Show 1
    RSCOLPL.Requery
End Sub

Private Sub Command1_Click()
    On Error GoTo ERRGENRAPLAN
    Dim YAEXISTE As Boolean, STRCAD As String
    YAEXISTE = False
    Screen.MousePointer = 11
    If ExisteTabla(REGSISTEMA.TABLAPLAN) Then
        'CAMBIAMOS EL NOMBRE
        If ExisteTabla("PLANAUX") Then
            MsgBox "La operacion anterior no fue completada con exito, se procedera a recuperar la informacion"
            DBSYSTEM.Execute "DROP TABLE PLANAUX"
        End If
        DBSYSTEM.Execute "SELECT * INTO PLANAUX FROM " & REGSISTEMA.TABLAPLAN
        DBSYSTEM.Execute "DROP TABLE " & REGSISTEMA.TABLAPLAN
        YAEXISTE = True
    End If
    STRCAD = "CREATE TABLE " & REGSISTEMA.TABLAPLAN & " (MES DATETIME,TIPOPLANILLA BIT, INUMBOL INT, CODTRAB VARCHAR(8), NOMBRES VARCHAR(35), TIPOTRAB VARCHAR(2), FECHAING DATETIME, SITUACION VARCHAR(2), CCOSTO VARCHAR(10), CENTROCOSTO VARCHAR(25), DEPARTAMENTO VARCHAR(25), CARGO VARCHAR(25), BASICO  Numeric(20,2) , FONDOPENS VARCHAR(2), FECHACESE DATETIME, CODSCTR VARCHAR(6), EPS VARCHAR(8), CARNETSEG VARCHAR(15), VACINI DATETIME, VACFIN DATETIME, CUSPP VARCHAR(12), REDONDEO  Numeric(20,2) "
    Dim RSAUX As New ADODB.Recordset
    RSAUX.Open "SELECT * FROM COLUMPL ORDER BY INDICE", DBSYSTEM, adOpenStatic
    Do While Not RSAUX.EOF
        STRCAD = STRCAD & ", " & RSAUX!Codigo & "  Numeric(20,2) "
        RSAUX.MoveNext
    Loop
    STRCAD = STRCAD & ")"
    DBSYSTEM.Execute STRCAD
    '-------------------------------------------------------------------------------------
    If YAEXISTE Then
        RSAUX.Close
        On Error GoTo ERRNOEXISTE
        RSAUX.Open "PLANAUX", DBSYSTEM, adOpenStatic
        Dim RSPL02 As New ADODB.Recordset
        RSPL02.Open REGSISTEMA.TABLAPLAN, DBSYSTEM, adOpenDynamic, adLockOptimistic
        Do While Not RSAUX.EOF
            RSPL02.AddNew
            For X = 0 To RSAUX.Fields.Count - 1
                RSPL02.Fields(Trim$(RSAUX.Fields(X).Name)).Value = RSAUX.Fields(X).Value
            Next
            RSPL02.Update
            RSAUX.MoveNext
        Loop
        Set RSPL02 = Nothing
        
    End If
    If ExisteTabla("PLANAUX") Then DBSYSTEM.Execute "DROP TABLE PLANAUX"
    'EMPIEZA LA INSERCIÓN DE LOS REGISTROS
    '-------------------------------------------------------------------------------------
    Set RSAUX = Nothing
    DBSYSTEM.Execute "DELETE FROM VERCOLPL"
    CREARPLANILLA
    Screen.MousePointer = 1
    Exit Sub
ERRNOEXISTE:
    Resume Next
ERRGENRAPLAN:
    Screen.MousePointer = 1
    MsgBox ERR.Description
End Sub

Private Sub Form_Activate()
    ActivarTools REGACT
End Sub

Private Sub Form_Load()
    RSCOLPL.Open "SELECT CODIGO, NOMBRE, INDICE FROM COLUMPL ORDER BY INDICE", DBSYSTEM, adOpenKeyset, adLockOptimistic
    Set dgColumnas.DataSource = RSCOLPL
    With REGACT
        .BUSCAR = False
        .EDITAR = True
        .ELIMINAR = True
        .FILTRAR = False
        .IMPRIMIR = False
        .NUEVO = True
        .PRELIMINAR = False
    End With
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    RSCOLPL.Close
    Set dgColumnas.DataSource = Nothing
    Set RSCOLPL = Nothing
End Sub

Public Sub COMANDOTOOLBAR(ByVal COMANDO As String)
    Select Case UCase(COMANDO)
        Case "NUEVO"
            CMADD_CLICK
        Case "EDITAR"
            CMEDITAR_CLICK
        Case "ELIMINAR"
            CMDEL_CLICK
        Case Else
            MsgBox "Funcion no implementada", vbCritical
    End Select
End Sub

Public Sub CREARPLANILLA()
    MsgBox "La creacion de la tabla de planilla concluyo satisfactoriamente", vbInformation, "INFORMACIÓN"
    Exit Sub
End Sub

