VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frAdminGrati 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cálculo de Gratificaciones"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8700
   Icon            =   "frAdminGrati.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   8700
   Begin VB.CommandButton cmAdelanto 
      Caption         =   "Dar Adelanto de Gratific."
      Height          =   465
      Left            =   1515
      TabIndex        =   12
      Top             =   4485
      Width           =   1155
   End
   Begin Crystal.CrystalReport Reporte 
      Left            =   3555
      Top             =   2265
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Planilla de Gratificación"
      Height          =   465
      Left            =   210
      TabIndex        =   9
      Top             =   4485
      Width           =   1155
   End
   Begin VB.CommandButton cmConsulta 
      Caption         =   "&Consulta"
      Height          =   405
      Left            =   7215
      TabIndex        =   2
      Top             =   2595
      Width           =   1380
   End
   Begin VB.CommandButton cmListado 
      Caption         =   "&Resumen"
      Height          =   465
      Left            =   5805
      TabIndex        =   10
      Top             =   4485
      Width           =   1155
   End
   Begin VB.CommandButton cmCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   405
      Left            =   7215
      TabIndex        =   6
      Top             =   4755
      Width           =   1380
   End
   Begin VB.CommandButton cmAceptar 
      Caption         =   "&Aceptar"
      Height          =   405
      Left            =   7215
      TabIndex        =   5
      Top             =   4215
      Width           =   1380
   End
   Begin VB.CommandButton cmEliminar 
      Caption         =   "&Eliminar"
      Height          =   405
      Left            =   7215
      TabIndex        =   3
      Top             =   3135
      Width           =   1380
   End
   Begin VB.CommandButton cmModificar 
      Caption         =   "&Modificar"
      Height          =   405
      Left            =   7215
      TabIndex        =   4
      Top             =   3675
      Width           =   1380
   End
   Begin VB.CommandButton cmPrueba 
      Caption         =   "&Prueba Cálculo"
      Height          =   405
      Left            =   7215
      TabIndex        =   1
      Top             =   2055
      Width           =   1380
   End
   Begin VB.CommandButton cmNuevo 
      Caption         =   "&Nuevo Cálculo"
      Height          =   405
      Left            =   7215
      TabIndex        =   0
      Top             =   1515
      Width           =   1380
   End
   Begin MSDataGridLib.DataGrid xData 
      Height          =   4200
      Left            =   210
      TabIndex        =   8
      Top             =   165
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   7408
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      Caption         =   "Planillas de Cálculo de Gratificaciones"
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "Codigo"
         Caption         =   "Código"
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
         DataField       =   "Nombre"
         Caption         =   "Nombre Descriptivo"
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
         DataField       =   "Soles"
         Caption         =   "Total S/."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00 "
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3195.213
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1005.165
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gratificación"
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
      Index           =   1
      Left            =   7365
      TabIndex        =   11
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Gratificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Index           =   0
      Left            =   7380
      TabIndex        =   7
      Top             =   735
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   7965
      Picture         =   "frAdminGrati.frx":08CA
      Top             =   195
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   5085
      Left            =   90
      Top             =   90
      Width           =   7005
   End
End
Attribute VB_Name = "frAdminGrati"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents RSCTS As ADODB.Recordset
Attribute RSCTS.VB_VarHelpID = -1

Private Sub CMACEPTAR_CLICK()
    VPTRASPRM = "" & RSCTS!Codigo
    frAceptaGrati.Show 1
    RSCTS.Requery
    Set xData.DataSource = RSCTS
End Sub

Private Sub CMADELANTO_CLICK()
    If RSCTS.BOF Or RSCTS.EOF Then Exit Sub
    VPTRASPRM = "" & RSCTS!Codigo
    frAdelantoGratif.Show 1
End Sub

Private Sub cmCerrar_Click()
    Unload Me
End Sub

Private Sub CMCONSULTA_CLICK()
    VPTAREA = "VISTA"
    VPTRASPRM = "" & RSCTS!Codigo
    frCalcGrati.Show 1
End Sub


Private Sub CMELIMINAR_CLICK()
    If RSCTS.EOF Or RSCTS.RecordCount = 0 Then
        MsgBox "No hay registro ha eliminar", vbInformation
    Else
        MsgBox "ADVERTENCIA: eliminará una planilla de Gratificaciones, sin posibilidad a recuperar su información", vbExclamation
        If MsgBox("Seguro de eliminar la planilla de Gratificaciones: " & RSCTS!NOMBRE & " . Seguro de Continuar", vbYesNo + vbQuestion) = vbNo Then Exit Sub
        DBSYSTEM.Execute "DELETE FROM GRATIFICACION WHERE CODIGO=" & RSCTS!Codigo
        DBSYSTEM.Execute "DELETE FROM PLANGRATI WHERE CODIGO=" & RSCTS!Codigo
        DBSYSTEM.Execute "DELETE FROM DETALLEGRATI WHERE CODIGO=" & RSCTS!Codigo
        RSCTS.Requery
        Set xData.DataSource = RSCTS
    End If
End Sub
Private Sub CMLISTADO_CLICK()
    If RSCTS.RecordCount = 0 Then
        MsgBox "No existen registros ha imprimir", vbCritical
        Exit Sub
    End If
    Call CambiaPanelBD(True)
    Screen.MousePointer = 11
    With Reporte
        .Reset
        .WindowTitle = "PLAN0054 - RESUMEN DE CÁLCULO DE GRATIFICACIONES"
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0054.RPT"
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
        .StoredProcParam(0) = REGSISTEMA.BASESQL
        .StoredProcParam(1) = "PLANGRATI"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "XMES='CORRESPONDIENTE A: " & RSCTS!NOMBRE & "'"
        If .Status <> 2 Then .Action = 1
    End With
    Call CambiaPanelBD(False)
    Screen.MousePointer = 1
End Sub

Private Sub CMMODIFICAR_CLICK()
    VPTAREA = "MODIFICAR"
    VPTRASPRM = "" & RSCTS!Codigo
    frCalcGrati.Show 1
    RSCTS.Requery
    Set xData.DataSource = RSCTS
End Sub

Private Sub CMNUEVO_CLICK()
    VPTAREA = "NUEVO"
    frCalcGrati.Show 1
    RSCTS.Requery
    Set xData.DataSource = RSCTS
End Sub

Private Sub CMPRUEBA_CLICK()
    VPTAREA = "PRUEBA"
    frCalcGrati.Show 1
End Sub

Private Sub Command1_Click()
    If RSCTS.RecordCount = 0 Then
        MsgBox "No existen registros ha imprimir", vbCritical
        Exit Sub
    End If
    Call CambiaPanelBD(True)
    Screen.MousePointer = 11
    If ExisteTablaAux(" [##TMPPLANGRATI" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPPLANGRATI" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "SELECT PG.NOMBRES, DG.* INTO  [##TMPPLANGRATI" & VGL_COMPUTER & "]  FROM PLANGRATI PG, DETALLEGRATI DG WHERE PG.CODIGO = DG.CODIGO AND PG.CODTRAB = DG.CODTRAB AND PG.CODIGO=" & RSCTS!Codigo
    DBSYSTEM.Execute "INSERT INTO  [##TMPPLANGRATI" & VGL_COMPUTER & "]  (CODTRAB, NOMBRES, CONCEPTO, IMPORTE) SELECT CODTRAB, NOMBRES, ' CALCULO DE GRATIF.' AS CONCEPTO,IMPORTEGRATI AS IMPORTE FROM PLANGRATI WHERE CODIGO=" & RSCTS!Codigo
    DBSYSTEM.Execute "UPDATE  [##TMPPLANGRATI" & VGL_COMPUTER & "]  SET IMPORTE=IMPORTE"
    With Reporte
        .Reset
        .WindowTitle = "PLAN0056 - DETALLE DE CÁLCULO DE GRATIFICACIONES"
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0056.RPT"
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=" & VGL_BASE & ""
        .StoredProcParam(0) = " [##TMPPLANGRATI" & VGL_COMPUTER & "] "
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "XMES='CORRESPONDIENTE A: " & RSCTS!NOMBRE & "'"
        If .Status <> 2 Then .Action = 1
    End With
    Call CambiaPanelBD(False)
    Screen.MousePointer = 1
End Sub

Private Sub Form_Load()
    Set RSCTS = New ADODB.Recordset
    RSCTS.Open "GRATIFICACION", DBSYSTEM, adOpenKeyset, adLockOptimistic
    Set xData.DataSource = RSCTS
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSCTS = Nothing
End Sub

Private Sub RSCTS_MOVECOMPLETE(ByVal ADREASON As ADODB.EventReasonEnum, ByVal PERROR As ADODB.Error, ADSTATUS As ADODB.EventStatusEnum, ByVal PRECORDSET As ADODB.Recordset)
    If RSCTS.EOF Or RSCTS.RecordCount = 0 Or RSCTS.BOF Then
        cmAceptar.Enabled = False
        cmEliminar.Enabled = False
        cmListado.Enabled = False
        cmModificar.Enabled = False
        cmConsulta.Enabled = False
    Else
        If RSCTS!CERRADO = 1 Then
            cmAceptar.Enabled = False
        Else
            cmAceptar.Enabled = True
        End If
        cmEliminar.Enabled = True
        cmListado.Enabled = True
        cmModificar.Enabled = True
        cmConsulta.Enabled = True
    End If
End Sub

