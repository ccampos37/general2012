VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frAdminUtil 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Depósitos de UTILIDADES"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8610
   Icon            =   "frAdminUtil.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   8610
   Begin VB.CommandButton cmFormulas 
      Caption         =   "Fórmulas de Utilidades"
      Height          =   465
      Left            =   1470
      TabIndex        =   11
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
      Caption         =   "Planilla de Utilidades"
      Height          =   465
      Left            =   210
      TabIndex        =   9
      Top             =   4485
      Width           =   1155
   End
   Begin VB.CommandButton cmConsulta 
      Caption         =   "&Consulta"
      Height          =   405
      Left            =   7239
      TabIndex        =   2
      Top             =   2610
      Width           =   1275
   End
   Begin VB.CommandButton cmCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   405
      Left            =   7239
      TabIndex        =   6
      Top             =   4770
      Width           =   1275
   End
   Begin VB.CommandButton cmAceptar 
      Caption         =   "&Aceptar"
      Height          =   405
      Left            =   7239
      TabIndex        =   5
      Top             =   4230
      Width           =   1275
   End
   Begin VB.CommandButton cmEliminar 
      Caption         =   "&Eliminar"
      Height          =   405
      Left            =   7239
      TabIndex        =   3
      Top             =   3150
      Width           =   1275
   End
   Begin VB.CommandButton cmModificar 
      Caption         =   "&Modificar"
      Height          =   405
      Left            =   7239
      TabIndex        =   4
      Top             =   3690
      Width           =   1275
   End
   Begin VB.CommandButton cmPrueba 
      Caption         =   "&Prueba Cálculo"
      Height          =   405
      Left            =   7239
      TabIndex        =   1
      Top             =   2070
      Width           =   1275
   End
   Begin VB.CommandButton cmNuevo 
      Caption         =   "&Nuevo Cálculo"
      Height          =   405
      Left            =   7239
      TabIndex        =   0
      Top             =   1530
      Width           =   1275
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
      Caption         =   "Planilla de Utilidades"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "Codigo"
         Caption         =   "Código"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
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
         DataField       =   "Utilidad"
         Caption         =   "Utilidad"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,###,##0.00 "
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "PorPart"
         Caption         =   "% Partip."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0 "
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "PartDist"
         Caption         =   "Partic a Distrib."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,###,##0.00 "
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
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3195.213
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column04 
            WrapText        =   -1  'True
            ColumnWidth     =   1005.165
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Utilidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   1
      Left            =   7755
      TabIndex        =   10
      Top             =   705
      Width           =   660
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Utilidad"
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
      Left            =   7755
      TabIndex        =   7
      Top             =   720
      Width           =   660
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   7935
      Picture         =   "frAdminUtil.frx":08CA
      Top             =   165
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
Attribute VB_Name = "frAdminUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents RsUTIL As ADODB.Recordset
Attribute RsUTIL.VB_VarHelpID = -1

Private Sub CMACEPTAR_CLICK()
    VPTRASPRM = "" & RsUTIL!Codigo
    frAceptaUtil.Show 1
    RsUTIL.Requery
    Set xData.DataSource = RsUTIL
End Sub

Private Sub cmCerrar_Click()
    Unload Me
End Sub

Private Sub CMCONSULTA_CLICK()
    VPTAREA = "Vista"
    VPTRASPRM = "" & RsUTIL!Codigo
    frCalcUtil.Show 1
End Sub

Private Sub CMCUSTODIA_CLICK()

End Sub

Private Sub CMELIMINAR_CLICK()
    If RsUTIL.EOF Or RsUTIL.RecordCount = 0 Then
        MsgBox "No existe nada por eliminar}"
    Else
        MsgBox "ADVERTENCIA: eliminará una planilla de UTILIDAD, sin posibilidad a recuperar su información", vbExclamation
        If MsgBox("Seguro de eliminar la planilla de UTILIDAD: " & RsUTIL!NOMBRE & " . Seguro de Continuar", vbYesNo + vbQuestion) = vbNo Then Exit Sub
        DBSYSTEM.Execute "DELETE FROM UTIL Where codigo=" & RsUTIL!Codigo
        DBSYSTEM.Execute "DELETE FROM PlanUTIL WHERE Codigo=" & RsUTIL!Codigo
        DBSYSTEM.Execute "DELETE FROM DetalleUTIL WHERE Codigo=" & RsUTIL!Codigo
        RsUTIL.Requery
        Set xData.DataSource = RsUTIL
    End If
End Sub

Private Sub cmFormulas_Click()
     frFormulasUTIL.Show 1
End Sub

Private Sub CMLISTADO_CLICK()

End Sub

Private Sub CMMODIFICAR_CLICK()
    VPTAREA = "Modificar"
    VPTRASPRM = "" & RsUTIL!Codigo
    frCalcUtil.Show 1
    RsUTIL.Requery
    Set xData.DataSource = RsUTIL
End Sub

Private Sub CMNUEVO_CLICK()
    VPTAREA = "Nuevo"
    frCalcUtil.Show 1
    RsUTIL.Requery
    Set xData.DataSource = RsUTIL
End Sub

Private Sub CMPRUEBA_CLICK()
    VPTAREA = "Prueba"
    frCalcUtil.Show 1
End Sub

Private Sub Command1_Click()
    If RsUTIL.RecordCount = 0 Then
        MsgBox "No existen registros para imprimir", vbCritical
        Exit Sub
    End If
    With Reporte
        frWait.Show 1
        .WindowTitle = "PLAN0058 - Planilla de Pago de UTILIDADES"
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0058.RPT"
        .Connect = "DSN=" & VGL_SERVER & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL & ""
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .SelectionFormula = "{PLANUTIL.Codigo}=" & RsUTIL!Codigo
        .Formulas(0) = "xEmpresa='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "xRuc='RUC N° " & REGSISTEMA.RUC & "'"
        .Formulas(2) = "xMes='Planilla de pagos de UTILIDAD : " & RsUTIL!NOMBRE & "'"
        If .Status <> 2 Then .Action = 1
    End With
End Sub

Private Sub Form_Load()
    Set RsUTIL = New ADODB.Recordset
    RsUTIL.Open "UTIL", DBSYSTEM, adOpenKeyset, adLockOptimistic
    Set xData.DataSource = RsUTIL
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RsUTIL = Nothing
End Sub

Private Sub RsUTIL_MoveComplete(ByVal ADREASON As ADODB.EventReasonEnum, ByVal PERROR As ADODB.Error, ADSTATUS As ADODB.EventStatusEnum, ByVal PRECORDSET As ADODB.Recordset)
    If RsUTIL.EOF Or RsUTIL.RecordCount = 0 Or RsUTIL.BOF Then
        cmAceptar.Enabled = False
        cmEliminar.Enabled = False
        cmModificar.Enabled = False
        cmConsulta.Enabled = False
    Else
        If RsUTIL!CERRADO = 1 Then
            cmAceptar.Enabled = False
        Else
            cmAceptar.Enabled = True
        End If
        cmAceptar.Enabled = True
        cmEliminar.Enabled = True
        cmModificar.Enabled = True
        cmConsulta.Enabled = True
    End If
End Sub
