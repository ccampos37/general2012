VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frPDTSunat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PDT Seguro Complementario de Trabajo de Riesgo"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   Icon            =   "frPDTSunat.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   6750
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport Reporte 
      Left            =   2625
      Top             =   2805
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   390
      Left            =   5070
      TabIndex        =   14
      Top             =   5460
      Width           =   1395
   End
   Begin VB.CommandButton cmExportar 
      Caption         =   "&Exportar"
      Height          =   390
      Left            =   5070
      TabIndex        =   13
      Top             =   4950
      Width           =   1395
   End
   Begin VB.CommandButton cmImprimir 
      Caption         =   "&Imprimir"
      Height          =   390
      Left            =   5070
      TabIndex        =   12
      Top             =   4425
      Width           =   1395
   End
   Begin VB.CommandButton cmQuitar 
      Caption         =   "&Quitar Trab."
      Height          =   390
      Left            =   5070
      TabIndex        =   11
      Top             =   3915
      Width           =   1395
   End
   Begin MSDataGridLib.DataGrid DGLista 
      Height          =   3225
      Left            =   75
      TabIndex        =   1
      Top             =   570
      Width           =   6630
      _ExtentX        =   11695
      _ExtentY        =   5689
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
   Begin VB.Label Total2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00 "
      Height          =   285
      Left            =   3390
      TabIndex        =   10
      Top             =   4500
      Width           =   1305
   End
   Begin VB.Label Total1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00 "
      Height          =   285
      Left            =   2070
      TabIndex        =   9
      Top             =   4500
      Width           =   1305
   End
   Begin VB.Label df 
      BackColor       =   &H80000010&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total General a Pagar"
      Height          =   285
      Left            =   105
      TabIndex        =   8
      Top             =   4500
      Width           =   1950
   End
   Begin VB.Label Tri1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00 "
      Height          =   285
      Index           =   1
      Left            =   3390
      TabIndex        =   7
      Top             =   4200
      Width           =   1305
   End
   Begin VB.Label Tot1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00 "
      Height          =   285
      Index           =   1
      Left            =   2070
      TabIndex        =   6
      Top             =   4200
      Width           =   1305
   End
   Begin VB.Label C1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Imp. Extraordinario de Sol."
      Height          =   285
      Index           =   1
      Left            =   105
      TabIndex        =   5
      Top             =   4200
      Width           =   1950
   End
   Begin VB.Label Tri1 
      Alignment       =   2  'Center
      BackColor       =   &H80000010&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Importe a pagar"
      Height          =   285
      Index           =   0
      Left            =   3390
      TabIndex        =   4
      Top             =   3900
      Width           =   1305
   End
   Begin VB.Label Tot1 
      Alignment       =   2  'Center
      BackColor       =   &H80000010&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Base Imponible"
      Height          =   285
      Index           =   0
      Left            =   2070
      TabIndex        =   3
      Top             =   3900
      Width           =   1305
   End
   Begin VB.Label C1 
      Alignment       =   2  'Center
      BackColor       =   &H80000010&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Concepto"
      Height          =   285
      Index           =   0
      Left            =   105
      TabIndex        =   2
      Top             =   3900
      Width           =   1950
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Resultados de la Exportación de Planillas para el PDT SCTR Sunat - Formulario 0610 para la Versión PDT Sunat 2.0"
      ForeColor       =   &H8000000E&
      Height          =   405
      Left            =   615
      TabIndex        =   0
      Top             =   60
      Width           =   5895
   End
   Begin VB.Image Image1 
      Height          =   510
      Left            =   15
      Picture         =   "frPDTSunat.frx":030A
      Top             =   0
      Width           =   510
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      Height          =   510
      Left            =   525
      Top             =   0
      Width           =   6270
   End
End
Attribute VB_Name = "frPDTSunat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSPDT As ADODB.Recordset

Private Sub cmCerrar_Click()
    Unload Me
End Sub

Private Sub CMEXPORTAR_CLICK()
    On Error GoTo Err1
            Dim xFile As String, CADPDT As String
            frSelDir.Show 1
            If VPTAREA = "" Then Exit Sub
            If Right(VPTAREA, 1) <> "\" Then VPTAREA = VPTAREA & "\"
            xFile = VPTAREA & "\0610" & Right(frPlans.LPlans.SelectedItem.SubItems(1), 4) & Mid(frPlans.LPlans.SelectedItem.SubItems(1), 4, 2) & REGSISTEMA.RUC & ".SCT"
            If Dir$(xFile) <> "" Then
                If MsgBox("Ya existe en esta ruta un archivo correspondiente al PDT SUNAT - SCTR, Desea Ud. reemplazar el archivo por el nuevo que está procesando", vbYesNo + vbQuestion) = vbNo Then Exit Sub
                Kill xFile
            End If
            Open xFile For Append As #1
            Do While Not RSPDT.EOF
                CADPDT = ""
                CADPDT = Val(RSPDT!TIPDOC) & "|" & RSPDT!DOCIDEN & "|" & RSPDT!RUC & "|" & RSPDT!CORRELATIVO & "|" & RSPDT!TASA & "|" & IIf(RSPDT!REMUSCTR = 0, "", RSPDT!REMUSCTR) & "|"
                Print #1, CADPDT
                RSPDT.MoveNext
            Loop
            Close #1
            RSPDT.MoveFirst
            MsgBox "Proceso completado. ingrese al PDT SUNAT y escoja la opción importar del Menú Declaraciones, dentro del módulo 0610 Seguro Complementario de Trabajo de Riesgo", vbInformation
            Exit Sub
Err1:
            MsgBox ERR.Description
            Exit Sub
End Sub

Private Sub CMIMPRIMIR_CLICK()
MsgBox "NO DISPONIBLE"
Exit Sub
    With Reporte
        .WindowTitle = "PLAN0027 - REPORTE DE EXPORTACIÓN DE DATOS AL PDT SUNAT"
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0027.RPT"
        .DataFiles(0) = App.PATH & "\BDAUXCOM.MDB"
        .Destination = crptToWindow
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowState = crptMaximized
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "XRUC='RUC N° " & REGSISTEMA.RUC & "'"
        .Formulas(2) = "XMES='CORRESPONDIENTE AL MES DE " & frPlans.LPlans.SelectedItem.Text & "'"
        If .Status <> 2 Then .Action = 1
    End With
End Sub

Private Sub CMQUITAR_CLICK()
    If RSPDT.RecordCount = 0 Or RSPDT.EOF Then
        MsgBox "Debe existir un trabajador", vbCritical
        Exit Sub
    End If
    If MsgBox("Realmente desea quitar de la exportación para el PDT SUNAT SCTR al trabajador " & RSPDT!NOMBRES, vbYesNo + vbQuestion) = vbNo Then Exit Sub
    DBSYSTEM.Execute "DELETE FROM PDTSUNATSCTR WHERE CODTRAB='" & RSPDT!CODTRAB & "'"
    REFRESCARDG
End Sub

Private Sub DGLISTA_AFTERUPDATE()
    Dim RSAUX As ADODB.Recordset
    Set RSAUX = New ADODB.Recordset
    RSAUX.Open "SELECT SUM(REMUSCTR) AS SUMAAFECTO,SUM(REMUSCTR*TASA/100) AS PAGO FROM ##PDTSUNATSCTR", DBSYSTEM, adOpenStatic
    If RSAUX.RecordCount > 0 Then
        Tot1(1).Caption = Format(Round(RSAUX!SUMAAFECTO, 0), "0.00 ")
        Tri1(1).Caption = Format(Round(RSAUX!PAGO, 0), "0.00 ")
        Total1.Caption = Tot1(1).Caption
        Total2.Caption = Tri1(1).Caption
    End If
    Set RSAUX = Nothing
End Sub

Private Sub DGLISTA_HEADCLICK(ByVal ColIndex As Integer)
    RSPDT.Sort = DGLista.Columns(ColIndex).Caption
End Sub

Private Sub FORM_LOAD()
    Set RSPDT = New ADODB.Recordset
    RSPDT.Open "##PDTSUNATSCTR", DBSYSTEM, adOpenStatic, adLockOptimistic
    REFRESCARDG
End Sub

Public Sub REFRESCARDG()
    RSPDT.Requery
    Set DGLista.DataSource = RSPDT
    With DGLista
        .Columns("NOMBRES").Width = 2600
        .Columns("NOMBRES").Locked = True
        .Columns("REMUSCTR").NumberFormat = "0.00 "
        .Columns("REMUSCTR").Alignment = dbgRight
    End With
    DGLISTA_AFTERUPDATE
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSPDT = Nothing
End Sub

